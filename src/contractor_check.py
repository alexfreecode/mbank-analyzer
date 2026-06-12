"""
contractor_check.py — сравнение клиентов из выписки со справочником
контрагентов, выгруженным из Saldeo Smart (CSV).

Цель: предупредить пользователя ДО импорта фактур о ситуациях, которые
приводят к появлению в Saldeo дублирующих карточек одного и того же
контрагента (опечатка, другой порядок имени/фамилии, другой регистр и т.п.),
а также сообщить о клиентах, которых в базе Saldeo ещё нет вовсе —
для них при импорте будет создана новая карточка.

Использование:
    contractors = load_saldeo_contractors("eksport_kontrahentow.csv")
    results = check_clients(["Jan Kowalski", ...], contractors)
    # results: [{"name": ..., "status": "exact"|"similar"|"new", "matched": ...}, ...]
"""

import csv
import re
import unicodedata
from difflib import SequenceMatcher

# Порог схожести строк (0..1) для отметки «подозрительно похожее имя».
# Подобран эмпирически: ловит опечатки и мелкие отличия, но не путает
# разных людей с похожими, но разными фамилиями.
SIMILARITY_THRESHOLD = 0.84


# ── Нормализация имён для сравнения ──────────────────────────────────────────

def _normalize(s: str) -> str:
    """Приводит имя к виду для сравнения: без диакритики, регистра и лишних пробелов."""
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", s.strip().lower())


def _tokens(s: str) -> frozenset:
    """Множество слов имени — чтобы ловить «Ковальски Ян» == «Ян Ковальски»."""
    return frozenset(_normalize(s).split())


def _sorted_norm(s: str) -> str:
    """
    Нормализованное имя со словами в алфавитном порядке — нужно, чтобы
    нечёткое сравнение «видело» сходство даже тогда, когда одновременно
    и порядок слов другой, И есть небольшая опечатка (например,
    «Jan Kowalski» vs «Kowalski Joan»).
    Обычное посимвольное сравнение в таких случаях даёт низкий коэффициент
    из-за перестановки слов, а сравнение множеств токенов требует точного
    совпадения слов и не ловит опечатку.
    """
    return " ".join(sorted(_normalize(s).split()))


# ── Загрузка справочника контрагентов Saldeo ─────────────────────────────────

def load_saldeo_contractors(csv_path: str) -> list[dict]:
    """
    Загружает CSV-выгрузку справочника контрагентов из Saldeo Smart
    (раздел «Kontrahenci» → «Eksportuj»).

    Сравнение ведётся по колонкам «Nazwa skrócona:» и «Nazwa pełna» —
    их индексы определяются по заголовку файла (на случай изменения
    порядка колонок в будущих версиях Saldeo).

    Возвращает список {"short": ..., "full": ...} с оригинальными
    значениями (нужны для показа пользователю).
    """
    contractors: list[dict] = []

    # encoding="utf-8-sig" — выгрузки Saldeo обычно содержат BOM
    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        header = next(reader, None)
        if not header:
            return contractors

        def _find(col_prefix: str):
            for i, h in enumerate(header):
                if h.strip().lower().startswith(col_prefix.lower()):
                    return i
            return None

        idx_short = _find("Nazwa skrócona")
        idx_full  = _find("Nazwa pełna")
        if idx_short is None and idx_full is None:
            return contractors

        for row in reader:
            if not row or all(not cell.strip() for cell in row):
                continue
            short = row[idx_short].strip() if idx_short is not None and idx_short < len(row) else ""
            full  = row[idx_full].strip()  if idx_full  is not None and idx_full  < len(row) else ""
            if short or full:
                contractors.append({"short": short, "full": full})

    return contractors


# ── Сравнение клиентов со справочником ───────────────────────────────────────

def check_clients(client_names: list[str], contractors: list[dict]) -> list[dict]:
    """
    Сравнивает имена клиентов из выписки со справочником контрагентов Saldeo.

    Возвращает список {"name", "status", "matched"} для каждого имени:
      "exact"   — точное совпадение (без учёта регистра/пробелов/диакритики) —
                  фактура подвяжется к существующей карточке контрагента;
      "similar" — найдено подозрительно похожее, но не идентичное имя —
                  велика вероятность, что это тот же человек, и при импорте
                  появится дублирующая карточка;
      "new"     — совпадений не найдено вовсе — Saldeo создаст новую карточку.

    Логика сравнения использует «Nazwa pełna» как основную цель поиска —
    она всегда ближе к реальному имени человека в банковской выписке.
    «Nazwa skrócona» — произвольный псевдоним ограниченной длины, задаётся
    вручную и может быть чем угодно (аббревиатурой, NIP-ом и т.п.);
    он используется только как запасной вариант, когда полное имя отсутствует.

    Поле "matched" ВСЕГДА содержит «Nazwa skrócona» — это ключ, по которому
    Saldeo ищет существующую карточку при импорте фактур.
    """
    # ref: (canonical_short, full_form, short_form_fallback)
    #   canonical_short    — Nazwa skrócona; возвращается в matched (ключ импорта)
    #   full_form          — формы Nazwa pełna; основная цель для сравнения
    #   short_form_fallback — формы Nazwa skrócona; используется ТОЛЬКО когда
    #                         полное имя отсутствует или совпадает с коротким
    _Form = tuple[str, frozenset, str] | None
    ref: list[tuple[str, _Form, _Form]] = []
    seen: set[str] = set()
    for c in contractors:
        canonical = (c["short"] or c["full"]).strip()
        if not canonical or canonical in seen:
            continue
        seen.add(canonical)
        short = c["short"].strip()
        full  = c["full"].strip()
        # Основная форма для сравнения — полное имя; если его нет, берём короткое
        cmp_full  = full if full else short
        # Запасная форма — короткое имя, только если оно отличается от полного
        cmp_short = short if short and short != cmp_full else ""
        full_form  = (_normalize(cmp_full),  _tokens(cmp_full),  _sorted_norm(cmp_full))  if cmp_full  else None
        short_form = (_normalize(cmp_short), _tokens(cmp_short), _sorted_norm(cmp_short)) if cmp_short else None
        ref.append((canonical, full_form, short_form))

    results: list[dict] = []
    for name in client_names:
        norm_name   = _normalize(name)
        name_tokens = _tokens(name)
        status, matched = "new", None

        # 1) точное совпадение с ПОЛНЫМ именем — самый надёжный сигнал:
        #    «Nazwa pełna» всегда ближе к реальному написанию в выписке
        for canonical, full_form, _ in ref:
            if full_form and norm_name == full_form[0]:
                status, matched = "exact", canonical
                break

        # 2) точное совпадение с КОРОТКИМ именем (запасной вариант):
        #    срабатывает, если полное имя пустое, или короткий псевдоним
        #    случайно совпал с именем в выписке
        if status == "new":
            for canonical, _, short_form in ref:
                if short_form and norm_name == short_form[0]:
                    status, matched = "exact", canonical
                    break

        # 3) другой порядок слов — «Kowalski Jan» / «Jan Kowalski»
        #    проверяем сначала по полному имени, затем по короткому
        if status == "new":
            for canonical, full_form, short_form in ref:
                for form in (full_form, short_form):
                    if form and name_tokens and name_tokens == form[1]:
                        status, matched = "similar", canonical
                        break
                if status == "similar":
                    break

        # 4) полное имя Saldeo является подмножеством токенов выписки —
        #    в Saldeo двусоставное имя, в выписке трёхсоставное (с отчеством
        #    или вторым именем). Например: {wieczorek, sabina} ⊆ {wieczorek, sabina, dorota}.
        #    Проверяем только по полному имени (≥ 2 токенов).
        if status == "new":
            for canonical, full_form, _ in ref:
                if full_form and len(full_form[1]) >= 2 and full_form[1].issubset(name_tokens):
                    status, matched = "similar", canonical
                    break

        # 5) нечёткое сходство строк — опечатки, мелкие отличия в написании.
        #    Сравниваем по полному имени; если его нет — по короткому.
        if status == "new":
            best_ratio, best_match = 0.0, None
            for canonical, full_form, short_form in ref:
                form = full_form or short_form
                if not form:
                    continue
                ratio = SequenceMatcher(None, norm_name, form[0]).ratio()
                if ratio > best_ratio:
                    best_ratio, best_match = ratio, canonical
            if best_ratio >= SIMILARITY_THRESHOLD:
                status, matched = "similar", best_match

        # 6) другой порядок слов + опечатка одновременно.
        #    Сортируем токены алфавитно — перестановка перестаёт мешать,
        #    и опечатка становится единственным отличием.
        #    Сравниваем по полному имени; если его нет — по короткому.
        if status == "new":
            sorted_name = _sorted_norm(name)
            best_ratio, best_match = 0.0, None
            for canonical, full_form, short_form in ref:
                form = full_form or short_form
                if not form:
                    continue
                ratio = SequenceMatcher(None, sorted_name, form[2]).ratio()
                if ratio > best_ratio:
                    best_ratio, best_match = ratio, canonical
            if best_ratio >= SIMILARITY_THRESHOLD:
                status, matched = "similar", best_match

        results.append({"name": name, "status": status, "matched": matched})

    return results
