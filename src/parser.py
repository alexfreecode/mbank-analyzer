"""
parser.py — Анализатор банковской выписки mBank
Группирует входящие платежи от физических лиц по клиенту.
"""

import argparse
import re
import sys
from pathlib import Path

import pandas as pd

try:
    from tabulate import tabulate
    HAS_TABULATE = True
except ImportError:
    HAS_TABULATE = False


# ─── Константы ────────────────────────────────────────────────────────────────

HEADER_MARKER = "#Data operacji"

# Корень для определения входящих переводов.
# Все польские типы входящих операций содержат "przychodzący" (входящий):
#   PRZELEW ZEWNĘTRZNY/WEWNĘTRZNY PRZYCHODZĄCY
#   BLIK P2P-PRZYCHODZĄCY
#   PRZELEW EXPRESS ELIXIR PRZYCH.
#   ... и любые будущие варианты
# Корень PRZYCH покрывает все варианты без перечисления конкретных фраз.
INCOMING_ROOT = "PRZYCH"

# Ключевые слова юридических лиц и банковских операций — исключить
COMPANY_KEYWORDS = [
    "SPÓŁKA", "S.A.", "SP. Z O.O.", "SP.Z O.O.",
    "ZAKŁAD", "URZĄD", "FUNDUSZ", "TOWARZYSTWO",
    "UZNANIE NATYCH",
]


# ─── Парсинг данных ────────────────────────────────────────────────────────────

def find_header_row(file_path: str, encoding: str) -> int:
    """Находит номер строки с заголовком таблицы (0-indexed)."""
    with open(file_path, encoding=encoding, errors="replace") as f:
        for i, line in enumerate(f):
            if HEADER_MARKER in line:
                return i
    raise ValueError(
        f"Nie znaleziono wiersza nagłówka ze znacznikiem '{HEADER_MARKER}'. "
        "Upewnij się, że plik jest wyciągiem mBank."
    )


def parse_kwota(value: str) -> float:
    """Преобразует '1 300,00 PLN' → 1300.0, '-22,15 PLN' → -22.15"""
    cleaned = (
        str(value)
        .replace(" PLN", "")
        .replace("\xa0", "")   # неразрывный пробел
        .replace(" ", "")
        .replace(",", ".")
    )
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def _clean_address(name_part: str) -> str:
    """Убирает адрес из строки «ИМЯ [АДРЕС]», оставляя только имя."""
    # Убираем суффикс " ." (формат BLIK: «TIUPA YELYZAVETA .»)
    name_part = re.sub(r"\s+\.\s*$", "", name_part).strip()
    # Убираем адрес с явными префиксами улицы
    name_part = re.split(r"\s+(?:UL\.|AL\.|OS\.|PL\.)", name_part)[0].strip()
    # Убираем почтовый индекс и всё после него
    name_part = re.sub(r"\s+\d{2}-\d{3}.*$", "", name_part).strip()
    # Убираем «НАЗВАНИЕ_УЛИЦЫ НОМЕР_ДОМА» без префикса (в конце строки)
    # Пример: «BOHDAN HUDYMA GRANICZNA 53/74» → «BOHDAN HUDYMA»
    name_part = re.sub(r"\s+\S+\s+\d+[\w/]*\s*$", "", name_part).strip()
    return " ".join(name_part.split()).upper()


def extract_address(desc: str) -> tuple[str, str, str]:
    """
    Извлекает адрес из поля Opis operacji.
    Возвращает (улица, почтовый_индекс, город).
    Возвращает ("", "", "") если адреса нет (BLIK, Express Elixir и т.д.).

    Стандартный формат: «ИМЯ [УЛ. УЛИЦА НОМЕР] [КОД ГОРОД], ЗАГОЛОВОК ...»
    Адрес находится в блоке до первой запятой, после имени.
    """
    # Express Elixir (нет запятой) — адрес ненадёжен, пропускаем
    if "," not in desc:
        return ("", "", "")

    name_part = desc.split(",")[0].strip()
    # Убираем BLIK-суффикс " ."
    name_part = re.sub(r"\s+\.\s*$", "", name_part).strip()

    # ── Ищем почтовый индекс XX-XXX ──────────────────────────────────────────
    postal_match = re.search(r'\b(\d{2}-\d{3})\b', name_part)
    if postal_match:
        postal = postal_match.group(1)
        city   = name_part[postal_match.end():].strip().upper()
        before = name_part[:postal_match.start()].strip()

        # Улица с префиксом UL./AL./OS./PL.
        pref = re.search(r'(?:UL\.|AL\.|OS\.|PL\.)\s*(.+)$', before, re.I)
        if pref:
            street = "UL. " + pref.group(1).strip().upper()
        else:
            # Улица без префикса: последний «СЛОВО НОМЕР» в конце строки
            no_pref = re.search(r'(\S+)\s+(\d+[\w/]*)\s*$', before)
            street  = (no_pref.group(1) + " " + no_pref.group(2)).upper() if no_pref else ""

        return (street, postal, city)

    # ── Только префикс без почтового индекса ─────────────────────────────────
    pref = re.search(r'(?:UL\.|AL\.|OS\.|PL\.)\s*(.+)$', name_part, re.I)
    if pref:
        return ("UL. " + pref.group(1).strip().upper(), "", "")

    return ("", "", "")


def extract_name(desc: str) -> str:
    """
    Извлекает имя клиента из поля Opis operacji.

    Два формата:

    1. Стандартный (PRZELEW / BLIK P2P):
       «ИМЯ [АДРЕС], ОПИСАНИЕ   ТИП_ОПЕРАЦИИ   НОМЕР_СЧЁТА»
       → имя берётся до первой запятой

    2. Express Elixir (нет запятой):
       «PRZELEW EXPRESS ELIXIR PRZYCH.  ИМЯ  АДРЕС  .  ...  PRZYCH.  СЧЁТ»
       → имя берётся из второго блока (после двойного пробела)
    """
    if "," not in desc:
        # Express Elixir и подобные форматы без запятой
        blocks = [
            b.strip()
            for b in re.split(r"\s{2,}", desc.strip())
            if b.strip() and b.strip() != "."
        ]
        # blocks[0] — тип операции, blocks[1] — имя (возможно с адресом)
        name_part = blocks[1] if len(blocks) >= 2 else (blocks[0] if blocks else "")
    else:
        # Стандартный формат
        name_part = desc.split(",")[0].strip()

    result = _clean_address(name_part)
    return result if result else "NIEZNANY"


def extract_title(desc: str) -> str:
    """
    Извлекает краткое описание платежа из поля Opis operacji.

    Стандартный формат: берёт текст после первой запятой, до двойного пробела.
      Пример: «..., FV 1/03/2026  KOWALSKA...» → «FV 1/03/2026»

    Express Elixir (нет запятой): возвращает «Express Elixir».
    """
    if "," not in desc:
        # Express Elixir — в описании нет отдельного назначения платежа
        return "Express Elixir"

    after_comma = desc.split(",", 1)[1].strip()
    # Убираем суффикс «.» (BLIK)
    after_comma = re.sub(r"\s+\.\s*$", "", after_comma)
    # Берём первый токен до 2+ пробелов
    parts = re.split(r"\s{2,}", after_comma)
    title = parts[0].strip().rstrip(".").strip()
    return title


def is_individual(desc: str) -> bool:
    """True, если платёж от физического лица (не компании).

    Два независимых признака — если хотя бы один срабатывает, это юрлицо:
    1. Чёрный список ключевых слов (SPÓŁKA, S.A., URZĄD, ...)
    2. NIP — 10-значный налоговый номер польского юрлица.
       У физлиц NIP в банковских описаниях не встречается.
       Паттерн: ровно 10 цифр, не являющихся частью более длинного числа
       (IBAN из 26 цифр под это не попадает — там цифры не прерываются).
    """
    desc_upper = desc.upper()
    if any(kw in desc_upper for kw in COMPANY_KEYWORDS):
        return False
    if re.search(r'(?<!\d)\d{10}(?!\d)', desc):
        return False
    return True


def is_incoming(desc: str) -> bool:
    """True, если операция — входящий платёж.
    Проверяет корень PRZYCH, который присутствует в любом польском
    типе входящего перевода независимо от конкретной формулировки.
    """
    return INCOMING_ROOT in desc.upper()


# ─── Форматирование вывода ─────────────────────────────────────────────────────

def fmt_amount(value: float) -> str:
    """1300.0 → '1 300,00'"""
    formatted = f"{value:,.2f}".replace(",", " ").replace(".", ",")
    # Python форматирует 1300.0 как «1,300.00», поправляем вручную
    # Более надёжный способ:
    integer_part = int(abs(value))
    frac_part = round(abs(value) % 1 * 100)
    sign = "-" if value < 0 else ""

    # Разбивка на группы по 3 цифры
    s = str(integer_part)
    groups = []
    while s:
        groups.append(s[-3:])
        s = s[:-3]
    integer_str = "\u202f".join(reversed(groups))  # неразрывный тонкий пробел

    return f"{sign}{integer_str},{frac_part:02d}"


def print_separator(char: str = "─", width: int = 62) -> None:
    print("  " + char * width)


def print_header(char: str = "=", width: int = 64) -> None:
    print(char * width)


# ─── Основная логика ───────────────────────────────────────────────────────────

def load_transactions(file_path: str, encoding: str) -> pd.DataFrame:
    """Загружает CSV-выписку mBank, возвращает DataFrame с нужными колонками."""
    header_row = find_header_row(file_path, encoding)

    df = pd.read_csv(
        file_path,
        sep=";",
        encoding=encoding,
        skiprows=header_row,
        dtype=str,
        index_col=False,   # данные имеют лишний столбец — без авто-индекса
    )

    # Убираем лишние пробелы из названий колонок
    df.columns = [c.strip() for c in df.columns]

    required = ["#Data operacji", "#Opis operacji", "#Kwota"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            f"W pliku nie znaleziono kolumn: {missing}\n"
            f"Dostępne kolumny: {list(df.columns)}"
        )

    # Убираем пустые строки
    df = df.dropna(subset=["#Data operacji", "#Kwota"])

    return df


def analyze(file_path: str, encoding: str) -> pd.DataFrame:
    """
    Загружает выписку, фильтрует платежи от физ. лиц,
    возвращает DataFrame с колонками: date, name, amount, title.
    """
    df = load_transactions(file_path, encoding)

    # Парсим сумму
    df["amount"] = df["#Kwota"].apply(parse_kwota)

    # Фильтр: входящие (amount > 0) + ключевые слова + не компания
    mask = (
        (df["amount"] > 0)
        & df["#Opis operacji"].apply(lambda x: is_incoming(str(x)))
        & df["#Opis operacji"].apply(lambda x: is_individual(str(x)))
    )
    df_clients = df[mask].copy()

    if df_clients.empty:
        return df_clients

    # Извлекаем имя, описание и адрес
    df_clients["name"]  = df_clients["#Opis operacji"].apply(
        lambda x: extract_name(str(x))
    )
    df_clients["title"] = df_clients["#Opis operacji"].apply(
        lambda x: extract_title(str(x))
    )
    df_clients["date"]  = df_clients["#Data operacji"].str.strip()

    addr = df_clients["#Opis operacji"].apply(
        lambda x: extract_address(str(x))
    )
    df_clients["addr_street"] = addr.apply(lambda t: t[0])
    df_clients["addr_postal"] = addr.apply(lambda t: t[1])
    df_clients["addr_city"]   = addr.apply(lambda t: t[2])

    return df_clients[
        ["date", "name", "amount", "title",
         "addr_street", "addr_postal", "addr_city"]
    ].reset_index(drop=True)


def categorize_transaction(amount: float, desc: str) -> str:
    """
    Относит операцию к одной из трёх категорий:
      "Klienci"          — приход от физлица (то, что попадает в df_clients)
      "Pozostałe wpływy" — любой другой приход (от фирм, возвраты, проценты и т.п.)
      "Wydatki"          — любой расход
    """
    if amount > 0:
        if is_incoming(desc) and is_individual(desc):
            return "Klienci"
        return "Pozostałe wpływy"
    return "Wydatki"


def _collapse_desc(desc: str) -> str:
    """Сворачивает все пробельные последовательности в одиночный пробел."""
    return " ".join(str(desc).split())


def _shorten(text: str, max_len: int = 70) -> str:
    if len(text) > max_len:
        return text[: max_len - 1] + "…"
    return text


def print_reconciliation_report(file_path: str, encoding: str, output_lines: list) -> None:
    """
    Выводит отдельный отчёт «Kontrola kompletności wyciągu» — разбивает
    ВСЕ операции выписки на три категории (Klienci / Pozostałe wpływy /
    Wydatki), чтобы пользователь мог проверить, что ничего не потеряно
    и общая сумма сходится с выпиской.

    Внутри каждой категории операции группируются по точному совпадению
    текста «Opis operacji» — повторяющиеся описания (например, комиссии
    банка) сворачиваются в одну строку с количеством и суммой.
    """

    def out(line: str = "") -> None:
        try:
            print(line)
        except Exception:
            pass
        output_lines.append(line)

    df = load_transactions(file_path, encoding)
    df["amount"] = df["#Kwota"].apply(parse_kwota)
    df["category"] = df.apply(
        lambda r: categorize_transaction(r["amount"], str(r["#Opis operacji"])),
        axis=1,
    )

    out("═" * 64)
    out("  KONTROLA KOMPLETNOŚCI WYCIĄGU")
    out("═" * 64)
    out()
    incoming = df[df["amount"] > 0]
    out(f"  Wszystkich operacji w wyciągu: {len(df)}")
    out(f"  Suma wpływów:                  {fmt_amount(incoming['amount'].sum())} PLN"
        f"  ({len(incoming)} operacji)")
    out(f"  Suma wszystkich operacji:      {fmt_amount(df['amount'].sum())} PLN")

    for category in ("Klienci", "Pozostałe wpływy", "Wydatki"):
        sub = df[df["category"] == category]

        out()
        out("─" * 64)
        out(f"  {category.upper()}  —  {len(sub)} operacji,"
            f"  {fmt_amount(sub['amount'].sum())} PLN")
        out("─" * 64)

        if sub.empty:
            out("  (brak operacji)")
            continue

        sub = sub.copy()
        sub["_desc"] = sub["#Opis operacji"].apply(_collapse_desc)

        grouped = (
            sub.groupby("_desc", sort=False)
            .agg(n=("amount", "count"),
                 s=("amount", "sum"),
                 d=("#Data operacji", "first"))
            .reset_index()
        )
        grouped = grouped.reindex(
            grouped["s"].abs().sort_values(ascending=False).index
        )

        for _, row in grouped.iterrows():
            label = _shorten(row["_desc"])
            amount_str = fmt_amount(row["s"]) + " PLN"
            if row["n"] == 1:
                out(f"  {str(row['d']).strip():<12}  {amount_str:>16}   {label}")
            else:
                out(f"  {'':<12}  {amount_str:>16}   {label}  (× {int(row['n'])})")

    out()
    out("═" * 64)
    out(f"  SUMA KONTROLNA (wszystkie operacje):  "
        f"{fmt_amount(df['amount'].sum())} PLN")
    out("═" * 64)


def print_report(df: pd.DataFrame, output_lines: list) -> None:
    """Выводит подробный отчёт по клиентам + сводную таблицу."""

    def out(line: str = "") -> None:
        # print() пишет в sys.stdout — это нужно только для CLI-режима.
        # В GUI-режиме (а тем более в собранном «--windowed» .exe, где
        # sys.stdout вообще равен None) эта печать не нужна — отчёт и так
        # собирается в output_lines и выводится в текстовое поле/файл.
        # При этом она может УПАСТЬ: если кодировка консоли не «понимает»
        # польские диакритики (например, cp1252/„charmap” не содержит
        # Ą Ć Ę Ł Ń Ś Ź Ż), print() бросает UnicodeEncodeError и весь
        # отчёт обрывается с ошибкой — поэтому глушим любые сбои печати.
        try:
            print(line)
        except Exception:
            pass
        output_lines.append(line)

    if df.empty:
        out("Nie znaleziono płatności od osób fizycznych.")
        return

    # Определяем дату первого платежа каждого клиента
    first_payment = df.groupby("name")["date"].min().rename("first_date")
    df_sorted = df.join(first_payment, on="name")

    # Сортируем: клиенты — по дате первого платежа, внутри клиента — по дате
    df_sorted = df_sorted.sort_values(["first_date", "name", "date"])

    # ── Блоки по клиентам (sort=False — порядок уже задан выше) ──
    for name, group in df_sorted.groupby("name", sort=False):
        group = group.sort_values("date")
        total = group["amount"].sum()
        count = len(group)

        out()
        out("=" * 64)
        out(f"  KLIENT: {name}")
        out("-" * 64)

        # Таблица транзакций
        rows = []
        for _, row in group.iterrows():
            rows.append([row["date"], fmt_amount(row["amount"]) + " PLN", row["title"]])

        if HAS_TABULATE:
            table = tabulate(
                rows,
                headers=["Data", "Kwota", "Opis"],
                tablefmt="plain",
            )
            for line in table.splitlines():
                out("  " + line)
        else:
            out(f"  {'Data':<14}  {'Kwota':<16}  Opis")
            out("  " + "─" * 60)
            for r in rows:
                out(f"  {r[0]:<14}  {r[1]:<16}  {r[2]}")

        out("  " + "─" * 44)
        out(f"  RAZEM:  {fmt_amount(total) + ' PLN':>16}  ({count} transakcji)")

    # ── Сводная таблица ──
    out()
    out()
    out("═" * 64)
    out("  TABELA ZBIORCZA")
    out("═" * 64)

    summary = (
        df.groupby("name")
        .agg(transakcji=("amount", "count"), razem=("amount", "sum"))
        .sort_values("razem", ascending=False)
        .reset_index()
    )
    summary.rename(columns={"name": "Klient"}, inplace=True)

    summary_rows = [
        [row["Klient"], int(row["transakcji"]), fmt_amount(row["razem"]) + " PLN"]
        for _, row in summary.iterrows()
    ]

    if HAS_TABULATE:
        table = tabulate(
            summary_rows,
            headers=["Klient", "Transakcji", "Kwota"],
            tablefmt="plain",
            colalign=("left", "right", "right"),
        )
        for line in table.splitlines():
            out("  " + line)
    else:
        out(f"  {'Klient':<35}  {'Transakcji':>10}  {'Kwota':>16}")
        out("  " + "─" * 65)
        for r in summary_rows:
            out(f"  {r[0]:<35}  {r[1]:>10}  {r[2]:>16}")

    out("  " + "═" * 65)
    total_all = df["amount"].sum()
    total_count = len(df)
    clients_count = df["name"].nunique()
    out(f"  {'RAZEM':<35}  {total_count:>10}  {fmt_amount(total_all) + ' PLN':>16}")
    out()
    out(f"  Unikalnych klientów: {clients_count}")
    out("═" * 64)


# ─── CLI ───────────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Analizator wyciągu bankowego mBank — płatności od osób fizycznych",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Przykłady:
  python parser.py --file lista_operacji_260301_260331.csv
  python parser.py --file statement.csv --output report.txt
  python parser.py --file statement.csv --encoding cp1250
""",
    )
    parser.add_argument(
        "--file", "-f",
        required=True,
        help="Ścieżka do pliku CSV z wyciągiem mBank",
    )
    parser.add_argument(
        "--encoding", "-e",
        default="utf-8-sig",
        help="Kodowanie pliku (domyślnie: utf-8-sig). "
             "Dla starszych wyciągów spróbuj cp1250.",
    )
    parser.add_argument(
        "--output", "-o",
        default=None,
        help="Zapisz raport do pliku tekstowego (opcjonalnie)",
    )

    args = parser.parse_args()

    # Проверяем файл
    if not Path(args.file).exists():
        print(f"Błąd: plik nie znaleziony: {args.file}", file=sys.stderr)
        sys.exit(1)

    # Анализ
    try:
        df = analyze(args.file, args.encoding)
    except Exception as exc:
        print(f"Błąd podczas odczytu pliku: {exc}", file=sys.stderr)
        sys.exit(1)

    # Вывод
    output_lines: list = []
    print_report(df, output_lines)

    # Сохранение в файл
    if args.output:
        try:
            with open(args.output, "w", encoding="utf-8") as f:
                f.write("\n".join(output_lines))
            print(f"\nRaport zapisany: {args.output}")
        except OSError as exc:
            print(f"Nie udało się zapisać pliku: {exc}", file=sys.stderr)


if __name__ == "__main__":
    # Переключаем stdout/stderr на UTF-8 для Windows-терминала
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    main()
