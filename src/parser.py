"""
parser.py — Suma Wpłat: analiza wyciągu bankowego mBank
Grupuje wpłaty przychodzące od osób fizycznych według klienta.
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


# ─── Stałe ────────────────────────────────────────────────────────────────────

HEADER_MARKER = "#Data operacji"

# Rdzeń do rozpoznawania przelewów przychodzących.
# Wszystkie polskie typy operacji przychodzących zawierają „przychodzący”:
#   PRZELEW ZEWNĘTRZNY/WEWNĘTRZNY PRZYCHODZĄCY
#   BLIK P2P-PRZYCHODZĄCY
#   PRZELEW EXPRESS ELIXIR PRZYCH.
#   ... i wszelkie przyszłe warianty
# Rdzeń PRZYCH obejmuje wszystkie warianty bez wymieniania konkretnych fraz.
INCOMING_ROOT = "PRZYCH"

# Słowa kluczowe firm i operacji bankowych — do wykluczenia
COMPANY_KEYWORDS = [
    "SPÓŁKA", "S.A.", "SP. Z O.O.", "SP.Z O.O.",
    "ZAKŁAD", "URZĄD", "FUNDUSZ", "TOWARZYSTWO",
    "UZNANIE NATYCH",
]


# ─── Parsowanie danych ─────────────────────────────────────────────────────────

def find_header_row(file_path: str, encoding: str) -> int:
    """Znajduje numer wiersza z nagłówkiem tabeli (indeksowany od 0)."""
    with open(file_path, encoding=encoding, errors="replace") as f:
        for i, line in enumerate(f):
            if HEADER_MARKER in line:
                return i
    raise ValueError(
        f"Nie znaleziono wiersza nagłówka ze znacznikiem '{HEADER_MARKER}'. "
        "Upewnij się, że plik jest wyciągiem mBank."
    )


def parse_kwota(value: str) -> float:
    """Zamienia '1 300,00 PLN' → 1300.0, '-22,15 PLN' → -22.15"""
    cleaned = (
        str(value)
        .replace(" PLN", "")
        .replace("\xa0", "")   # twarda spacja
        .replace(" ", "")
        .replace(",", ".")
    )
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def _clean_address(name_part: str) -> str:
    """Usuwa adres z tekstu „IMIĘ [ADRES]”, pozostawiając samo imię i nazwisko."""
    # Usuwamy sufiks " ." (format BLIK: „TIUPA YELYZAVETA .”)
    name_part = re.sub(r"\s+\.\s*$", "", name_part).strip()
    # Usuwamy adres z wyraźnym prefiksem ulicy
    name_part = re.split(r"\s+(?:UL\.|AL\.|OS\.|PL\.)", name_part)[0].strip()
    # Usuwamy kod pocztowy i wszystko po nim
    name_part = re.sub(r"\s+\d{2}-\d{3}.*$", "", name_part).strip()
    # Usuwamy „NAZWA_ULICY NUMER_DOMU” bez prefiksu (na końcu tekstu)
    # Przykład: „BOHDAN HUDYMA GRANICZNA 53/74” → „BOHDAN HUDYMA”
    name_part = re.sub(r"\s+\S+\s+\d+[\w/]*\s*$", "", name_part).strip()
    return " ".join(name_part.split()).upper()


def extract_address(desc: str) -> tuple[str, str, str]:
    """
    Wyodrębnia adres z pola Opis operacji.
    Zwraca (ulica, kod_pocztowy, miejscowość).
    Zwraca ("", "", "") gdy adresu brak (BLIK, Express Elixir itd.).

    Format standardowy: „IMIĘ [UL. ULICA NUMER] [KOD MIASTO], TYTUŁ ...”
    Adres znajduje się w bloku przed pierwszym przecinkiem, po imieniu.
    """
    # Express Elixir (brak przecinka) — adres niewiarygodny, pomijamy
    if "," not in desc:
        return ("", "", "")

    name_part = desc.split(",")[0].strip()
    # Usuwamy sufiks BLIK " ."
    name_part = re.sub(r"\s+\.\s*$", "", name_part).strip()

    # ── Szukamy kodu pocztowego XX-XXX ───────────────────────────────────────
    postal_match = re.search(r'\b(\d{2}-\d{3})\b', name_part)
    if postal_match:
        postal = postal_match.group(1)
        city   = name_part[postal_match.end():].strip().upper()
        before = name_part[:postal_match.start()].strip()

        # Ulica z prefiksem UL./AL./OS./PL.
        pref = re.search(r'(?:UL\.|AL\.|OS\.|PL\.)\s*(.+)$', before, re.I)
        if pref:
            street = "UL. " + pref.group(1).strip().upper()
        else:
            # Ulica bez prefiksu: ostatnie „SŁOWO NUMER” na końcu tekstu
            no_pref = re.search(r'(\S+)\s+(\d+[\w/]*)\s*$', before)
            street  = (no_pref.group(1) + " " + no_pref.group(2)).upper() if no_pref else ""

        return (street, postal, city)

    # ── Sam prefiks bez kodu pocztowego ──────────────────────────────────────
    pref = re.search(r'(?:UL\.|AL\.|OS\.|PL\.)\s*(.+)$', name_part, re.I)
    if pref:
        return ("UL. " + pref.group(1).strip().upper(), "", "")

    return ("", "", "")


def extract_name(desc: str) -> str:
    """
    Wyodrębnia nazwę klienta z pola Opis operacji.

    Dwa formaty:

    1. Standardowy (PRZELEW / BLIK P2P):
       „IMIĘ [ADRES], TYTUŁ   TYP_OPERACJI   NUMER_RACHUNKU”
       → imię i nazwisko brane do pierwszego przecinka

    2. Express Elixir (brak przecinka):
       „PRZELEW EXPRESS ELIXIR PRZYCH.  IMIĘ  ADRES  .  ...  PRZYCH.  RACHUNEK”
       → imię i nazwisko brane z drugiego bloku (po podwójnej spacji)
    """
    if "," not in desc:
        # Express Elixir i podobne formaty bez przecinka
        blocks = [
            b.strip()
            for b in re.split(r"\s{2,}", desc.strip())
            if b.strip() and b.strip() != "."
        ]
        # blocks[0] — typ operacji, blocks[1] — imię i nazwisko (ewent. z adresem)
        name_part = blocks[1] if len(blocks) >= 2 else (blocks[0] if blocks else "")
    else:
        # Format standardowy
        name_part = desc.split(",")[0].strip()

    result = _clean_address(name_part)
    return result if result else "NIEZNANY"


def extract_title(desc: str) -> str:
    """
    Wyodrębnia krótki tytuł płatności z pola Opis operacji.

    Format standardowy: bierze tekst po pierwszym przecinku, do podwójnej spacji.
      Przykład: „..., FV 1/03/2026  KOWALSKA...” → „FV 1/03/2026”

    Express Elixir (brak przecinka): zwraca „Express Elixir”.
    """
    if "," not in desc:
        # Express Elixir — w opisie nie ma osobnego tytułu płatności
        return "Express Elixir"

    after_comma = desc.split(",", 1)[1].strip()
    # Usuwamy sufiks „.” (BLIK)
    after_comma = re.sub(r"\s+\.\s*$", "", after_comma)
    # Bierzemy pierwszy człon do 2+ spacji
    parts = re.split(r"\s{2,}", after_comma)
    title = parts[0].strip().rstrip(".").strip()
    return title


def is_individual(desc: str) -> bool:
    """True, gdy płatność pochodzi od osoby fizycznej (nie od firmy).

    Dwa niezależne sygnały — jeśli którykolwiek zadziała, to firma:
    1. Czarna lista słów kluczowych (SPÓŁKA, S.A., URZĄD, ...)
    2. NIP — 10-cyfrowy numer podatkowy polskiej firmy.
       U osób fizycznych NIP nie występuje w opisach bankowych.
       Wzorzec: dokładnie 10 cyfr niebędących częścią dłuższej liczby
       (IBAN z 26 cyfr tu nie pasuje — tam ciąg cyfr jest nieprzerwany).
    """
    desc_upper = desc.upper()
    if any(kw in desc_upper for kw in COMPANY_KEYWORDS):
        return False
    if re.search(r'(?<!\d)\d{10}(?!\d)', desc):
        return False
    return True


def is_incoming(desc: str) -> bool:
    """True, gdy operacja jest płatnością przychodzącą.
    Sprawdza rdzeń PRZYCH, obecny w każdym polskim typie przelewu
    przychodzącego niezależnie od konkretnego sformułowania.
    """
    return INCOMING_ROOT in desc.upper()


# ─── Formatowanie wyniku ───────────────────────────────────────────────────────

def fmt_amount(value: float) -> str:
    """1300.0 → '1 300,00'"""
    formatted = f"{value:,.2f}".replace(",", " ").replace(".", ",")
    # Python formatuje 1300.0 jako „1,300.00” — poprawiamy ręcznie
    # Bardziej niezawodny sposób:
    integer_part = int(abs(value))
    frac_part = round(abs(value) % 1 * 100)
    sign = "-" if value < 0 else ""

    # Podział na grupy po 3 cyfry
    s = str(integer_part)
    groups = []
    while s:
        groups.append(s[-3:])
        s = s[:-3]
    integer_str = " ".join(reversed(groups))  # wąska twarda spacja

    return f"{sign}{integer_str},{frac_part:02d}"


def print_separator(char: str = "─", width: int = 62) -> None:
    print("  " + char * width)


def print_header(char: str = "=", width: int = 64) -> None:
    print(char * width)


# ─── Logika główna ─────────────────────────────────────────────────────────────

def load_transactions(file_path: str, encoding: str) -> pd.DataFrame:
    """Wczytuje wyciąg CSV z mBanku, zwraca DataFrame z potrzebnymi kolumnami."""
    header_row = find_header_row(file_path, encoding)

    df = pd.read_csv(
        file_path,
        sep=";",
        encoding=encoding,
        skiprows=header_row,
        dtype=str,
        index_col=False,   # dane mają nadmiarową kolumnę — bez auto-indeksu
    )

    # Usuwamy zbędne spacje z nazw kolumn
    df.columns = [c.strip() for c in df.columns]

    required = ["#Data operacji", "#Opis operacji", "#Kwota"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            f"W pliku nie znaleziono kolumn: {missing}\n"
            f"Dostępne kolumny: {list(df.columns)}"
        )

    # Usuwamy puste wiersze
    df = df.dropna(subset=["#Data operacji", "#Kwota"])

    return df


def analyze(file_path: str, encoding: str) -> pd.DataFrame:
    """
    Wczytuje wyciąg, filtruje płatności od osób fizycznych,
    zwraca DataFrame z kolumnami: date, name, amount, title.
    """
    df = load_transactions(file_path, encoding)

    # Parsujemy kwotę
    df["amount"] = df["#Kwota"].apply(parse_kwota)

    # Filtr: przychodzące (amount > 0) + słowa kluczowe + nie firma
    mask = (
        (df["amount"] > 0)
        & df["#Opis operacji"].apply(lambda x: is_incoming(str(x)))
        & df["#Opis operacji"].apply(lambda x: is_individual(str(x)))
    )
    df_clients = df[mask].copy()

    if df_clients.empty:
        return df_clients

    # Wyodrębniamy nazwę, tytuł i adres
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
    Przypisuje operację do jednej z trzech kategorii:
      "Klienci"          — wpłata od osoby fizycznej (to, co trafia do df_clients)
      "Pozostałe wpływy" — każdy inny wpływ (od firm, zwroty, odsetki itp.)
      "Wydatki"          — każdy wydatek
    """
    if amount > 0:
        if is_incoming(desc) and is_individual(desc):
            return "Klienci"
        return "Pozostałe wpływy"
    return "Wydatki"


def _collapse_desc(desc: str) -> str:
    """Zwija wszystkie ciągi białych znaków do pojedynczej spacji."""
    return " ".join(str(desc).split())


def _shorten(text: str, max_len: int = 70) -> str:
    if len(text) > max_len:
        return text[: max_len - 1] + "…"
    return text


def print_reconciliation_report(file_path: str, encoding: str, output_lines: list) -> None:
    """
    Generuje osobny raport „Kontrola kompletności wyciągu” — dzieli
    WSZYSTKIE operacje wyciągu na trzy kategorie (Klienci / Pozostałe
    wpływy / Wydatki), aby użytkownik mógł sprawdzić, że nic nie zginęło
    i suma łączna zgadza się z wyciągiem.

    Wewnątrz każdej kategorii operacje grupowane są po identycznym tekście
    „Opis operacji” — powtarzające się opisy (np. prowizje bankowe)
    zwijane są do jednego wiersza z liczbą operacji i sumą.
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
    """Generuje szczegółowy raport per klient + tabelę zbiorczą."""

    def out(line: str = "") -> None:
        # print() pisze do sys.stdout — potrzebne tylko w trybie CLI.
        # W trybie GUI (a tym bardziej w zbudowanym „--windowed” .exe, gdzie
        # sys.stdout jest wręcz równy None) ten wydruk nie jest potrzebny —
        # raport i tak zbierany jest w output_lines i trafia do pola
        # tekstowego/pliku. Przy tym wydruk może SIĘ WYSYPAĆ: jeśli kodowanie
        # konsoli nie „rozumie” polskich znaków (np. cp1252/„charmap” nie
        # zawiera Ą Ć Ę Ł Ń Ś Ź Ż), print() rzuca UnicodeEncodeError i cały
        # raport urywa się błędem — dlatego wyciszamy wszelkie błędy druku.
        try:
            print(line)
        except Exception:
            pass
        output_lines.append(line)

    if df.empty:
        out("Nie znaleziono płatności od osób fizycznych.")
        return

    # Ustalamy datę pierwszej wpłaty każdego klienta
    first_payment = df.groupby("name")["date"].min().rename("first_date")
    df_sorted = df.join(first_payment, on="name")

    # Sortujemy: klienci — wg daty pierwszej wpłaty, w ramach klienta — wg daty
    df_sorted = df_sorted.sort_values(["first_date", "name", "date"])

    # ── Bloki per klient (sort=False — kolejność już ustalona wyżej) ──
    for name, group in df_sorted.groupby("name", sort=False):
        group = group.sort_values("date")
        total = group["amount"].sum()
        count = len(group)

        out()
        out("=" * 64)
        out(f"  KLIENT: {name}")
        out("-" * 64)

        # Tabela transakcji
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

    # ── Tabela zbiorcza ──
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
        description="Suma Wpłat — wpłaty od klientów indywidualnych z wyciągu mBank",
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

    # Sprawdzamy plik
    if not Path(args.file).exists():
        print(f"Błąd: plik nie znaleziony: {args.file}", file=sys.stderr)
        sys.exit(1)

    # Analiza
    try:
        df = analyze(args.file, args.encoding)
    except Exception as exc:
        print(f"Błąd podczas odczytu pliku: {exc}", file=sys.stderr)
        sys.exit(1)

    # Wynik
    output_lines: list = []
    print_report(df, output_lines)

    # Zapis do pliku
    if args.output:
        try:
            with open(args.output, "w", encoding="utf-8") as f:
                f.write("\n".join(output_lines))
            print(f"\nRaport zapisany: {args.output}")
        except OSError as exc:
            print(f"Nie udało się zapisać pliku: {exc}", file=sys.stderr)


if __name__ == "__main__":
    # Przełączamy stdout/stderr na UTF-8 dla terminala Windows
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    main()
