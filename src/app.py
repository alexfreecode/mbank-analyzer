"""
app.py — nakładka GUI dla analizatora wyciągu mBank
Uruchomienie: python app.py  (lub pythonw app.py — bez konsoli)
"""

import json
import os
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

# Importujemy logikę analizy i generowania faktur
from parser import analyze, print_report, print_reconciliation_report
from saldeo_export import VAT_BASIS, generate_saldeo_xlsx
from contractor_check import load_saldeo_contractors, check_clients

# ─── Konfiguracja (zachowywana między uruchomieniami) ─────────────────────────


def _config_dir() -> Path:
    """Folder dla config.json.

    Przy uruchomieniu zbudowanego .exe (PyInstaller --onefile) __file__
    wskazuje na folder tymczasowy, usuwany po zamknięciu — nie można tam
    przechowywać ustawień. Dlatego dla „zamrożonej” aplikacji używamy
    folderu profilu użytkownika (%APPDATA%), a przy uruchomieniu zwykłego
    .py — folderu obok skryptu (wygodne przy programowaniu).
    """
    if getattr(sys, "frozen", False):
        base = Path(os.environ.get("APPDATA", Path.home()))
        d = base / "mBank Analyzer"
        try:
            d.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass
        return d
    return Path(__file__).parent


_CONFIG_PATH = _config_dir() / "config.json"

def _load_config() -> dict:
    try:
        return json.loads(_CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}

def _save_config(data: dict) -> None:
    try:
        _CONFIG_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2),
                                encoding="utf-8")
    except Exception:
        pass


# ─── Stałe ────────────────────────────────────────────────────────────────────

TITLE     = "Analizator wyciągu mBank"
FONT_MONO = ("Courier New", 10)
FONT_UI   = ("Segoe UI", 10)
FONT_BTN  = ("Segoe UI", 10)
WIN_W, WIN_H = 860, 680
PAD = 10

# ─── Tekst pomocy (wyświetlany w oknie „Pomoc”) ───────────────────────────────

HELP_TEXT = """\
POMOC — Analizator wyciągu mBank
══════════════════════════════════════════════════════════════════════════════

1. DO CZEGO SŁUŻY TEN PROGRAM?
─────────────────────────────────────────────────────────────────────────────
Program wczytuje listę operacji wyeksportowaną z bankowości internetowej
mBank (plik CSV), rozpoznaje wpłaty od poszczególnych klientów, przygotowuje
czytelny raport zbiorczy oraz — w razie potrzeby — plik gotowy do
zaimportowania jako faktury do Saldeo Smart.

Typowy cykl pracy: co miesiąc pobierasz z bankowości internetowej (sekcja
„Historia” → „Eksportuj listę”) listę operacji w formacie CSV → wskazujesz ją
w programie → uruchamiasz analizę → (opcjonalnie) generujesz na jej podstawie
faktury do importu w Saldeo.


2. PODSTAWOWA ANALIZA LISTY OPERACJI
─────────────────────────────────────────────────────────────────────────────
  1. Kliknij „Wybierz...” przy polu „Plik wejściowy CSV” i wskaż wyeksportowaną
     z bankowości internetowej listę operacji (plik .csv).
  2. Jeśli chcesz dodatkowo zapisać wynik do pliku tekstowego — wskaż jego
     ścieżkę w polu „Plik wyjściowy TXT (opcjonalnie)” (krok ten można pominąć).
  3. Sprawdź „Kodowanie pliku” — zwykle właściwe jest „utf-8-sig”; jeżeli polskie
     litery (ą, ę, ś, ż...) wyświetlają się jako „krzaki”, zmień na „cp1250”
     i uruchom analizę ponownie.
  4. Kliknij „▶ Uruchom analizę”. W oknie poniżej pojawi się raport: lista
     wpłat pogrupowana według klientów wraz z sumami miesięcznymi i łączną
     tabelą zbiorczą.


3. GENEROWANIE FAKTUR DLA SALDEO SMART
─────────────────────────────────────────────────────────────────────────────
Po pomyślnie zakończonej analizie aktywuje się przycisk „Faktury Saldeo...”.
Otwiera on okno, w którym można:

  • zaznaczyć / odznaczyć klientów, dla których mają zostać wystawione faktury
    (przyciski „Zaznacz wszystko” / „Odznacz wszystko” ułatwiają pracę przy
    dłuższych listach — Twój wybór jest zapamiętywany do następnego razu),
  • ustawić numer początkowy faktury („Lp.”),
  • wybrać „Podstawę zastosowania stawki ZW” (a113, a43, a82, du, iz — zgodnie
    z tym, jak rozliczasz zwolnienie z VAT z urzędem skarbowym),
  • zdecydować, czy oznaczyć faktury jako zapłacone (pole „Zapłacono”).
    Domyślnie WYŁĄCZONE — faktury trafiają do Saldeo jako nieopłacone, dzięki
    czemu możesz je później uzgodnić z wpłatami z wyciągu. Włącz tę opcję tylko,
    jeśli chcesz od razu zaznaczyć je jako opłacone,
  • wskazać plik wynikowy .xlsx,
  • opcjonalnie wskazać bazę kontrahentów Saldeo, by uniknąć duplikatów —
    patrz punkt 5 poniżej.

Po kliknięciu „Generuj” program tworzy plik Excel gotowy do zaimportowania:
w Saldeo Smart → Faktury → Importuj z pliku.

Uwaga: pole „Data dostawy” Saldeo nie przyjmuje zakresu dat — jeśli klient
płacił kilka razy w miesiącu, program wpisuje datę jego OSTATNIEJ wpłaty w tym
miesiącu. Jeżeli klient w jednym miesiącu zapłacił różne kwoty, program zapisze
to jako jedną pozycję „1 szt. × suma wpłat” i wypisze stosowne ostrzeżenie —
taką fakturę warto sprawdzić ręcznie po imporcie.


4. DANE SPRZEDAWCY
─────────────────────────────────────────────────────────────────────────────
Przycisk „⚙ Dane sprzedawcy...” otwiera okno, w którym można wprowadzić lub
poprawić w dowolnej chwili:

  • imię / nazwę sprzedawcy — pojawia się na fakturze jako „Wystawca faktury”,
  • numer konta bankowego — pole „Konto bankowe” na fakturze,
  • nazwę usługi — pole „Nazwa towaru” na fakturze (np. „Konsultacja
    psychologiczna”).

Dane te zapisywane są lokalnie na Twoim komputerze i wykorzystywane przy
każdym kolejnym generowaniu faktur — nie trzeba wpisywać ich za każdym razem.
Nikt poza Tobą nie ma do nich dostępu — nie są nigdzie wysyłane.


5. JAK UNIKNĄĆ DUPLIKATÓW KONTRAHENTÓW W SALDEO
─────────────────────────────────────────────────────────────────────────────
To opcjonalna, ale zalecana funkcja. Chroni przed sytuacją, w której w bazie
kontrahentów Saldeo z czasem pojawia się wiele kart dla tej samej osoby —
np. raz zapisanej jako „Jan Kowalski”, innym razem jako „Kowalski Jan”, a
jeszcze innym razem z drobną literówką. Powoduje to rozdrobnienie historii
płatności klienta na kilka kart i utrudnia później porządki w Saldeo.

Jak z tego skorzystać — krok po kroku:

  a) Zaloguj się do Saldeo Smart, przejdź do sekcji „Kontrahenci” i wyeksportuj
     listę kontrahentów do pliku CSV (zwykle przycisk „Eksportuj”).

  b) Wróć do naszego programu — w oknie „Faktury Saldeo...”, w polu „Baza
     kontrahentów Saldeo — plik CSV (opcjonalnie)”, kliknij „Wczytaj...”
     i wskaż plik pobrany w poprzednim kroku. Ścieżka do niego zostanie
     zapamiętana — przy kolejnych generowaniach wystarczy ją od czasu do
     czasu odświeżyć nową wersją wyeksportowanej listy.

  c) Kliknij „Generuj”. Program porówna nazwy klientów z listy operacji z bazą
     kontrahentów Saldeo i obsłuży trzy sytuacje:

       • DOKŁADNE DOPASOWANIE — nic się nie dzieje, faktura trafi prosto do
         istniejącej karty kontrahenta;

       • PODOBNA, ALE NIE IDENTYCZNA NAZWA — otworzy się dodatkowe okno
         „Podobni kontrahenci w bazie Saldeo”. Dla każdej takiej pary zobaczysz
         nazwę z listy operacji i podobną nazwę już istniejącą w Saldeo, a obok —
         pole wyboru „To ten sam kontrahent — użyj nazwy z bazy Saldeo”:
            – zaznacz je, jeśli to ta sama osoba — wówczas faktura zostanie
              zapisana pod nazwą DOKŁADNIE TAKĄ, jak w bazie Saldeo, dzięki
              czemu przy imporcie trafi do istniejącej karty (bez duplikatu);
            – zostaw niezaznaczone, jeśli to inna osoba — nazwa pozostanie
              taka, jak na liście operacji (Saldeo utworzy dla niej osobną kartę);
         przycisk „Anuluj” w tym oknie przerywa całe generowanie faktur, jeśli
         wolisz najpierw wyjaśnić wątpliwości;

       • ZUPEŁNIE NOWY KLIENT — program tylko informuje, że dla takiej osoby
         Saldeo utworzy nową kartę kontrahenta. To normalne i oczekiwane przy
         pojawieniu się nowego klienta.

  d) Po zapisaniu pliku zobaczysz podsumowanie wszystkich tych ustaleń w
     oknie „Gotowe (z uwagami)”, np.:
        ✓  „Kowalski Jan” zostanie zapisane w fakturze jako
           „JAN KOWALSKI” (zgodnie z bazą Saldeo)...
        ℹ  Nowy kontrahent: „Anna Nowak” — zostanie utworzona dla niego
           nowa karta...

Jeśli nie wskażesz pliku z bazą kontrahentów — program po prostu pominie ten
krok i będzie działał tak, jak dotychczas.


6. NAJCZĘSTSZE PYTANIA
─────────────────────────────────────────────────────────────────────────────
P: Program pokazuje klienta jako „NIEZNANY” — co to znaczy?
O: Nie udało się rozpoznać nazwiska nadawcy w opisie operacji. Zdarza się to
   przy nietypowych formatach przelewów. Taką pozycję warto sprawdzić i w razie
   potrzeby poprawić ręcznie po imporcie do Saldeo.

P: Dlaczego w polu „Data dostawy” jest tylko jedna data, choć klient płacił
   kilka razy w miesiącu?
O: Saldeo nie przyjmuje zakresu dat w tym polu — program celowo wpisuje datę
   ostatniej wpłaty z danego miesiąca. To zgodne z przyjętą praktyką pracy.

P: Czy moje dane (imię, numer konta) są gdzieś wysyłane?
O: Nie. Wszystkie dane — Twoje dane sprzedawcy oraz wczytywane pliki — są
   przetwarzane wyłącznie lokalnie, na Twoim komputerze, i nigdzie nie są
   przesyłane.


7. WSPARCIE AUTORA (całkowicie dobrowolne)
─────────────────────────────────────────────────────────────────────────────
Program jest darmowy i takim pozostanie — to w żaden sposób się nie zmieni.
Jeśli jednak zaoszczędził Ci czasu i miał(a)byś ochotę w jakiś sposób
podziękować autorowi, możesz wysłać dowolną kwotę (choćby symboliczną,
„na kawę”) przez Revolut:

    revolut.me/oleksa49b        (RevTag: @oleksa49b)

To czysto symboliczny gest — funkcje programu w żaden sposób od niego nie
zależą.
"""


# ─── Kreator pierwszego uruchomienia (dane sprzedawcy) ────────────────────────

class SellerSetupDialog(tk.Toplevel):
    """
    Pokazywany raz — przy pierwszym uruchomieniu programu (gdy w config.json
    nie ma jeszcze danych sprzedawcy). Pyta o imię/nazwę sprzedawcy, numer
    rachunku bankowego i nazwę usługi — te dane wstawiane są do generowanych
    faktur Saldeo. Wpisane wartości zapisywane są lokalnie w config.json
    i nigdy więcej nie są pytane.

    Dzięki temu w kodzie programu nie ma ani jednej zaszytej na stałe danej
    osobistej — program można swobodnie przekazywać innym użytkownikom.
    """

    def __init__(self, parent: tk.Tk, first_run: bool = True):
        super().__init__(parent)
        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self._first_run = first_run

        self.title("Konfiguracja — pierwsze uruchomienie" if first_run
                   else "Dane sprzedawcy")
        self.resizable(False, False)
        self.configure(bg="#f0f0f0")

        outer = tk.Frame(self, bg="#f0f0f0", padx=18, pady=16)
        outer.pack(fill=tk.BOTH, expand=True)

        if first_run:
            intro = ("Witaj! Zanim zaczniesz, podaj swoje dane sprzedawcy —\n"
                     "będą automatycznie wpisywane na fakturach Saldeo.\n"
                     "Dane zapiszą się lokalnie na tym komputerze i nie będą\n"
                     "już pytane przy kolejnych uruchomieniach.")
        else:
            intro = ("Tutaj możesz poprawić swoje dane sprzedawcy —\n"
                     "np. jeśli pomyliłeś się przy pierwszym uruchomieniu.\n"
                     "Zmiany zostaną zapisane lokalnie na tym komputerze.")

        tk.Label(
            outer, text=intro,
            font=FONT_UI, bg="#f0f0f0", justify=tk.LEFT,
        ).pack(anchor="w", pady=(0, 14))

        cfg = _load_config()

        def _field(label_text, default=""):
            row = tk.Frame(outer, bg="#f0f0f0")
            row.pack(fill=tk.X, pady=(0, 8))
            tk.Label(row, text=label_text, font=FONT_UI, bg="#f0f0f0",
                     width=24, anchor="w").pack(side=tk.LEFT)
            var = tk.StringVar(value=default)
            tk.Entry(row, textvariable=var, font=FONT_UI, width=40).pack(
                side=tk.LEFT, fill=tk.X, expand=True)
            return var

        self.name_var    = _field("Imię / nazwa sprzedawcy:",
                                  cfg.get("seller_name", ""))
        self.account_var = _field("Numer konta bankowego:",
                                  cfg.get("seller_account", ""))
        self.service_var = _field("Nazwa usługi (na fakturze):",
                                  cfg.get("service_name", "Usługa"))

        btn_row = tk.Frame(outer, bg="#f0f0f0")
        btn_row.pack(pady=(10, 0))

        tk.Button(btn_row,
                  text=("  Zapisz i kontynuuj  " if first_run else "  Zapisz  "),
                  font=("Segoe UI", 11, "bold"),
                  bg="#107c10", fg="white",
                  activebackground="#0a5b0a", activeforeground="white",
                  relief=tk.FLAT, cursor="hand2", padx=8, pady=4,
                  command=self._save).pack(side=tk.LEFT)

        if not first_run:
            tk.Button(btn_row, text="Anuluj", font=FONT_BTN, width=10,
                      command=self.destroy).pack(side=tk.LEFT, padx=(10, 0))

        # Wyśrodkowanie względem okna rodzica
        self.update_idletasks()
        pw = parent.winfo_x() + parent.winfo_width()  // 2
        ph = parent.winfo_y() + parent.winfo_height() // 2
        self.geometry(f"+{pw - self.winfo_width()//2}+{ph - self.winfo_height()//2}")

        self.wait_window()

    def _save(self):
        name    = self.name_var.get().strip()
        account = self.account_var.get().strip()
        service = self.service_var.get().strip()

        if not name or not account:
            messagebox.showerror(
                "Błąd",
                "Podaj przynajmniej nazwę sprzedawcy i numer konta —\n"
                "te dane są niezbędne do wygenerowania faktur.",
                parent=self,
            )
            return

        cfg = _load_config()
        cfg["seller_name"]    = name
        cfg["seller_account"] = account
        cfg["service_name"]   = service or "Usługa"
        _save_config(cfg)

        self.destroy()

    def _on_close(self):
        # Nie zamykamy użytkownika na siłę — ale bez danych sprzedawcy
        # generowanie faktur będzie po prostu wstawiać puste wartości,
        # dopóki ich nie uzupełni (kreator można otworzyć ponownie,
        # restartując program po usunięciu seller_name/seller_account
        # z config.json).
        self.destroy()


# ─── Okno „Podobni kontrahenci w bazie Saldeo” ────────────────────────────────

class ContractorMatchDialog(tk.Toplevel):
    """
    Modalne okno rozstrzygania „podejrzanie podobnych” dopasowań do bazy
    kontrahentów Saldeo.

    Dla każdej pary (nazwa z wyciągu ↔ podobna nazwa z bazy Saldeo) użytkownik
    decyduje: to ta sama osoba czy nie.
      • Jeśli „tak” — w fakturze zostanie użyta nazwa Z BAZY SALDEO, aby
        Saldeo przy imporcie podpięło fakturę pod istniejącą kartę kontrahenta
        (a nie utworzyło zduplikowanej).
      • Jeśli „nie” — nazwa pozostanie jak w wyciągu (powstanie nowa karta).

    Wynik — self.result:
      dict {nazwa_z_wyciągu: nazwa_z_bazy_Saldeo} tylko dla potwierdzonych par,
      albo None, gdy użytkownik kliknął „Anuluj” (generowanie należy przerwać).
    """

    def __init__(self, parent: tk.Tk, matches: list[dict]):
        super().__init__(parent)
        self.transient(parent)
        self.grab_set()

        self.title("Podobni kontrahenci w bazie Saldeo")
        self.resizable(False, False)
        self.configure(bg="#f0f0f0")

        self._matches = matches
        self._vars: list[tk.BooleanVar] = []
        self.result: dict[str, str] | None = None

        self._build_ui()

        self.update_idletasks()
        pw = parent.winfo_x() + parent.winfo_width()  // 2
        ph = parent.winfo_y() + parent.winfo_height() // 2
        self.geometry(f"+{pw - self.winfo_width()//2}+{ph - self.winfo_height()//2}")

        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.wait_window()

    def _build_ui(self):
        outer = tk.Frame(self, bg="#f0f0f0", padx=16, pady=14)
        outer.pack(fill=tk.BOTH, expand=True)

        tk.Label(outer,
                 text=("Niektóre nazwy klientów z listy operacji są bardzo podobne do\n"
                       "kontrahentów już istniejących w bazie Saldeo. Zaznacz, jeśli\n"
                       "to ta sama osoba — wtedy faktura zostanie wystawiona pod\n"
                       "nazwą z bazy Saldeo, co pozwoli uniknąć powstania\n"
                       "zduplikowanej karty kontrahenta."),
                 font=FONT_UI, bg="#f0f0f0", justify="left",
                 anchor="w").pack(anchor="w", pady=(0, 12))

        list_frame = tk.Frame(outer, bg="#f0f0f0")
        list_frame.pack(fill=tk.BOTH, expand=True)

        for m in self._matches:
            var = tk.BooleanVar(value=False)
            self._vars.append(var)

            row = tk.Frame(list_frame, bg="white", relief=tk.GROOVE, bd=1)
            row.pack(fill=tk.X, pady=3)

            tk.Label(row, text=f"Z listy operacji: „{m['name']}”",
                     font=FONT_UI, bg="white", anchor="w").pack(fill=tk.X, padx=8, pady=(6, 0))
            tk.Label(row, text=f"W bazie Saldeo:  „{m['matched']}”",
                     font=FONT_UI, bg="white", fg="#0078d4", anchor="w").pack(fill=tk.X, padx=8)
            tk.Checkbutton(row, variable=var, bg="white", anchor="w",
                           font=FONT_UI,
                           text="To ten sam kontrahent — użyj nazwy z bazy Saldeo"
                           ).pack(fill=tk.X, padx=6, pady=(0, 6))

        btn_row = tk.Frame(outer, bg="#f0f0f0")
        btn_row.pack(pady=(14, 0))

        tk.Button(btn_row, text="  Kontynuuj generowanie  ",
                  font=("Segoe UI", 11, "bold"),
                  bg="#107c10", fg="white",
                  activebackground="#0a5b0a", activeforeground="white",
                  relief=tk.FLAT, cursor="hand2", padx=8, pady=4,
                  command=self._confirm).pack(side=tk.LEFT)

        tk.Button(btn_row, text="Anuluj", font=FONT_BTN, width=10,
                  command=self._cancel).pack(side=tk.LEFT, padx=(10, 0))

    def _confirm(self):
        self.result = {
            m["name"]: m["matched"]
            for m, var in zip(self._matches, self._vars)
            if var.get()
        }
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()


# ─── Okno „Pomoc” ─────────────────────────────────────────────────────────────

class HelpDialog(tk.Toplevel):
    """Okno pomocy — pokazuje HELP_TEXT w przewijanym polu tekstowym."""

    def __init__(self, parent: tk.Tk):
        super().__init__(parent)
        self.transient(parent)

        self.title("Pomoc — Analizator wyciągu mBank")
        self.configure(bg="#f0f0f0")
        self.geometry("760x620")
        self.minsize(480, 360)

        outer = tk.Frame(self, bg="#f0f0f0", padx=12, pady=12)
        outer.pack(fill=tk.BOTH, expand=True)

        text = scrolledtext.ScrolledText(
            outer, font=FONT_MONO, wrap=tk.WORD,
            bg="white", fg="#1e1e1e",
            relief=tk.SUNKEN, bd=1,
        )
        text.pack(fill=tk.BOTH, expand=True)
        text.insert("1.0", HELP_TEXT)
        text.configure(state=tk.DISABLED)

        tk.Button(outer, text="Zamknij", font=FONT_BTN, width=12,
                  command=self.destroy).pack(pady=(10, 0))

        # Wyśrodkowanie względem okna rodzica
        self.update_idletasks()
        pw = parent.winfo_x() + parent.winfo_width()  // 2
        ph = parent.winfo_y() + parent.winfo_height() // 2
        self.geometry(f"+{pw - self.winfo_width()//2}+{ph - self.winfo_height()//2}")


# ─── Okno „Kontrola kompletności wyciągu” ─────────────────────────────────────

class ReconciliationDialog(tk.Toplevel):
    """Okno osobnego raportu „Kontrola kompletności” —
    podział WSZYSTKICH operacji wyciągu na Klienci / Pozostałe wpływy / Wydatki."""

    def __init__(self, parent: tk.Tk, report_text: str, default_dir: str, default_name: str):
        super().__init__(parent)
        self.transient(parent)

        self._report_text = report_text
        self._default_dir = default_dir
        self._default_name = default_name

        self.title("Kontrola kompletności wyciągu")
        self.configure(bg="#f0f0f0")
        self.geometry("760x620")
        self.minsize(480, 360)

        outer = tk.Frame(self, bg="#f0f0f0", padx=12, pady=12)
        outer.pack(fill=tk.BOTH, expand=True)

        text = scrolledtext.ScrolledText(
            outer, font=FONT_MONO, wrap=tk.NONE,
            bg="#1e1e1e", fg="#d4d4d4",
            insertbackground="white",
            relief=tk.SUNKEN, bd=1,
        )
        text.pack(fill=tk.BOTH, expand=True)
        text.insert("1.0", report_text)
        text.configure(state=tk.DISABLED)

        btn_row = tk.Frame(outer, bg="#f0f0f0")
        btn_row.pack(pady=(10, 0))

        tk.Button(btn_row, text="Zapisz jako...", font=FONT_BTN, width=14,
                  command=self._save).pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(btn_row, text="Zamknij", font=FONT_BTN, width=12,
                  command=self.destroy).pack(side=tk.LEFT)

        # Wyśrodkowanie względem okna rodzica
        self.update_idletasks()
        pw = parent.winfo_x() + parent.winfo_width()  // 2
        ph = parent.winfo_y() + parent.winfo_height() // 2
        self.geometry(f"+{pw - self.winfo_width()//2}+{ph - self.winfo_height()//2}")

    def _save(self):
        path = filedialog.asksaveasfilename(
            title="Zapisz raport jako...",
            defaultextension=".txt",
            filetypes=[("Pliki tekstowe", "*.txt"), ("Wszystkie pliki", "*.*")],
            initialdir=self._default_dir,
            initialfile=self._default_name,
        )
        if not path:
            return
        try:
            Path(path).write_text(self._report_text, encoding="utf-8")
        except OSError as exc:
            messagebox.showwarning("Nie udało się zapisać pliku", str(exc))


# ─── Okno „Generowanie faktur Saldeo” ─────────────────────────────────────────

class SaldeoDialog(tk.Toplevel):
    """Modalne okno konfiguracji i generowania pliku Excel importu Saldeo."""

    def __init__(self, parent: tk.Tk, df):
        super().__init__(parent)
        self.transient(parent)
        self.grab_set()

        self.title("Generowanie faktur Saldeo")
        self.resizable(False, False)
        self.configure(bg="#f0f0f0")

        self._df = df
        self._client_vars: dict[str, tk.BooleanVar] = {}
        self._result_path: str | None = None

        self._build_ui()

        # Wyśrodkowanie względem okna rodzica
        self.update_idletasks()
        pw = parent.winfo_x() + parent.winfo_width()  // 2
        ph = parent.winfo_y() + parent.winfo_height() // 2
        self.geometry(f"+{pw - self.winfo_width()//2}+{ph - self.winfo_height()//2}")

        self.wait_window()

    # ── Budowa interfejsu ──────────────────────────────────────────────────────

    def _build_ui(self):
        outer = tk.Frame(self, bg="#f0f0f0", padx=14, pady=12)
        outer.pack(fill=tk.BOTH, expand=True)

        # ── Lista klientów ──
        tk.Label(outer, text="Klienci do uwzględnienia w fakturach:",
                 font=FONT_UI, bg="#f0f0f0").pack(anchor="w")

        list_frame = tk.Frame(outer, bg="white", relief=tk.SUNKEN, bd=1)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(4, 10))

        canvas    = tk.Canvas(list_frame, bg="white", highlightthickness=0, height=200)
        scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL, command=canvas.yview)
        inner     = tk.Frame(canvas, bg="white")

        inner.bind("<Configure>",
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT,  fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Wypełniamy listę klientów; wcześniej odznaczonych przywracamy z konfiguracji
        saved_excluded = set(_load_config().get("excluded_clients", []))
        all_clients = sorted(self._df["name"].unique())
        for client in all_clients:
            var = tk.BooleanVar(value=client not in saved_excluded)
            self._client_vars[client] = var

            row = tk.Frame(inner, bg="white")
            row.pack(fill=tk.X, padx=6, pady=1)

            tk.Checkbutton(row, text=client, variable=var,
                           bg="white", font=FONT_UI, anchor="w").pack(side=tk.LEFT)

        # Przyciski „Zaznacz wszystko / Odznacz wszystko”
        btn_row = tk.Frame(outer, bg="#f0f0f0")
        btn_row.pack(fill=tk.X, pady=(0, 10))
        tk.Button(btn_row, text="Zaznacz wszystko",   font=FONT_BTN, width=14,
                  command=lambda: self._toggle_all(True)).pack(side=tk.LEFT)
        tk.Button(btn_row, text="Odznacz wszystko", font=FONT_BTN, width=16,
                  command=lambda: self._toggle_all(False)).pack(side=tk.LEFT,
                                                                padx=(6, 0))

        # ── Początkowy numer faktury ──
        num_row = tk.Frame(outer, bg="#f0f0f0")
        num_row.pack(fill=tk.X, pady=(0, 8))

        tk.Label(num_row, text="Początkowy numer faktury (Lp.):",
                 font=FONT_UI, bg="#f0f0f0").pack(side=tk.LEFT)
        self.inv_num_var = tk.IntVar(value=1)
        tk.Spinbox(num_row, from_=1, to=9999,
                   textvariable=self.inv_num_var,
                   width=6, font=FONT_UI).pack(side=tk.LEFT, padx=(8, 0))

        # ── Podstawa zastosowania stawki ZW ──
        zw_row = tk.Frame(outer, bg="#f0f0f0")
        zw_row.pack(fill=tk.X, pady=(0, 10))

        tk.Label(zw_row, text="Podstawa zastosowania stawki ZW:",
                 font=FONT_UI, bg="#f0f0f0").pack(side=tk.LEFT)

        cfg = _load_config()
        saved_basis = cfg.get("vat_basis", VAT_BASIS)
        self.vat_basis_var = tk.StringVar(value=saved_basis)
        ttk.Combobox(zw_row, textvariable=self.vat_basis_var,
                     values=["a113", "a43", "a82", "du", "iz"],
                     state="readonly", width=8,
                     font=FONT_UI).pack(side=tk.LEFT, padx=(8, 0))

        # ── Oznaczenie faktur jako zapłacone (domyślnie WYŁĄCZONE) ──
        self.mark_paid_var = tk.BooleanVar(value=cfg.get("mark_paid", False))
        paid_row = tk.Frame(outer, bg="#f0f0f0")
        paid_row.pack(fill=tk.X, pady=(0, 2))
        tk.Checkbutton(
            paid_row,
            text="Oznacz faktury jako zapłacone (wypełnij kolumnę „Zapłacono”)",
            variable=self.mark_paid_var,
            bg="#f0f0f0", font=FONT_UI, anchor="w",
        ).pack(anchor="w")
        tk.Label(outer,
                 text=("Domyślnie wyłączone — faktury importują się jako nieopłacone,\n"
                       "dzięki czemu można je później uzgodnić z wpłatami z wyciągu."),
                 font=("Segoe UI", 8), fg="#666666", bg="#f0f0f0",
                 justify="left").pack(anchor="w", pady=(0, 10))

        # ── Porównanie z bazą kontrahentów Saldeo (opcjonalnie) ──
        tk.Label(outer, text="Baza kontrahentów Saldeo — plik CSV (opcjonalnie):",
                 font=FONT_UI, bg="#f0f0f0").pack(anchor="w")

        ref_row = tk.Frame(outer, bg="#f0f0f0")
        ref_row.pack(fill=tk.X, pady=(4, 2))

        self.contractors_csv_var = tk.StringVar(
            value=cfg.get("saldeo_contractors_csv", ""))
        tk.Entry(ref_row, textvariable=self.contractors_csv_var,
                 font=FONT_UI, width=46).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(ref_row, text="Wczytaj...", font=FONT_BTN, width=12,
                  command=self._browse_contractors_csv).pack(side=tk.LEFT, padx=(6, 0))

        tk.Label(outer,
                 text=("Wskaż plik wyeksportowany z Saldeo (Kontrahenci → Eksportuj),\n"
                       "a program ostrzeże o nowych kontrahentach i podobnych nazwach —\n"
                       "by uniknąć powstawania zduplikowanych kart w bazie Saldeo."),
                 font=("Segoe UI", 8), fg="#666666", bg="#f0f0f0",
                 justify="left").pack(anchor="w", pady=(0, 10))

        # ── Plik wyjściowy ──
        tk.Label(outer, text="Plik wyjściowy Excel (.xlsx):",
                 font=FONT_UI, bg="#f0f0f0").pack(anchor="w")

        out_row = tk.Frame(outer, bg="#f0f0f0")
        out_row.pack(fill=tk.X, pady=(4, 10))

        self.out_var = tk.StringVar()
        tk.Entry(out_row, textvariable=self.out_var,
                 font=FONT_UI, width=52).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(out_row, text="Zapisz...", font=FONT_BTN, width=12,
                  command=self._browse_output).pack(side=tk.LEFT, padx=(6, 0))

        # ── Przyciski akcji ──
        action_row = tk.Frame(outer, bg="#f0f0f0")
        action_row.pack(pady=(4, 0))

        tk.Button(action_row,
                  text="  Generuj  ",
                  font=("Segoe UI", 11, "bold"),
                  bg="#107c10", fg="white",
                  activebackground="#0a5b0a", activeforeground="white",
                  relief=tk.FLAT, cursor="hand2", padx=8, pady=4,
                  command=self._generate).pack(side=tk.LEFT)

        tk.Button(action_row, text="Anuluj", font=FONT_BTN, width=10,
                  command=self.destroy).pack(side=tk.LEFT, padx=(10, 0))

    # ── Obsługa zdarzeń ────────────────────────────────────────────────────────

    def _toggle_all(self, state: bool):
        for var in self._client_vars.values():
            var.set(state)

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Zapisz plik importu Saldeo",
            defaultextension=".xlsx",
            filetypes=[("Pliki Excel", "*.xlsx"), ("Wszystkie pliki", "*.*")],
        )
        if path:
            self.out_var.set(path)

    def _browse_contractors_csv(self):
        path = filedialog.askopenfilename(
            title="Wskaż plik wyeksportowany z Saldeo (lista kontrahentów, CSV)",
            filetypes=[("Pliki CSV", "*.csv"), ("Wszystkie pliki", "*.*")],
        )
        if path:
            self.contractors_csv_var.set(path)

    def _generate(self):
        out_path = self.out_var.get().strip()
        if not out_path:
            messagebox.showerror("Błąd", "Podaj ścieżkę zapisu pliku.",
                                 parent=self)
            return
        # Dodajemy .xlsx, gdy nie podano rozszerzenia
        if not Path(out_path).suffix:
            out_path += ".xlsx"
            self.out_var.set(out_path)

        excluded  = {name for name, var in self._client_vars.items() if not var.get()}
        start_num = self.inv_num_var.get()
        vat_basis = self.vat_basis_var.get().strip()
        contractors_csv = self.contractors_csv_var.get().strip()
        mark_paid = self.mark_paid_var.get()

        # Zapisujemy ustawienia na następne uruchomienie
        cfg = _load_config()
        cfg["vat_basis"]               = vat_basis
        cfg["excluded_clients"]        = sorted(excluded)   # odznaczone pola
        cfg["saldeo_contractors_csv"]  = contractors_csv
        cfg["mark_paid"]               = mark_paid
        _save_config(cfg)

        # ── Porównanie z bazą kontrahentów Saldeo (gdy wskazano plik) ──
        contractor_warnings: list[str] = []
        name_overrides: dict[str, str] = {}
        if contractors_csv:
            try:
                contractors = load_saldeo_contractors(contractors_csv)
                if not contractors:
                    contractor_warnings.append(
                        "⚠ Nie udało się odczytać kontrahentów ze wskazanego pliku CSV — "
                        "upewnij się, że to plik wyeksportowany z Saldeo "
                        "(Kontrahenci → Eksportuj)."
                    )
                else:
                    included = sorted(
                        name for name, var in self._client_vars.items() if var.get()
                    )
                    check_results = check_clients(included, contractors)
                    similar = [r for r in check_results if r["status"] == "similar"]
                    new     = [r for r in check_results if r["status"] == "new"]

                    # Podejrzanie podobne nazwy — decyzję podejmuje użytkownik:
                    # „czy to ten sam kontrahent?” → jeśli tak, w fakturze
                    # zostanie użyta nazwa z bazy Saldeo (jak w bazie),
                    # by nie tworzyć zduplikowanej karty.
                    if similar:
                        dlg = ContractorMatchDialog(self, similar)
                        if dlg.result is None:
                            return  # użytkownik kliknął „Anuluj” — przerywamy generowanie
                        name_overrides = dlg.result

                    for r in similar:
                        if r["name"] in name_overrides:
                            contractor_warnings.append(
                                f"✓ „{r['name']}” zostanie zapisane w fakturze jako "
                                f"„{name_overrides[r['name']]}” — zgodnie z bazą Saldeo, "
                                "by uniknąć powstania zduplikowanej karty kontrahenta."
                            )
                        else:
                            contractor_warnings.append(
                                f"⚠ „{r['name']}” jest podobne do kontrahenta "
                                f"„{r['matched']}” w bazie Saldeo, ale wskazano, że to "
                                "inna osoba — zostanie utworzona osobna karta kontrahenta."
                            )

                    for r in new:
                        contractor_warnings.append(
                            f"ℹ Nowy kontrahent: „{r['name']}” — nie znaleziono go "
                            "w bazie kontrahentów Saldeo. Przy imporcie zostanie "
                            "utworzona dla niego nowa karta."
                        )
            except Exception as exc:
                contractor_warnings.append(
                    f"⚠ Błąd podczas wczytywania bazy kontrahentów Saldeo: {exc}"
                )

        try:
            _, warnings = generate_saldeo_xlsx(
                self._df,
                out_path,
                start_inv_num=start_num,
                exclude_names=excluded,
                name_overrides=name_overrides or None,
                vat_basis=vat_basis or None,
                seller_name=cfg.get("seller_name") or None,
                seller_account=cfg.get("seller_account") or None,
                service_name=cfg.get("service_name") or None,
                mark_paid=mark_paid,
            )
        except Exception as exc:
            messagebox.showerror("Błąd podczas generowania", str(exc), parent=self)
            return

        all_warnings = contractor_warnings + warnings

        msg = f"Plik zapisany:\n{out_path}"
        if all_warnings:
            msg += "\n\nUwagi:\n" + "\n".join(all_warnings)
            messagebox.showwarning("Gotowe (z uwagami)", msg, parent=self)
        else:
            messagebox.showinfo("Gotowe", msg, parent=self)

        self._result_path = out_path
        self.destroy()


# ─── Okno główne ───────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(TITLE)
        self.resizable(True, True)
        self.minsize(640, 480)

        # Wyśrodkowanie okna na ekranie
        self.update_idletasks()
        x = (self.winfo_screenwidth()  - WIN_W) // 2
        y = (self.winfo_screenheight() - WIN_H) // 2
        self.geometry(f"{WIN_W}x{WIN_H}+{x}+{y}")

        self._df = None   # wynik ostatniej analizy

        self._build_ui()

        # Kreator pierwszego uruchomienia: gdy dane sprzedawcy nie są jeszcze
        # zapisane w config.json — pytamy o nie raz, przed rozpoczęciem pracy.
        cfg = _load_config()
        if not cfg.get("seller_name") or not cfg.get("seller_account"):
            self.update_idletasks()
            SellerSetupDialog(self, first_run=True)

    # ── Budowa interfejsu ──────────────────────────────────────────────────────

    def _build_ui(self):
        self.configure(bg="#f0f0f0")

        # ── Panel sterowania (góra) ──
        ctrl = tk.Frame(self, bg="#f0f0f0", padx=PAD, pady=PAD)
        ctrl.pack(fill=tk.X)

        # Plik wejściowy
        tk.Label(ctrl, text="Plik wejściowy CSV:", font=FONT_UI,
                 bg="#f0f0f0").grid(row=0, column=0, sticky="w", pady=(0, 2))

        self.input_var = tk.StringVar()
        tk.Entry(ctrl, textvariable=self.input_var,
                 font=FONT_UI, width=68).grid(row=1, column=0,
                                              sticky="ew", padx=(0, 6))
        tk.Button(ctrl, text="Wybierz...", font=FONT_BTN, width=12,
                  command=self._browse_input).grid(row=1, column=1)

        # Plik wyjściowy TXT
        tk.Label(ctrl, text="Plik wyjściowy TXT (opcjonalnie):", font=FONT_UI,
                 bg="#f0f0f0").grid(row=2, column=0, sticky="w", pady=(PAD, 2))

        self.output_var = tk.StringVar()
        tk.Entry(ctrl, textvariable=self.output_var,
                 font=FONT_UI, width=68).grid(row=3, column=0,
                                              sticky="ew", padx=(0, 6))
        tk.Button(ctrl, text="Zapisz...", font=FONT_BTN, width=12,
                  command=self._browse_output).grid(row=3, column=1)

        # Kodowanie
        enc_row = tk.Frame(ctrl, bg="#f0f0f0")
        enc_row.grid(row=4, column=0, columnspan=2, sticky="w", pady=(PAD, 0))

        tk.Label(enc_row, text="Kodowanie pliku:", font=FONT_UI,
                 bg="#f0f0f0").pack(side=tk.LEFT)
        self.enc_var = tk.StringVar(value="utf-8-sig")
        ttk.Combobox(enc_row, textvariable=self.enc_var,
                     values=["utf-8-sig", "cp1250"],
                     state="readonly", width=12,
                     font=FONT_UI).pack(side=tk.LEFT, padx=(8, 0))

        # ── Przyciski akcji ──
        btn_row = tk.Frame(ctrl, bg="#f0f0f0")
        btn_row.grid(row=5, column=0, columnspan=2, pady=(PAD + 4, 0))

        tk.Button(btn_row,
                  text="  ▶  Uruchom analizę  ",
                  font=("Segoe UI", 11, "bold"),
                  bg="#0078d4", fg="white",
                  activebackground="#005a9e", activeforeground="white",
                  relief=tk.FLAT, cursor="hand2", padx=8, pady=4,
                  command=self._run).pack(side=tk.LEFT)

        self._saldeo_btn = tk.Button(
            btn_row,
            text="  Faktury Saldeo...  ",
            font=("Segoe UI", 11),
            bg="#5c2d91", fg="white",
            activebackground="#3b1a5e", activeforeground="white",
            relief=tk.FLAT, cursor="hand2", padx=8, pady=4,
            state=tk.DISABLED,
            command=self._open_saldeo_dialog,
        )
        self._saldeo_btn.pack(side=tk.LEFT, padx=(12, 0))

        tk.Button(
            btn_row,
            text="  ✓ Kontrola kompletności  ",
            font=FONT_BTN,
            relief=tk.FLAT, cursor="hand2", padx=8, pady=4,
            command=self._open_reconciliation,
        ).pack(side=tk.LEFT, padx=(12, 0))

        tk.Button(
            btn_row,
            text="  ⚙ Dane sprzedawcy...  ",
            font=FONT_BTN,
            relief=tk.FLAT, cursor="hand2", padx=8, pady=4,
            command=self._open_seller_settings,
        ).pack(side=tk.LEFT, padx=(12, 0))

        tk.Button(
            btn_row,
            text="  ?  Pomoc  ",
            font=FONT_BTN,
            relief=tk.FLAT, cursor="hand2", padx=8, pady=4,
            command=self._open_help,
        ).pack(side=tk.LEFT, padx=(12, 0))

        ctrl.columnconfigure(0, weight=1)

        # ── Separator ──
        ttk.Separator(self, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=PAD, pady=2)

        # ── Pole tekstowe wyniku ──
        result_frame = tk.Frame(self, bg="#f0f0f0")
        result_frame.pack(fill=tk.BOTH, expand=True, padx=PAD, pady=(0, 4))

        tk.Label(result_frame, text="Wynik analizy:",
                 font=FONT_UI, bg="#f0f0f0").pack(anchor="w")

        self.result_text = scrolledtext.ScrolledText(
            result_frame,
            font=FONT_MONO,
            bg="#1e1e1e", fg="#d4d4d4",
            insertbackground="white",
            wrap=tk.NONE,
            state=tk.DISABLED,
        )
        self.result_text.pack(fill=tk.BOTH, expand=True)

        # Poziomy pasek przewijania
        h_scroll = tk.Scrollbar(result_frame, orient=tk.HORIZONTAL,
                                 command=self.result_text.xview)
        h_scroll.pack(fill=tk.X)
        self.result_text.configure(xscrollcommand=h_scroll.set)

        # ── Pasek stanu ──
        self.status_var = tk.StringVar(
            value="Wybierz plik z listą operacji i kliknij «Uruchom analizę»"
                  "   •   ☕ Spodobał się program? → zakładka „Pomoc”")
        tk.Label(self, textvariable=self.status_var,
                 font=("Segoe UI", 9), bg="#e0e0e0",
                 anchor="w", padx=PAD, pady=3,
                 relief=tk.SUNKEN).pack(fill=tk.X, side=tk.BOTTOM)

    # ── Obsługa przycisków ─────────────────────────────────────────────────────

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="Wybierz plik z listą operacji mBank",
            filetypes=[("Pliki CSV", "*.csv"), ("Wszystkie pliki", "*.*")],
        )
        if not path:
            return
        self.input_var.set(path)
        p = Path(path)
        self.output_var.set(str(p.with_suffix(".txt")))
        self.status_var.set(f"Wybrano plik: {p.name}")

    def _browse_output(self):
        initial_dir  = ""
        initial_file = ""
        if self.input_var.get():
            p = Path(self.input_var.get())
            initial_dir  = str(p.parent)
            initial_file = p.stem + "_raport.txt"

        path = filedialog.asksaveasfilename(
            title="Zapisz raport jako...",
            defaultextension=".txt",
            filetypes=[("Pliki tekstowe", "*.txt"), ("Wszystkie pliki", "*.*")],
            initialdir=initial_dir,
            initialfile=initial_file,
        )
        if path:
            self.output_var.set(path)

    def _run(self):
        input_path = self.input_var.get().strip()
        if not input_path:
            messagebox.showerror("Błąd", "Wybierz plik wejściowy CSV.")
            return
        if not Path(input_path).exists():
            messagebox.showerror("Błąd", f"Plik nie znaleziony:\n{input_path}")
            return

        encoding    = self.enc_var.get()
        output_path = self.output_var.get().strip()

        self.status_var.set("Analizuję...")
        self.update_idletasks()

        try:
            df = analyze(input_path, encoding)
        except Exception as exc:
            messagebox.showerror("Błąd podczas odczytu pliku", str(exc))
            self.status_var.set("Błąd — zobacz komunikat powyżej")
            return

        output_lines: list = []
        try:
            print_report(df, output_lines)
        except Exception as exc:
            messagebox.showerror("Błąd podczas tworzenia raportu", str(exc))
            self.status_var.set("Błąd — zobacz komunikat powyżej")
            return

        # Wyświetlamy w polu tekstowym
        report_text = "\n".join(output_lines)
        self.result_text.configure(state=tk.NORMAL)
        self.result_text.delete("1.0", tk.END)
        self.result_text.insert(tk.END, report_text)
        self.result_text.configure(state=tk.DISABLED)
        self.result_text.see("1.0")

        # Zapis do pliku
        saved_msg = ""
        if output_path:
            try:
                Path(output_path).write_text(report_text, encoding="utf-8")
                saved_msg = f"  •  Zapisano: {Path(output_path).name}"
            except OSError as exc:
                messagebox.showwarning("Nie udało się zapisać pliku", str(exc))

        # Zapamiętujemy wynik i odblokowujemy przycisk Saldeo
        self._df = df
        self._saldeo_btn.configure(
            state=tk.NORMAL if not df.empty else tk.DISABLED
        )

        # Pasek stanu
        if df.empty:
            self.status_var.set("Nie znaleziono płatności od osób fizycznych")
        else:
            clients = df["name"].nunique()
            txns    = len(df)
            total   = df["amount"].sum()
            self.status_var.set(
                f"Gotowe  •  {clients} klientów  •  "
                f"{txns} transakcji  •  "
                f"{total:,.2f} PLN".replace(",", " ").replace(".", ",")
                + saved_msg
            )

    def _open_reconciliation(self):
        input_path = self.input_var.get().strip()
        if not input_path:
            messagebox.showerror("Błąd", "Wybierz plik wejściowy CSV.")
            return
        if not Path(input_path).exists():
            messagebox.showerror("Błąd", f"Plik nie znaleziony:\n{input_path}")
            return

        encoding = self.enc_var.get()

        output_lines: list = []
        try:
            print_reconciliation_report(input_path, encoding, output_lines)
        except Exception as exc:
            messagebox.showerror("Błąd podczas tworzenia raportu", str(exc))
            return

        p = Path(input_path)
        ReconciliationDialog(
            self,
            "\n".join(output_lines),
            default_dir=str(p.parent),
            default_name=p.stem + "_kontrola.txt",
        )

    def _open_saldeo_dialog(self):
        if self._df is None or self._df.empty:
            messagebox.showinfo("Brak danych",
                                "Najpierw uruchom analizę listy operacji.")
            return
        SaldeoDialog(self, self._df)

    def _open_seller_settings(self):
        """Otwiera okno edycji danych sprzedawcy —
        dostępne w każdej chwili, nie tylko przy pierwszym uruchomieniu."""
        SellerSetupDialog(self, first_run=False)

    def _open_help(self):
        """Otwiera okno wbudowanej pomocy użytkownika."""
        HelpDialog(self)


# ─── Punkt wejścia ────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
