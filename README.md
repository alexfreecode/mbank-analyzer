# mBank Analyzer

Proste narzędzie dla osób prowadzących jednoosobową działalność (terapeutów,
korepetytorów, freelancerów), które co miesiąc muszą przejrzeć wpłaty od
klientów na koncie mBank i wystawić za nie faktury w Saldeo Smart.
Program automatyzuje tę żmudną, ręczną pracę.

> Projekt nieoficjalny — nie jest powiązany z mBank S.A. ani Saldeo Smart.
> Program jedynie odczytuje pliki CSV/Excel eksportowane ręcznie przez
> użytkownika z tych serwisów.

![mBank Analyzer — okno główne](docs/screenshot.png)

## Co program potrafi

- **Wczytuje listę operacji z mBanku** (CSV z sekcji „Historia” → „Eksportuj
  listę”) i automatycznie rozpoznaje wpłaty od osób prywatnych — odróżniając
  je od przelewów własnych, opłat bankowych i płatności od firm.
- **Tworzy czytelny raport zbiorczy** — listę klientów z datami i kwotami
  wpłat, sumą miesięczną każdego klienta i tabelą zbiorczą. Raport można
  obejrzeć w programie albo zapisać do pliku tekstowego.
- **Kontrola kompletności** — osobny raport rozbijający WSZYSTKIE operacje
  wyciągu na trzy kategorie (Klienci / Pozostałe wpływy / Wydatki) z sumami
  kontrolnymi, dzięki któremu łatwo sprawdzisz, że nic nie zostało pominięte
  przy przygotowywaniu dokumentów dla księgowości.
- **Generuje plik importu faktur do Saldeo Smart** — gotowy Excel ze
  wszystkimi danymi (nabywca, kwoty, daty, forma płatności), do wczytania
  w Saldeo jednym kliknięciem. Można wybrać, których klientów uwzględnić.
- **Chroni przed dublowaniem kontrahentów w Saldeo** — porównuje klientów
  z wyciągu z wyeksportowaną bazą kontrahentów Saldeo i ostrzega, gdy
  nazwisko jest zapisane nieco inaczej (inna kolejność imienia i nazwiska,
  literówka), zanim import utworzy zdublowaną kartę.
- **Zapamiętuje dane sprzedawcy** — imię/nazwę, numer konta i nazwę usługi
  wpisujesz raz; można je zmienić w dowolnej chwili.
- **Wbudowana pomoc** — instrukcja krok po kroku dostępna z poziomu programu
  (menu „Pomoc”).

## Prywatność

Program działa **w 100% lokalnie** — żadne dane (Twoje ani Twoich klientów)
nie są nigdzie wysyłane. Nie ma żadnej komunikacji z internetem. Kod źródłowy
jest otwarty — każdy może to zweryfikować.

## Instalacja

1. Pobierz `mBank Analyzer Setup.exe` z sekcji
   [Releases](../../releases/latest).
2. Uruchom instalator — instalacja nie wymaga uprawnień administratora.
3. Przy pierwszym uruchomieniu program poprosi o dane sprzedawcy
   (potrzebne tylko do generowania faktur).

> **Uwaga:** Windows SmartScreen może wyświetlić ostrzeżenie przy pierwszym
> uruchomieniu (program nie ma płatnego podpisu cyfrowego). Kliknij
> „Więcej informacji” → „Uruchom mimo to”. Kod źródłowy jest otwarty —
> możesz sprawdzić, co dokładnie robi program.

## Jak używać

1. W serwisie mBank: **Historia** → **Eksportuj listę** → format **CSV**.
2. W programie: wskaż pobrany plik i kliknij **„Uruchom analizę”**.
3. Obejrzyj raport, w razie potrzeby zapisz go do pliku.
4. (Opcjonalnie) **„Kontrola kompletności”** — sprawdź, czy wszystkie
   operacje wyciągu zostały uwzględnione.
5. (Opcjonalnie) **„Faktury Saldeo...”** — wygeneruj plik Excel i zaimportuj
   go w Saldeo Smart.

Szczegółowa instrukcja — przycisk **„Pomoc”** w programie.

## Budowanie ze źródeł

Wymagania: Python 3.10+, [Inno Setup 6](https://jrsoftware.org/isinfo.php)
(tylko dla instalatora).

```powershell
pip install -r src/requirements.txt
python src/app.py          # uruchomienie z kodu źródłowego
./build_all.ps1            # zbudowanie .exe + instalatora (dist/)
```

## Jak to powstało

Ten program to przykład tzw. „vibe codingu”: pomysł, potrzebę i wszystkie
wymagania sformułował człowiek — osoba, która co miesiąc zmagała się
z dokładnie tym samym żmudnym zadaniem co Ty. Natomiast całą resztę —
projekt, kod, testy, poprawki, instalator i tę instrukcję — od początku
do końca napisała sztuczna inteligencja (Claude, Anthropic), prowadzona
kolejnymi, jasno postawionymi zadaniami.

---

## English summary

**mBank Analyzer** is a small Windows desktop app for Polish sole traders
(therapists, tutors, freelancers) who receive client payments into an mBank
account and invoice them monthly via Saldeo Smart.

It parses the bank statement CSV exported from mBank online banking,
automatically identifies incoming payments from individual clients, produces
a per-client monthly report, a completeness-control report (all statement
transactions split into clients / other income / expenses, with control
sums), and generates a ready-to-import invoice Excel file for Saldeo Smart —
including a duplicate-contractor check against the Saldeo contractor export.

Everything runs **100% locally** — no data ever leaves the user's computer.
Built with Python (tkinter, pandas, openpyxl), packaged with PyInstaller
and Inno Setup. The UI and documentation are in Polish, since the tool is
only useful to mBank + Saldeo users.
