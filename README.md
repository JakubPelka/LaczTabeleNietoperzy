# LaczTabeleNietoperzy

Narzędzia do łączenia tabel XLSX i CSV z rejestracjami nietoperzy oraz tworzenia
zestawień i wykresów. Repozytorium zawiera dwa klasyczne algorytmy Tkinter oraz
niezależną aplikację `parallel-graph` do interaktywnego porównywania źródeł.

## Który algorytm wybrać?

| Algorytm | Wspólne pliki Excel | Wykresy zbiorcze | Wykresy noc-po-nocy |
|---|---:|---:|---:|
| `merge_summary_charts.py` | tak | tak | nie |
| `merge_summary_and_nightly_charts.py` | tak | opcjonalnie | opcjonalnie |

Oba algorytmy:

- łączą wybrane pliki XLSX i CSV o zgodnej strukturze;
- tworzą warianty `<nazwa>_NVI.xlsx` i `<nazwa>_ART.xlsx`;
- podsumowują `MANUAL ID` jako `Socialt`, `Födosökande` lub
  `Förbiflygande`;
- obliczają liczbę rejestracji, nocy, rejestracji na noc i wszystkich dźwięków;
- mogą generować wykresy liniowe oraz skumulowane wykresy ART/NVI;
- używają wspólnej skali Y dla porównywalnych wykresów.

Wariant nocny dodatkowo dzieli rekordy na noce terenowe. Rekord z godziną przed
12:00 jest przypisywany do poprzedniej daty, np. `2026-07-18 01:30` należy do
nocy `17/18.07`. Każda noc otrzymuje osobne katalogi wykresów i etykietę zakresu
dat. W GUI można niezależnie włączyć wykresy zbiorcze i noc-po-nocy.

## Uruchomienie offline (także z USB)

Starter nie tworzy środowiska, nie uruchamia `pip` i nie łączy się z internetem.
Korzysta z tego interpretera Python, którym został otwarty. Portable Python na
USB musi więc już zawierać `pandas`, `openpyxl`, `matplotlib` oraz Tkinter.

Windows: otwórz dwukrotnie plik `start.py` albo uruchom go za pomocą Pythona z
USB:

```text
X:\sciezka\do\python.exe start.py
```

Można również zachować dotychczasowy sposób uruchamiania przez
`run_algorithms.bat`. Starter `.bat` jedynie uruchamia `start.py` dostępnym
Pythonem — nie tworzy środowiska i niczego nie pobiera.

Linux/macOS:

```bash
python3 start.py
# albo:
./run_algorithms.sh
```

Starter pokazuje okno wyboru wariantu algorytmu. Jeśli brakuje biblioteki,
wyświetla jej nazwę i kończy pracę bez próby pobierania. `requirements.txt` jest
wyłącznie listą bibliotek potrzebną przy przygotowywaniu kompletnego Pythona USB.

Algorytm można również uruchomić bezpośrednio:

```bash
python algorithms/merge_summary_charts.py
python algorithms/merge_summary_and_nightly_charts.py
```

## Dane wejściowe

Pierwszy arkusz pliku XLSX lub tabela w pliku CSV powinny zawierać:

- `MANUAL ID` — klasyfikację lub kilka klasyfikacji oddzielonych przecinkiem;
- `DATE` lub `Datum` — datę;
- `TIME` lub `Tid` — czas jako wartość Excela, datetime lub `HH:MM[:SS]`.

`SOC` oznacza zachowanie społeczne, `FOD` żerowanie, a pozostałe wpisy są
traktowane jako przelot. `NOISE` jest pomijane w zestawieniu gatunków.

Nie dodawaj do repozytorium rzeczywistych danych lokalizacyjnych ani nagrań.
Przykłady powinny być syntetyczne lub bezpieczne do publicznego udostępnienia.

## Parallel Activity — interaktywne porównanie źródeł

Katalog [`parallel-graph`](parallel-graph/) zawiera odrębną aplikację GUI do
porównywania aktywności z dowolnej liczby plików XLSX i CSV. Program odczytuje
kolumny `DATE`, `TIME` i `MANUAL ID`, rozdziela klasyfikacje zapisane po
przecinku, pomija puste wartości oraz `Noise`, a następnie tworzy samodzielny
interaktywny plik HTML. Dane źródłowe nie są modyfikowane.

### Funkcje widoku HTML

- interfejs, wykres i komunikaty w języku polskim, angielskim lub szwedzkim;
- osobne kolory i symbole dla źródeł oraz pełne etykiety `MANUAL ID`;
- niezależne filtry źródeł i klasyfikacji, działające również z legendą;
- przełączana rozdzielczość czasu: 1, 2, 3, 5, 10, 15, 30 lub 60 minut;
- dynamiczne sumowanie obserwacji w wybranych przedziałach czasu;
- zoom, przesuwanie osi i suwak ograniczania widocznego zakresu;
- zbiorcze dymki dla obserwacji przypadających na ten sam czas;
- wyróżnianie klasyfikacji nakładających się w tym samym punkcie wykresu;
- tabela automatycznie synchronizowana z filtrami, zoomem i zaznaczeniem;
- wybór punktów narzędziami Lasso i Box Select oraz przycisk czyszczenia;
- eksport aktualnego, przefiltrowanego widoku do PNG;
- nakładka ładowania podczas przeliczania filtrów i rozdzielczości czasu;
- działanie gotowego HTML bez serwera i bez połączenia z internetem.

Oprócz `parallel_bat_activity.html` aplikacja zapisuje zagregowane dane CSV oraz
tekstowy raport importu z liczbą odczytanych i pominiętych rekordów.

### Przygotowanie bibliotek i praca offline

Z dostępem do internetu należy uruchomić skrypt odpowiedni dla systemu:

```bash
cd parallel-graph
./parallel.sh  # Linux/macOS
```

```text
cd parallel-graph
parallel.bat  # Windows
```

Oba skrypty pobierają `openpyxl`, `plotly` i ich zależności do lokalnego
katalogu `parallel-graph/vendor`, bez instalowania ich globalnie i bez tworzenia
środowiska `.venv`, po czym uruchamiają aplikację. Uruchomienie `.sh` lub `.bat`
zakłada dostęp do internetu i aktualizuje lokalne pakiety.

Później cały katalog `parallel-graph`, razem z `vendor`, można skopiować na dysk
USB. Trybem w pełni offline jest bezpośrednie uruchomienie `parallel.py`
odpowiednim interpreterem Pythona.

Tkinter musi być składnikiem używanej instalacji Pythona, ponieważ nie jest
pakietem instalowanym przez `pip`. Jeśli lokalnych bibliotek brakuje,
`parallel.py` pokazuje instrukcję ich przygotowania zamiast próbować połączyć
się z internetem.

Pełny opis danych wejściowych, obsługi wykresu i plików wynikowych znajduje się
w [`parallel-graph/README.md`](parallel-graph/README.md).

## Struktura

```text
LaczTabeleNietoperzy/
├── algorithms/
│   ├── merge_summary_charts.py
│   └── merge_summary_and_nightly_charts.py
├── parallel-graph/
│   ├── src/parallel_graph/
│   │   ├── __init__.py
│   │   ├── __main__.py
│   │   ├── aggregation.py
│   │   ├── app.py
│   │   ├── chart.py
│   │   ├── excel_reader.py
│   │   ├── export.py
│   │   ├── models.py
│   │   └── translations.py
│   ├── vendor/              # lokalne pakiety, tworzone przez .sh lub .bat
│   ├── parallel.py          # uruchomienie offline
│   ├── parallel.sh          # przygotowanie online: Linux/macOS
│   ├── parallel.bat         # przygotowanie online: Windows
│   ├── pyproject.toml
│   ├── requirements.txt
│   └── README.md
├── SampleData/
├── requirements.txt
├── start.py
├── run_algorithms.bat
├── run_algorithms.sh
└── README.md
```

Wyniki powstają w katalogu `Results/` wybranym w GUI lub `results/` przy uruchomieniu bezobsługowym. Plików wynikowych nie należy commitować.

## Tryb bezobsługowy (Headless CLI & API)

Algorytm `merge_summary_and_nightly_charts.py` oraz narzędzie `parallel-graph/parallel.py` obsługują uruchamianie w trybie bezobsługowym (headless CLI).

### Przykłady użycia CLI

```bash
python3 algorithms/merge_summary_and_nightly_charts.py \
  --headless \
  --base-dir /path/to/output_directory \
  --base-name sammanstallning_fladdermus \
  --time-mode manual \
  --time-start 21:00 \
  --time-end 04:30 \
  --ymax-mode zoomed \
  --enable-html \
  --input-files file1.csv file2.csv
```

Opcje CLI:
- `--headless`: Uruchomienie bez interfejsu GUI Tkinter.
- `--base-dir`: Katalog wyjściowy dla wyników.
- `--base-name`: Nazwa bazowa plików podsumowania Excel (`_NVI.xlsx` / `_ART.xlsx`).
- `--time-mode`: `manual` (domyślne okno nagrywania, np. `21:00` -> `04:30`) lub `auto` (wyznaczany dynamicznie z danych).
- `--time-start`: Czas rozpoczęcia w formacie `HH:MM` (dwucyfrowy godzinowo, np. `21:00`).
- `--time-end`: Czas zakończenia w formacie `HH:MM` (dwucyfrowy godzinowo, np. `04:30`). Zakresy przechodzące przez północ (`stop < start`) są w fully obsługiwane.
- `--ymax-mode` / `--y-axis-mode`: `fixed` (domyślny wspólny Y-max w całym zbiorze + 10%) lub `zoomed` (Y-max wyliczany osobno z danych wykresu + 10% zapasu, min 1).
- `--enable-html`: Generuje interaktywny plik HTML oraz towarzyszący raport CSV i TXT (`parallel_bat_activity.*`).

### Nowe funkcjonalności (#7, #8, #9)

- **#7 Tryb Osi Y**: Wybór pomiędzy `fixed` (stała globalna skala Y) a `zoomed` (dynamiczna skala osi Y dopasowana do faktycznie przedstawionych danych na danym wykresie + 10% nagłówka).
- **#8 Klasa SOF i kompatybilność**: Rejestracje `SOF` (`socialt - flyg`) oraz `SOC` (`socialt - läte`). Jeśli dataset nie zawiera `SOF`, zachowanie i paleta `SOC` pozostają 100% wstecznie kompatybilne. Przy wykryciu `SOF` w zbiorze, `SOC` przyjmuje nowy domyślny odcień (#9A0B08 NVI / #8D0B82 ART), a `SOF` przejmuje poprzednią barwę `SOC`.
- **#9 Wykres zbiorczy wszystkich gatunków**: Tworzony jako `alla_arter.png` w podkatalogach `stacked_ART/` i `stacked_NVI/` zarówno dla podsumowania zbiorczego, jak i każdej nocy biologicznej. Zachowuje układy słupków per gatunek w 15-minutowych interwałach oraz pionowe etykiety gatunków.

### Gwarancja bezobsługowości offline (Desktop Offline Guarantee)

Aplikacja Tkinter działa w 100% offline bez żadnych zależności sieciowych, zapytań HTTP ani modułów specyficznych dla Perun Works.

### Struktura pakietu wyników (Formalized Result Package)

```text
results/
  combined/
    <base_name>_NVI.xlsx
    <base_name>_ART.xlsx
    parallel_bat_activity.html        (gdy włączono --enable-html)
    parallel_bat_activity_data.csv    (gdy włączono --enable-html)
    parallel_bat_activity_report.txt  (gdy włączono --enable-html)
  <input_stem>/
    summary/
      line/
        alla_arter.png
        <species>.png
      stacked_ART/
        alla_arter.png
        <species>.png
      stacked_NVI/
        alla_arter.png
        <species>.png
    nights/
      <night>/
        line/
          alla_arter.png
          <species>.png
        stacked_ART/
          alla_arter.png
          <species>.png
        stacked_NVI/
          alla_arter.png
          <species>.png
```

Dodatkowa zależność zewnętrzna: `plotly>=5.20` jest wymagana przez moduł Parallel Graph (`requirements.txt`).
