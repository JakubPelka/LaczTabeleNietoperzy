# LaczTabeleNietoperzy

Narzędzia do łączenia tabel Excel z rejestracjami nietoperzy oraz tworzenia
zestawień i wykresów. Repozytorium zawiera dwa klasyczne algorytmy Tkinter oraz
niezależną aplikację `parallel-graph` do interaktywnego porównywania źródeł.

## Który algorytm wybrać?

| Algorytm | Wspólne pliki Excel | Wykresy zbiorcze | Wykresy noc-po-nocy |
|---|---:|---:|---:|
| `merge_summary_charts.py` | tak | tak | nie |
| `merge_summary_and_nightly_charts.py` | tak | opcjonalnie | opcjonalnie |

Oba algorytmy:

- łączą wybrane skoroszyty o zgodnej strukturze;
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

Pierwszy arkusz każdego pliku powinien zawierać:

- `MANUAL ID` — klasyfikację lub kilka klasyfikacji oddzielonych przecinkiem;
- `DATE` lub `Datum` — datę;
- `TIME` lub `Tid` — czas jako wartość Excela, datetime lub `HH:MM[:SS]`.

`SOC` oznacza zachowanie społeczne, `FOD` żerowanie, a pozostałe wpisy są
traktowane jako przelot. `NOISE` jest pomijane w zestawieniu gatunków.

Nie dodawaj do repozytorium rzeczywistych danych lokalizacyjnych ani nagrań.
Przykłady powinny być syntetyczne lub bezpieczne do publicznego udostępnienia.

## Parallel Activity — interaktywne porównanie źródeł

Katalog [`parallel-graph`](parallel-graph/) zawiera odrębną aplikację GUI do
porównywania aktywności z dowolnej liczby skoroszytów Excel. Program odczytuje
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

Przy pierwszym uruchomieniu z dostępem do internetu należy wykonać:

```bash
cd parallel-graph
./parallel.sh
```

Skrypt pobiera `openpyxl`, `plotly` i ich zależności do lokalnego katalogu
`parallel-graph/vendor`, bez instalowania ich globalnie i bez tworzenia
środowiska `.venv`. Później cały katalog `parallel-graph`, razem z `vendor`,
można skopiować na dysk USB i uruchamiać offline przez `parallel.py` lub
`parallel.bat`.

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
├── SampleData/
├── requirements.txt
├── start.py
├── run_algorithms.bat
├── run_algorithms.sh
└── README.md
```

Wyniki powstają w katalogu `Results/` wybranym w GUI. Plików wynikowych nie
należy commitować.
