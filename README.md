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

## Parallel Bat Graph

Katalog [`parallel-graph`](parallel-graph/) zawiera odrębną, wielojęzyczną
aplikację GUI. Porównuje dowolną liczbę plików, agreguje wpisy do minut i tworzy
samodzielny HTML z filtrami, tabelą, zoomem, suwakiem czasu oraz eksportem PNG,
CSV i raportu importu. Szczegóły znajdują się w
[`parallel-graph/README.md`](parallel-graph/README.md).

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
