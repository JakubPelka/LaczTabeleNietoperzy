# Parallel Bat Graph

Niezależna aplikacja do porównywania rejestracji nietoperzy z dowolnej liczby
arkuszy Excel. Program agreguje wpisy z kolumny `MANUAL ID` do pełnych minut i
tworzy samodzielny, interaktywny wykres HTML.

Projekt nie modyfikuje plików wejściowych ani istniejącego narzędzia
`LaczTabeleNietoperzy`.

## Najważniejsze funkcje

- dowolna liczba źródeł `.xlsx`;
- interfejs i wykres w języku polskim, angielskim lub szwedzkim;
- edytowalne nazwy źródeł;
- agregacja `minuta × źródło × MANUAL ID`;
- zachowanie pełnych klasyfikacji, w tym `FOD` i `SOC`;
- rozdzielanie identyfikacji wielogatunkowych zapisanych po przecinku na osobne
  wystąpienia w tym samym punkcie czasu;
- pomijanie pustych klasyfikacji oraz `Noise`;
- osobny kolor i symbol dla każdego źródła;
- liczba wewnątrz powiększonego markera, gdy w minucie występuje więcej niż
  jedna rejestracja danej klasyfikacji;
- niezależne filtry źródeł i klasyfikacji `MANUAL ID`;
- dynamiczna tabela danych zsynchronizowana z wykresem;
- zoom, przesuwanie, wybór zakresu i suwak czasu;
- przycisk eksportu PNG zapisujący aktualnie widoczny stan wykresu;
- samodzielny HTML działający bez serwera i bez internetu;
- dodatkowy CSV z zagregowanymi danymi oraz raport importu.

## Wymagania danych

Program wyszukuje w skoroszycie arkusz i wiersz nagłówkowy zawierający:

- `DATE` — data rejestracji;
- `TIME` — czas rejestracji;
- `MANUAL ID` — końcowa klasyfikacja gatunku lub zachowania.

Nagłówki mogą znajdować się w jednym z pierwszych 50 wierszy. Pozostałe kolumny
nie wpływają na analizę. Program obsługuje daty i czasy zapisane jako wartości
Excela oraz najczęściej spotykane wartości tekstowe.

## Uruchomienie w systemie Linux

Nadaj starterowi prawo wykonania (w repozytorium jest już ustawione), a następnie
uruchom:

```bash
./parallel.sh
```

Starter automatycznie utworzy `.venv` i zainstaluje zależności przy pierwszym
uruchomieniu. Jeśli brakuje Tkinter, na Ubuntu/Debian zainstaluj go poleceniem:

```bash
sudo apt install python3-tk python3-venv
```

## Uruchomienie w systemie Windows

Zainstaluj Python 3.10 lub nowszy z [python.org](https://www.python.org/), razem
z opcjonalnym komponentem Tcl/Tk. Następnie kliknij dwukrotnie `parallel.bat`.
Starter sam utworzy `.venv` i zainstaluje potrzebne biblioteki.

## Obsługa

1. Kliknij **Dodaj pliki Excel…** i wybierz wszystkie porównywane pliki.
2. W prawym górnym rogu wybierz `Polski`, `English` albo `Svenska`.
3. W razie potrzeby zmień nazwy źródeł w lewej kolumnie listy.
4. Wybierz folder docelowy.
5. Kliknij **Generuj wykres**.
6. Otwórz utworzony HTML w przeglądarce.

Zmiana języka natychmiast aktualizuje interfejs Tkinter. Język wybrany w chwili
generowania określa teksty w HTML i raporcie: tytuł, osie, filtry, przyciski,
legendę oraz komunikaty. Nazwy źródeł i wartości `MANUAL ID` nie są tłumaczone.

Liczba plików nie jest ograniczona. Style źródeł są przydzielane automatycznie;
po wykorzystaniu palety kolorów i symboli ich kombinacje zaczynają się powtarzać.

## Pliki wynikowe

- `parallel_bat_activity.html` — interaktywny wykres;
- `parallel_bat_activity_data.csv` — zagregowane dane;
- `parallel_bat_activity_report.txt` — podsumowanie importu i pominiętych wpisów.

CSV jest zapisany jako UTF-8 z BOM, aby szwedzkie i polskie znaki były poprawnie
rozpoznawane przez Excel.

## Interpretacja wykresu

Każdy punkt oznacza liczbę rejestracji jednej wartości `MANUAL ID`, w jednym
źródle i w jednej minucie. Kolor oraz kształt oznaczają źródło. Pełna etykieta
klasyfikacji, np. `Nyctalus noctula FOD`, jest widoczna w filtrze, legendzie i po
najechaniu na punkt.

Jeżeli komórka zawiera kilka identyfikacji oddzielonych przecinkiem, np.
`Myotis daubentonii, Pipistrellus pygmaeus`, każda z nich jest liczona jako osobne
wystąpienie z tym samym czasem i źródłem.

Dolny suwak wykresu służy wyłącznie do zawężania i przesuwania widocznego zakresu
czasu. Nie zmienia danych ani filtrów źródeł i gatunków.

Przycisk **Drukuj / PNG**, **Print / PNG** lub **Skriv ut / PNG** zapisuje
bieżący widok wykresu jako obraz PNG w
podwyższonej rozdzielczości. Eksport uwzględnia aktywne filtry, widoczne serie,
zoom i zakres czasu. Panel checkboxów nie jest umieszczany na obrazie.

Filtry źródeł i gatunków działają łącznie. Można więc pozostawić wszystkie
źródła i jeden gatunek albo jedno źródło i wiele gatunków.

Pod wykresem znajduje się tabela z kolumnami: minuta, źródło, `MANUAL ID` i
liczba. Pokazuje wyłącznie dane widoczne w aktualnym zakresie wykresu. Aktualizuje
się automatycznie po zmianie checkboxów, ukryciu serii w legendzie, użyciu zoomu
albo przesunięciu dolnego suwaka czasu. Nagłówek pozostaje widoczny podczas
przewijania dłuższej tabeli.

Narzędzia **Lasso Select** i **Box Select** dodatkowo ograniczają tabelę do
zaznaczonych markerów. Informacja o aktywnym zaznaczeniu pojawia się obok liczby
wierszy. Wyczyszczenie zaznaczenia na wykresie przywraca tabelę dla całego
aktualnie widocznego zakresu.

Pasek narzędzi wykresu zawiera dodatkowy przycisk **Wyczyść zaznaczenie** /
**Clear selection** / **Rensa markering** oznaczony ikoną `X`. Usuwa on
zaznaczenie Lasso lub Box Select bez zmiany zoomu i zakresu czasu. Przycisk jest
umieszczony w osobnej grupie tego samego wiersza co standardowe narzędzia
Plotly. Czyszczenie usuwa jednocześnie obrys zaznaczenia, stan zaznaczonych
punktów i filtr tabeli.

Po wskazaniu minuty pojawia się zbiorczy dymek ze wszystkimi widocznymi
rejestracjami z tej minuty. Dzięki temu gatunki zarejestrowane równocześnie przez
ten sam box są widoczne nawet wtedy, gdy ich markery mają tę samą wartość i
nakładają się na wykresie. Dymek respektuje aktywne filtry źródeł i gatunków.

Jeśli w tym samym źródle i minucie wystąpi łącznie więcej niż jedna rejestracja,
jeden reprezentatywny marker jest większy i zawiera ich łączną liczbę. Dotyczy
to zarówno kilku rejestracji tego samego gatunku, jak i różnych gatunków, np.
`1× Nyctalus noctula + 1× Vespertilio murinus` daje cyfrę `2`. Liczba jest
dynamicznie przeliczana po zmianie filtrów lub ukryciu serii w legendzie, więc
obejmuje wyłącznie aktualnie widoczne dane. Przy pojedynczym wystąpieniu symbol
pozostaje bez cyfry. Wykres ma zwiększoną wysokość, aby większe symbole
pozostawały czytelne również przy wielu równoległych seriach.

## Testy

Po utworzeniu środowiska uruchom:

```bash
.venv/bin/python -m unittest discover -s tests -v
```

W systemie Windows:

```bat
.venv\Scripts\python.exe -m unittest discover -s tests -v
```

## Struktura projektu

```text
parallel-graph/
├── parallel.sh
├── parallel.bat
├── pyproject.toml
├── requirements.txt
├── README.md
├── sampleData/
├── src/parallel_graph/
└── tests/
```
