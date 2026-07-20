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
- wyróżnianie różnych klasyfikacji nakładających się w tej samej minucie na tej
  samej wartości osi Y;
- niezależne filtry źródeł i klasyfikacji `MANUAL ID`;
- dynamiczna tabela danych zsynchronizowana z wykresem;
- przełączana agregacja czasu: 1, 2, 3, 5, 10, 15, 30 lub 60 minut;
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

## Przygotowanie i uruchomienie offline (także z USB)

Przy dostępie do internetu uruchom odpowiedni skrypt przygotowawczy:

```bash
./parallel.sh  # Linux/macOS
```

```text
parallel.bat  # Windows
```

Startery pobierają `openpyxl`, `plotly` i wszystkie ich zależności do lokalnego
folderu `vendor`, a następnie uruchamiają program. Nie tworzą `.venv` i nie
instalują pakietów globalnie. Każde kolejne uruchomienie `.sh` lub `.bat`
aktualizuje lokalny zestaw zgodnie z `requirements.txt` i wymaga dostępu do
internetu.

Po skopiowaniu całego katalogu `parallel-graph` (razem z `vendor`) można
uruchamiać aplikację bez internetu bezpośrednio przez `parallel.py`. Starter
zawsze dodaje lokalny folder bibliotek do ścieżki importu.

Windows: otwórz dwukrotnie `parallel.py` albo wskaż wprost interpreter z USB:

```text
X:\sciezka\do\python.exe parallel.py
```

Linux/macOS:


```bash
python3 parallel.py  # offline, po wcześniejszym przygotowaniu vendor
```

Tkinter nie jest pakietem `pip`: musi należeć do używanej instalacji Pythona.
Na Ubuntu/Debian można go doinstalować pakietem `python3-tk`; instalator Pythona
dla Windows powinien zawierać opcjonalny komponent Tcl/Tk.

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
Zakres osi czasu, przesuwania i suwaka jest ograniczony do przedziału od 24 godzin
przed najwcześniejszą obserwacją do 24 godzin po najpóźniejszej obserwacji.

Przycisk **Drukuj / PNG**, **Print / PNG** lub **Skriv ut / PNG** zapisuje
bieżący widok wykresu jako obraz PNG w
podwyższonej rozdzielczości. Eksport uwzględnia aktywne filtry, widoczne serie,
zoom i zakres czasu. Panel checkboxów nie jest umieszczany na obrazie.

Filtry źródeł i gatunków działają łącznie. Można więc pozostawić wszystkie
źródła i jeden gatunek albo jedno źródło i wiele gatunków.

Przełącznik rozdzielczości czasu sumuje punkty w przedziałach rozpoczynających
się o pełnej wielokrotności wybranej wartości. Przykładowo dla 15 minut są to
`01:00–01:14`, `01:15–01:29` itd. Domyślna rozdzielczość to 1 minuta.
Zmiana działa bez ponownego wczytywania plików i aktualizuje wykres, tabelę,
dymki, oznaczenia nakładających się klasyfikacji oraz eksport PNG. Filtry i
aktualny zakres czasu pozostają zachowane.

Podczas zmiany rozdzielczości oraz filtrów źródeł lub gatunków wyświetlana jest
nakładka ładowania. Znika po zakończeniu przeliczenia wykresu i tabeli.

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

Liczebność danej klasyfikacji pokazuje wyłącznie położenie punktu na osi Y.
Przykładowo trzy rejestracje jednego gatunku dają zwykły punkt na Y=3 bez cyfry
w markerze.

Różne klasyfikacje są wykrywane jako nakładające się tylko wtedy, gdy
pochodzą z tego samego źródła, tej samej minuty i mają tę samą liczebność, czyli
leżą w tej samej współrzędnej wykresu. Badge pokazuje liczbę nakładających się
klasyfikacji, a nie liczbę rejestracji. Zarówno `1× Nyctalus noctula + 1×
Vespertilio murinus` na Y=1, jak i `2× Nyctalus noctula + 2× Eptesicus nilssonii`
na Y=2 dają badge `2`. Punkty o różnych wartościach Y pozostają oddzielne.
Wyróżnienie jest dynamicznie przeliczane po zmianie filtrów lub ukryciu serii w
legendzie. Wykres ma zwiększoną wysokość, aby większe symbole pozostawały
czytelne również przy wielu równoległych seriach.

## Testy

Uruchom Pythonem zawierającym wymagane biblioteki:

```bash
PYTHONPATH="vendor:src" python3 -m unittest discover -s tests -v
```

W systemie Windows, używając Pythona z USB:

```text
set PYTHONPATH=vendor;src
X:\sciezka\do\python.exe -m unittest discover -s tests -v
```

## Struktura projektu

```text
parallel-graph/
├── vendor/          # lokalne biblioteki utworzone przez .sh lub .bat
├── parallel.sh
├── parallel.bat
├── parallel.py
├── pyproject.toml
├── requirements.txt
├── README.md
├── sampleData/
├── src/parallel_graph/
└── tests/
```
