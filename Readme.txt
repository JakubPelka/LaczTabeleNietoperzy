Skrypt Python do łączenia wyników obserwacji nietoperzy zapisanych w wielu plikach Excel w jedną zbiorczą tabelę.

Funkcjonalność:
- Użytkownik wybiera jeden lub wiele plików Excel (o tej samej strukturze).
- Dane z każdego pliku są kopiowane do osobnego arkusza w pliku wynikowym. Nazwa arkusza odpowiada nazwie pliku wejściowego.
- W arkuszu "Översikt" tworzona jest zbiorcza tabela, zawierająca:
  * listę wszystkich gatunków nietoperzy występujących w danych,
  * rozbicie obserwacji na typy zachowań (Förbiflygande, Socialt, Fodosökande),
  * liczby obserwacji dla poszczególnych plików,
  * sumaryczne zestawienie ("Fladdermusregistreringar") oraz całkowitą liczbę nagrań ("Total antal ljud").
- Kolumna z nazwą gatunku zawiera zarówno nazwę szwedzką jak i łacińską.
- Gatunki są posortowane alfabetycznie według nazwy łacińskiej, z wyjątkiem kategorii specjalnych ("Nyctaloid" i "Chiroptera"), które zawsze umieszczane są na końcu.
- Tabela jest estetycznie sformatowana:
  * kolory tła wierszy odpowiadają typom zachowań,
  * scalone komórki dla gatunków,
  * pogrubione linie oddzielające bloki gatunków i kolumny plików,
  * wiersze sum wyróżnione pogrubioną czcionką.

Wymagania:
- Python 3.10+,
- biblioteki: pandas, openpyxl, tkinter (standardowo dostępny w CPython na Windows).

Instalacja bibliotek:
    python -m pip install --user pandas openpyxl

Uruchomienie:
    python Sla_ihop_exceL_fladdermus_tabell.py

Po zakończeniu działania skrypt zapisuje nowy plik Excel ze wszystkimi arkuszami i automatycznie go otwiera.
