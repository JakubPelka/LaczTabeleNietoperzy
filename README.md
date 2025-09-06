# README – sammanställning & diagram för fladdermusdata (uppdaterad)

> **Script:** `Sla_ihop_exceL_fladdermus_tabell_och_skapa_diagram.py`
>
> **Syfte:** Sammanställa flera Excel-filer till en översikt (två versioner: *\_NVI* och *\_ART*) samt – valfritt – generera linje- och stapeldiagram. All interaktion sker nu i **ett GUI-fönster** (Tkinter).

---

## Nyheter (senaste uppdateringar)

* **Nytt GUI** för hela flödet: val av indatafiler, utdata, generering av diagram, tidsläge (**auto/manuellt**) och **färgval via palett**.
* **Färgpalett**: obok każdego pola koloru dostępny jest podgląd oraz przycisk **„Välj…”** (Tkinter `colorchooser`). Możesz też ręcznie wpisać HEX.
* **Två identiska sammanställningsfiler** sparas i `Results/`:

  * `<basnamn>_NVI.xlsx` – tabellfärger matchar NVI-staplar (röd/gul/ljusgrå i tabellen).
  * `<basnamn>_ART.xlsx` – tabellfärger matchar ART-staplar (rosa/lila/ljusgrå i tabellen).
  * **Förbiflygande** har **samma ljusgrå tabellfärg** i NVI och ART (`#E7E6E6`).
* **Mappen `Results/`** skapas automatiskt i vald utdata-katalog; båda Excel-filer sparas där.
* **Diagram**: om du väljer att generera, domyślna basmapp to `Results/diagramer` (lub własna, wskazana w GUI).
* **Global Y-skala** (linje + stapel) beräknas över **alla valda indatafiler** inom valt tidsintervall ⇒ lätt att jämföra mellan boxar.
* **Svenska etiketter/kommentarer** i koden. Titel: *Svenskt namn (Latinskt), antal observerade beteenden: NN* (fallback: bara latinskt).
* **Sammanställning**: dodane wiersze **Antal nätter** i **Antal registreringar / natt** (to drugie liczone formułą w Excelu).

---

## Förutsättningar

* **Python**: 3.9–3.12.
* **Bibliotek**: se `requirements.txt` (`pandas`, `openpyxl`, `matplotlib`). Tkinter/`colorchooser` är częścią standardowej biblioteki Pythona.

Installation (exempel):

```bash
python -m venv .venv
# Windows
. .venv/Scripts/activate
# macOS/Linux
# source .venv/bin/activate
pip install -r requirements.txt
```

---

## Inmatning (krav på Excel)

Första bladet i varje indatafil ska minst innehålla:

* **`MANUAL ID`** – lista art(er) + ev. typkod, separerade komma.

  * Tolkning av typkoder:

    * `SOC` → **Socialt**
    * `FOD` → **Födosökande**
    * annars → **Förbiflygande**
  * `NOISE` ignoreras.
* **Datum**: kolumn **`DATE`** eller **`Datum`**.
* **Tid**: kolumn **`TIME`** eller **`Tid`** (Excel-tid, `HH:MM(:SS)` eller datetime funkar).

> **Obs:** tider **före 12:00** tillhör **föregående fältnatt** vid beräkning av *Antal nätter*.

---

## Användning (GUI – steg-för-steg)

1. **Starta** skrypt: `python Sla_ihop_exceL_fladdermus_tabell_och_skapa_diagram.py`.
2. I GUI:

   * **Indatafiler**: *Lägg till filer…* (możesz dodać wiele `*.xlsx`).
   * **Utdatakatalog & basnamn**: wskaż folder, do którego zostanie dodany podfolder **`Results/`**; wpisz bazową nazwę pliku.
   * **Diagram (valfritt)**:

     * Zaznacz **Generera diagram** albo odznacz, jeśli chcesz tylko Excela.
     * **X-axel**: *automatisk* (na podstawie danych) lub *eget intervall* – wpisz `HH:MM`–`HH:MM`. Skrypt zaokrągla do 15 min.
     * **Bas-mapp för diagram** (valfritt): jeśli pusta, użyjemy `Results/diagramer`.
   * **Färger**:

     * NVI: Socialt / Födosökande / Förbiflygande.
     * ART: Socialt / Födosökande / Förbiflygande.
     * Wpisz HEX (np. `#FF0000`) **lub** użyj **Välj…** i palety; podgląd aktualizuje się na bieżąco.
3. Kliknij **Starta**. Skrypt utworzy `Results/` i zapisze `<basnamn>_NVI.xlsx` oraz `<basnamn>_ART.xlsx`. Jeśli wybrano diagramy, powstaną podkatalogi `diagramer/…`.

---

## Co generuje się w Excelu (Översikt)

* Kolumny: `Art`, `Beteendetyper`, następnie po jednej kolumnie dla **każdego** źródłowego arkusza/plikun.
* **Kolejność typów** w blokach: `Socialt`, `Födosökande`, `Förbiflygande`.
* **Etykieta artu**: *svenski*, nowa linia, *łaciński* (jeśli znany). Dla grup `Nyctaloid`/`Chiroptera` – bez prefiksu.
* **Podsumowania** (na dole, w tej kolejności):

  1. **Fladdermusregistreringar** – suma (bez `NOISE`).
  2. **Antal nätter** – unikalne noce z pola `DATE/Datum` + `TIME/Tid` (<12 → poprzednia noc).
  3. **Antal registreringar / natt** – formuła Excel: `Fladdermusregistreringar / Antal nätter` (z `IFERROR`, 1 miejsce po przecinku).
  4. **Total antal ljud** – całkowita liczba wierszy (łącznie z `NOISE`).

### Färger i tabellen

* **NVI.xlsx**: Socialt `#FF0000`, Födosökande `#FFC000`, Förbiflygande **tabell** `#E7E6E6`.
* **ART.xlsx**: Socialt `#EB09D8`, Födosökande `#D98FD3`, Förbiflygande **tabell** `#E7E6E6`.

> W **diagramach** kolor `Förbiflygande` jest szary dla kontrastu: NVI `#A9A9A9`, ART `#ABAAA9`.

---

## Diagram

* **Global Y-skala** (wspólne maksimum z marginesem) liczona dla **wszystkich wybranych plików** w wybranym przedziale czasowym; dotyczy **linii i słupków**.
* **X-axel**: przedział 15-minutowy – automatycznie z danych lub z podanego zakresu.
* **Tytuły** (per art): `Svenskt namn (Latinskt), antal observerade beteenden: NN` (jeśli brak nazwy szwedzkiej – tylko łacińska).

Przykładowa struktura wyjścia:

```
Results/
├─ <basnamn>_NVI.xlsx
├─ <basnamn>_ART.xlsx
└─ diagramer/
   ├─ <fil1>_linjediagram/
   │  ├─ alla_arter.png
   │  └─ <Art 1>.png, <Art 2>.png, …
   ├─ <fil1>_stapeldiagram_ART/
   │  └─ <Art X>.png, …
   ├─ <fil1>_stapeldiagram_NVI/
   │  └─ <Art X>.png, …
   └─ <fil2>_… (analogt)
```

---

## Tips & ograniczenia

* Brak `MANUAL ID` → arkusz źródłowy nadal kopiowany, ale bez wartości „artowych”.
* Braki w kolumnie czasu mogą uniemożliwić generowanie diagramu dla danej próbki.
* Własny zakres czasu musi być w formacie `HH:MM` (24h); skrypt zaokrągla do pełnych 15 min.
* Nietypowe wpisy w `MANUAL ID` mogą dawać nieoczekiwane wyniki (sprawdź dane surowe).

---

## Felsökning

* **Puste wykresy / X-axel niepoprawna**: sprawdź `TIME/Tid` i formatowanie.
* **Puste „Antal registreringar / natt”**: jeśli `Antal nätter` = 0/brak, formuła zwróci pustą komórkę (działa `IFERROR`).
* **Błąd zapisu**: zamknij otwarte pliki Excel przed uruchomieniem skryptu.

---

## Licens & kontakt

Skrypt projektowy do raportowania. Dostosuj w razie potrzeby. Zgłoszenia/uwagi: poprzez issues w repo lub kontakt z zespołem projektowym.


## Licens & kontakt

Intern projektskript för rapportering. Justera efter behov. Frågor/förslag: öppna ett ärende i repo eller kontakta Paulina Wietrzyk-Pelka.


## Förslag på `.gitignore`

```
diagramer/
*_linjediagram/
*_stapeldiagram_*/
sammanstallning_fladdermus.xlsx
.venv/
__pycache__/
```

---

# requirements.txt

```
pandas>=2.0
openpyxl>=3.1
matplotlib>=3.7
```
