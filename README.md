# README.md

# README – sammanställning & diagram för fladdermusdata

> **Script:** `Sla_ihop_exceL_fladdermus_tabell_och_skapa_diagram.py`
>
> **Syfte:** Sammanställa flera Excel-filer till en översikt (två versioner, \*\_NVI och *\_ART*), samt (valfritt) generera linje- och stapeldiagram. All funktionalitet presenteras via dialogrutor (Tkinter).

---

## Nyheter (senaste uppdateringar)

* **Mapp `Results/`** skapas automatiskt på den plats du väljer för sammanställningen. **Båda Excel-filerna** sparas där: `*_NVI.xlsx` och `*_ART.xlsx`.
* **Två identiska sammanställningsfiler**:

  * **NVI** – färger i tabellen matchar stapeldiagram *NVI (röd/gul/grå)*.
  * **ART** – färger i tabellen matchar stapeldiagram *ART (rosa/ljuslila/grå)*.
  * **Förbiflygande** har **samma ljusgrå färg i tabellerna** för NVI och ART (`#E7E6E6`).
* **Global Y-skala** för **alla** diagram (linje + stapel), beräknad över **samtliga valda indatafiler** och det valda tidsintervallet. Gör att diagram blir jämförbara mellan boxar.
* **Svenska kommentarer/texter** i koden; etiketter/titlar använder **svenskt + latinskt namn** (om svenskt namn finns).

---

## Förutsättningar

* **Python**: 3.9–3.12 (rekommenderat).
* **Bibliotek**: se `requirements.txt` (t.ex. `pandas`, `openpyxl`, `matplotlib`, `tk`).

Installation (exempel):

```bash
python -m venv .venv
. .venv/Scripts/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

---

## Inmatning (krav på Excel)

Varje indatafil (första bladet) förväntas innehålla:

* **`MANUAL ID`** – textfält med art + ev. beteendetypskod, separerad med kommatecken vid flera värden.

  * Tolkning av beteendetyper i koden:

    * `SOC` → **Socialt**
    * `FOD` → **Födosökande**
    * annars → **Förbiflygande**
  * Poster med `NOISE` ignoreras.
* **Datum**: kolumn **`DATE`** eller **`Datum`**.
* **Tid**: kolumn **`TIME`** eller **`Tid`** (kan vara Excel-tid, `HH:MM(:SS)` eller datetime).

> **Obs!** Tider **före kl 12:00** räknas till **föregående fältnatt** när antalet nätter beräknas.

---

## Användning (steg-för-steg)

1. **Välj indatafiler** (`.xlsx`). Du kan markera flera.
2. **Välj plats och basnamn** för sammanställningen.

   * Scriptet skapar **`Results/`** och sparar **båda** filerna där:

     * `<basnamn>_NVI.xlsx`
     * `<basnamn>_ART.xlsx`
3. **(Valfritt) Generera diagram**:

   * Förvald målmapp i dialogen är **`Results/`** (om den finns), annars **mappen för första indatafilen**.
   * Välj **automatisk** X-axel (från data) **eller** ange **eget intervall** (t.ex. `22:15` till `02:45`).
   * Diagram sparas i en undermapp `diagramer/` inne i vald målmapp. För varje indatafil skapas:

     * `<filnamn>_linjediagram/`
     * `<filnamn>_stapeldiagram_ART/`
     * `<filnamn>_stapeldiagram_NVI/`

Körning (exempel):

```bash
python Sla_ihop_exceL_fladdermus_tabell_och_skapa_diagram.py
```

---

## Vad innehåller sammanställningen (Översikt)

Arket **`Översikt`** byggs med följande:

* Kolumner: `Art`, `Beteendetyper`, följt av **en kolumn per indatafil** (namngivna efter fil/blad).
* **Radsortering i artblock**: `Socialt`, `Födosökande`, `Förbiflygande`.
* **Artetikett**: Svenskt namn + radbrytning + latinskt namn (om svenskt namn finns i ordboken). För grupper som **Nyctaloid/Chiroptera** visas endast originalnamnet.
* **Summeringsrader** (i slutet, i ordning):

  1. **Fladdermusregistreringar** – antal poster efter tolkning (exkl. `NOISE`).
  2. **Antal nätter** – unika fältnätter beräknade från datum + tid (<12 → föregående natt).
  3. **Antal registreringar / natt** – **formel i Excel**: `Fladdermusregistreringar / Antal nätter` (med `IFERROR` och en decimal).
  4. **Total antal ljud** – totala rader inkl. `NOISE` (för kontroll).

### Färger i tabellen

* **NVI.xlsx**: Socialt (**röd** `#FF0000`), Födosökande (**gul** `#FFC000`), Förbiflygande (**ljusgrå tabell** `#E7E6E6`).
* **ART.xlsx**: Socialt (**rosa** `#EB09D8`), Födosökande (**ljuslila** `#D98FD3`), Förbiflygande (**ljusgrå tabell** `#E7E6E6`).

> **Notera:** Förbiflygande har samma ljusgrå i båda tabellversionerna. I **diagram** kan Förbiflygande vara **grå** (`#A9A9A9`) för tydligare kontrast.

---

## Diagram

* **Global Y-skala**: Ett gemensamt max (med liten marginal) beräknas över alla valda filer **inom det aktuella tidsintervallet**. Gäller **både** linje- och stapeldiagram.
* **X-axel**: 15-minutersintervall, automatiskt från datan eller enligt manuellt angivet intervall.
* **Titlar (per art)**: `Svenskt namn (Latinskt), antal observerade beteenden: NN` (om svenskt namn finns, annars bara latinskt).
* **Färger (stapeldiagram)**:

  * **NVI**: Socialt `#FF0000`, Födosökande `#FFC000`, Förbiflygande `#A9A9A9`.
  * **ART**: Socialt `#EB09D8`, Födosökande `#D98FD3`, Förbiflygande `#ABAAA9`.

Utdataträd (exempel):

```
Results/
├─ <basnamn>_NVI.xlsx
├─ <basnamn>_ART.xlsx
└─ diagramer/
   ├─ <fil1>_linjediagram/
   │  ├─ alla_arter.png
   │  └─ <Art 1>.png, <Art 2>.png, ...
   ├─ <fil1>_stapeldiagram_ART/
   │  └─ <Art X>.png, ...
   ├─ <fil1>_stapeldiagram_NVI/
   │  └─ <Art X>.png, ...
   └─ <fil2>_... (analogt)
```

---

## Begränsningar & tips

* Om **`MANUAL ID`** saknas i en fil skapas ändå bladkopian, men raden ger inga artvärden.
* Om **Tid** saknas i alla rader för en fil kan diagram för den filen hoppas över (inget giltigt intervall).
* **Manuellt intervall** måste anges som `HH:MM` (24h). Scriptet rundar till närmaste **15-minutare**.
* **Särskrivningar/konstiga värden** i `MANUAL ID` kan ge oväntad tolkning; se över källdata.

---

## Felsökning

* **Tomma diagram eller felaktig X-axel**: kontrollera `TIME/Tid`-kolumnen och formatet.
* **Division by zero i Excel**: om `Antal nätter` blir tomt/0; formeln använder `IFERROR` och visar då tomt värde.
* **Behörighetsfel vid skrivning**: stäng öppna Excel-filer innan körning.

---

## Licens & kontakt

Intern projektskript för rapportering. Justera efter behov. Frågor/förslag: öppna ett ärende i repo eller kontakta projektgruppen.


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
