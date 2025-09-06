# README.md

# Sammanställning och diagram för fladdermusdata (Excel)

Detta repo innehåller ett skript som:

1. Slår ihop valda `.xlsx`-filer till ett översiktsblad i Excel.
2. Lägger till summeringar:

   * **Fladdermusregistreringar** (utan NOISE)
   * **Antal nätter** (kan justeras manuellt)
   * **Antal registreringar / natt** (Excel-formel, uppdateras när ”Antal nätter” ändras)
   * **Total antal ljud** (inkl. NOISE)
3. *(Valfritt)* Genererar linje- och stapeldiagram per fil.

## Filer

* **Huvudskript:** `Sla_ihop_exceL_fladdermus_tabell_och_skapa_diagram.py`
* **Utdata (Excel):** `sammanstallning_fladdermus.xlsx`
* **Utdata (diagram):**

```
<vald_mapp>/diagramer/
├─ <källfil>_linjediagram/
│  ├─ alla_arter.png
│  └─ <art>.png
├─ <källfil>_stapeldiagram_ART/
│  └─ <art>.png
└─ <källfil>_stapeldiagram_NVI/
   └─ <art>.png
```

## Systemkrav

* Python 3.9+ (rekommenderat 3.10/3.11)
* Bibliotek: se `requirements.txt`
* Tkinter för dialogrutor (ingår i standard-Python, men kräver systemets Tk)

## Installation (exempel)

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

pip install -r requirements.txt
```

## Körning

```bash
python Sla_ihop_exceL_fladdermus_tabell_och_skapa_diagram.py
```

1. Välj indatafiler (.xlsx).
2. Välj plats/namn för sammanställning (Excel).
3. Vill du generera diagram:

   * Välj målmapp → underkatalog **diagramer** skapas automatiskt
   * Välj tidsintervall: **Auto** (från filer) eller **Manuellt** (`HH:MM–HH:MM`)

## Indataformat (kolumner)

**Obligatoriskt för tabellen**

* `MANUAL ID`: art + ev. typ, flera poster kan separeras med komma.
  Ex: `Myotis daubentonii FOD, Eptesicus nilssonii SOC`

  * Typkoder: `FOD` → *Födosökande*, `SOC` → *Socialt*, annars → *Förbiflygande*.
  * Rader där art är `Noise` ignoreras.
* För **Antal nätter**: datum och (helst) tid

  * Datumkolumn (en av): `DATE` eller `Datum`
  * Tidskolumn (en av): `TIME` eller `Tid`
  * **Regel:** tider `< 12:00` räknas till föregående natt.

**Obligatoriskt för diagram**

* `MANUAL ID` (se ovan)
* `TIME` (exakt namn) som:

  * `HH:MM` eller `HH:MM:SS`, **eller**
  * Excel-tid som decimalt dygn (t.ex. `0.5` = 12:00)
* Om `TIME` inte kan tolkas försöker skriptet använda de genererade 15-minutersintervallen som fallback.

## Logik & format

* Beteendetyper i tabellen: **Socialt**, **Födosökande**, **Förbiflygande**.
* Kolumn **Art**: *Svenskt namn* + radbrytning + *Latinskt namn* (om känt).
  För **Chiroptera** / **Nyctaloid** visas bara originaltext.
* **Antal nätter** beräknas automatiskt men kan ändras manuellt i Excel.
* **Antal registreringar / natt** = `IFERROR(Fladdermusregistreringar / Antal nätter, "")` (1 decimal).
* Diagram:

  * Gemensam Y-skala per källfil – baserad på arten med högsta staplade toppsumma.
  * Linjediagram: ett samlingsdiagram `alla_arter.png` + ett per art.
  * Stapeldiagram: två set per art: `_ART` (rosa/ljus) och `_NVI` (röd/gul/grå).
  * Legend/ordning: **Socialt**, **Födosökande**, **Förbiflygande**.

## Felsökning

* Inga dialogrutor? Kontrollera Tkinter-installation.
* Tiden kan inte tolkas? Säkerställ kolumnnamnet `TIME` och giltigt format (`HH:MM(:SS)` eller Excel-float).
* Tomma/konstiga diagram? Kontrollera `MANUAL ID` och att artnamn inte är `Noise`.

## Licens

Intern projektfil. Lägg gärna till licens vid behov.

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
