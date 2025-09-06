# LaczTabeleNietoperzy

Skrypt Pythona łączący wiele plików Excel z danymi o nietoperzach w jedną zbiorczą tabelę z podsumowaniami i formatowaniem w Excelu.

Skrypt ze "skapa_diagram" w nazwie implementuje takze repozytyorium tworzenia wykresow dla danych. 
po utworzeniu tabeli z danymi pyta nas czy utworzyc wykresy i jesli tak - odpala ich kreacje od razu.

Projekt: Sammanställning och diagram för fladdermusdata (Excel)

ÖVERSIKT
Detta skript:
1) Slår ihop valda .xlsx-filer till ett översiktsblad i Excel.
2) Lägger till summeringar: 
   - Fladdermusregistreringar (utan NOISE)
   - Antal nätter (kan justeras manuellt)
   - Antal registreringar / natt (Excel-formel, uppdateras när "Antal nätter" ändras)
   - Total antal ljud (inkl. NOISE)
3) (Valfritt) Genererar linje- och stapeldiagram per fil.

FIL(ER)
- Huvudskript: Sla_ihop_exceL_fladdermus_tabell_och_skapa_diagram.py
- Utdata (Excel): sammanstallning_fladdermus.xlsx
- Utdata (diagram): 
  <vald_mapp>/diagramer/
    └─ <kälfil>_linjediagram/
       ├─ alla_arter.png
       └─ <art>.png
    └─ <kälfil>_stapeldiagram_ART/
       └─ <art>.png
    └─ <kälfil>_stapeldiagram_NVI/
       └─ <art>.png

SYSTEMKRAV
- Python 3.9+ (rekommenderat 3.10/3.11)
- Bibliotek: se requirements.txt
- Tkinter för dialogrutor (ingår i standard-Python, men kräver systemets Tk)

INSTALLATION (exempel)
1) Skapa (valfritt) virtuell miljö:
   python -m venv .venv
   .venv\Scripts\activate   (Windows)
   source .venv/bin/activate (macOS/Linux)

2) Installera beroenden:
   pip install -r requirements.txt

KÖRNING
1) Starta:
   python Sla_ihop_exceL_fladdermus_tabell_och_skapa_diagram.py

2) Välj:
   - Indatafiler (.xlsx) med samma kolumnstruktur
   - Plats/namn för sammanställning (Excel)
   - Om du vill generera diagram:
     • Välj målmapp → skriptet skapar underkatalog "diagramer"
     • Välj tidsintervall: auto (från filer) eller manuellt (HH:MM–HH:MM)

INDATAFORMAT (kolumner)
Obligatoriskt för tabellen:
- "MANUAL ID": art + ev. typ, flera poster kan separeras komma. Exempel:
    "Myotis daubentonii FOD, Eptesicus nilssonii SOC"
  • Typkoder: FOD → "Födosökande", SOC → "Socialt", annars → "Förbiflygande".
  • Rader där art är "Noise" ignoreras.
- För "Antal nätter": datum och (helst) tid:
  • Datumkolumn (en av): "DATE" eller "Datum" (skiftlägesokänsligt)
  • Tidskolumn (en av): "TIME" eller "Tid" (skiftlägesokänsligt)
  • REGEL: tider < 12:00 räknas till föregående natt.

Obligatoriskt för diagram:
- "MANUAL ID" (se ovan)
- "TIME" (exakt namn) – tid i format:
  • "HH:MM" eller "HH:MM:SS", eller 
  • Excel-tid som decimalt dygn (t.ex. 0.5 = 12:00)
- (Om "TIME" ej kan tolkas) skriptet försöker använda de redan skapade 15-minutersintervallen som fallback.

LOGIK & FORMAT
- Sortering av beteendetyper i tabellen: Socialt, Födosökande, Förbiflygande.
- Kolumnen "Art" visar: Svenskt namn + radbrytning + Latinskt namn (om känt). För Chiroptera/Nyctaloid visas bara originaltext.
- "Antal nätter" beräknas automatiskt men kan ändras manuellt i Excel.
- "Antal registreringar / natt" = IFERROR(Fladdermusregistreringar / Antal nätter, "") som Excel-formel (1 decimal).
- Diagram:
  • Gemensam Y-skala inom varje källfil – baseras på arten med högst staplad toppsumma i 15-min intervallen.
  • Linjediagram: ett samlingsdiagram ("alla_arter.png") + ett per art.
  • Stapeldiagram: två uppsättningar per art: _ART (rosa/ljus) och _NVI (röd/gul/grå).
  • Beteendetypernas ordning i legenden: Socialt, Födosökande, Förbiflygande.

FEL & FELSÖKNING
- Dialogrutor visas inte: kontrollera att Tkinter är installerat i din Python.
- Tiden kan inte tolkas: kontrollera att kolumnen heter "TIME" (för diagram) eller att tiderna kan läsas som "HH:MM(:SS)" eller Excel-float.
- Tomma/konstiga diagram: kontrollera att MANUAL ID inte är tomt och att artnamn inte är "Noise".

LICENS
Intern projektfil. Lägg gärna till licens efter behov.

FÖRSLAG PÅ .gitignore
diagramer/
*_linjediagram/
*_stapeldiagram_*/
sammanstallning_fladdermus.xlsx
.venv/
__pycache__/
