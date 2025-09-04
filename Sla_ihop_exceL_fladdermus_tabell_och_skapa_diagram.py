import os
import re
import sys
import math
import subprocess
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
from matplotlib.ticker import MaxNLocator

import tkinter as tk
from tkinter import filedialog, messagebox

from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter

# ================== UI ==================
TITLE_OPEN = "Välj en eller flera Excel-filer (.xlsx, samma struktur)"
TITLE_SAVE = "Välj var du vill spara sammanställningen"
DEFAULT_OUT = "sammanstallning_fladdermus.xlsx"

# ================== Kolory / style ==================
COLOR_FORBI = "FFE7E6E6"     # Förbiflygande (jasnoszary)
COLOR_SOC   = "FFFF0000"     # Socialt (czerwony)
COLOR_FODO  = "FFFFC000"     # Födosökande (żółty)
COLOR_HEADER_BG = "FF595959" # nagłówek tło
COLOR_HEADER_FG = "FFFFFFFF" # nagłówek font

FILL_FORBI = PatternFill("solid", fgColor=COLOR_FORBI)
FILL_SOC   = PatternFill("solid", fgColor=COLOR_SOC)
FILL_FODO  = PatternFill("solid", fgColor=COLOR_FODO)
FILL_HDR   = PatternFill("solid", fgColor=COLOR_HEADER_BG)

BORDER_MEDIUM = Side(style="medium", color="FF000000")

# ================== Słownik: Łacina → Szwedzki ==================
LATIN_TO_SV = {
    "Barbastella barbastellus": "Barbastell",
    "Eptesicus nilssonii": "Nordfladdermus",
    "Eptesicus serotinus": "Sydfladdermus",
    "Myotis alcathoe": "Nymffladdermus",
    "Myotis bechsteinii": "Bechsteins fladdermus",
    "Myotis brandtii": "Tajgafladdermus",
    "Myotis dasycneme": "Dammfladdermus",
    "Myotis daubentonii": "Vattenfladdermus",
    "Myotis myotis": "Större musöra",
    "Myotis mystacinus": "Mustaschfladdermus",
    "Myotis nattereri": "Fransfladdermus",
    "Myotis mystacinus/brandtii": "Mustasch/Tajgafladdermus",
    "Nyctalus leisleri": "Mindre brunfladdermus",
    "Nyctalus noctula": "Större brunfladdermus",
    "Pipistrellus kuhlii": "Parkpipistrell",
    "Pipistrellus nathusii": "Trollpipistrell",
    "Pipistrellus pipistrellus": "Sydpipistrell",
    "Pipistrellus pygmaeus": "Dvärgpipistrell",
    "Plecotus auritus": "Brunlångöra",
    "Plecotus austriacus": "Grålångöra",
    "Vespertilio murinus": "Gråskimlig fladdermus",
    "Nyctaloid": None,
    "Chiroptera": None,
}
SPECIAL_TAIL = {"nyctaloid", "chiroptera"}  # Nyctaloid/Chiroptera na koniec listy

# ================== Pomocnicze (wspólne) ==================
def safe_sheet_name(path: str, used: set) -> str:
    base = os.path.splitext(os.path.basename(path))[0]
    base = re.sub(r'[:\\\/\?\*\[\]]', '_', base).strip()[:31] or "Ark"
    cand = base; i = 2
    while cand in used or not cand:
        suf = f"_{i}"
        cand = (base[: (31 - len(suf))] + suf) if len(base) + len(suf) > 31 else (base + suf)
        i += 1
    used.add(cand); return cand

def display_label_multiline(latin: str) -> str:
    """Zwraca 'Szwedzka,\\nŁacińska' lub samo 'Łacińska' (dla Nyctaloid/Chiroptera)."""
    latin = str(latin).strip()
    sv = LATIN_TO_SV.get(latin, None)
    if sv:
        return f"{sv.capitalize()},\n{latin}"
    return latin

def extract_species_and_type(manual_id_value):
    """
    Rozbija MANUAL ID na (gatunek, typ). Typy: Förbiflygande / Socialt / Födosökande.
    Uwaga: zawsze zwracamy 'Födosökande' z 'ö' (spójność z wykresami).
    """
    if pd.isna(manual_id_value): return []
    out = []
    for raw in str(manual_id_value).split(","):
        entry = raw.strip()
        if not entry: continue
        upper = entry.upper()
        if "FOD" in upper: typ = "Födosökande"
        elif "SOC" in upper: typ = "Socialt"
        else: typ = "Förbiflygande"
        species = re.sub(r'\b(FOD|SOC)\b', "", entry, flags=re.IGNORECASE).strip(" ,;")
        out.append((species, typ))
    return out

def open_file(path):
    try:
        if sys.platform.startswith("win"): os.startfile(path)
        elif sys.platform == "darwin": subprocess.run(["open", path])
        else: subprocess.run(["xdg-open", path])
    except Exception as e:
        print(f"Kan inte öppna filen automatiskt: {e}")

def species_sort_key(latin: str):
    return (str(latin).strip().lower() in SPECIAL_TAIL, str(latin).casefold())

def fill_for_type(t):
    # Akceptujemy oba warianty zapisu (z/bez 'ö'), ale używamy 'Födosökande' dalej
    if t in ("Födosökande", "Fodosökande"): return FILL_FODO
    if t == "Socialt": return FILL_SOC
    if t == "Förbiflygande": return FILL_FORBI
    return None

def safe_filename(s):
    return re.sub(r'[\\/:\*\?"<>\|]', '_', str(s))

def str_to_dt(time_str):
    # Przyjmuje HH:MM:SS (string lub excelowy czas rzutowany do stringa)
    t = pd.to_datetime(str(time_str), format="%H:%M:%S")
    fake_date = "2000-01-02" if t.hour < 12 else "2000-01-01"
    return datetime.strptime(f"{fake_date} {t.hour:02d}:{t.minute:02d}", "%Y-%m-%d %H:%M")

def round_down_15(dt):
    return dt.replace(minute=(dt.minute // 15) * 15, second=0)

def round_up_15(dt):
    if dt.minute % 15 != 0 or dt.second > 0:
        dt = dt + timedelta(minutes=15 - (dt.minute % 15))
    return dt.replace(second=0)

def format_title(species_latin: str, total_count: int) -> str:
    """
    Tytuł wykresu pojedynczego gatunku: 'svensk (latin), antal observerade beteenden: NN'
    lub 'latin, antal ...' gdy brak nazwy szwedzkiej (None/brak w słowniku).
    """
    sv = LATIN_TO_SV.get(species_latin)
    return (f"{sv} ({species_latin}), antal observerade beteenden: {int(total_count)}"
            if sv else
            f"{species_latin}, antal observerade beteenden: {int(total_count)}")

# ================== KROK 1: Wybór plików i zapis zbiorczy ==================
root = tk.Tk(); root.withdraw()

input_files = filedialog.askopenfilenames(
    title=TITLE_OPEN, filetypes=[("Excel-filer (.xlsx)", "*.xlsx")]
)
if not input_files:
    print("Ingen fil vald. Avslutar.")
    sys.exit(0)
input_files = [p for p in input_files if p.lower().endswith(".xlsx")]
if not input_files:
    print("Endast .xlsx stöds. Avslutar.")
    sys.exit(0)

out_path = filedialog.asksaveasfilename(
    title=TITLE_SAVE, defaultextension=".xlsx",
    initialfile=DEFAULT_OUT, filetypes=[("Excel-fil (.xlsx)", "*.xlsx")]
)
if not out_path:
    print("Ingen plats vald. Avslutar.")
    sys.exit(0)

# ---- Zbieranie danych do tabeli „Översikt” ----
used_sheet_names = set()
sheets_to_write = []       # [(sheet_name, df)]
counts_per_file = {}       # {sheet_name: {(latin, type): count}}
total_ljud_per_file = {}   # {sheet_name: total_rows}
all_species_latin = set()

for path in input_files:
    sheet_name = safe_sheet_name(path, used_sheet_names)
    df = pd.read_excel(path, sheet_name=0, engine="openpyxl")
    sheets_to_write.append((sheet_name, df))
    total_ljud_per_file[sheet_name] = int(len(df))  # z NOISE

    if "MANUAL ID" not in df.columns:
        counts_per_file[sheet_name] = {}
        continue

    df["__list"] = df["MANUAL ID"].map(extract_species_and_type)
    long = df.explode("__list")
    long = long[long["__list"].notna()]
    if long.empty:
        counts_per_file[sheet_name] = {}
        continue

    long[["ArtLatin", "Beteendetyper"]] = pd.DataFrame(long["__list"].tolist(), index=long.index)
    # wyklucz 'Noise'
    mask_ok = long["ArtLatin"].astype(str).str.strip().str.lower() != "noise"
    long = long[mask_ok]

    all_species_latin.update(long["ArtLatin"].astype(str).str.strip().tolist())

    grp = long.groupby(["ArtLatin", "Beteendetyper"]).size()
    counts_per_file[sheet_name] = {(sp, typ): int(n) for (sp, typ), n in grp.items()}

# ---- Budowa DataFrame „Översikt” ----
type_order_overview = ["Förbiflygande", "Socialt", "Födosökande"]  # zachowaj kolejność
species_sorted_latin = sorted(all_species_latin, key=species_sort_key)
file_cols = list(counts_per_file.keys())

rows_data = []
for latin in species_sorted_latin:
    disp = display_label_multiline(latin)
    for typ in type_order_overview:
        row = {"Art": disp, "Beteendetyper": typ}
        for col in file_cols:
            val = counts_per_file.get(col, {}).get((latin, typ), 0)
            row[col] = ("" if val == 0 else int(val))  # pusto zamiast 0
        rows_data.append(row)

# wiersze sum:
sum_row = {"Art": "", "Beteendetyper": "Fladdermusregistreringar"}
for col in file_cols:
    sum_row[col] = int(sum(counts_per_file.get(col, {}).values()))
rows_data.append(sum_row)

tot_row = {"Art": "", "Beteendetyper": "Total antal ljud"}
for col in file_cols:
    tot_row[col] = int(total_ljud_per_file.get(col, 0))
rows_data.append(tot_row)

overview_df = pd.DataFrame(rows_data, columns=["Art", "Beteendetyper"] + file_cols)

# ---- Zapis wstępny ----
with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
    overview_df.to_excel(writer, sheet_name="Översikt", index=False)
    for sheet_name, df_orig in sheets_to_write:
        df_orig.to_excel(writer, sheet_name=sheet_name, index=False)

# ---- Stylowanie openpyxl ----
wb = load_workbook(out_path)
ws = wb["Översikt"]

max_row = ws.max_row
max_col = ws.max_column
header_row = 1
data_start = header_row + 1

num_species_rows = len(species_sorted_latin) * 3
sum_row_idx   = data_start + num_species_rows
total_row_idx = sum_row_idx + 1

# Nagłówek
for c in range(1, max_col + 1):
    cell = ws.cell(row=header_row, column=c)
    cell.fill = FILL_HDR
    cell.font = Font(color=COLOR_HEADER_FG, bold=True)
    cell.alignment = Alignment(vertical="center", horizontal="center")
ws.row_dimensions[header_row].height = 18

# Szerokości kolumn
ws.column_dimensions["A"].width = 44
ws.column_dimensions["B"].width = 20
for idx, col_name in enumerate(file_cols, start=3):
    header_text = str(col_name)
    width = max(12, min(50, int(len(header_text) * 1.1)))
    ws.column_dimensions[get_column_letter(idx)].width = width

# Merge i wrap kolumny A
if num_species_rows > 0:
    current = data_start
    merge_end_limit = sum_row_idx - 1
    while current <= merge_end_limit:
        art_val = ws[f"A{current}"].value
        if not art_val:
            current += 1
            continue
        end = current
        while end + 1 <= merge_end_limit and ws[f"A{end+1}"].value == art_val:
            end += 1
        if end > current:
            ws.merge_cells(start_row=current, start_column=1, end_row=end, end_column=1)
        ws.cell(row=current, column=1).alignment = Alignment(vertical="center", wrap_text=True)
        current = end + 1

# Kolorowanie wg kategorii
for r in range(data_start, sum_row_idx):
    typ = ws.cell(row=r, column=2).value
    fill = fill_for_type(typ)
    if fill:
        ws.cell(row=r, column=2).fill = fill
        for c in range(3, max_col + 1):
            val = ws.cell(row=r, column=c).value
            if val not in (None, "", 0):
                ws.cell(row=r, column=c).fill = fill

# Wiersze sum – bold
for r in (sum_row_idx, total_row_idx):
    for c in range(1, max_col + 1):
        ws.cell(row=r, column=c).font = Font(bold=True)
        ws.cell(row=r, column=c).alignment = Alignment(vertical="center")

# Grubsze linie
for i in range(len(species_sorted_latin)):
    top_row = data_start + i * 3
    for c in range(1, max_col + 1):
        old = ws.cell(row=top_row, column=c).border
        ws.cell(row=top_row, column=c).border = Border(
            left=old.left, right=old.right, top=BORDER_MEDIUM, bottom=old.bottom
        )

for c in range(1, max_col + 1):
    old = ws.cell(row=sum_row_idx, column=c).border
    ws.cell(row=sum_row_idx, column=c).border = Border(
        left=old.left, right=old.right, top=BORDER_MEDIUM, bottom=old.bottom
    )
    old = ws.cell(row=total_row_idx, column=c).border
    ws.cell(row=total_row_idx, column=c).border = Border(
        left=old.left, right=old.right, top=old.top, bottom=BORDER_MEDIUM
    )

for c in range(3, max_col + 1):
    for r in range(header_row, max_row + 1):
        old = ws.cell(row=r, column=c).border
        ws.cell(row=r, column=c).border = Border(
            left=BORDER_MEDIUM, right=old.right, top=old.top, bottom=old.bottom
        )
for r in range(header_row, max_row + 1):
    oldA = ws.cell(row=r, column=1).border
    ws.cell(row=r, column=1).border = Border(
        left=BORDER_MEDIUM, right=oldA.right, top=oldA.top, bottom=oldA.bottom
    )
    oldB = ws.cell(row=r, column=2).border
    ws.cell(row=r, column=2).border = Border(
        left=BORDER_MEDIUM, right=oldB.right, top=oldB.top, bottom=oldB.bottom
    )
    oldR = ws.cell(row=r, column=max_col).border
    ws.cell(row=r, column=max_col).border = Border(
        left=oldR.left, right=BORDER_MEDIUM, top=oldR.top, bottom=oldR.bottom
    )

wb.save(out_path)
wb.close()

open_file(out_path)
print(f"Klar! Sparad fil: {out_path}")

# ================== KROK 2: (opcjonalnie) Wykresy nietoperzy ==================
def generate_bat_diagrams(input_files_list, diagrams_root, custom_time_range):
    """
    Generuje wykresy (linia + słupki) dla listy plików.
    - diagrams_root: katalog główny zapisu (już istniejący)
    - custom_time_range: None (auto) lub tuple("HH:MM","HH:MM")
    """
    for input_file in input_files_list:
        print(f"\nBearbetar: {os.path.basename(input_file)}")

        stem = os.path.splitext(os.path.basename(input_file))[0]
        output_dir_lines  = os.path.join(diagrams_root, f"{stem}_linjediagram")
        output_dir_stacks = os.path.join(diagrams_root, f"{stem}_stapeldiagram")
        os.makedirs(output_dir_lines, exist_ok=True)

        # --- wczytaj i przygotuj ---
        df = pd.read_excel(input_file)
        df["species_type_list"] = df["MANUAL ID"].map(extract_species_and_type)

        def time_to_interval(time_str):
            try:
                t = pd.to_datetime(str(time_str), format="%H:%M:%S")
                minutes = int((t.minute // 15) * 15)
                return f"{t.hour:02d}:{minutes:02d}"
            except Exception:
                return ""
        df["interval"] = df["TIME"].map(time_to_interval)

        df_long = df.explode("species_type_list")
        df_long = df_long[df_long["species_type_list"].notna()]
        if df_long.empty:
            print("Inga data efter tolkning. Hoppar över.")
            continue
        df_long[["species", "obs_type"]] = pd.DataFrame(df_long["species_type_list"].tolist(), index=df_long.index)
        df_long = df_long[df_long["species"].astype(str).str.strip().str.lower() != "noise"]
        df_long = df_long[df_long["species"].astype(str).str.strip() != ""]

        # Zakres czasu
        if custom_time_range:
            min_dt = str_to_dt(custom_time_range[0] + ":00")
            max_dt = str_to_dt(custom_time_range[1] + ":00")
            min_dt = round_down_15(min_dt)
            max_dt = round_up_15(max_dt)
        else:
            df["__dt"] = df["TIME"].apply(str_to_dt)
            min_dt = round_down_15(df["__dt"].min())
            max_dt = round_up_15(df["__dt"].max())

        all_intervals = []
        t = min_dt
        while t <= max_dt:
            all_intervals.append(t.strftime("%H:%M"))
            t += timedelta(minutes=15)
        all_intervals = list(dict.fromkeys(all_intervals))

        # Agregaty
        agg = df_long.groupby(["interval", "species", "obs_type"]).size().reset_index(name="antal")
        agg["interval"] = pd.Categorical(agg["interval"], categories=all_intervals, ordered=True)
        species_list = sorted(df_long["species"].unique())

        # Wspólny limit OY (referencja: gatunek z największym „stackiem”)
        def compute_ymax_for_species(species_name, column_order):
            if not species_name:
                return 0
            pd_sp = (
                agg[agg["species"] == species_name]
                .pivot(index="interval", columns="obs_type", values="antal")
                .fillna(0)
                .reindex(all_intervals, fill_value=0)
                .reindex(columns=column_order, fill_value=0)
            )
            if pd_sp.empty:
                return 0
            return int(pd_sp.sum(axis=1).max())

        type_order = ["Socialt", "Födosökande", "Förbiflygande"]
        species_totals = df_long.groupby("species").size().sort_values(ascending=False)
        ref_species = species_totals.index[0] if not species_totals.empty else None
        y_max_ref = compute_ymax_for_species(ref_species, type_order)
        y_lim = max(1, math.ceil(y_max_ref * 1.05))
        print(f"Ref art: {ref_species} | gemensam Y-max (linje + stapel): {y_lim}")

        # LINIE
        agg_line = df_long.groupby(["interval", "species"]).size().reset_index(name="antal")
        agg_line["interval"] = pd.Categorical(agg_line["interval"], categories=all_intervals, ordered=True)
        pivot_line = agg_line.pivot(index="interval", columns="species", values="antal").fillna(0)
        pivot_line = pivot_line.reindex(all_intervals, fill_value=0)

        total_obs = df_long.shape[0]

        # zbiorczy (bez zmiany tytułu)
        plt.figure(figsize=(12, 6))
        ax_all = plt.gca()
        pivot_line.plot(ax=ax_all, marker='o')
        ax_all.set_xlabel("Tid (15-minutersintervall)")
        ax_all.set_ylabel("Antal ljudfiler")
        ax_all.set_title(f"Fladdermusobservationer – alla arter, antal observerade beteenden: {total_obs}")
        ax_all.legend(title="Art", bbox_to_anchor=(1.05, 1), loc='upper left')
        ax_all.yaxis.set_major_locator(MaxNLocator(integer=True))
        ax_all.set_xticks(range(len(pivot_line.index)))
        ax_all.set_xticklabels(pivot_line.index, rotation=270)
        ax_all.set_ylim(0, y_lim)
        plt.tight_layout()
        plt.grid(True, axis='y')
        plt.savefig(os.path.join(output_dir_lines, "alla_arter.png"))
        plt.close()

        # per gatunek
        for species in species_list:
            total_sp = int(pivot_line[species].sum())
            plt.figure(figsize=(10, 4))
            ax = plt.gca()
            ax.plot(pivot_line.index, pivot_line[species], marker='o')
            ax.set_xlabel("Tid (15-minutersintervall)")
            ax.set_ylabel("Antal ljudfiler")
            ax.set_title(format_title(species, total_sp))
            ax.yaxis.set_major_locator(MaxNLocator(integer=True))
            ax.set_xticks(range(len(pivot_line.index)))
            ax.set_xticklabels(pivot_line.index, rotation=270)
            ax.set_ylim(0, y_lim)
            plt.grid(True, axis='y')
            plt.tight_layout()
            plt.savefig(os.path.join(output_dir_lines, f"{safe_filename(species)}.png"))
            plt.close()

        # SŁUPKI – ART
        output_dir_stacks_art = output_dir_stacks + "_ART"
        os.makedirs(output_dir_stacks_art, exist_ok=True)
        color_art = ["#eb09d8", "#d98fd3", "#abaaa9"]  # Socialt, Födosökande, Förbiflygande

        for species in species_list:
            plot_data = (
                agg[agg["species"] == species]
                .pivot(index="interval", columns="obs_type", values="antal")
                .fillna(0)
                .reindex(all_intervals, fill_value=0)
                .reindex(columns=type_order, fill_value=0)
            )
            ax = plot_data.plot(kind="bar", stacked=True, color=color_art, figsize=(14, 6))
            plt.xlabel("Tid (15-minutersintervall)")
            plt.ylabel("Antal ljudfiler")
            plt.title(format_title(species, int(plot_data.values.sum())))
            plt.legend(title="Beteendetyper")
            ax.yaxis.set_major_locator(MaxNLocator(integer=True))
            ax.set_ylim(0, y_lim)
            plt.xticks(rotation=270)
            plt.tight_layout()
            plt.grid(True, axis='y')
            plt.savefig(os.path.join(output_dir_stacks_art, f"{safe_filename(species)}.png"))
            plt.close()

        print(f"Stapeldiagram ART sparade i mappen: {output_dir_stacks_art}")

        # SŁUPKI – NVI
        output_dir_stacks_nvi = output_dir_stacks + "_NVI"
        os.makedirs(output_dir_stacks_nvi, exist_ok=True)
        color_nvi = [(255/255,0/255,0/255), (255/255,190/255,0/255), (169/255,169/255,169/255)]
        type_order_nvi = ["Socialt", "Födosökande", "Förbiflygande"]

        for species in species_list:
            plot_data = (
                agg[agg["species"] == species]
                .pivot(index="interval", columns="obs_type", values="antal")
                .fillna(0)
                .reindex(all_intervals, fill_value=0)
                .reindex(columns=type_order_nvi, fill_value=0)
            )
            ax = plot_data.plot(kind="bar", stacked=True, color=color_nvi, figsize=(14, 6))
            plt.xlabel("Tid (15-minutersintervall)")
            plt.ylabel("Antal ljudfiler")
            plt.title(format_title(species, int(plot_data.values.sum())))
            plt.legend(title="Beteendetyper")
            ax.yaxis.set_major_locator(MaxNLocator(integer=True))
            ax.set_ylim(0, y_lim)
            plt.xticks(rotation=270)
            plt.tight_layout()
            plt.grid(True, axis='y')
            plt.savefig(os.path.join(output_dir_stacks_nvi, f"{safe_filename(species)}.png"))
            plt.close()

        print(f"Stapeldiagram NVI sparade i mappen: {output_dir_stacks_nvi}")
        print(f"Linjediagram sparade i mappen: {output_dir_lines}")
        print(f"Tidsintervall: {all_intervals[0]} – {all_intervals[-1]}")

# ---- Pytanie o generowanie wykresów ----
try:
    if messagebox.askyesno(
        "Generera fladdermusdiagram",
        "Vill du generera fladdermusdiagram nu?\n\n"
        "Du väljer mål-mapp i nästa steg."
    ):
        # Wybór katalogu bazowego na wykresy
        chosen_base = filedialog.askdirectory(title="Välj mapp där resultaten ska sparas")
        if chosen_base:
            diagrams_root = os.path.join(chosen_base, "diagramer")
        else:
            # fallback: folder pliku zbiorczego
            diagrams_root = os.path.join(os.path.dirname(out_path), "diagramer")
        os.makedirs(diagrams_root, exist_ok=True)
        print(f"Resultat kommer att sparas i: {diagrams_root}")

        # Ustal X-axelns intervall (auto / manual)
        mode = messagebox.askquestion(
            "Välj X-axelns intervall",
            "Vill du basera X-axeln på inspelningstid (auto)?\n\n"
            "Välj 'Ja' för automatisk start/slut (utifrån ljudfiler).\n"
            "Välj 'Nej' för att ange eget tidsintervall."
        )
        if mode == "yes":
            custom_time_range = None
        else:
            while True:
                start = tk.simpledialog.askstring("Starttid", "Ange starttid (t.ex. 22:15):")
                end = tk.simpledialog.askstring("Sluttid", "Ange sluttid (t.ex. 02:45):")
                try:
                    pd.to_datetime(start, format="%H:%M")
                    pd.to_datetime(end, format="%H:%M")
                    break
                except Exception:
                    messagebox.showerror("Fel", "Felaktigt tidsformat. Ange t.ex. 22:15 eller 02:45.")
            custom_time_range = (start, end)

        # Generuj wykresy dla TEJ samej listy plików wejściowych
        generate_bat_diagrams(input_files, diagrams_root, custom_time_range)
except Exception as e:
    print(f"Kunde inte visa dialogen eller köra diagramgenerering: {e}")
