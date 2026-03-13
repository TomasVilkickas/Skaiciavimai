import openpyxl
from openpyxl.styles import Border, Side, PatternFill
from copy import copy

def spalvinti_greitis(kaminas):
    file_name = "Rezultatai.xlsx"
    sheet_name = "Greitis"
    
    try:
        wb = openpyxl.load_workbook(file_name)
    except FileNotFoundError:
        print(f"Klaida: Failas {file_name} nerastas.")
        return

    ws = wb[sheet_name]
    
    # Spalvų ir rėmelių nustatymai
    zalia_fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    geltona_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    thin_side = Side(style="thin")
    pilnas_remelis = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # 1. D6 langelio spalvinimas žaliai
    ws["D6"].fill = zalia_fill
    ws["D6"].border = pilnas_remelis

    # 2. Frazių paieška ir pirminis 6-os eilutės spalvinimas
    zalia_frazes = [
        "diferencinis slėgis", 
        "Pito vamzdelio koeficientas",
        "Temperatūra ortakyje",
        "Atmosferinis slėgis",
        "Statinis slėgis"
    ]
    
    geltona_frazes = [
        "Dujų slėgis kamine",
        "Dujų mol. masė",
        "Dujų srauto greitis matavimo taškuose"
    ]

    paskutinis_duomenu_stulpelis = 4
    for cell in ws[5]:
        if cell.value:
            verte = str(cell.value).lower().strip()
            col = cell.column
            if col > paskutinis_duomenu_stulpelis:
                paskutinis_duomenu_stulpelis = col

            target_cell = ws.cell(row=6, column=col)
            if any(f.lower() in verte for f in zalia_frazes):
                target_cell.fill = zalia_fill
                target_cell.border = pilnas_remelis
            if any(f.lower() in verte for f in geltona_frazes):
                target_cell.fill = geltona_fill
                target_cell.border = pilnas_remelis

    # 3. Formato kopijavimas žemyn (Identinis spalvų ir rėmų klonavimas)
    start_row = 6
    tasku_skaicius = getattr(kaminas, 'tasku_skaicius', 0)
    end_row = start_row + tasku_skaicius - 1

    if tasku_skaicius > 1:
        for col in range(4, paskutinis_duomenu_stulpelis + 1):
            source_cell = ws.cell(row=start_row, column=col)
            if source_cell.fill and source_cell.fill.fill_type is not None:
                source_fill = copy(source_cell.fill)
                for r in range(start_row + 1, end_row + 1):
                    new_cell = ws.cell(row=r, column=col)
                    new_cell.fill = source_fill
                    new_cell.border = pilnas_remelis

    # 4. A-C stulpelių suliejimas (nuo 6 eilutės)
    if tasku_skaicius > 0:
        for merged in list(ws.merged_cells.ranges):
            if merged.min_col <= 3 and merged.min_row >= start_row:
                ws.unmerge_cells(str(merged))
        for col in range(1, 4):
            ws.merge_cells(start_row=start_row, start_column=col, end_row=end_row, end_column=col)
            for r in range(start_row, end_row + 1):
                ws.cell(row=r, column=col).border = pilnas_remelis

    # --- STACIONARŪS STULPELIAI ---
    # Šis rėmelis atsiras 6-oje eilutėje, iškart po paskutinio užpildyto duomenų stulpelio
    paskutinis_spalvotas = 0

    for cell in ws[6]:
        if cell.fill and cell.fill.fill_type == "solid":
            paskutinis_spalvotas = cell.column

    stulpelis_po_lenteles = paskutinis_spalvotas + 1

    target_extra_col = ws.cell(row=6, column=stulpelis_po_lenteles)
    target_extra_col.border = pilnas_remelis

    # --- STACIONARIOS EILUTĖS ---
    # Šis rėmelis atsiras A stulpelyje, iškart po paskutinės duomenų eilutės
    eilute_po_lenteles = end_row + 1
    target_extra_row = ws.cell(row=eilute_po_lenteles, column=1)
    target_extra_row.border = pilnas_remelis

    wb.save(file_name)