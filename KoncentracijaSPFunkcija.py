import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment, PatternFill
from MatavimoVieta import Kaminas

def spalvinti_koncentracija(kaminas_obj: Kaminas):
    file_path = "Rezultatai.xlsx"
    sheet_name = "Koncentracija"
    
    wb = load_workbook(file_path)
    if sheet_name not in wb.sheetnames:
        wb.create_sheet(sheet_name)
    ws = wb[sheet_name]

    # Stiliai
    thin_side = Side(border_style="thin", color="000000")
    geltona_uzpilda = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    centruoti = Alignment(horizontal="center", vertical="center")

    # --- ŠALINAME buvusį 1 ir 2 punktus (A, B, C stulpelių tvarkymą) ---
    # Paliekame tik tai, kas prasideda nuo D stulpelio (4-as indeksas)

    # 3. D, E, F, G, H, I, J, K stulpeliai: visi langeliai aprėminti atskirai
    geltoni_cols = [5, 6, 7, 8, 9, 10, 11] # E iki K
    visi_cols = range(4, 12) # D(4) iki K(11)
    
    for c in visi_cols:
        for r in range(6, 9):
            cell = ws.cell(row=r, column=c)
            cell.border = Border(top=thin_side, left=thin_side, right=thin_side, bottom=thin_side)
            if c in geltoni_cols:
                cell.fill = geltona_uzpilda

    # 4. L, M, N, O stulpeliai: sulieti 6, 7, 8 ir aprėminti
    for c in range(12, 16): # L, M, N, O
        # Prieš liejant, saugiau išvalyti senus suliejimus, jei jie egzistuoja
        for range_ in list(ws.merged_cells.ranges):
            if ws.cell(row=6, column=c).coordinate in range_:
                ws.unmerge_cells(str(range_))
                
        ws.merge_cells(start_row=6, start_column=c, end_row=8, end_column=c)
        cell = ws.cell(row=6, column=c)
        cell.border = Border(top=thin_side, left=thin_side, right=thin_side, bottom=thin_side)
        cell.alignment = centruoti
        if c in [12, 13]: # L ir M
            cell.fill = geltona_uzpilda

    # 5. P stulpelis: 6, 7, 8 langeliai aprėminti atskirai
    for r in range(6, 9):
        ws.cell(row=r, column=16).border = Border(top=thin_side, left=thin_side, right=thin_side, bottom=thin_side)

    # 6. Teksto įrašymas į O stulpelį
    ws.cell(row=6, column=15).value = kaminas_obj.filtru_skaicius
    ws.cell(row=6, column=15).alignment = centruoti

    # 7. Teksto įrašymas į P stulpelį
    linijos = kaminas_obj.liniju_skaicius
    filtrai = kaminas_obj.filtru_skaicius
    lin_tekstas = "linija" if linijos == 1 else "linijos"
    
    for i in range(1, filtrai + 1):
        if i > 3: break
        irasas = f"{linijos} {lin_tekstas}, {i} filtras"
        target_row = 5 + i
        ws.cell(row=target_row, column=16).value = irasas
        ws.cell(row=target_row, column=16).alignment = Alignment(horizontal="left", vertical="center")
    
    wb.save(file_path)