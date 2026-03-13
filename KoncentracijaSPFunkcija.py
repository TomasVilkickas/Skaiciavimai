import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment, PatternFill
from MatavimoVieta import Kaminas

def spalvinti_koncentracija(kaminas_obj: Kaminas):
    file_path = "Rezultatai.xlsx"
    sheet_name = "Koncentracija"
    
    # Užkrauname darbaknygę
    wb = load_workbook(file_path)
    if sheet_name not in wb.sheetnames:
        wb.create_sheet(sheet_name)
    ws = wb[sheet_name]

    # Stiliai
    thin_side = Side(border_style="thin", color="000000")
    geltona_uzpilda = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    centruoti = Alignment(horizontal="center", vertical="center")

    def reminti_isore(ws, start_row, end_row, col_idx):
        """Aprėmina stulpelio langelių grupę vienu bendru išoriniu rėmu"""
        for r in range(start_row, end_row + 1):
            cell = ws.cell(row=r, column=col_idx)
            # Nustatome rėmus tik kraštinėms
            left = thin_side if col_idx else None
            right = thin_side
            top = thin_side if r == start_row else None
            bottom = thin_side if r == end_row else None
            
            # Kadangi tai vienas stulpelis, kairė ir dešinė visada "thin"
            cell.border = Border(top=top, bottom=bottom, left=thin_side, right=thin_side)

    # 1. A ir B stulpeliai: 6, 7, 8 langeliai aprėminti bendru išoriniu rėmu
    reminti_isore(ws, 6, 8, 1) # A
    reminti_isore(ws, 6, 8, 2) # B

    # 2. C stulpelis: sulieti 6, 7, 8 ir aprėminti
    ws.merge_cells(start_row=6, start_column=3, end_row=8, end_column=3)
    ws.cell(row=6, column=3).border = Border(top=thin_side, left=thin_side, right=thin_side, bottom=thin_side)
    ws.cell(row=6, column=3).alignment = centruoti

    # 3. D, E, F, G, H, I, J, K stulpeliai: visi langeliai aprėminti atskirai
    # E, F, G, H, I, J, K spalvinami geltonai
    geltoni_cols = [5, 6, 7, 8, 9, 10, 11] # E iki K
    visi_cols = range(4, 12) # D(4) iki K(11)
    
    for c in visi_cols:
        for r in range(6, 9):
            cell = ws.cell(row=r, column=c)
            cell.border = Border(top=thin_side, left=thin_side, right=thin_side, bottom=thin_side)
            if c in geltoni_cols:
                cell.fill = geltona_uzpilda

    # 4. L, M, N, O stulpeliai: sulieti 6, 7, 8 ir aprėminti
    # L(12) ir M(13) spalvinami geltonai
    for c in range(12, 16): # L, M, N, O
        ws.merge_cells(start_row=6, start_column=c, end_row=8, end_column=c)
        cell = ws.cell(row=6, column=c)
        cell.border = Border(top=thin_side, left=thin_side, right=thin_side, bottom=thin_side)
        cell.alignment = centruoti
        if c in [12, 13]: # L ir M
            cell.fill = geltona_uzpilda

    # 5. P stulpelis: 6, 7, 8 langeliai aprėminti atskirai
    for r in range(6, 9):
        ws.cell(row=r, column=16).border = Border(top=thin_side, left=thin_side, right=thin_side, bottom=thin_side)

    # 6. Teksto įrašymas į O stulpelį (Sulietas O6:O8)
    # Įrašome bendrą filtrų skaičių
    ws.cell(row=6, column=15).value = kaminas_obj.filtru_skaicius
    ws.cell(row=6, column=15).alignment = centruoti

    # 7. Teksto įrašymas į P stulpelį (P6, P7, P8)
    linijos = kaminas_obj.liniju_skaicius
    filtrai = kaminas_obj.filtru_skaicius

    # Žodžių galūnių derinimas (paprastas variantas)
    lin_tekstas = "linija" if linijos == 1 else "linijos"
    
    # Einame per eilutes (6, 7, 8), bet tik tiek, kiek turime filtrų (iki 3)
    for i in range(1, filtrai + 1):
        if i > 3: break  # Kadangi turime tik 3 eilutes (6, 7, 8)
        
        fil_tekstas = "filtras" if i == 1 else "filtrai" # Arba tiesiog "filtras" pagal jūsų pvz.
        irasas = f"{linijos} {lin_tekstas}, {i} filtras"
        
        # Rašome į P stulpelį (16 stulpelis), eilutės 6, 7, 8
        target_row = 5 + i
        ws.cell(row=target_row, column=16).value = irasas
        ws.cell(row=target_row, column=16).alignment = Alignment(horizontal="left", vertical="center")
    
        wb.save(file_path)