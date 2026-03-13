import openpyxl
from openpyxl.styles import Alignment, Font
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

    start_col = paskutinis_spalvotas + 1
    start_row_block = 6
    
    # 5 stulpelių konfigūracija: (stulpelio_indeksas, spalvų_sąrašas)
    # spalvų_sąrašas: True = spalvoti, False = be spalvos
    konfigūracija = {
        0: ([True, False, False, False], geltona_fill), # 1 stulpelis
        1: ([True, True, True, False], zalia_fill),     # 2 stulpelis
        2: ([True, True, False, False], geltona_fill),  # 3 stulpelis
        3: ([True, True, False, False], geltona_fill),  # 4 stulpelis
        4: ([True, True, False, False], zalia_fill)     # 5 stulpelis
    }

    for c_idx in range(5):
        dabartinis_col = start_col + c_idx
        spalvu_planas, spalva = konfigūracija[c_idx]
        
        for r_idx in range(4):
            dabartine_eil = start_row_block + r_idx
            target_cell = ws.cell(row=dabartine_eil, column=dabartinis_col)
            
            # Aprėminame visus bloko langelius
            target_cell.border = pilnas_remelis
            
            # Spalviname pagal planą
            if spalvu_planas[r_idx]:
                target_cell.fill = spalva

        extra_start_col = start_col + 5  # Pradedame ten, kur baigėsi 5x4 blokas
    
    # 7 stulpelių spalvų planas (True = žalia, False = geltona)
    # Seka: žalia, geltona, žalia, žalia, žalia, geltona, geltona
    sekos_spalvos = [zalia_fill, geltona_fill, zalia_fill, zalia_fill, zalia_fill, geltona_fill, geltona_fill]

    for i, spalva in enumerate(sekos_spalvos):
        curr_col = extra_start_col + i
        for r_idx in range(2):  # Dvi eilutės (6 ir 7)
            target_cell = ws.cell(row=6 + r_idx, column=curr_col)
            target_cell.fill = spalva
            target_cell.border = pilnas_remelis

    # --- SPECIFINĖ STACIONARIŲ STULPELIŲ PABAIGA ---
    # 1. Vienas langelis pirmoje eilutėje (6 eilutė) žaliai
    zalias_langelis_col = extra_start_col + 7
    zalias_langelis = ws.cell(row=6, column=zalias_langelis_col)
    zalias_langelis.fill = zalia_fill
    zalias_langelis.border = pilnas_remelis
    
    # 2. Sekantis stulpelis: sulietos abi eilutės (6 ir 7), be spalvos
    merge_col = zalias_langelis_col + 1
    ws.merge_cells(start_row=6, start_column=merge_col, end_row=7, end_column=merge_col)
    
    # Rėmelio pritaikymas sulietam langeliui
    for r_idx in range(2):
        ws.cell(row=6 + r_idx, column=merge_col).border = pilnas_remelis

    # --- STACIONARIOS EILUTĖS ---
   # 1. PIRMA STACIONARI EILUTĖ (Tik rėmai iki greičio stulpelio)
    eilute_po_lenteles = end_row + 1
    ribinis_stulpelis = 1
    for cell in ws[5]:
        if cell.value and "Vidutinis dujų srauto greitis ortakyje w, m/s" in str(cell.value):
            ribinis_stulpelis = cell.column
            break
    
    for col in range(1, ribinis_stulpelis):
        ws.cell(row=eilute_po_lenteles, column=col).border = pilnas_remelis

    # 2. ANTROJI (ORANŽINĖ) EILUTĖ
    dabartine_eil = eilute_po_lenteles + 1
    
    # A, B, C stulpeliai (rėmeliai)
    for col in range(1, 4):
        ws.cell(row=dabartine_eil, column=col).border = pilnas_remelis
    
    # D stulpelis (RPU***)
    rpu_cell = ws.cell(row=dabartine_eil, column=4)
    rpu_cell.value = "RPU***"
    rpu_cell.font = Font(bold=True)
    rpu_cell.border = pilnas_remelis
    rpu_cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # Oranžinis blokas (E stulpelis iki tos pačios ribos, kur baigiasi pirma eilutė)
    # Naudojame ribinis_stulpelis - 1, kad oranžinė dalis neviršytų pagrindinės lentelės pločio
    oranz_end_col = ribinis_stulpelis - 1
    if oranz_end_col >= 5:
        ws.merge_cells(start_row=dabartine_eil, start_column=5, end_row=dabartine_eil, end_column=oranz_end_col)
        oranzinis_blokas = ws.cell(row=dabartine_eil, column=5)
        oranzinis_blokas.value = "Kai bus 2 linijos, RPU*** yra 2 linijos vidurinis taškas"
        oranzinis_blokas.fill = PatternFill(start_color="FF9900", end_color="FF9900", fill_type="solid")
        oranzinis_blokas.font = Font(bold=True)
        oranzinis_blokas.alignment = Alignment(horizontal="center", vertical="center")
        
        # Rėmeliai oranžiniam blokui
        for col in range(5, oranz_end_col + 1):
            ws.cell(row=dabartine_eil, column=col).border = pilnas_remelis

    # 3. TRYS PASKUTINĖS EILUTĖS (3x3 blokelis A, B, C stulpeliuose)
    for i in range(1, 4):
        kitas_eil = dabartine_eil + i
        for col in range(1, 4):
            ws.cell(row=kitas_eil, column=col).border = pilnas_remelis

    wb.save(file_name)