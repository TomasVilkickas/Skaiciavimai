import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment
from MatavimoVieta import Kaminas

def spalvinti_sverimas(kaminas: Kaminas):
    # Naudojame pandas ir openpyxl pagal susitarimą [cite: 2026-03-03]
    try:
        wb = openpyxl.load_workbook("Rezultatai.xlsx")
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        
    if "Svėrimas" not in wb.sheetnames:
        wb.create_sheet("Svėrimas")
    ws = wb["Svėrimas"]

    # Stiliai
    zalia_fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    geltona_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    thin_side = Side(style='thin')
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    centravimas = Alignment(horizontal="center", vertical="center")

    # 1. SPALVINIMAS PAGAL FILTRŲ SKAIČIŲ
    zali_stulpeliai = [5, 6, 7, 10, 11, 12] # E, F, G, J, K, L
    
    # Filtrų dalis (nuo 6 eilutės)
    for i in range(kaminas.filtru_skaicius):
        eilute = 6 + i
        cell_c = ws.cell(row=eilute, column=3)
        cell_c.value = i + 1
        cell_c.alignment = centravimas
        
        for col in zali_stulpeliai:
            ws.cell(row=eilute, column=col).fill = zalia_fill

    # PAPILDYMAS: Nuspalvinama 9-ta eilutė (E9, F9, G9, J9, K9, L9)
    for col in zali_stulpeliai:
        ws.cell(row=9, column=col).fill = zalia_fill

    # Antra žalia sekcija (nuo 11 iki 15 eilutės)
    for eilute in range(11, 16):
        for col in zali_stulpeliai:
            ws.cell(row=eilute, column=col).fill = zalia_fill

    # Geltona spalva M stulpelyje (13-as stulpelis)
    for eilute in list(range(6, 10)) + list(range(11, 17)):
        ws.cell(row=eilute, column=13).fill = geltona_fill

    # --- RĖMAI IR SULIEJIMAS (6-16 eilutės) ---
    s_row, e_row = 6, 16

    # Funkcija bendram išoriniam rėmui (A, D, H, I, N)
    def apply_column_frame(col):
        for r in range(s_row, e_row + 1):
            cell = ws.cell(row=r, column=col)
            cell.border = Border(
                left=thin_side, 
                right=thin_side, 
                top=(thin_side if r == s_row else None),
                bottom=(thin_side if r == e_row else None)
            )

    for c in [1, 4, 8, 9, 14]: # A, D, H, I, N
        apply_column_frame(c)

    # B STULPELIS: Suliejimas ir rėmas (su apatiniu brūkšniu)
    ws.merge_cells(start_row=s_row, start_column=2, end_row=e_row, end_column=2)
    for r in range(s_row, e_row + 1):
        cell = ws.cell(row=r, column=2)
        cell.border = Border(
            left=thin_side, 
            right=thin_side, 
            top=(thin_side if r == s_row else None),
            bottom=(thin_side if r == e_row else None)
        )

    # C, E, F, G, J, K, L, M, O stulpeliai: Kiekvienas langelis atskirai
    atskiri = [3, 5, 6, 7, 10, 11, 12, 13, 15]
    for r in range(s_row, e_row + 1):
        for c in atskiri:
            ws.cell(row=r, column=c).border = thin_border

# --- TEKSTŲ UŽPILDYMAS ---
    
    # E9: Tuščiasis ėminys
    ws.cell(row=9, column=15).value = "Tuščiasis ėminys (mt)"
    ws.cell(row=9, column=15).alignment = Alignment(wrap_text=True, horizontal="left", vertical="center")

    # O stulpelis: Papildomi aprašymai
    papildomi_tekstai = {
        11: "Išplautos nuosėdos (mn)",
        12: "Išvalytos ėminių sistemos tuščiasis ėminys (mIt)",
        13: "Tušti indai svorio kontrolei (mIt1)",
        14: "Tušti indai svorio kontrolei (mIt2)",
        15: "Tušti indai svorio kontrolei (mIt3)",
        16: "Tuščių indų ir išvalytos ėminių sistemos tuščiojo ėminio svorio pokytis (mItp)"
    }

    for eil, tekstas in papildomi_tekstai.items():
        cell = ws.cell(row=eil, column=15) # 15 yra O stulpelis
        cell.value = tekstas
        cell.alignment = Alignment(wrap_text=True, horizontal="left", vertical="center")

    # O6: Linijų ir filtrų skaičius (su lietuviškomis galūnėmis)
    def gauti_galune(n, vns, dgs_kilm, dgs_vard):
        # 11-19 visada kilmininkas (linijų, filtrų)
        if 11 <= n % 100 <= 19:
            return f"{n} {dgs_kilm}"
        # Jei baigiasi 1 (bet ne 11) - vardininkas vienaskaita
        if n % 10 == 1:
            return f"{n} {vns}"
        # Jei baigiasi 2-9 - vardininkas daugiskaita
        if 2 <= n % 10 <= 9:
            return f"{n} {dgs_vard}"
        # Likusieji (pvz. 20, 30) - kilmininkas
        return f"{n} {dgs_kilm}"

    linijos_str = gauti_galune(kaminas.liniju_skaicius, "linija", "linijų", "linijos")
    filtrai_str = gauti_galune(kaminas.filtru_skaicius, "filtras", "filtrų", "filtrai")
    
    ws.cell(row=6, column=15).value = f"{linijos_str}, {filtrai_str}"
    ws.cell(row=6, column=15).alignment = Alignment(horizontal="left", vertical="center")
    # Nustatome O stulpelio (15-as stulpelis) plotį
    ws.column_dimensions['O'].width = 38
   
    # --- PAPILDOMI ĮRAŠAI C STULPELYJE ---
    c_stulpelio_duomenys = {
        9: "2",
        11: "41",
        12: "00",
        13: "01",
        14: "02",
        15: "03"
    }

    for eil, reiksme in c_stulpelio_duomenys.items():
        cell = ws.cell(row=eil, column=3)
        cell.value = reiksme
        cell.alignment = centravimas # Naudojame jūsų kode jau apibrėžtą centravimą

    wb.save("Rezultatai.xlsx")