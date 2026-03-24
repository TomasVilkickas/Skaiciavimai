from openpyxl import load_workbook
from MatavimoVieta import Kaminas
from PaemimasSP2Funkcija import spalvinti_paemimas2

def spalvinti_paemimas1(kaminas_obj: Kaminas):
    file_path = "Rezultatai.xlsx"
    
    # 1. Atidarome failą
    try:
        wb = load_workbook(file_path)
    except FileNotFoundError:
        from openpyxl import Workbook
        wb = Workbook()

    # 2. APIBRĖŽIAME 'ws' kintamąjį (kad išvengtume UnboundLocalError)
    if "Paėmimas" in wb.sheetnames:
        ws = wb["Paėmimas"]
    else:
        ws = wb.create_sheet("Paėmimas")

    dabartine_eilute = 6
    tarpas = 4
    ar_pirma_lentele = True

    # 3. Sukame ciklus
    for f in range(kaminas_obj.filtru_skaicius):
        for l in range(kaminas_obj.liniju_skaicius):
            
            yra_kopija = not ar_pirma_lentele
            
            # Kviečiame darbininką, dabar 'ws' jau tikrai egzistuoja
            spalvinti_paemimas2(
                ws=ws, 
                kaminas_obj=kaminas_obj, 
                pradzios_eilute=dabartine_eilute, 
                yra_kopija=yra_kopija
            )
            
            # Perstumiame eilutę žemyn
            dabartine_eilute += kaminas_obj.tasku_skaicius + tarpas
            ar_pirma_lentele = False

    # 4. Išsaugojame
    wb.save(file_path)