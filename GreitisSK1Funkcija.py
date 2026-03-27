import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment

def skaiciuoti_greitis1(kaminas_obj):
    # 1. Failo pavadinimą nustatome čia, kad nereikėtų jo siųsti iš pagrindinės programos
    failo_pavadinimas = 'Rezultatai.xlsx'
    
    try:
        # 2. Atidarome failą (naudojame openpyxl, kaip nurodyta jūsų taisyklėse)
        wb = load_workbook(failo_pavadinimas)
        
        # Patikriname, ar yra lapas „Greitis“, jei ne - naudojame pirmą aktyvų
        if 'Greitis' in wb.sheetnames:
            ws = wb['Greitis']
        else:
            ws = wb.active
            
    except FileNotFoundError:
        print(f"Klaida: Failas {failo_pavadinimas} nerastas.")
        return

    # Pagalbinė funkcija surasti frazės koordinates
    def rasti_koordinates(frazė):
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and str(frazė) in str(cell.value):
                    return cell.row, cell.column
        return None, None

    # 3. Surandame pagrindinių parametrų vietas
    atm_row, atm_col = rasti_koordinates("Atmosferinis slėgis P, hPa")
    stat_row, stat_col = rasti_koordinates("Statinis slėgis ortakyje ± ΔP, hPa")
    pk_row, pk_col = rasti_koordinates("Dujų slėgis kamine Pk= P+ ΔP, hPa")

    # Tikriname, ar viską radome
    if not all([atm_row, stat_row, pk_row]):
        print("Klaida: Nepavyko rasti visų reikiamų frazių Excel faile.")
        return

    # 4. Skaičiavimai pagal kaminas_obj parametrus
    for i in range(1, kaminas_obj.tasku_skaicius + 1):
        # Pasiimame koordinates
        src_atm_cell = ws.cell(row=atm_row + i, column=atm_col).coordinate
        src_stat_cell = ws.cell(row=stat_row + i, column=stat_col).coordinate
        
        # Įrašome formulę į Pk stulpelį
        target_cell = ws.cell(row=pk_row + i, column=pk_col)
        target_cell.value = f"={src_atm_cell}+{src_stat_cell}"
        
        # Formatuojame
        target_cell.alignment = Alignment(horizontal='center')
        target_cell.number_format = '0.0'

    # --- Dujų molinės masės Ms skaičiavimas ---
    qn_row, qn_col = rasti_koordinates("Sausų dujų tankis normaliosioms sąlygomis qn, (kg/m3)")
    ms_row, ms_col = rasti_koordinates("Dujų mol. masė Ms= qn *22.4, kg/")

    if qn_row and ms_row:
        # Paimame tikslų langelį po fraze (pvz., AA5)
        qn_langele = ws.cell(row=qn_row + 1, column=qn_col)
        qn_adresas = qn_langele.coordinate # Gauname adresą, pvz., 'AA5'
        
        for i in range(1, kaminas_obj.tasku_skaicius + 1):
            target_ms_cell = ws.cell(row=ms_row + i, column=ms_col)
            
            # SUPAPRASTINTA FORMULĖ: be papildomų vienučių kabučių, kurios dažnai sugadina failą
            target_ms_cell.value = f"={qn_adresas}*22.4"
            
            target_ms_cell.alignment = Alignment(horizontal='center')
            target_ms_cell.number_format = '0.0000'

# ---Greičio wi skaičiavimai ---

    # 1. Surandame pagrindinių parametrų koordinates (naudojame unikalias frazių dalis)
    k_row, k_col = rasti_koordinates("Pito vamzdelio koeficientas")
    t_row, t_col = rasti_koordinates("Temperatūra ortakyje t")
    atm_row, atm_col = rasti_koordinates("Atmosferinis slėgis P")
    stat_row, stat_col = rasti_koordinates("Statinis slėgis ortakyje")
    ms_row, ms_col = rasti_koordinates("Dujų mol. masė Ms")

    # K koeficientas (langelis po fraze)
    k_addr = ws.cell(row=k_row + 1, column=k_col).coordinate if k_row else "1.0"

    # 2. Ciklas per linijas (iš kaminas_obj)
    for l_idx in range(1, kaminas_obj.liniju_skaicius + 1):
        
        # Paieška konkrečiai linijai (naudojame fragmentus, kad išvengtume tarpų klaidų)
        wi_frazė = f"srauto greitis matavimo taškuose, {l_idx} linija"
        pdi_frazė = f"diferencinis slėgis taškuose" # Papildomai tikrinsime linijos nr. žemiau
        
        wi_res_row, wi_res_col = rasti_koordinates(wi_frazė)
        
        # Pdi paieška konkrečiai linijai
        pdi_row, pdi_col = None, None
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and "diferencinis slėgis" in str(cell.value) and f"{l_idx} linija" in str(cell.value):
                    pdi_row, pdi_col = cell.row, cell.column
                    break
            if pdi_row: break

        if wi_res_row and pdi_row:
            # 3. Einame per matavimo taškus žemyn
            for i in range(1, kaminas_obj.tasku_skaicius + 1):
                
                # Dinaminiai langelių adresai kiekvienam taškui
                pdi_addr = ws.cell(row=pdi_row + i, column=pdi_col).coordinate
                t_addr = ws.cell(row=t_row + i, column=t_col).coordinate
                atm_addr = ws.cell(row=atm_row + i, column=atm_col).coordinate
                stat_addr = ws.cell(row=stat_row + i, column=stat_col).coordinate
                ms_addr = ws.cell(row=ms_row + i, column=ms_col).coordinate
                
                target_cell = ws.cell(row=wi_res_row + i, column=wi_res_col)
                
                # Jūsų nurodyta formulė:
                # =K*129*(SQRT(Pdi*(273+t)/(Atm+Stat)/Ms))
                tiksli_formule = (
                    f"={k_addr}*129*(SQRT({pdi_addr}*(273+{t_addr})/"
                    f"({atm_addr}+{stat_addr})/{ms_addr}))"
                )
                
                target_cell.value = tiksli_formule
                
                # Formatavimas
                target_cell.alignment = Alignment(horizontal='center')
                target_cell.number_format = '0.00'
    
    # 5. Išsaugome
    wb.save(failo_pavadinimas)