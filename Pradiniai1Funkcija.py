import pandas as pd
import os
from MatavimoVieta import Kaminas
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, Font

def sukurti_sablona(kaminas_obj):
    """Sukuriamas Pradiniai.xlsx failas su suformatuota lentele duomenų įvedimui."""
    failo_pavadinimas = "Pradiniai.xlsx"
    lapas_pavadinimas = "Pradiniai"
    
    # Jei failas jau egzistuoja, jo neperrašome, kad neištrintume įvestų duomenų
    if os.path.exists(failo_pavadinimas):
        print(f"Pastaba: {failo_pavadinimas} jau egzistuoja, naujas šablonas nebus kuriamas.")
        return 

    # 1. Stulpelių pavadinimai
    stulpeliai = [
        "Matavimo data", 
        "Ėminių registracijos Nr. T-107-2026-E-", 
        "Objekto pavadinimas, adresas, taršos šaltinio Nr."
    ]
    
    # Sukuriame tuščią DataFrame su šiais stulpeliais
    df_template = pd.DataFrame(columns=stulpeliai)
    
    # 2. Įrašome į Excel naudojant pandas ir openpyxl [cite: 2026-03-03]
    with pd.ExcelWriter(failo_pavadinimas, engine='openpyxl') as writer:
        for pav in ["Pradiniai", "Greitis", "Paėmimas"]:
            if pav == lapas_pavadinimas:
                # Pradedame nuo 5-os eilutės (startrow=4)
                df_template.to_excel(writer, sheet_name=pav, startrow=4, index=False)
            else:
                pd.DataFrame().to_excel(writer, sheet_name=pav, index=False)
        
        # 3. Formatuojame išvaizdą (rėmeliai, BOLD, centravimas)
        ws = writer.sheets[lapas_pavadinimas]
        
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'), 
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        bold_font = Font(bold=True)
        
        # Stulpelių pločiai
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 40
        ws.column_dimensions['C'].width = 60
        
        # Eilučių aukščiai
        ws.row_dimensions[5].height = 40 # Antraštė
        ws.row_dimensions[6].height = 50 # Vieta įvedimui
        
        # Pritaikome stilius A, B ir C stulpeliams
        for col_num in range(1, 4):
            # Antraštės (5 eilutė)
            cell_header = ws.cell(row=5, column=col_num)
            cell_header.alignment = alignment
            cell_header.border = thin_border
            cell_header.font = bold_font
            
            # Įvedimo langeliai (6 eilutė)
            cell_input = ws.cell(row=6, column=col_num)
            cell_input.alignment = alignment
            cell_input.border = thin_border

    print(f"Sėkmingai sukurtas šablonas: {failo_pavadinimas}")

def nuskaityti_ir_perkelti_duomenis(kaminas_obj):
    """Nuskaito ranka įvestus duomenis iš Pradiniai.xlsx ir perkelia į Rezultatai.xlsx."""
    failo_pradiniai = "Pradiniai.xlsx"
    failo_rezultatai = "Rezultatai.xlsx"
    
    # 1. NUSKAITYMAS: Pasiimame duomenis iš Excel į Python atmintį
    # header=4 nurodo, kad stulpelių pavadinimai yra 5-oje eilutėje
    df = pd.read_excel(failo_pradiniai, sheet_name="Pradiniai", header=4)
    
    if df.empty:
        print("Klaida: Nerasta duomenų faile Pradiniai.xlsx!")
        return

    # 2. ŽODYNAS: Sudedame duomenis į "krepšelį" (dictionary)
    # iloc[0, x] paima pirmąją duomenų eilutę (tą, kurią jūs užpildėte)
   # Pasiimame datą ir ją sutvarkome
    zalia_data = df.iloc[0, 0]
    
    # Jei tai datos objektas, paverčiame į tekstą YYYY-MM-DD
    if pd.api.types.is_datetime64_any_dtype(zalia_data) or not isinstance(zalia_data, str):
        tikra_data = pd.to_datetime(zalia_data).strftime('%Y-%m-%d')
    else:
        tikra_data = zalia_data

    duomenys = {
        'data': tikra_data,  # Naudojame sutvarkytą datą
        'reg_nr': df.iloc[0, 1],
        'objektas': df.iloc[0, 2]
    }

    # 3. ĮRAŠYMAS: Atidarome Rezultatai.xlsx ir įrašome į konkrečias vietas [cite: 2026-03-03]
    wb = load_workbook(failo_rezultatai)
    ws = wb["Greitis"]
    
    centravimas = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Įrašome reikšmes į langelius A6, B6, C6
    ws["A6"] = duomenys['data']
    ws["B6"] = duomenys['reg_nr']
    ws["C6"] = duomenys['objektas']

    # Suformatuojame tuos langelius
    for coord in ["A6", "B6", "C6"]:
        ws[coord].alignment = centravimas

    # --- PERKĖLIMAS Į LAPĄ "PAĖMIMAS" ---
    ws_paemimas = wb["Paėmimas"]
    tasku_sk = kaminas_obj.tasku_skaicius
    
    # Apibrėžiame stulpelius ir jų atitinkamas reikšmes
    perkeliami_duomenys = {
        "A": duomenys['data'],
        "B": duomenys['reg_nr'],
        "C": duomenys['objektas']
    }

    for col_let, reiksme in perkeliami_duomenys.items():
        # 1. Įrašome reikšmę į 6-ąją eilutę
        cell = ws_paemimas[f"{col_let}6"]
        cell.value = reiksme
        cell.alignment = centravimas
        
        # 2. Jei taškų skaičius > 1, suliejame langelius nuo 6 iki (6 + taškų_sk - 1)
        if tasku_sk > 1:
            pabaigos_eilute = 6 + tasku_sk - 1
            merge_range = f"{col_let}6:{col_let}{pabaigos_eilute}"
            
            # Prieš suliejant įsitikiname, kad stilius bus pritaikytas visam rėžiui
            ws_paemimas.merge_cells(merge_range)
            
            # Openpyxl centre-alignment po suliejimo reikalauja, kad viršutinis kairysis langelis turėtų stilių
            ws_paemimas[f"{col_let}6"].alignment = centravimas

    # Išsaugome Rezultatai.xlsx failą
    wb.save(failo_rezultatai)
    print(f"Duomenys perkelti: {duomenys['data']}, {duomenys['reg_nr']}")