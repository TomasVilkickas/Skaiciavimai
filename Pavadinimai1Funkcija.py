import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side
from MatavimoVieta import Kaminas

def irasyti_antrastes(kaminas_obj: Kaminas):
    failas = "Rezultatai.xlsx"
    n_lin = kaminas_obj.liniju_skaicius
    
    # 1. Grąžinti visi pilni jūsų nurodyti pavadinimai
    stulpeliai = [
        "Matavimo data",
        "Ėminių registracijos Nr. T-107-2026-E-",
        "Objekto pavadinimas, adresas, taršos šaltinio Nr.",
        "Matavimo taškai ortakyje nuo vidinės sienelės, cm"
    ]
    
    # Diferencinis slėgis (dinamiškas kiekis)
    for i in range(1, n_lin + 1):
        stulpeliai.append(f"Išmatuotas diferencinis slėgis taškuose {i} linija, Pdi, hPa")
        
    stulpeliai.extend([
        "Pito vamzdelio koeficientas, K", 
        "Temperatūra ortakyje t, oC", 
        "Atmosferinis slėgis P, hPa",
        "Statinis slėgis ortakyje ± ΔP, hPa", 
        "Dujų slėgis kamine Pk= P+ ΔP, hPa", 
        "Dujų mol. masė Ms= qn *22.4, kg/"
    ])
    
    # Srauto greitis (dinamiškas kiekis)
    for i in range(1, n_lin + 1):
        stulpeliai.append(f"Dujų srauto greitis matavimo taškuose, {i} linija, wi = K*129 *√t *√Pdi/ √Pk/√Ms, m/s")
        
    stulpeliai.extend([
        "Vidutinis dujų srauto greitis ortakyje w, m/s", 
        "Ortakio diametras, matmenys, m", 
        "Ortakio skerspjūvio plotas F, m2",
        "Dujų tūrio debitas realiomis sąlygomis Vk = wvid × F, m3/s", 
        "Dujų tūrio debitas normaliosiomis sąlygomis Vdr n.s =Vk *0,269*(P±ΔP)/(273+t), m3/s",
        "Vandens kondensato masė m H2O, kg", 
        "Vandens garų konc. dujose x=mH2O/Vmn*qn, kg/kg",
        "Sausų dujų tūrio debitas normaliosiomis sąlygomis Vn= Vdr n.s *((1/(1+(x*qn/0.8038)))",
        "Prasiurbtas dujų tūris normaliosiomis sąlygomis Vmn, Nm3", 
        "Išplėstinės neapibrėžties koef. A",
        "Išplėstinė neapibrėžtis (dujų tūrio debitas) Uv", 
        "Išplėstinė neapibrėžtis (dujų srauto greitis) Uw",
        "Sausų dujų tankis normaliosioms sąlygomis qn, (kg/m3)", 
        "Skaičiavimus atliko (inicialai, data)"
    ])

    wb = load_workbook(failas)
    ws = wb["Greitis"]
    
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))

    # 2. Antraštė 2-oje eilutėje (tik tekstas pirmame langelyje, be suliejimo)
    ws.cell(row=2, column=1, value="Debitas pagal matuotą diferencinį slėgį").font = Font(bold=True)

    # 3. Stulpelių formavimas (4 eilutė)
    for idx, pavadinimas in enumerate(stulpeliai, start=1):
        cell = ws.cell(row=4, column=idx, value=pavadinimas)
        cell.font = Font(bold=False) # Nuimtas paryškinimas
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        
        # Stulpelio plotis - parinktas 15, kad būtų kompaktiškiau, bet tilptų tekstas
        col_letter = cell.column_letter
        ws.column_dimensions[col_letter].width = 15

    # 4 eilutės aukštis (didiname, nes pavadinimai ilgi ir turi talpintis langelyje)
    ws.row_dimensions[4].height = 130 

    # 4. Matmenų įrašymas po stulpeliu "Ortakio diametras, matmenys, m"
    for idx, pavadinimas in enumerate(stulpeliai, start=1):
        if "Ortakio diametras, matmenys, m" in pavadinimas:
            if kaminas_obj.forma == 'A':
                reiksme = f"{(kaminas_obj.skersmuo / 100):.2f}".replace('.', ',')
                ws.cell(row=5, column=idx, value=reiksme).border = thin_border
            else:
                g_m = f"{(kaminas_obj.gylis / 100):.2f}".replace('.', ',')
                p_m = f"{(kaminas_obj.plotis / 100):.2f}".replace('.', ',')
                ws.cell(row=5, column=idx, value=g_m).border = thin_border
                ws.cell(row=6, column=idx, value=p_m).border = thin_border

    wb.save(failas)
    print("SĖKMINGA: Lentelė suformuota su pilnais pavadinimais, be suliejimų.")