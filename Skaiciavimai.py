import pandas as pd
import os
import time
from MatavimoVieta import Kaminas
from AtstumaiFunkcija import vykdyti_atstumai
from Pavadinimai1Funkcija import irasyti_antrastes
from Pavadinimai2Funkcija import irasyti_antrastes2
from Pradiniai1Funkcija import sukurti_sablona, nuskaityti_ir_perkelti_duomenis
from Pradiniai2Funkcija import nuskaityti_ir_perkelti_Greitis, paruosti_H2O_lapa, perkelti_greitis_duomenis, perkelti_H2O_duomenis
from Pradiniai3Funkcija import nuskaityti_ir_perkelti_Paemimas, perkelti_paemimas_duomenis 
from Pradiniai4Funkcija import sukurti_Paemimas_komplektus, perkelti_paemimas_komplektus
from H2OSK1Funkcija import skaiciuoti_H2O1
from PaemimasSK1Funkcija import skaiciuoti_paemimas1
from AerodinamikaSK1Funkcija import skaiciuoti_aerodinamika1
from H2OSK2Funkcija import skaiciuoti_H2O2
from H2OSK3Funkcija import skaiciuoti_H2O3
from AerodinamikaSK2Funkcija import skaiciuoti_aerodinamika2
from GreitisSK1Funkcija import skaiciuoti_greitis1
from GreitisSPFunkcija import spalvinti_greitis
from H2OSPFunkcija import spalvinti_H2O
from PaemimasSP1Funkcija import spalvinti_paemimas1
from AerodinamikaSP1Funkcija import spalvinti_aerodinamika1
from KoncentracijaSPFunkcija import spalvinti_koncentracija
from SverimasSPFunkcija import spalvinti_sverimas

def sukurti_rezultatu_faila():
    failo_pavadinimas = "Rezultatai.xlsx"
    # Jei failas jau yra, mes jo nebeperrašome, kad neištrintume duomenų
    if os.path.exists(failo_pavadinimas):
        print(f"1. Failas {failo_pavadinimas} jau egzistuoja.")
        return
     
    lapai = ["Greitis", "H2O", "Paėmimas", "Aerodinamika", "Koncentracija", "Svėrimas", "Koncentracijos ribinės vertės"]
    with pd.ExcelWriter(failo_pavadinimas, engine='openpyxl') as writer:
        for lapas in lapai:
            pd.DataFrame().to_excel(writer, sheet_name=lapas, index=False)
    print(f"1. Sukurtas naujas failas: {failo_pavadinimas}")

def vykdyti_apklausa():
    print("\n2. --- Matavimo vietos parametrų įvedimas ---")
    Kaminas.forma = input("Kokia matavimo vietos skerspjūvio forma (Apvalus/Stačiakampis A/S): ").upper()
    if Kaminas.forma == 'A':
        Kaminas.skersmuo = int(input("Koks ortakio skersmuo, cm: "))
        Kaminas.gylis = Kaminas.skersmuo
        Kaminas.liniju_skaicius = int(input("Kiek matavimo linijų (1 arba 2): "))
    else:
        Kaminas.gylis = int(input("Koks ortakio gylis, cm: "))
        Kaminas.plotis = int(input("Koks ortakio plotis, cm: "))
        Kaminas.liniju_skaicius = int(input("Kiek matavimo linijų (1-5): "))
    Kaminas.filtru_skaicius = int(input("Kiek filtrų (1, 2, 3): "))

def pagrindine_programa():
    sukurti_rezultatu_faila()
    vykdyti_apklausa()
    # SVARBU: vykdyti_atstumai turi wb.save("Rezultatai.xlsx") viduje
    vykdyti_atstumai(Kaminas) 
    irasyti_antrastes(Kaminas)
    irasyti_antrastes2(Kaminas)
    
    # 2. ĮTERPIAME PAUZĘ IR PRADINIŲ DUOMENŲ PILDYMĄ ČIA
    print("\n" + "="*60)
    print(">>> ANTRAŠTĖS SUFORMUOTOS. DABAR RUOŠIAMAS 'Pradiniai.xlsx' <<<")
    
    # Sukuriame šabloną (jei reikia, su patikra, kad neperrašytų esamo)
    sukurti_sablona(Kaminas)
    nuskaityti_ir_perkelti_Greitis(Kaminas)
    paruosti_H2O_lapa()
    nuskaityti_ir_perkelti_Paemimas(Kaminas)
    sukurti_Paemimas_komplektus(Kaminas)
    zodynas = sukurti_Paemimas_komplektus(Kaminas)
    
    # Atidarome failą vartotojui
    failas_pradiniai = "Pradiniai.xlsx"
    if os.path.exists(failas_pradiniai):
        os.startfile(failas_pradiniai)
        print(f"Failas '{failas_pradiniai}' atidarytas. Užpildykite ir išsaugokite.")
    
    print("="*60)
    
    # Stabdomas Python procesas
    input("\nKai užbaigsite pildyti 'Pradiniai.xlsx', paspauskite ENTER čia...")
    nuskaityti_ir_perkelti_duomenis(Kaminas)
    perkelti_greitis_duomenis(Kaminas)
    perkelti_H2O_duomenis()
    perkelti_paemimas_duomenis(Kaminas)
    perkelti_paemimas_komplektus(Kaminas, zodynas)

    # 3. TĘSIAME TOLIAU (Spalvinimas ir kiti skaičiavimai)
    print("\nTęsiama programa: spalvinami lapai ir baigiami skaičiavimai...")

    skaiciuoti_H2O1()
    skaiciuoti_paemimas1(Kaminas)
    skaiciuoti_aerodinamika1(Kaminas)
    skaiciuoti_H2O2(Kaminas)
    skaiciuoti_H2O3(Kaminas)
    skaiciuoti_aerodinamika2(Kaminas)
    skaiciuoti_greitis1(Kaminas)
    spalvinti_greitis(Kaminas)
    spalvinti_H2O()
    spalvinti_paemimas1(Kaminas)
    spalvinti_aerodinamika1(Kaminas)
    spalvinti_koncentracija(Kaminas)
    spalvinti_sverimas(Kaminas)
    print("\n=== PROGRAMOS PABAIGA (Patikrinkite Rezultatai.xlsx) ===")

if __name__ == "__main__":
    pagrindine_programa()