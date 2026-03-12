import pandas as pd
import os
import time
from MatavimoVieta import Kaminas
from AtstumaiFunkcija import vykdyti_atstumai
from Pavadinimai1Funkcija import irasyti_antrastes
from Pavadinimai2Funkcija import irasyti_antrastes2
from GreitisSPFunkcija import spalvinti_greitis
from SverimasSPFunkcija import spalvinti_sverimas

def sukurti_rezultatu_faila():
    failo_pavadinimas = "Rezultatai.xlsx"
    # Jei failas jau yra, mes jo nebeperrašome, kad neištrintume duomenų
    if os.path.exists(failo_pavadinimas):
        print(f"1. Failas {failo_pavadinimas} jau egzistuoja.")
        return

    time.sleep(2)

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
    spalvinti_greitis(Kaminas)
    spalvinti_sverimas(Kaminas)
    print("\n=== PROGRAMOS PABAIGA (Patikrinkite Rezultatai.xlsx) ===")

if __name__ == "__main__":
    pagrindine_programa()