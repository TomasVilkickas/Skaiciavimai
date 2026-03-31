import os
import win32com.client as win32
import pandas as pd
from openpyxl import load_workbook
import psutil

def uzdaryti_excel_procesus():
    """Priverstinai išjungia visus fone veikiančius Excel procesus."""
    for proc in psutil.process_iter(['name']):
        try:
            if proc.info['name'] == "EXCEL.EXE":
                proc.kill()
                print("[INFO] Fone veikęs Excel procesas buvo išjungtas.")
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass

def optimizuoti_izokinetiskuma(kaminas):
    # SVARBU: Išvalome viską prieš pradedant
    uzdaryti_excel_procesus()
    
    originalus_failas = 'Rezultatai.xlsx'
    pilnas_kelias = os.path.abspath(originalus_failas)
    
    if not os.path.exists(pilnas_kelias):
        print(f"[KLAIDA] Nerastas failas: {pilnas_kelias}")
        return

    # 1. NAUDOJAME openpyxl TIK EILUČIŲ GAVIMUI IR IŠKART UŽDAROME
    wb_oxl = load_workbook(originalus_failas, read_only=True, data_only=False)
    start_row = 6
    matu_eilutes = gauti_matu_eilutes(start_row, kaminas)
    wb_oxl.close() # <--- BŪTINA UŽDARYTI ČIA, kad neatimtų teisių iš Excel programos

    # 2. ATIDAROME EXCEL PROGRAMĄ
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False 

    try:
        wb_excel = excel.Workbooks.Open(pilnas_kelias)
        ws_p_excel = wb_excel.Worksheets('Paėmimas')
        ws_a_excel = wb_excel.Worksheets('Aerodinamika')

        geriausias_da = None
        geriausi_vo = {}
        geriausias_max_nuokrypis = float('inf')
        geriausios_izo_reiksmes = {}

        print("\n--- Pradedama izokinetiškumo optimizacija ---")

        for da in range(4, 15): 
            vo_sprendinys = {}
            izo_reiksmes_siam_da = {}
            max_nuokrypis_siam_da = 0.0
            visi_tinka = True

            for row in matu_eilutes:
                einamasis_geriausias_vo = None
                einamasis_maziausias_nuokrypis = float('inf')
                einamoji_izo = None

                for vo in range(10, 17): 
                    ws_p_excel.Cells(row, 11).Value = da
                    ws_p_excel.Cells(row, 14).Value = vo
                    
                    langelio_val = ws_a_excel.Cells(row, 11).Value
                    izo = pd.to_numeric(langelio_val, errors='coerce')

                    if pd.isna(izo):
                        continue

                    nuokrypis = abs(1.0 - izo)

                    if nuokrypis < einamasis_maziausias_nuokrypis:
                        einamasis_maziausias_nuokrypis = nuokrypis
                        einamasis_geriausias_vo = vo
                        einamoji_izo = izo

                if einamasis_geriausias_vo is None:
                    visi_tinka = False
                    break

                vo_sprendinys[row] = einamasis_geriausias_vo
                izo_reiksmes_siam_da[row] = einamoji_izo
                max_nuokrypis_siam_da = max(max_nuokrypis_siam_da, einamasis_maziausias_nuokrypis)

            if visi_tinka and max_nuokrypis_siam_da < geriausias_max_nuokrypis:
                geriausias_max_nuokrypis = max_nuokrypis_siam_da
                geriausias_da = da
                geriausi_vo = vo_sprendinys.copy()
                geriausios_izo_reiksmes = izo_reiksmes_siam_da.copy()
                print(f"Rastas geresnis variantas: DA={da}mm, Max nuokrypis={geriausias_max_nuokrypis:.4f}")

        if geriausias_da is not None:
            print(f"\n[SĖKMĖ] Rezultatai rasti!")
            for row in matu_eilutes:
                v_opt = geriausi_vo[row]
                i_opt = geriausios_izo_reiksmes[row]
                ws_p_excel.Cells(row, 11).Value = geriausias_da
                ws_p_excel.Cells(row, 14).Value = v_opt
                print(f"Eilutė {row}: DA={geriausias_da}, Vo={v_opt}, Izo={i_opt:.4f}")

            # IŠSAUGOJIMAS
            wb_excel.Save()
            print("\n[INFO] Failas sėkmingai išsaugotas.")
        else:
            print("[KLAIDA] Nepavyko rasti tinkamų parametrų.")

    except Exception as e:
        print(f"[KRITINĖ KLAIDA] {e}")
    
    finally:
        if 'wb_excel' in locals():
            wb_excel.Close(SaveChanges=True)
        excel.Quit()

def gauti_matu_eilutes(start_row, kaminas):
    eilutes = []
    blokai = kaminas.liniju_skaicius * kaminas.filtru_skaicius
    taskai = kaminas.tasku_skaicius
    for blokas in range(blokai):
        blokas_start = start_row + blokas * (taskai + 4)
        for t in range(taskai):
            eilutes.append(blokas_start + t)
    return eilutes