import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io
import datetime
from openpyxl.styles import PatternFill

st.set_page_config(page_title="Dialyse Dienstplan", layout="wide")
st.title("🏥 Dienstplan-Generator: Dialyse (Stunden-Optimiert)")
st.markdown("Laden Sie Ihre Excel-Datei hoch. Die KI erstellt den Plan, achtet strikt auf die +/- 25h Grenze und verteilt die Stunden fair.")

uploaded_file = st.file_uploader("Excel-Tabelle hochladen (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_haupt = pd.read_excel(uploaded_file, sheet_name=0)
        df_haupt.fillna("Leer", inplace=True)
        
        try:
            df_regeln = pd.read_excel(uploaded_file, sheet_name=1)
            df_regeln.fillna("", inplace=True)
        except:
            st.error("Fehler: Konnte das zweite Tabellenblatt 'Sonderregeln' nicht finden. Bitte legen Sie es an.")
            st.stop()
            
        st.success("Tabelle und Sonderregeln erfolgreich eingelesen!")
        
        if st.button("Dienstplan jetzt berechnen"):
            with st.spinner("Die KI balanciert nun die Stundenkonten aus... Dies kann einen Moment dauern."):
                
                model = cp_model.CpModel()
                
                mitarbeiter_liste = df_haupt['Name'].tolist()
                tage = [col for col in df_haupt.columns if col not in ['Name', 'Stundenausmaß', 'Berechnung Soll-Arbeitszeit', 'Übertrag Vormonat']]
                num_tage = len(tage)
                schichten = ['D1', 'V1', 'SL', 'D7', 'Frei', 'U', 'ZB', 'ÜZA', 'BA', 'FB']
                
                dienst_vars = {}
                for m in mitarbeiter_liste:
                    for t_idx in range(num_tage):
                        for s in schichten:
                            dienst_vars[(m, t_idx, s)] = model.NewBoolVar(f"{m}_{t_idx}_{s}")
                
                # --- 0. ZAHLEN SICHER EINLESEN ---
                def parse_num(val):
                    try:
                        if pd.isna(val) or str(val).strip() == "Leer" or str(val).strip() == "": return 0.0
                        return float(str(val).replace(',', '.').replace('+', '').strip())
                    except:
                        return 0.0

                ma_daten = {}
                for index, row in df_haupt.iterrows():
                    m = row['Name']
                    ma_daten[m] = {
                        'ausmass_int': int(parse_num(row.get('Stundenausmaß', 0)) * 100),
                        'soll_int': int(parse_num(row.get('Berechnung Soll-Arbeitszeit', 0)) * 100),
                        'uebertrag_int': int(parse_num(row.get('Übertrag Vormonat', 0)) * 100)
                    }

                # --- 1. ROBUSTE WOCHENTAGS-ERKENNUNG ---
                wochentage_idx = []
                samstage_idx = []
                for t_idx, tag_val in enumerate(tage):
                    if isinstance(tag_val, datetime.datetime):
                        wt = tag_val.weekday()
                    else:
                        try:
                            dt = pd.to_datetime(str(tag_val), dayfirst=True)
                            wt = dt.weekday()
                        except:
                            wt = t_idx % 7 
                            
                    wochentage_idx.append(wt)
                    if wt == 5: 
                        samstage_idx.append(t_idx)

                # --- 2. GRUNDREGELN & EXCEL-INPUT ---
                feste_eintraege = {} 
                
                for m in mitarbeiter_liste:
                    for t_idx in range(num_tage):
                        model.AddExactlyOne([dienst_vars[(m, t_idx, s)] for s in schichten])
                        
                for index, row in df_haupt.iterrows():
                    m = row['Name']
                    for t_idx, tag in enumerate(tage):
                        if wochentage_idx[t_idx] == 6: 
                            continue
                            
                        eintrag = str(row[tag]).strip()
                        if eintrag != "Leer":
                            feste_eintraege[(m, t_idx)] = eintrag 
                            if eintrag == "F": 
                                model.Add(dienst_vars[(m, t_idx, 'Frei')] == 1)
                            elif eintrag in schichten:
                                model.Add(dienst_vars[(m, t_idx, eintrag)] == 1)

                # --- 3. DAS ÜBERDRUCK-VENTIL ---
                straf_variablen = []
                fehlende_D1_vars = {}
                fehlende_V1_vars = {}
                
                for t_idx in range(num_tage):
                    wt = wochentage_idx[t_idx]
                    if wt == 6: 
                        for m in mitarbeiter_liste:
                            model.Add(dienst_vars[(m, t_idx, 'Frei')] == 1)
                        continue
                        
                    fehlend_d1 = model.NewIntVar(0, 7, f'fehlend_D1_{t_idx}')
                    fehlende_D1_vars[t_idx] = fehlend_d1
                    straf_variablen.append(fehlend_d1 * 1000000)
                    
                    if wt in [0, 1]: 
                        model.Add(sum(dienst_vars[(m, t_idx, 'D1')] for m in mitarbeiter_liste) + fehlend_d1 == 7)
                        
                        fehlend_v1 = model.NewIntVar(0, 1, f'fehlend_V1_{t_idx}')
                        fehlende_V1_vars[t_idx] = fehlend_v1
                        straf_variablen.append(fehlend_v1 * 1000000)
                        model.Add(sum(dienst_vars[(m, t_idx, 'V1')] for m in mitarbeiter_liste) + fehlend_v1 == 1)
                    else: 
                        model.Add(sum(dienst_vars[(m, t_idx, 'D1')] for m in mitarbeiter_liste) + fehlend_d1 == 7)
                        model.Add(sum(dienst_vars[(m, t_idx, 'V1')] for m in mitarbeiter_liste) == 0)

                # --- 4. FIXE GLOBALE REGELN ---
                # Identifiziere alle Montage für die Wochen-Regel
                mondays = [t for t, wt in enumerate(wochentage_idx) if wt == 0]
                
                for m in mitarbeiter_liste:
                    for t in range(num_tage - 3):
                        model.Add(sum(dienst_vars[(m, t+i, 'D1')] for i in range(4)) <= 3)
                        
                    for t in range(num_tage - 6):
                        model.Add(sum(dienst_vars[(m, t+i, s)] for i in range(7) for s in ['D1', 'V1', 'SL', 'D7']) <= 4)
                        
                    model.Add(sum(dienst_vars[(m, t, 'V1')] for t in range(num_tage)) <= 1)
                        
                    if len(samstage_idx) > 0:
                        model.Add(sum(dienst_vars[(m, t, 'D1')] for t in samstage_idx) <= 3) 
                        dritter_sat = model.NewBoolVar(f"DritterSat_{m}")
                        model.Add(sum(dienst_vars[(m, t, 'D1')] for t in samstage_idx) <= 2 + dritter_sat)
                        straf_variablen.append(dritter_sat * 50000) 

                    # NEU: Zwingend 2 aufeinanderfolgende freie Tage pro Kalenderwoche (Mo-So)
                    for mo in mondays:
                        if mo + 6 < num_tage: # Nur für vollständige Wochen im Monat anwenden
                            week_pairs = []
                            for i in range(6):
                                t1 = mo + i
                                t2 = mo + i + 1
                                
                                work1 = model.NewBoolVar(f"work_{m}_{t1}")
                                model.Add(work1 == sum(dienst_vars[(m, t1, s)] for s in ['D1', 'V1', 'SL', 'D7']))
                                
                                work2 = model.NewBoolVar(f"work_{m}_{t2}")
                                model.Add(work2 == sum(dienst_vars[(m, t2, s)] for s in ['D1', 'V1', 'SL', 'D7']))
                                
                                pair_off = model.NewBoolVar(f"pairoff_{m}_{t1}")
                                model.Add(work1 + work2 == 0).OnlyEnforceIf(pair_off)
                                model.Add(work1 + work2 > 0).OnlyEnforceIf(pair_off.Not())
                                
                                week_pairs.append(pair_off)
                            
                            # Es muss mindestens ein solches Paar in dieser Woche geben
                            model.AddBoolOr(week_pairs)

                # --- 5. INDIVIDUELLE REGELN ---
                wt_map = {"Montag": 0, "Dienstag": 1, "Mittwoch": 2, "Donnerstag": 3, "Freitag": 4, "Samstag": 5, "Sonntag": 6}
                
                for index, row in df_regeln.iterrows():
                    m = row['Name']
                    if m not in mitarbeiter_liste: continue
                    
                    freier_tag_str = str(row.get('Fester freier Tag', '')).strip()
                    if freier_tag_str:
                        feste_freie_tage = [t.strip() for t in freier_tag_str.split(',')]
                        for tag_name in feste_freie_tage:
                            if tag_name in wt_map:
                                ziel_wt = wt_map[tag_name]
                                for t_idx in range(num_tage):
                                    if wochentage_idx[t_idx] == ziel_wt:
                                        if (m, t_idx) not in feste_eintraege:
                                            model.Add(dienst_vars[(m, t_idx, 'Frei')] == 1)

                    fester_dienst_str = str(row.get('Fester Tag Dienst', '')).strip()
                    if fester_dienst_str:
                        feste_dienst_tage = [t.strip() for t in fester_dienst_str.split(',')]
                        for tag_name in feste_dienst_tage:
                            if tag_name in wt_map:
                                ziel_wt = wt_map[tag_name]
                                for t_idx in range(num_tage):
                                    if wochentage_idx[t_idx] == ziel_wt:
                                        if (m, t_idx) not in feste_eintraege:
                                            model.Add(dienst_vars[(m, t_idx, 'D1')] == 1)
                                
                    if str(row.get('Keine V1', '')).strip().lower() == 'ja':
                        for t_idx in range(num_tage):
                            model.Add(dienst_vars[(m, t_idx, 'V1')] == 0)

                    if str(row.get('Immer ein V1-Dienst', '')).strip().lower() == 'ja':
                        model.Add(sum(dienst_vars[(m, t_idx, 'V1')] for t_idx in range(num_tage)) == 1)
                            
                    if str(row.get('Max 1 D1 am Stück', '')).strip().lower
