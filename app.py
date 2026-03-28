import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io
import datetime

st.set_page_config(page_title="Dialyse Dienstplan", layout="wide")
st.title("🏥 Dienstplan-Generator: Dialyse (Diagnose-Modus)")
st.markdown("Laden Sie Ihre Excel-Datei hoch. Die KI erstellt den Plan auch bei Unterbesetzung und zeigt Engpässe an.")

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
            with st.spinner("Die KI jongliert nun mit dem Personal und sucht nach Engpässen..."):
                
                model = cp_model.CpModel()
                
                mitarbeiter_liste = df_haupt['Name'].tolist()
                tage = [col for col in df_haupt.columns if col not in ['Name', 'Stundenausmaß', 'Berechnung Soll-Arbeitszeit', 'Übertrag Vormonat']]
                num_tage = len(tage)
                schichten = ['D1', 'V1', 'SL', 'D7', 'Frei', 'U', 'ZB', 'ÜZA', 'BA', 'Fobi']
                
                dienst_vars = {}
                for m in mitarbeiter_liste:
                    for t_idx in range(num_tage):
                        for s in schichten:
                            dienst_vars[(m, t_idx, s)] = model.NewBoolVar(f"{m}_{t_idx}_{s}")
                
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
                        if wochentage_idx[t_idx] == 6: # Sonntag ignorieren
                            continue
                            
                        eintrag = str(row[tag]).strip()
                        if eintrag != "Leer":
                            feste_eintraege[(m, t_idx)] = eintrag 
                            if eintrag == "F": 
                                model.Add(dienst_vars[(m, t_idx, 'Frei')] == 1)
                            elif eintrag in schichten:
                                model.Add(dienst_vars[(m, t_idx, eintrag)] == 1)

                # --- 3. DAS ÜBERDRUCK-VENTIL (Flexible Bedarfsplanung) ---
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
                    straf_variablen.append(fehlend_d1 * 10000)
                    
                    if wt in [0, 1]: 
                        model.Add(sum(dienst_vars[(m, t_idx, 'D1')] for m in mitarbeiter_liste) + fehlend_d1 == 7)
                        
                        fehlend_v1 = model.NewIntVar(0, 1, f'fehlend_V1_{t_idx}')
                        fehlende_V1_vars[t_idx] = fehlend_v1
                        straf_variablen.append(fehlend_v1 * 10000)
                        model.Add(sum(dienst_vars[(m, t_idx, 'V1')] for m in mitarbeiter_liste) + fehlend_v1 == 1)
                    else: 
                        model.Add(sum(dienst_vars[(m, t_idx, 'D1')] for m in mitarbeiter_liste) + fehlend_d1 == 7)
                        model.Add(sum(dienst_vars[(m, t_idx, 'V1')] for m in mitarbeiter_liste) == 0)

                # --- 4. FIXE GLOBALE REGELN ---
                for m in mitarbeiter_liste:
                    # Maximal 3x D1 am Stück
                    for t in range(num_tage - 3):
                        model.Add(sum(dienst_vars[(m, t+i, 'D1')] for i in range(4)) <= 3)
                        
                    # Maximal 4 Dienste pro Woche
                    for t in range(num_tage - 6):
                        model.Add(sum(dienst_vars[(m, t+i, s)] for i in range(7) for s in ['D1', 'V1', 'SL', 'D7']) <= 4)
                        
                    # NEU: Maximal 1 V1-Dienst pro Monat (Gerechte Verteilung)
                    model.Add(sum(dienst_vars[(m, t, 'V1')] for t in range(num_tage)) <= 1)
                        
                    # Samstags-Limit
                    if len(samstage_idx) > 0:
                        model.Add(sum(dienst_vars[(m, t, 'D1')] for t in samstage_idx) <= 3) 
                        dritter_sat = model.NewBoolVar(f"DritterSat_{m}")
                        model.Add(sum(dienst_vars[(m, t, 'D1')] for t in samstage_idx) <= 2 + dritter_sat)
                        straf_variablen.append(dritter_sat * 50) 

                # --- 5. INDIVIDUELLE REGELN ---
                wt_map = {"Montag": 0, "Dienstag": 1, "Mittwoch": 2, "Donnerstag": 3, "Freitag": 4, "Samstag": 5, "Sonntag": 6}
                
                for index, row in df_regeln.iterrows():
                    m = row['Name']
                    if m not in mitarbeiter_liste: continue
                    
                    # Fester freier Tag
                    freier_tag_str = str(row.get('Fester freier Tag', '')).strip()
                    if freier_tag_str in wt_map:
                        ziel_wt = wt_map[freier_tag_str]
                        for t_idx in range(num_tage):
                            if wochentage_idx[t_idx] == ziel_wt:
                                if (m, t_idx) not in feste_eintraege:
                                    model.Add(dienst_vars[(m, t_idx, 'Frei')] == 1)
                                
                    # Keine V1
                    if str(row.get('Keine V1', '')).strip().lower() == 'ja':
                        for t_idx in range(num_tage):
                            model.Add(dienst_vars[(m, t_idx, 'V1')] == 0)

                    # NEU: Immer ein V1-Dienst
                    if str(row.get('Immer ein V1-Dienst', '')).strip().lower() == 'ja':
                        model.Add(sum(dienst_vars[(m, t_idx, 'V1')] for t_idx in range(num_tage)) == 1)
                            
                    # Max 1 D1 am Stück
                    if str(row.get('Max 1 D1 am Stück', '')).strip().lower() == 'ja':
                        for t_idx in range(num_tage - 1):
                            model.Add(dienst_vars[(m, t_idx, 'D1')] + dienst_vars[(m, t_idx+1, 'D1')] <= 1)

                    # Freitag frei vor Samstag-D1
                    if str(row.get('Freitag frei vor Samstag-D1', '')).strip().lower() == 'ja':
                        for t_idx in samstage_idx:
                            if t_idx > 0: 
                                model.AddImplication(dienst_vars[(m, t_idx, 'D1')], dienst_vars[(m, t_idx-1, 'Frei')])

                    # NEU: Keine Samstag/Montag Konstellation
                    if str(row.get('Keine Samstag/Montag Konstellation', '')).strip().lower() == 'ja':
                        for t_idx in samstage_idx:
                            if t_idx + 2 < num_tage: 
                                # Wenn am Samstag gearbeitet wird (D1), ist der Montag zwingend frei
                                model.AddImplication(dienst_vars[(m, t_idx, 'D1')], dienst_vars[(m, t_idx+2, 'Frei')])

                    # NEU: Bevorzuge 2er D1-Blöcke (Straft einzelne D1 und 3er-Blöcke ab)
                    if str(row.get('Bevorzuge 2er D1-Blöcke', '')).strip().lower() == 'ja':
                        for t_idx in range(num_tage):
                            d_heute = dienst_vars[(m, t_idx, 'D1')]
                            d_gestern = dienst_vars[(m, t_idx-1, 'D1')] if t_idx > 0 else 0
                            d_morgen = dienst_vars[(m, t_idx+1, 'D1')] if t_idx < num_tage - 1 else 0
                            
                            # Bestrafe isolierte Schichten
                            val_iso = model.NewIntVar(-2, 1, f"iso_{m}_{t_idx}")
                            model.Add(val_iso == d_heute - d_gestern - d_morgen)
                            is_iso = model.NewBoolVar(f"is_iso_bool_{m}_{t_idx}")
                            model.Add(val_iso == 1).OnlyEnforceIf(is_iso)
                            model.Add(val_iso != 1).OnlyEnforceIf(is_iso.Not())
                            straf_variablen.append(is_iso * 15)
                            
                            # Bestrafe 3er-Blöcke
                            val_3 = model.NewIntVar(0, 3, f"3blk_{m}_{t_idx}")
                            model.Add(val_3 == d_gestern + d_heute + d_morgen)
                            is_3blk = model.NewBoolVar(f"is_3blk_bool_{m}_{t_idx}")
                            model.Add(val_3 == 3).OnlyEnforceIf(is_3blk)
                            model.Add(val_3 != 3).OnlyEnforceIf(is_3blk.Not())
                            straf_variablen.append(is_3blk * 10)

                # --- 6. WEICHE REGELN & OPTIMIERUNG ---
                for m in mitarbeiter_liste:
                    for i in range(len(samstage_idx) - 1):
                        sat1 = samstage_idx[i]
                        sat2 = samstage_idx[i+1]
                        doppel_sat = model.NewBoolVar(f"DoppelSat_{m}_{sat1}")
                        model.Add(dienst_vars[(m, sat1, 'D1')] + dienst_vars[(m, sat2, 'D1')] == 2).OnlyEnforceIf(doppel_sat)
                        straf_variablen.append(doppel_sat * 10) 

                model.Minimize(sum(straf_variablen))

                # --- 7. BERECHNUNG ---
                solver = cp_model.CpSolver()
                solver.parameters.max_time_in_seconds = 60.0 
                status = solver.Solve(model)

                if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
                    
                    warnungen = []
                    for t_idx in fehlende_D1_vars:
                        fehlt_d1 = solver.Value(fehlende_D1_vars[t_idx])
                        if fehlt_d1 > 0:
                            datum = tage[t_idx]
                            if isinstance(datum, datetime.datetime):
                                datum = datum.strftime('%d.%m.%Y')
                            warnungen.append(f"⚠️ **{datum}**: Es fehlen **{fehlt_d1}** Mitarbeiter für den D1-Dienst.")
                            
                    for t_idx in fehlende_V1_vars:
                        fehlt_v1 = solver.Value(fehlende_V1_vars[t_idx])
                        if fehlt_v1 > 0:
                            datum = tage[t_idx]
                            if isinstance(datum, datetime.datetime):
                                datum = datum.strftime('%d.%m.%Y')
                            warnungen.append(f"⚠️ **{datum}**: Der V1-Dienst konnte nicht besetzt werden.")

                    if len(warnungen) > 0:
                        st.warning("Der Dienstplan wurde erstellt, aber es gibt personelle Engpässe! Die Station ist an folgenden Tagen unterbesetzt:")
                        for w in warnungen:
                            st.write(w)
                    else:
                        st.success("🎉 Plan erfolgreich berechnet! Die Station ist an allen Tagen zu 100 % besetzt.")

                    ausgabe_df = df_haupt.copy()
                    for index, row in df_haupt.iterrows():
                        m = row['Name']
                        for t_idx, tag in enumerate(tage):
                            for s in schichten:
                                if solver.Value(dienst_vars[(m, t_idx, s)]) == 1:
                                    ausgabe_df.at[index, tag] = s if s != 'Frei' else 'F'
                                    
                    st.dataframe(ausgabe_df.head(10))
                    
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        ausgabe_df.to_excel(writer, index=False)
                    
                    st.download_button(label="📥 Fertigen Dienstplan herunterladen", data=buffer.getvalue(), file_name="Dienstplan_Fertig.xlsx")
                else:
                    st.error("🚨 Kritischer Fehler! Die Mathematik widerspricht sich komplett.")
    except Exception as e:
        st.error(f"Ein Fehler ist aufgetreten: {e}")
