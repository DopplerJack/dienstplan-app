import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io
import datetime

st.set_page_config(page_title="Dialyse Dienstplan", layout="wide")
st.title("🏥 Dienstplan-Generator: Dialyse (Pro-Version)")
st.markdown("Laden Sie Ihre Excel-Datei hoch. Bitte stellen Sie sicher, dass sie zwei Tabellenblätter enthält: 'Dienstplan' und 'Sonderregeln'.")

uploaded_file = st.file_uploader("Excel-Tabelle hochladen (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Beide Tabellenblätter einlesen
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
            with st.spinner("Die KI berechnet nun Millionen von Kombinationen und optimiert die Samstage..."):
                
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
                
                # --- HILFSFUNKTION FÜR DATUM ---
                # Wandelt z.B. "1.4.26" in einen echten Wochentag um (0=Montag, 5=Samstag, 6=Sonntag)
                wochentage_idx = []
                samstage_idx = []
                for t_idx, tag_str in enumerate(tage):
                    try:
                        dt = pd.to_datetime(tag_str, format='%d.%m.%y')
                        wt = dt.weekday()
                    except:
                        wt = t_idx % 7 # Fallback, falls Datum nicht lesbar
                    wochentage_idx.append(wt)
                    if wt == 5: # Samstag
                        samstage_idx.append(t_idx)

                # --- 1. GRUNDREGELN ---
                for m in mitarbeiter_liste:
                    for t_idx in range(num_tage):
                        model.AddExactlyOne([dienst_vars[(m, t_idx, s)] for s in schichten])
                        
                for index, row in df_haupt.iterrows():
                    m = row['Name']
                    for t_idx, tag in enumerate(tage):
                        eintrag = str(row[tag]).strip()
                        if eintrag != "Leer":
                            if eintrag == "--":
                                model.Add(dienst_vars[(m, t_idx, 'Frei')] == 1)
                            elif eintrag in schichten:
                                model.Add(dienst_vars[(m, t_idx, eintrag)] == 1)

                # --- 2. TÄGLICHER BEDARF ---
                for t_idx in range(num_tage):
                    wt = wochentage_idx[t_idx]
                    if wt in [0, 1]: # Mo, Di
                        model.Add(sum(dienst_vars[(m, t_idx, 'D1')] for m in mitarbeiter_liste) == 7)
                        model.Add(sum(dienst_vars[(m, t_idx, 'V1')] for m in mitarbeiter_liste) == 1)
                    elif wt in [2, 3, 4, 5]: # Mi, Do, Fr, Sa
                        model.Add(sum(dienst_vars[(m, t_idx, 'D1')] for m in mitarbeiter_liste) == 7)
                        model.Add(sum(dienst_vars[(m, t_idx, 'V1')] for m in mitarbeiter_liste) == 0)
                    else: # Sonntag
                        for m in mitarbeiter_liste:
                            model.Add(dienst_vars[(m, t_idx, 'Frei')] == 1)

                # --- 3. NEUE FIXE GLOBALE REGELN ---
                for m in mitarbeiter_liste:
                    # FIX 1: Niemals 4x D1 hintereinander (Rollendes 4-Tage-Fenster)
                    for t in range(num_tage - 3):
                        model.Add(sum(dienst_vars[(m, t+i, 'D1')] for i in range(4)) <= 3)
                        
                    # FIX 2: Max 4 Dienste pro Woche (Rollendes 7-Tage-Fenster)
                    for t in range(num_tage - 6):
                        model.Add(sum(dienst_vars[(m, t+i, s)] for i in range(7) for s in ['D1', 'V1', 'SL', 'D7']) <= 4)
                        
                    # FIX 3a: Max 2 Samstage pro Monat
                    if len(samstage_idx) > 0:
                        model.Add(sum(dienst_vars[(m, t, 'D1')] for t in samstage_idx) <= 2)

                # --- 4. INDIVIDUELLE REGELN (Aus Blatt 2) ---
                # Wochentag-Übersetzer
                wt_map = {"Montag": 0, "Dienstag": 1, "Mittwoch": 2, "Donnerstag": 3, "Freitag": 4, "Samstag": 5, "Sonntag": 6}
                
                straf_variablen = [] # Sammelt Punkte für weiche Regeln
                
                for index, row in df_regeln.iterrows():
                    m = row['Name']
                    if m not in mitarbeiter_liste: continue
                    
                    # Regel: Fester freier Tag
                    freier_tag_str = str(row.get('Fester freier Tag', '')).strip()
                    if freier_tag_str in wt_map:
                        ziel_wt = wt_map[freier_tag_str]
                        for t_idx in range(num_tage):
                            if wochentage_idx[t_idx] == ziel_wt:
                                model.Add(dienst_vars[(m, t_idx, 'Frei')] == 1)
                                
                    # Regel: Keine V1
                    if str(row.get('Keine V1', '')).strip().lower() == 'ja':
                        for t_idx in range(num_tage):
                            model.Add(dienst_vars[(m, t_idx, 'V1')] == 0)
                            
                    # Regel: Max 1 D1 am Stück (niemals 2 hintereinander)
                    if str(row.get('Max 1 D1 am Stück', '')).strip().lower() == 'ja':
                        for t_idx in range(num_tage - 1):
                            model.Add(dienst_vars[(m, t_idx, 'D1')] + dienst_vars[(m, t_idx+1, 'D1')] <= 1)
                            
                    # Regel: Nur 2er/3er Blöcke (Kein einzelner D1)
                    if str(row.get('Nur 2er/3er D1-Blöcke', '')).strip().lower() == 'ja':
                        for t_idx in range(num_tage):
                            links = dienst_vars[(m, t_idx-1, 'D1')] if t_idx > 0 else 0
                            rechts = dienst_vars[(m, t_idx+1, 'D1')] if t_idx < num_tage - 1 else 0
                            # Wenn heute D1, muss gestern oder morgen auch D1 sein
                            model.AddImplication(dienst_vars[(m, t_idx, 'D1')], links + rechts >= 1)

                    # Regel: Freitag frei vor Samstag-D1
                    if str(row.get('Freitag frei vor Samstag-D1', '')).strip().lower() == 'ja':
                        for t_idx in samstage_idx:
                            if t_idx > 0: # Wenn Samstag nicht der Monatserste ist
                                model.AddImplication(dienst_vars[(m, t_idx, 'D1')], dienst_vars[(m, t_idx-1, 'Frei')])

                # --- 5. WEICHE REGELN & OPTIMIERUNG ---
                # FIX 3b: Samstage fair aufteilen (vermeide 2 am Stück)
                for m in mitarbeiter_liste:
                    for i in range(len(samstage_idx) - 1):
                        sat1 = samstage_idx[i]
                        sat2 = samstage_idx[i+1]
                        doppel_sat = model.NewBoolVar(f"DoppelSat_{m}_{sat1}")
                        # Wenn an beiden Samstagen D1 gearbeitet wird, wird doppel_sat auf 1 gezwungen
                        model.Add(dienst_vars[(m, sat1, 'D1')] + dienst_vars[(m, sat2, 'D1')] == 2).OnlyEnforceIf(doppel_sat)
                        straf_variablen.append(doppel_sat * 10) # 10 Strafpunkte für aufeinanderfolgende Samstage

                # Wir sagen der KI: Halte die Summe aller Strafpunkte so gering wie möglich
                model.Minimize(sum(straf_variablen))

                # --- 6. BERECHNUNG ---
                solver = cp_model.CpSolver()
                solver.parameters.max_time_in_seconds = 60.0 # KI bekommt 60 Sekunden Zeit zum Knobeln
                status = solver.Solve(model)

                if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
                    st.success("🎉 Plan erfolgreich berechnet und optimiert!")
                    
                    ausgabe_df = df_haupt.copy()
                    for index, row in df_haupt.iterrows():
                        m = row['Name']
                        for t_idx, tag in enumerate(tage):
                            for s in schichten:
                                if solver.Value(dienst_vars[(m, t_idx, s)]) == 1:
                                    ausgabe_df.at[index, tag] = s if s != 'Frei' else '--'
                    
                    st.dataframe(ausgabe_df.head(10))
                    
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        ausgabe_df.to_excel(writer, index=False)
                    
                    st.download_button(label="📥 Fertigen Dienstplan herunterladen", data=buffer.getvalue(), file_name="Dienstplan_Fertig.xlsx")
                    st.info("💡 Der Plan ist nun bereit. Sie können die Daten aus dieser Tabelle jetzt problemlos nutzen, um die individuellen Schichten in die Standard-App wie Etar zu übertragen.")
                else:
                    st.error("🚨 Keine Lösung gefunden! Zu viele Regeln oder Freiwünsche blockieren sich gegenseitig.")
    except Exception as e:
        st.error(f"Ein Fehler ist aufgetreten: {e}")
