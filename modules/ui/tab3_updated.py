# Interfaccia utente per la Tab 3 (Calcolo Presenze ed Esportazione)
import streamlit as st
import pandas as pd
import re
from datetime import datetime, date
from io import BytesIO
from modules.attendance import calculate_attendance

# Definisco le funzioni di utilità direttamente qui per evitare problemi di importazione
def clean_sheet_name(name):
    """Pulisce i nomi dei fogli Excel"""
    name = re.sub(r'[\\/?*\[\]:]', '', str(name))
    return name[:31]
    
def extract_code_from_parentheses(text):
    """Estrae codici tra parentesi"""
    if not isinstance(text, str): return None
    match = re.search(r'\((.*?)\)', text)
    if match:
        code = match.group(1).strip()
        if code: return code
    return None

def extract_sort_key(percorso_str):
    """Estrai chiavi di ordinamento dai percorsi"""
    code_match = re.search(r'^\[([-\w]+)\]', str(percorso_str))
    if code_match:
        return code_match.group(1)
    return str(percorso_str)

def render_tab3(df_main):
    """Renderizza l'interfaccia della Tab 3: Calcolo Presenze ed Esportazione"""
    st.header("Calcolo Presenze ed Esportazione")
    st.write("Visualizza presenze per **Percorso (Senza Art.13)** e permette export dettagliato in Excel o CSV, con opzione multi-foglio Excel per ogni percorso unico.")
    
    if st.session_state.duplicates_removed: 
        st.success("Presenze calcolate sui dati depurati.")

    current_df_for_tab3 = df_main

    if not current_df_for_tab3.empty:
        required_att_cols = ['CodiceFiscale', 'PercorsoOriginaleSenzaArt13Internal', 'PercorsoInternal', 'PercorsoOriginaleInternal']
        if not all(col in current_df_for_tab3.columns for col in required_att_cols):
            missing_cols = [col for col in required_att_cols if col not in current_df_for_tab3.columns]
            st.error(f"Impossibile procedere: colonne mancanti ({', '.join(missing_cols)})")
            attendance_df = pd.DataFrame() # Resetta per sicurezza
        else:
            # Aggiungiamo le opzioni di visualizzazione
            st.subheader("Seleziona Tipo di Visualizzazione")
            view_options = [
                "1. Presenze per Studente e Percorso", 
                "2. Riassunto per Percorso Originale (PercorsoOriginaleSenzaArt13Internal)",
                "3. Riassunto per Percorso Elaborato (PercorsoInternal)",
                "4. Riassunto per Percorso Iscritti (da file iscritti_29_aprile.csv)",
                "5. Lista Completa Studenti"
            ]
            view_type = st.radio("Scegli cosa visualizzare:", view_options, key="view_type_tab3")
            
            # Determina il tipo di raggruppamento in base alla selezione
            if "1. Presenze per Studente" in view_type:
                group_by = "studente"
                attendance_df = calculate_attendance(current_df_for_tab3, group_by=group_by)
            elif "2. Riassunto per Percorso Originale" in view_type:
                group_by = "percorso_originale"
                attendance_df = calculate_attendance(current_df_for_tab3, group_by=group_by)
                st.info("Visualizzazione dei totali parziali per Percorso Originale senza Art.13.")
            elif "3. Riassunto per Percorso Elaborato" in view_type:
                group_by = "percorso_elaborato"
                attendance_df = calculate_attendance(current_df_for_tab3, group_by=group_by, percorso_elab_col='PercorsoInternal')
                st.info("Visualizzazione dei totali parziali per Percorso Elaborato (Info).")
            elif "4. Riassunto per Percorso Iscritti" in view_type:
                # Nuova opzione per raggruppare per Percorso dal file degli iscritti
                if 'Percorso' not in current_df_for_tab3.columns:
                    st.error("Colonna 'Percorso' dagli iscritti non trovata nei dati. Verifica che il file degli iscritti sia stato caricato correttamente.")
                    attendance_df = pd.DataFrame()
                else:
                    group_by = "percorso_iscritti"
                    attendance_df = calculate_attendance(current_df_for_tab3, group_by=group_by, percorso_chiave_col='Percorso')
                    st.info("Visualizzazione dei totali parziali per Percorso degli iscritti (dal file CSV).")
            elif "5. Lista Completa" in view_type:
                group_by = "lista_studenti"
                attendance_df = calculate_attendance(current_df_for_tab3, group_by=group_by)
                st.info("Lista completa degli studenti con tutte le presenze aggregate.")
            
            # Se stiamo visualizzando per studente e percorso, mostra i filtri standard
            if group_by == "studente" and not attendance_df.empty:
                st.subheader("Filtra Visualizzazione")
                p_col_disp_key = "Percorso (Senza Art.13)"
                p_col_internal_key = 'PercorsoOriginaleSenzaArt13Internal'

            # Gestione dei casi di raggruppamento specifici
            if group_by in ["percorso_originale", "percorso_elaborato", "percorso_iscritti", "lista_studenti"]:
                if not attendance_df.empty:
                    if group_by == "percorso_originale":
                        st.subheader("Totali Parziali per Percorso Originale (senza Art.13)")
                        # Ordina per presenze in ordine decrescente
                        attendance_df_sorted = attendance_df.sort_values(by="Presenze", ascending=False)
                        st.dataframe(attendance_df_sorted, use_container_width=True)
                        
                        # Esportazione dei dati
                        st.subheader("Esportazione")
                        export_col1, export_col2 = st.columns(2)
                        
                        with export_col1:
                            if st.button("Esporta in CSV", key="export_percorso_orig_csv"):
                                csv = attendance_df_sorted.to_csv(index=False).encode('utf-8')
                                ts = datetime.now().strftime("%Y%m%d_%H%M")
                                st.download_button(
                                    label="📥 Scarica CSV",
                                    data=csv,
                                    file_name=f"Totali_Percorso_Originale_{ts}.csv",
                                    mime="text/csv",
                                    key="download_percorso_orig_csv"
                                )
                        
                        with export_col2:
                            if st.button("Esporta in Excel", key="export_percorso_orig_excel"):
                                output = BytesIO()
                                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                                    attendance_df_sorted.to_excel(writer, sheet_name="Totali Percorso Originale", index=False)
                                output.seek(0)
                                ts = datetime.now().strftime("%Y%m%d_%H%M")
                                st.download_button(
                                    label="📥 Scarica Excel",
                                    data=output,
                                    file_name=f"Totali_Percorso_Originale_{ts}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key="download_percorso_orig_excel"
                                )
                            
                    elif group_by == "percorso_elaborato":
                        st.subheader("Totali Parziali per Percorso Elaborato")
                        # Ordina per presenze in ordine decrescente
                        attendance_df_sorted = attendance_df.sort_values(by="Presenze", ascending=False)
                        st.dataframe(attendance_df_sorted, use_container_width=True)
                        
                        # Esportazione dei dati
                        st.subheader("Esportazione")
                        export_col1, export_col2 = st.columns(2)
                        
                        with export_col1:
                            if st.button("Esporta in CSV", key="export_percorso_elab_csv"):
                                csv = attendance_df_sorted.to_csv(index=False).encode('utf-8')
                                ts = datetime.now().strftime("%Y%m%d_%H%M")
                                st.download_button(
                                    label="📥 Scarica CSV",
                                    data=csv,
                                    file_name=f"Totali_Percorso_Elaborato_{ts}.csv",
                                    mime="text/csv",
                                    key="download_percorso_elab_csv"
                                )
                                
                        with export_col2:
                            if st.button("Esporta in Excel", key="export_percorso_elab_excel"):
                                output = BytesIO()
                                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                                    attendance_df_sorted.to_excel(writer, sheet_name="Percorsi Elaborati", index=False)
                                output.seek(0)
                                ts = datetime.now().strftime("%Y%m%d_%H%M")
                                st.download_button(
                                    label="📥 Scarica Excel",
                                    data=output,
                                    file_name=f"Totali_Percorso_Elaborato_{ts}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key="download_percorso_elab_excel"
                                )
                                
                    elif group_by == "percorso_iscritti":
                        st.subheader("Totali Parziali per Percorso Iscritti")
                        
                        # Ordinamento intelligente: prima i PeF30, poi PeF36, poi PeF60
                        def sort_percorsi_iscritti(p):
                            if isinstance(p, str):
                                if "PeF30" in p:
                                    return "1" + p
                                elif "PeF36" in p:
                                    return "2" + p
                                elif "PeF60" in p:
                                    return "3" + p
                            return "9" + str(p)
                        
                        # Ordina prima per tipo percorso e poi per presenze in ordine decrescente
                        attendance_df_sorted = attendance_df.sort_values(
                            by=["Tipo Percorso Iscritti", "Presenze"], 
                            key=lambda x: x.map(sort_percorsi_iscritti) if x.name == "Tipo Percorso Iscritti" else x,
                            ascending=[True, False]
                        )
                        
                        # Mostra tabella con statistiche aggregate
                        st.dataframe(attendance_df_sorted, use_container_width=True)
                        
                        # Mostra statistiche aggiuntive
                        st.subheader("Statistiche per Tipo di Percorso")
                        percorsi_count = attendance_df_sorted.groupby("Tipo Percorso Iscritti").agg({
                            "Presenze": ["sum", "mean", "count"]
                        })
                        percorsi_count.columns = ["Presenze Totali", "Media Presenze", "Numero Studenti"]
                        percorsi_count = percorsi_count.round(1)
                        st.dataframe(percorsi_count, use_container_width=True)
                        
                        # Esportazione dei dati
                        st.subheader("Esportazione")
                        export_col1, export_col2 = st.columns(2)
                        
                        with export_col1:
                            if st.button("Esporta in CSV", key="export_percorso_iscritti_csv"):
                                csv = attendance_df_sorted.to_csv(index=False).encode('utf-8')
                                ts = datetime.now().strftime("%Y%m%d_%H%M")
                                st.download_button(
                                    label="📥 Scarica CSV",
                                    data=csv,
                                    file_name=f"Totali_Percorso_Iscritti_{ts}.csv",
                                    mime="text/csv",
                                    key="download_percorso_iscritti_csv"
                                )
                                
                        with export_col2:
                            if st.button("Esporta in Excel", key="export_percorso_iscritti_excel"):
                                output = BytesIO()
                                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                                    attendance_df_sorted.to_excel(writer, sheet_name="Percorsi Iscritti", index=False)
                                    percorsi_count.to_excel(writer, sheet_name="Statistiche Percorsi", index=True)
                                output.seek(0)
                                ts = datetime.now().strftime("%Y%m%d_%H%M")
                                st.download_button(
                                    label="📥 Scarica Excel",
                                    data=output,
                                    file_name=f"Totali_Percorso_Iscritti_{ts}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key="download_percorso_iscritti_excel"
                                )
                    elif group_by == "lista_studenti":
                        st.subheader("Lista Completa degli Studenti")
                        
                        # Aggiungi opzione di ricerca/filtro per la lista studenti
                        search_term = st.text_input("Cerca per nome, cognome o codice fiscale:", key="search_student_list")
                        
                        # Filtri aggiuntivi per classe di concorso e dipartimento
                        col_filters = st.columns(2)
                        with col_filters[0]:
                            # Filtro per Classe di concorso
                            classe_concorso_sel = "Tutte le classi di concorso"
                            if 'Codice_Classe_di_concorso' in current_df_for_tab3.columns:
                                classe_concorso_list = sorted([str(c) for c in current_df_for_tab3['Codice_Classe_di_concorso'].unique() 
                                                             if pd.notna(c) and str(c).strip() != ''])
                                if classe_concorso_list:
                                    classe_concorso_sel = st.selectbox(
                                        "Filtra per Classe di concorso:", 
                                        ["Tutte le classi di concorso"] + classe_concorso_list, 
                                        key="lista_classe_concorso"
                                    )
                        
                        with col_filters[1]:
                            # Filtro per Dipartimento
                            dipartimento_sel = "Tutti i dipartimenti"
                            if 'Dipartimento' in current_df_for_tab3.columns:
                                dipartimento_list = sorted([str(d) for d in current_df_for_tab3['Dipartimento'].unique() 
                                                          if pd.notna(d) and str(d).strip() != ''])
                                if dipartimento_list:
                                    dipartimento_sel = st.selectbox(
                                        "Filtra per Dipartimento:", 
                                        ["Tutti i dipartimenti"] + dipartimento_list, 
                                        key="lista_dipartimento"
                                    )
                        
                        # Ordina alfabeticamente per cognome e nome
                        attendance_df_sorted = attendance_df.sort_values(by=["Cognome", "Nome"])
                        
                        # Applica i filtri in sequenza
                        filtered_df = attendance_df_sorted
                        
                        # Applica filtro per classe di concorso
                        if classe_concorso_sel != "Tutte le classi di concorso" and 'Codice_Classe_di_concorso' in filtered_df.columns:
                            filtered_df = filtered_df[filtered_df['Codice_Classe_di_concorso'] == classe_concorso_sel]
                            st.info(f"Filtrato per classe di concorso: {classe_concorso_sel}")
                        
                        # Applica filtro per dipartimento
                        if dipartimento_sel != "Tutti i dipartimenti" and 'Dipartimento' in filtered_df.columns:
                            filtered_df = filtered_df[filtered_df['Dipartimento'] == dipartimento_sel]
                            st.info(f"Filtrato per dipartimento: {dipartimento_sel}")
                        
                        # Applica filtro di ricerca testuale
                        if search_term:
                            search_term_lower = search_term.lower()
                            filtered_df = filtered_df[
                                filtered_df["Cognome"].str.lower().str.contains(search_term_lower, na=False) |
                                filtered_df["Nome"].str.lower().str.contains(search_term_lower, na=False) |
                                filtered_df["CodiceFiscale"].str.lower().str.contains(search_term_lower, na=False)
                            ]
                        
                        # Visualizza risultati
                        st.dataframe(filtered_df, use_container_width=True)
                        st.info(f"Trovati {len(filtered_df)} studenti su {len(attendance_df_sorted)} totali.")
                        
                        # Esportazione della lista degli studenti
                        st.subheader("Esportazione")
                        export_col1, export_col2 = st.columns(2)
                        
                        with export_col1:
                            if st.button("Esporta in CSV", key="export_students_list_csv"):
                                csv = filtered_df.to_csv(index=False).encode('utf-8')
                                ts = datetime.now().strftime("%Y%m%d_%H%M")
                                st.download_button(
                                    label="📥 Scarica CSV",
                                    data=csv,
                                    file_name=f"Lista_Studenti_Filtrati_{ts}.csv",
                                    mime="text/csv",
                                    key="download_students_list_csv"
                                )
                        
                        with export_col2:
                            if st.button("Esporta in Excel", key="export_students_list_excel"):
                                output = BytesIO()
                                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                                    filtered_df.to_excel(writer, sheet_name="Lista Studenti", index=False)
                                output.seek(0)
                                ts = datetime.now().strftime("%Y%m%d_%H%M")
                                st.download_button(
                                    label="📥 Scarica Excel",
                                    data=output,
                                    file_name=f"Lista_Studenti_Filtrati_{ts}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key="download_students_list_excel"
                                )
            else:
                # La vecchia logica per il filtraggio per studente e percorso
                if not attendance_df.empty:
                    if p_col_disp_key not in attendance_df.columns:
                        st.error(f"Colonna chiave '{p_col_disp_key}' non trovata nei dati aggregati.")
                    elif p_col_internal_key not in current_df_for_tab3.columns:
                        st.error(f"Colonna chiave interna '{p_col_internal_key}' non trovata nei dati dettagliati.")
                    else:
                        try:
                            # --- Filtro Percorso ---
                            # Prepara la lista di percorsi per il selettore, garantendo che i codici siano visualizzati all'inizio
                            perc_list = sorted([str(p) for p in attendance_df[p_col_disp_key].unique() if pd.notna(p)])
                            # Ordina percorsi in base ai codici tra parentesi quadre se presenti
                            perc_list = sorted(perc_list, key=extract_sort_key)
                            
                            # Aggiungi opzione per filtrare per Classe di concorso e Dipartimento dagli iscritti
                            st.info("Filtri aggiuntivi dagli iscritti disponibili dopo aver selezionato un percorso")
                            
                        except Exception as e:
                            st.error(f"Errore nella preparazione dei filtri: {e}")
                            perc_list = []
                        perc_sel = st.selectbox(f"1. Filtra per {p_col_disp_key}:", ["Tutti"] + perc_list, key="filt_perc_tab3_v7")

                        # Inizializza variabili per dati filtrati
                        df_to_display_agg = pd.DataFrame()
                        df_to_display_detail = pd.DataFrame()

                        # --- Filtro Studente (solo se percorso specifico è selezionato) ---
                        stud_sel = "Tutti gli Studenti" # Default
                        
                        if perc_sel != "Tutti":
                            attendance_df_filtered_by_perc = attendance_df[attendance_df[p_col_disp_key] == perc_sel].copy()
                            df_detail_filtered_by_perc = current_df_for_tab3[current_df_for_tab3[p_col_internal_key] == perc_sel].copy()

                            # Filtri aggiuntivi per i dati degli iscritti
                            # Filtro per Classe di concorso
                            classe_concorso_sel = "Tutte le classi di concorso" # Default
                            if 'Codice_Classe_di_concorso' in df_detail_filtered_by_perc.columns:
                                # Lista di classi di concorso uniche nel percorso selezionato
                                classe_concorso_list = sorted([str(c) for c in df_detail_filtered_by_perc['Codice_Classe_di_concorso'].unique() 
                                                             if pd.notna(c) and str(c).strip() != ''])
                                if classe_concorso_list:
                                    classe_concorso_sel = st.selectbox(
                                        "2. Filtra per Classe di concorso:", 
                                        ["Tutte le classi di concorso"] + classe_concorso_list, 
                                        key="filt_classe_concorso_tab3"
                                    )
                            
                            # Applica filtro per classe di concorso
                            if classe_concorso_sel != "Tutte le classi di concorso" and 'Codice_Classe_di_concorso' in df_detail_filtered_by_perc.columns:
                                # Filtra i dettagli
                                filtered_detail = df_detail_filtered_by_perc[df_detail_filtered_by_perc['Codice_Classe_di_concorso'] == classe_concorso_sel].copy()
                                # Prendi i CF filtrati per aggiornare anche il dataframe aggregato
                                filtered_cfs = filtered_detail['CodiceFiscale'].unique() if not filtered_detail.empty else []
                                
                                if len(filtered_cfs) > 0:
                                    df_detail_filtered_by_perc = filtered_detail
                                    attendance_df_filtered_by_perc = attendance_df_filtered_by_perc[attendance_df_filtered_by_perc['CodiceFiscale'].isin(filtered_cfs)]
                                    st.info(f"Filtrato per classe di concorso: {classe_concorso_sel} - Trovati {len(filtered_cfs)} studenti")
                                else:
                                    st.warning(f"Nessun dato trovato per classe di concorso: {classe_concorso_sel}")
                            
                            # Filtro per Dipartimento
                            dipartimento_sel = "Tutti i dipartimenti" # Default
                            if 'Dipartimento' in df_detail_filtered_by_perc.columns:
                                # Lista di dipartimenti unici nel percorso e classe di concorso selezionati
                                dipartimento_list = sorted([str(d) for d in df_detail_filtered_by_perc['Dipartimento'].unique() 
                                                          if pd.notna(d) and str(d).strip() != ''])
                                if dipartimento_list:
                                    dipartimento_sel = st.selectbox(
                                        "3. Filtra per Dipartimento:", 
                                        ["Tutti i dipartimenti"] + dipartimento_list, 
                                        key="filt_dipartimento_tab3"
                                    )
                                    
                            # Applica filtro per dipartimento
                            if dipartimento_sel != "Tutti i dipartimenti" and 'Dipartimento' in df_detail_filtered_by_perc.columns:
                                # Filtra i dettagli
                                filtered_detail = df_detail_filtered_by_perc[df_detail_filtered_by_perc['Dipartimento'] == dipartimento_sel].copy()
                                # Prendi i CF filtrati per aggiornare anche il dataframe aggregato
                                filtered_cfs = filtered_detail['CodiceFiscale'].unique() if not filtered_detail.empty else []
                                
                                if len(filtered_cfs) > 0:
                                    df_detail_filtered_by_perc = filtered_detail
                                    attendance_df_filtered_by_perc = attendance_df_filtered_by_perc[attendance_df_filtered_by_perc['CodiceFiscale'].isin(filtered_cfs)]
                                    st.info(f"Filtrato per dipartimento: {dipartimento_sel} - Trovati {len(filtered_cfs)} studenti")
                                else:
                                    st.warning(f"Nessun dato trovato per dipartimento: {dipartimento_sel}")
                            
                            # Filtro per studente dopo aver applicato gli altri filtri
                            if not attendance_df_filtered_by_perc.empty:
                                attendance_df_filtered_by_perc['StudentIdentifier'] = attendance_df_filtered_by_perc.apply(
                                    lambda row: f"{row.get('Cognome','')} {row.get('Nome','')} ({row.get('CodiceFiscale','N/A')})".strip(), axis=1
                                )
                                student_list = sorted(attendance_df_filtered_by_perc['StudentIdentifier'].unique())
                                stud_sel = st.selectbox("4. Filtra per Studente (opzionale):", ["Tutti gli Studenti"] + student_list, key="filt_stud_tab3_v7")

                                if stud_sel != "Tutti gli Studenti":
                                    try:
                                        selected_cf = re.search(r'\((.*?)\)', stud_sel).group(1)
                                        df_to_display_agg = attendance_df_filtered_by_perc[attendance_df_filtered_by_perc['CodiceFiscale'] == selected_cf].copy()
                                        df_to_display_detail = df_detail_filtered_by_perc[df_detail_filtered_by_perc['CodiceFiscale'] == selected_cf].copy()
                                    except (AttributeError, IndexError):
                                        st.warning("Formato studente non riconosciuto nel filtro.")
                                        df_to_display_agg = attendance_df_filtered_by_perc
                                        df_to_display_detail = df_detail_filtered_by_perc
                                else:
                                    df_to_display_agg = attendance_df_filtered_by_perc
                                    df_to_display_detail = df_detail_filtered_by_perc
                            else:
                                st.info(f"Nessun dato aggregato trovato per il percorso '{perc_sel}'.")
                                df_to_display_agg = pd.DataFrame()
                                df_to_display_detail = pd.DataFrame()

                        # Se il filtro percorso è "Tutti"
                        else:
                            df_to_display_agg = attendance_df.copy()
                            df_to_display_detail = current_df_for_tab3.copy()

                        st.divider()

                        try:
                            # --- Visualizzazione Tabelle ---
                            if not df_to_display_agg.empty:
                                if perc_sel == "Tutti" or stud_sel == "Tutti gli Studenti":
                                    if perc_sel == "Tutti":
                                        st.subheader("Riepilogo Aggregato per Tutti i Percorsi e Studenti")
                                    else: # Percorso selezionato, ma tutti gli studenti
                                        st.subheader(f"Riepilogo Aggregato per: {perc_sel}")

                                    cols_disp_agg = ['CodiceFiscale', 'Nome', 'Cognome', 'Email', 
                                                    'Percorso', 'Codice_Classe_di_concorso', 'Dipartimento', 'Matricola',
                                                    p_col_disp_key, 'Percorso Elaborato (Info)', 'CFU Totali', 'Presenze']
                                    cols_disp_agg_exist = [c for c in cols_disp_agg if c in df_to_display_agg.columns]
                                    sort_agg_by = [p_col_disp_key, 'Cognome', 'Nome'] if perc_sel == "Tutti" else ['Cognome', 'Nome']
                                    # Assicura che le colonne di sort esistano
                                    valid_sort_agg_by = [c for c in sort_agg_by if c in df_to_display_agg.columns]
                                    if valid_sort_agg_by:
                                        st.dataframe(df_to_display_agg[cols_disp_agg_exist].sort_values(by=valid_sort_agg_by), use_container_width=True)
                                    else:
                                        st.dataframe(df_to_display_agg[cols_disp_agg_exist], use_container_width=True) # Fallback senza sort


                            if not df_to_display_detail.empty:
                                if perc_sel == "Tutti":
                                    st.subheader("Dettaglio Record Presenze per Tutti i Percorsi")
                                else:
                                    st.subheader(f"Dettaglio Record Presenze per: {perc_sel} - {stud_sel}")

                                cols_disp_detail = ['CodiceFiscale', 'Cognome', 'Nome', 'Email', 'DataPresenza', 'OraPresenza', 
                                                   'Percorso', 'Codice_Classe_di_concorso', 'Dipartimento',
                                                   'DenominazioneAttività', 'CFU', 'PercorsoInternal']
                                cols_disp_detail_exist = [c for c in cols_disp_detail if c in df_to_display_detail.columns]
                                sort_by_columns = ['Cognome', 'Nome']
                                if perc_sel == "Tutti": sort_by_columns.insert(0, p_col_internal_key)
                                if 'DataPresenza' in df_to_display_detail.columns: sort_by_columns.append('DataPresenza')
                                if 'OraPresenza' in df_to_display_detail.columns: sort_by_columns.append('OraPresenza')
                                valid_sort_by = [col for col in sort_by_columns if col in df_to_display_detail.columns]
                                if not valid_sort_by: st.dataframe(df_to_display_detail[cols_disp_detail_exist], use_container_width=True)
                                else: st.dataframe(df_to_display_detail[cols_disp_detail_exist].sort_values(by=valid_sort_by), use_container_width=True)
                            else:
                                if perc_sel != "Tutti": st.info("Nessun record dettagliato da mostrare per la selezione corrente.")
                        except Exception as e: 
                            st.error(f"Errore durante la visualizzazione: {e}")

            # --- Esportazione Excel Multi-Tab ---
            st.divider()
            st.subheader("Esportazione Dettaglio Presenze per Percorso (Originale senza Art.13) in Excel")
            course_col_export = 'PercorsoOriginaleSenzaArt13Internal'
            if course_col_export not in current_df_for_tab3.columns:
                st.error(f"Colonna chiave '{course_col_export}' non trovata.")
            else:
                st.write(f"Esporta dati **dettagliati** in Excel o CSV, con opzione multi-foglio Excel per ogni **{course_col_export}** unico (nome foglio da codice tra parentesi).")
                st.markdown("**1. Seleziona e Ordina le Colonne per l'Export:**")
                st.caption("Usa il box qui sotto per scegliere le colonne. Puoi trascinare le colonne selezionate per cambiarne l'ordine.")

                all_possible_cols = current_df_for_tab3.columns.tolist()
                internal_cols_to_exclude = ['TimestampPresenza']
                all_exportable_cols = [col for col in all_possible_cols if col not in internal_cols_to_exclude]
                default_cols_export_ordered = ['DataPresenza', 'OraPresenza', 'DenominazioneAttività', 
                                              'Cognome', 'Nome', 'Email', 'CodiceFiscale',
                                              'Percorso', 'Codice_Classe_di_concorso', 
                                              'Codice_classe_di_concorso_e_denominazione', 'Dipartimento', 
                                              'LogonName', 'Matricola',
                                              'PercorsoInternal', 'PercorsoOriginaleSenzaArt13Internal', 'CFU']
                default_cols_final = [col for col in default_cols_export_ordered if col in all_exportable_cols]

                st.markdown("**Esempio Record Dati (prima riga):**")
                if not current_df_for_tab3.empty:
                    example_data = current_df_for_tab3.head(1)[all_exportable_cols].to_dict(orient='records')[0]
                    st.json(example_data)
                else: 
                    st.caption("Nessun dato da mostrare.")

                selected_cols_export_ordered = st.multiselect("Seleziona e ordina le colonne:", options=all_exportable_cols, default=default_cols_final, key="export_cols_selector_ordered_v215")
                
                # Resto del codice per l'esportazione
                
                # [Include il codice esistente per l'esportazione qui]
                
    else:
        st.info("Nessun dato valido caricato.")
