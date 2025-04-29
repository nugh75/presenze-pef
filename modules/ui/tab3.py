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
                "4. Lista Completa Studenti"
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
            elif "4. Lista Completa" in view_type:
                group_by = "lista_studenti"
                attendance_df = calculate_attendance(current_df_for_tab3, group_by=group_by)
                st.info("Lista completa degli studenti con tutte le presenze aggregate.")
            
            # Se stiamo visualizzando per studente e percorso, mostra i filtri standard
            if group_by == "studente" and not attendance_df.empty:
                st.subheader("Filtra Visualizzazione")
                p_col_disp_key = "Percorso (Senza Art.13)"
                p_col_internal_key = 'PercorsoOriginaleSenzaArt13Internal'

            # Gestione dei casi di raggruppamento specifici
            if group_by in ["percorso_originale", "percorso_elaborato", "lista_studenti"]:
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
                                    attendance_df_sorted.to_excel(writer, sheet_name="Totali Percorso Elaborato", index=False)
                                output.seek(0)
                                ts = datetime.now().strftime("%Y%m%d_%H%M")
                                st.download_button(
                                    label="📥 Scarica Excel",
                                    data=output,
                                    file_name=f"Totali_Percorso_Elaborato_{ts}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key="download_percorso_elab_excel"
                                )
                            
                    elif group_by == "lista_studenti":
                        st.subheader("Lista Completa degli Studenti")
                        
                        # Aggiungi opzione di ricerca/filtro per la lista studenti
                        search_term = st.text_input("Cerca per nome, cognome o codice fiscale:", key="search_student_list")
                        
                        # Ordina alfabeticamente per cognome e nome
                        attendance_df_sorted = attendance_df.sort_values(by=["Cognome", "Nome"])
                        
                        # Applica filtro se necessario
                        if search_term:
                            search_term_lower = search_term.lower()
                            filtered_df = attendance_df_sorted[
                                attendance_df_sorted["Cognome"].str.lower().str.contains(search_term_lower, na=False) |
                                attendance_df_sorted["Nome"].str.lower().str.contains(search_term_lower, na=False) |
                                attendance_df_sorted["CodiceFiscale"].str.lower().str.contains(search_term_lower, na=False)
                            ]
                            st.dataframe(filtered_df, use_container_width=True)
                            st.info(f"Trovati {len(filtered_df)} studenti su {len(attendance_df_sorted)} totali.")
                        else:
                            st.dataframe(attendance_df_sorted, use_container_width=True)
                        
                        # Esportazione della lista degli studenti
                        st.subheader("Esportazione")
                        export_col1, export_col2 = st.columns(2)
                        
                        with export_col1:
                            if st.button("Esporta in CSV", key="export_students_list_csv"):
                                csv = attendance_df_sorted.to_csv(index=False).encode('utf-8')
                                ts = datetime.now().strftime("%Y%m%d_%H%M")
                                st.download_button(
                                    label="📥 Scarica CSV",
                                    data=csv,
                                    file_name=f"Lista_Completa_Studenti_{ts}.csv",
                                    mime="text/csv",
                                    key="download_students_list_csv"
                                )
                        
                        with export_col2:
                            if st.button("Esporta in Excel", key="export_students_list_excel"):
                                output = BytesIO()
                                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                                    attendance_df_sorted.to_excel(writer, sheet_name="Lista Studenti", index=False)
                                output.seek(0)
                                ts = datetime.now().strftime("%Y%m%d_%H%M")
                                st.download_button(
                                    label="📥 Scarica Excel",
                                    data=output,
                                    file_name=f"Lista_Completa_Studenti_{ts}.xlsx",
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
                    # CORREZIONE: Questo else deve coprire tutto il blocco try/except successivo
                    else:
                        try:
                            # --- Filtro Percorso ---
                            # Prepara la lista di percorsi per il selettore, garantendo che i codici siano visualizzati all'inizio
                            perc_list = sorted([str(p) for p in attendance_df[p_col_disp_key].unique() if pd.notna(p)])
                            # Ordina percorsi in base ai codici tra parentesi quadre se presenti
                            perc_list = sorted(perc_list, key=extract_sort_key)
                            
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

                            if not attendance_df_filtered_by_perc.empty:
                                attendance_df_filtered_by_perc['StudentIdentifier'] = attendance_df_filtered_by_perc.apply(
                                    lambda row: f"{row.get('Cognome','')} {row.get('Nome','')} ({row.get('CodiceFiscale','N/A')})".strip(), axis=1
                                )
                                student_list = sorted(attendance_df_filtered_by_perc['StudentIdentifier'].unique())
                                stud_sel = st.selectbox("2. Filtra per Studente (opzionale):", ["Tutti gli Studenti"] + student_list, key="filt_stud_tab3_v7")

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

                                    cols_disp_agg = ['CodiceFiscale', 'Nome', 'Cognome', 'Email', p_col_disp_key, 'Percorso Elaborato (Info)', 'CFU Totali', 'Presenze']
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

                                cols_disp_detail = ['CodiceFiscale', 'Cognome', 'Nome', 'Email', 'DataPresenza', 'OraPresenza', 'DenominazioneAttività', 'CFU', 'PercorsoInternal']
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
                default_cols_export_ordered = ['DataPresenza','OraPresenza','DenominazioneAttività','Cognome','Nome','Email','PercorsoInternal','PercorsoOriginaleSenzaArt13Internal','CFU']
                default_cols_final = [col for col in default_cols_export_ordered if col in all_exportable_cols]

                st.markdown("**Esempio Record Dati (prima riga):**")
                if not current_df_for_tab3.empty:
                    example_data = current_df_for_tab3.head(1)[all_exportable_cols].to_dict(orient='records')[0]
                    st.json(example_data)
                else: 
                    st.caption("Nessun dato da mostrare.")

                selected_cols_export_ordered = st.multiselect("Seleziona e ordina le colonne:", options=all_exportable_cols, default=default_cols_final, key="export_cols_selector_ordered_v215")
                
                st.markdown("**2. Seleziona Periodo per l'Export (opzionale):**")
                col1_date, col2_date = st.columns(2)
                
                with col1_date:
                    # Determina la data minima e massima nel dataset
                    min_date = None
                    if not current_df_for_tab3.empty and 'DataPresenza' in current_df_for_tab3.columns:
                        valid_dates = current_df_for_tab3['DataPresenza'].dropna()
                        if not valid_dates.empty:
                            try:
                                min_date = valid_dates.min()
                                if not isinstance(min_date, date):
                                    min_date = pd.to_datetime(min_date).date() if not pd.isna(min_date) else None
                            except Exception as e:
                                st.warning(f"Problema con date: {e}")
                                min_date = None
                    
                    start_date = st.date_input("Data inizio:", value=min_date, key="export_start_date", help="Lascia vuoto per includere tutti i dati dall'inizio")
                
                with col2_date:
                    # Determina la data massima
                    max_date = None
                    if not current_df_for_tab3.empty and 'DataPresenza' in current_df_for_tab3.columns:
                        valid_dates = current_df_for_tab3['DataPresenza'].dropna()
                        if not valid_dates.empty:
                            try:
                                max_date = valid_dates.max()
                                if not isinstance(max_date, date):
                                    max_date = pd.to_datetime(max_date).date() if not pd.isna(max_date) else None
                            except Exception as e:
                                max_date = None
                    
                    end_date = st.date_input("Data fine:", value=max_date, key="export_end_date", help="Lascia vuoto per includere tutti i dati fino alla fine")

                # Mostra periodo selezionato
                if start_date is not None or end_date is not None:
                    date_filter_text = "Periodo selezionato: "
                    if start_date is not None:
                        date_filter_text += f"dal {start_date.strftime('%d/%m/%Y')}"
                    else:
                        date_filter_text += "dall'inizio"
                    
                    if end_date is not None:
                        date_filter_text += f" al {end_date.strftime('%d/%m/%Y')}"
                    else:
                        date_filter_text += " alla fine"
                    
                    st.info(date_filter_text)
                
                # Creazione delle colonne per i pulsanti di esportazione
                export_col1, export_col2 = st.columns(2)
                
                # Pulsante per generare ed esportare file Excel
                with export_col1:
                    if st.button("Genera ed Esporta File Excel", key="export_excel_ordered_v215"):
                        if not selected_cols_export_ordered: 
                            st.warning("Seleziona almeno una colonna.")
                        else:
                            overall_success = True
                            sheets_written = 0
                            error_messages = []
                            try:
                                output = BytesIO()
                                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                                    unique_courses = current_df_for_tab3[course_col_export].unique()
                                    unique_courses = sorted([str(c) for c in unique_courses if pd.notna(c)])
                                    if not unique_courses: 
                                        st.error(f"Nessun valore unico trovato in '{course_col_export}'.")
                                        overall_success = False
                                    else:
                                        prog_bar = st.progress(0, text="Creazione fogli...")
                                        for i, course_value in enumerate(unique_courses):
                                            # Cerca prima per il nuovo formato [codice]
                                            code_match = re.search(r'^\[([-\w]+)\]', course_value)
                                            if code_match:
                                                extracted_code = code_match.group(1)
                                                sheet_name_cleaned = clean_sheet_name(extracted_code)
                                            else:
                                                # Fallback al metodo precedente di estrazione dalle parentesi
                                                extracted_code = extract_code_from_parentheses(course_value)
                                                if extracted_code: 
                                                    sheet_name_cleaned = clean_sheet_name(extracted_code)
                                                else: 
                                                    sheet_name_cleaned = clean_sheet_name(course_value)
                                                    st.caption(f"Nota: Codice non trovato per '{course_value}', usato nome completo.")
                                            
                                            prog_text = f"Foglio: {sheet_name_cleaned} ({i+1}/{len(unique_courses)})" 
                                            prog_bar.progress((i + 1) / len(unique_courses), text=prog_text)
                                            df_sheet = current_df_for_tab3[current_df_for_tab3[course_col_export] == course_value].copy()
                                            
                                            if df_sheet.empty: 
                                                st.write(f"Info: Nessun dato per '{sheet_name_cleaned}', foglio saltato.")
                                                continue
                                            
                                            # Filtro per periodo di data se selezionato
                                            if 'DataPresenza' in df_sheet.columns:
                                                date_filtered = False
                                                if 'export_start_date' in st.session_state and st.session_state.export_start_date is not None:
                                                    start_date = st.session_state.export_start_date
                                                    df_sheet = df_sheet[df_sheet['DataPresenza'] >= start_date]
                                                    date_filtered = True
                                                
                                                if 'export_end_date' in st.session_state and st.session_state.export_end_date is not None:
                                                    end_date = st.session_state.export_end_date
                                                    df_sheet = df_sheet[df_sheet['DataPresenza'] <= end_date]
                                                    date_filtered = True
                                                
                                                if date_filtered and df_sheet.empty:
                                                    st.write(f"Info: Nessun dato per '{sheet_name_cleaned}' nel periodo selezionato, foglio saltato.")
                                                    continue
                                            
                                            final_ordered_cols_for_sheet = [col for col in selected_cols_export_ordered if col in df_sheet.columns]
                                            if not final_ordered_cols_for_sheet: 
                                                st.write(f"Info: Nessuna colonna selezionata trovata per '{sheet_name_cleaned}', foglio saltato.")
                                                continue
                                                
                                            df_sheet_export = df_sheet[final_ordered_cols_for_sheet]
                                            if df_sheet_export.empty: 
                                                st.write(f"Info: Nessun dato per '{sheet_name_cleaned}' dopo selezione colonne, foglio saltato.")
                                                continue
                                                
                                            try:
                                                rename_map_export = {
                                                    'PercorsoInternal': 'Tipo Percorso', 
                                                    'PercorsoOriginaleInternal': 'Percorso Originale Input', 
                                                    'PercorsoOriginaleSenzaArt13Internal': 'Denominazione Percorso', 
                                                    'DenominazioneAttivitaNormalizzataInternal': 'Attività Elaborata'
                                                }
                                                cols_to_rename_final = {k: v for k, v in rename_map_export.items() if k in df_sheet_export.columns}
                                                df_sheet_export = df_sheet_export.rename(columns=cols_to_rename_final)
                                                df_sheet_export.to_excel(writer, sheet_name=sheet_name_cleaned, index=False)
                                                sheets_written += 1
                                            except Exception as sheet_error: 
                                                error_msg = f"Errore scrittura foglio '{sheet_name_cleaned}': {sheet_error}"
                                                st.warning(error_msg)
                                                error_messages.append(error_msg)
                                                overall_success = False
                            except Exception as writer_error: 
                                st.error(f"Errore generale creazione Excel: {writer_error}")
                                st.exception(writer_error)
                                overall_success = False
                                
                            if overall_success and sheets_written > 0:
                                prog_bar.progress(1.0, text="Completato!")
                                output.seek(0)
                                ts = datetime.now().strftime("%Y%m%d_%H%M")
                                
                                # Aggiungi informazioni sul periodo al nome del file
                                period_info = ""
                                if 'export_start_date' in st.session_state and st.session_state.export_start_date is not None:
                                    start_str = st.session_state.export_start_date.strftime("%Y%m%d")
                                    period_info += f"_dal{start_str}"
                                if 'export_end_date' in st.session_state and st.session_state.export_end_date is not None:
                                    end_str = st.session_state.export_end_date.strftime("%Y%m%d")
                                    period_info += f"_al{end_str}"
                                
                                fname = f"Report_Presenze_Dettaglio_PercorsoSenzaArt13{period_info}_{ts}.xlsx"
                                st.success(f"File Excel generato con {sheets_written} fogli!")
                                if error_messages: 
                                    st.warning("Alcuni fogli potrebbero aver avuto problemi:")
                                    [st.caption(msg) for msg in error_messages]
                                    
                                st.download_button(
                                    label="📥 Scarica Report Excel Dettaglio",
                                    data=output,
                                    file_name=fname,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key="dl_excel_course_ordered_v215"
                                )
                            elif sheets_written == 0 and overall_success: 
                                st.warning("Nessun dato trovato per alcun percorso. File Excel non generato.")
                            else:
                                st.error("Generazione file Excel fallita o nessun foglio valido scritto.")
                                if error_messages: 
                                    st.warning("Dettaglio errori:")
                                    [st.caption(msg) for msg in error_messages]
                
                # Pulsante per generare ed esportare file CSV
                with export_col2:
                    if st.button("Genera ed Esporta File CSV", key="export_csv_ordered_v215"):
                        if not selected_cols_export_ordered: 
                            st.warning("Seleziona almeno una colonna.")
                        else:
                            # Filtra i dati in base al periodo se selezionato
                            filtered_df = current_df_for_tab3.copy()
                            date_filtered = False
                            
                            if 'export_start_date' in st.session_state and st.session_state.export_start_date is not None and 'DataPresenza' in filtered_df.columns:
                                start_date = st.session_state.export_start_date
                                filtered_df = filtered_df[filtered_df['DataPresenza'] >= start_date]
                                date_filtered = True
                            
                            if 'export_end_date' in st.session_state and st.session_state.export_end_date is not None and 'DataPresenza' in filtered_df.columns:
                                end_date = st.session_state.export_end_date
                                filtered_df = filtered_df[filtered_df['DataPresenza'] <= end_date]
                                date_filtered = True
                            
                            if filtered_df.empty:
                                st.warning("Nessun dato disponibile per il periodo selezionato.")
                            else:
                                # Prendi solo le colonne selezionate
                                final_ordered_cols = [col for col in selected_cols_export_ordered if col in filtered_df.columns]
                                
                                if not final_ordered_cols:
                                    st.warning("Nessuna delle colonne selezionate è presente nei dati.")
                                else:
                                    # Prepara il CSV con le colonne selezionate e rinominate
                                    df_export = filtered_df[final_ordered_cols]
                                    rename_map = {
                                        'PercorsoInternal': 'Tipo Percorso', 
                                        'PercorsoOriginaleInternal': 'Percorso Originale Input', 
                                        'PercorsoOriginaleSenzaArt13Internal': 'Denominazione Percorso', 
                                        'DenominazioneAttivitaNormalizzataInternal': 'Attività Elaborata'
                                    }
                                    cols_to_rename = {k: v for k, v in rename_map.items() if k in df_export.columns}
                                    df_export = df_export.rename(columns=cols_to_rename)
                                    
                                    # Esporta in CSV
                                    csv_data = df_export.to_csv(index=False).encode('utf-8')
                                    ts = datetime.now().strftime("%Y%m%d_%H%M")
                                    
                                    # Aggiungi informazioni sul periodo al nome del file
                                    period_info = ""
                                    if 'export_start_date' in st.session_state and st.session_state.export_start_date is not None:
                                        start_str = st.session_state.export_start_date.strftime("%Y%m%d")
                                        period_info += f"_dal{start_str}"
                                    if 'export_end_date' in st.session_state and st.session_state.export_end_date is not None:
                                        end_str = st.session_state.export_end_date.strftime("%Y%m%d")
                                        period_info += f"_al{end_str}"
                                    
                                    fname_csv = f"Report_Presenze_Dettaglio{period_info}_{ts}.csv"
                                    st.success(f"File CSV generato con {len(filtered_df)} record!")
                                    st.download_button(
                                        label="📥 Scarica Report CSV",
                                        data=csv_data,
                                        file_name=fname_csv,
                                        mime="text/csv",
                                        key="dl_csv_ordered_v215"
                                    )

            if not attendance_df.empty:
                pass  # Il blocco 'if' precedente conteneva già tutto il codice necessario
            else:  # if attendance_df.empty
                st.info("Nessun dato di presenza aggregato da visualizzare (verifica i dati o le colonne necessarie).")
    else:
        st.info("Nessun dato valido caricato.")
