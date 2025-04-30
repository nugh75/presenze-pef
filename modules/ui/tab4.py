# Interfaccia utente per la Tab 4 (Frequenza Lezioni)
import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from modules.attendance import calculate_lesson_attendance

def render_tab4(df_main):
    """Renderizza l'interfaccia della Tab 4: Frequenza Lezioni"""
    st.header("Frequenza Lezioni")
    st.write("Visualizza il numero di partecipanti unici per ogni combinazione di data e attività. Puoi filtrare i risultati per data e/o per attività.")

    current_df_for_tab4 = df_main

    if not current_df_for_tab4.empty:
        date_col = 'DataPresenza'
        activity_col = 'DenominazioneAttivitaNormalizzataInternal'
        cf_col = 'CodiceFiscale'

        required_cols = [date_col, activity_col, cf_col]
        if not all(col in current_df_for_tab4.columns for col in required_cols):
            missing_cols = [col for col in required_cols if col not in current_df_for_tab4.columns]
            st.error(f"Impossibile procedere: colonne mancanti ({', '.join(missing_cols)})")
        else:
            st.subheader("Filtri")
            
            # Creazione di due colonne per i filtri indipendenti
            col1, col2 = st.columns(2)
            
            # Filtro per data
            with col1:
                # Ottieni tutte le date uniche e ordinate
                unique_dates = sorted(current_df_for_tab4[date_col].unique())
                date_filter = st.selectbox(
                    "Filtra per data:",
                    ["Tutte le date"] + [d for d in unique_dates if pd.notna(d)],
                    key="date_filter_tab4"
                )
            
            # Filtro per attività
            with col2:
                # Ottieni tutte le attività uniche e ordinate alfabeticamente
                unique_activities = sorted([a for a in current_df_for_tab4[activity_col].unique() if isinstance(a, str) and a.strip()])
                activity_filter = st.selectbox(
                    "Filtra per attività:",
                    ["Tutte le attività"] + unique_activities,
                    key="activity_filter_tab4"
                )
            
            # Calcola e visualizza i dati sulla frequenza delle lezioni
            if date_filter == "Tutte le date" and activity_filter == "Tutte le attività":
                st.subheader("Frequenza per tutte le lezioni")
            elif date_filter != "Tutte le date" and activity_filter == "Tutte le attività":
                st.subheader(f"Frequenza per le lezioni del {date_filter}")
            elif date_filter == "Tutte le date" and activity_filter != "Tutte le attività":
                st.subheader(f"Frequenza per l'attività: {activity_filter}")
            else:
                st.subheader(f"Frequenza per l'attività: {activity_filter} del {date_filter}")
            
            # Gestisci i parametri dei filtri
            date_param = date_filter if date_filter != "Tutte le date" else None
            activity_param = activity_filter if activity_filter != "Tutte le attività" else None
            
            # Calcola i dati della frequenza
            attendance_data = calculate_lesson_attendance(
                current_df_for_tab4,
                date_filter=date_param,
                activity_filter=activity_param,
                date_col=date_col,
                activity_col=activity_col,
                cf_column=cf_col
            )
            
            # Visualizza i risultati
            if not attendance_data.empty:
                # Rinomina le colonne per la visualizzazione
                display_cols = {
                    date_col: 'Data',
                    activity_col: 'Attività',
                    'Partecipanti': 'Partecipanti'
                }
                attendance_display = attendance_data.rename(columns=display_cols)
                
                # Visualizza la tabella con i dati
                st.dataframe(attendance_display, use_container_width=True)
                
                # Aggiunta di statistiche riepilogative
                st.subheader("Statistiche")
                col_stats1, col_stats2, col_stats3 = st.columns(3)
                
                with col_stats1:
                    total_lessons = len(attendance_data)
                    st.metric("Numero di lezioni", total_lessons)
                
                with col_stats2:
                    if not attendance_data.empty:
                        avg_attendance = round(attendance_data['Partecipanti'].mean(), 1)
                        st.metric("Media partecipanti", avg_attendance)
                
                with col_stats3:
                    if not attendance_data.empty:
                        total_attendance = attendance_data['Partecipanti'].sum()
                        st.metric("Totale presenze", total_attendance)
                
                # Aggiungi lista partecipanti per lezione specifica
                st.divider()
                if date_param is not None or activity_param is not None:
                    st.subheader("Lista dei Partecipanti")
                    
                    # Filtra i dati per ottenere i partecipanti alla lezione selezionata
                    participants_df = current_df_for_tab4.copy()
                    
                    if date_param is not None:
                        participants_df = participants_df[participants_df[date_col] == date_param]
                    
                    if activity_param is not None:
                        participants_df = participants_df[participants_df[activity_col] == activity_param]
                    
                    if not participants_df.empty:
                        # Estrai solo i record unici per ogni persona
                        participants_df = participants_df.sort_values(by=['Cognome', 'Nome', 'CodiceFiscale'])
                        participants_df = participants_df.drop_duplicates(subset=['CodiceFiscale'])
                        
                        # Seleziona solo le colonne necessarie, inclusi i dati degli iscritti e i percorsi
                        display_columns = ['Cognome', 'Nome', 'CodiceFiscale', 'Email', 
                                          'Percorso', 'Codice_Classe_di_concorso', 'Codice_classe_di_concorso_e_denominazione', 
                                          'Dipartimento', 'LogonName', 'Matricola',
                                          'Percorso']
                        columns_to_show = [col for col in display_columns if col in participants_df.columns]
                        
                        # Rinomina le colonne dei percorsi per una migliore visualizzazione
                        rename_map = {}
                        if 'Percorso' in columns_to_show:
                            rename_map['Percorso'] = 'Percorso'
                        
                        # Applica la rinomina se necessario
                        if rename_map:
                            participants_df = participants_df.rename(columns=rename_map)
                            # Aggiorna i nomi delle colonne per visualizzazione
                            columns_to_show = [rename_map.get(col, col) for col in columns_to_show]
                        
                        if columns_to_show:
                            # Visualizza la tabella con i dati dei partecipanti
                            st.dataframe(participants_df[columns_to_show], use_container_width=True)
                            st.info(f"Partecipanti totali: {len(participants_df)}")
                            
                            # Aggiungi opzione per esportare la lista dei partecipanti
                            st.subheader("Esporta lista partecipanti")
                            export_col1, export_col2 = st.columns(2)
                            
                            # Crea nome file
                            parts = ["Partecipanti"]
                            if date_param:
                                parts.append(f"Data_{str(date_param).replace('-', '')}")
                            if activity_param:
                                activity_safe = activity_param.replace(" ", "_").replace("/", "-")[:30]
                                parts.append(f"Attivita_{activity_safe}")
                            
                            ts = datetime.now().strftime("%Y%m%d_%H%M")
                            
                            with export_col1:
                                if st.button("Esporta in CSV", key="export_participants_csv"):
                                    # Usa le colonne rinominate per l'esportazione
                                    participants_csv = participants_df[columns_to_show].to_csv(index=False).encode('utf-8')
                                    filename_csv = f"{'_'.join(parts)}_{ts}.csv"
                                    
                                    st.download_button(
                                        label="📥 Scarica CSV",
                                        data=participants_csv,
                                        file_name=filename_csv,
                                        mime="text/csv",
                                        key="download_participants_list_csv"
                                    )
                            
                            with export_col2:
                                if st.button("Esporta in Excel", key="export_participants_excel"):
                                    output = BytesIO()
                                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                                        # Usa le colonne rinominate per l'esportazione
                                        participants_df[columns_to_show].to_excel(writer, sheet_name="Lista Partecipanti", index=False)
                                    
                                    output.seek(0)
                                    filename_excel = f"{'_'.join(parts)}_{ts}.xlsx"
                                    
                                    st.download_button(
                                        label="📥 Scarica Excel",
                                        data=output,
                                        file_name=filename_excel,
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        key="download_participants_list_excel"
                                    )
                        else:
                            st.warning("Dati anagrafici non disponibili per i partecipanti.")
                    else:
                        st.info("Nessun partecipante trovato per la selezione corrente.")
                
                # Esportazione dati
                st.divider()
                st.subheader("Esportazione dati")
                
                export_col1, export_col2 = st.columns(2)
                
                # Crea un nome file significativo in base ai filtri applicati
                filename_parts = ["Frequenza_Lezioni"]
                if date_param:
                    date_str = str(date_param).replace("-", "")
                    filename_parts.append(f"Data_{date_str}")
                if activity_param:
                    # Normalizza il nome dell'attività per il filename
                    activity_str = activity_param.replace(" ", "_").replace("/", "-")[:30]
                    filename_parts.append(f"Attivita_{activity_str}")
                
                ts = datetime.now().strftime("%Y%m%d_%H%M")
                
                with export_col1:
                    if st.button("Esporta in CSV", key="export_lesson_attendance_csv"):
                        csv = attendance_display.to_csv(index=False).encode('utf-8')
                        filename_csv = f"{'_'.join(filename_parts)}_{ts}.csv"
                        
                        st.download_button(
                            label="📥 Scarica CSV",
                            data=csv,
                            file_name=filename_csv,
                            mime="text/csv",
                            key="download_lesson_attendance_csv"
                        )
                
                with export_col2:
                    if st.button("Esporta in Excel", key="export_lesson_attendance_excel"):
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            attendance_display.to_excel(writer, sheet_name="Frequenza Lezioni", index=False)
                        
                        output.seek(0)
                        filename_excel = f"{'_'.join(filename_parts)}_{ts}.xlsx"
                        
                        st.download_button(
                            label="📥 Scarica Excel",
                            data=output,
                            file_name=filename_excel,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="download_lesson_attendance_excel"
                        )
            else:
                st.info("Nessun dato disponibile per i filtri selezionati.")
    else:
        st.info("Nessun dato valido caricato.")
