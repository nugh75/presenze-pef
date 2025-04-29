# Interfaccia utente per la Tab 2 (Gestione Duplicati)
import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from modules.duplicates import detect_duplicate_records

def render_tab2(df_main):
    """Renderizza l'interfaccia della Tab 2: Gestione Duplicati"""
    st.header("Gestione Record Potenzialmente Duplicati")
    st.markdown("Identifica cluster di timbrature ravvicinate (≤ 10 min) per stesso CF/giorno.")
    
    if not st.session_state.duplicates_removed and st.session_state.duplicate_detection_results[0].empty:
        with st.spinner("Rilevamento duplicati..."):
            df_per_detect = df_main
            if not df_per_detect.empty:
                required_dup_cols = ['TimestampPresenza', 'CodiceFiscale']
                if all(col in df_per_detect.columns for col in required_dup_cols):
                    st.session_state.duplicate_detection_results = detect_duplicate_records(df_per_detect)
                else:
                    missing_cols = [col for col in required_dup_cols if col not in df_per_detect.columns]
                    st.warning(f"Colonne necessarie ({', '.join(missing_cols)}) non trovate.")
                    st.session_state.duplicate_detection_results = (pd.DataFrame(), [], [])
            else:
                st.session_state.duplicate_detection_results = (pd.DataFrame(), [], [])
                
    duplicates_df_display, involved_indices, indices_to_drop_suggested = st.session_state.duplicate_detection_results
    duplicates_found = not duplicates_df_display.empty
    
    if duplicates_found:
        st.warning(f"Trovati **{len(duplicates_df_display)}** record potenzialmente duplicati in **{duplicates_df_display['GruppoDuplicati'].nunique()}** cluster.")
        st.markdown("---")
        st.subheader("Azione Rapida: Eliminazione Automatica Suggerita")
        
        indices_to_drop_suggested_auto = st.session_state.duplicate_detection_results[2]
        current_df_auto = df_main
        valid_indices_to_remove_auto = []
        
        if not current_df_auto.empty and indices_to_drop_suggested_auto:
            valid_indices_to_remove_auto = [idx for idx in indices_to_drop_suggested_auto if idx in current_df_auto.index]
            
        num_valid_to_remove = len(valid_indices_to_remove_auto)
        auto_remove_disabled = st.session_state.duplicates_removed or num_valid_to_remove == 0
        
        col_btn_auto, col_report_dl = st.columns([2,3])
        
        with col_btn_auto:
            if st.button(f"Elimina {num_valid_to_remove} Record Suggeriti", key="auto_remove_and_report", 
                         disabled=auto_remove_disabled, 
                         help="Rimuove tutti tranne il primo per cluster e prepara un CSV degli eliminati."):
                if valid_indices_to_remove_auto:
                    try:
                        with st.spinner(f"Eliminazione e preparazione report..."):
                            df_deleted_report = current_df_auto.loc[valid_indices_to_remove_auto].copy()
                            duplicates_df_display_report = st.session_state.duplicate_detection_results[0]
                            
                            if not duplicates_df_display_report.empty and 'GruppoDuplicati' in duplicates_df_display_report.columns:
                                try:
                                    valid_report_indices = duplicates_df_display_report.index.intersection(valid_indices_to_remove_auto)
                                    df_deleted_report['GruppoDuplicati'] = duplicates_df_display_report.loc[valid_report_indices, 'GruppoDuplicati'] if not valid_report_indices.empty else 0
                                except Exception:
                                    pass
                                    
                            cols_report_preferred = ['GruppoDuplicati', 'CodiceFiscale', 'Nome', 'Cognome', 
                                                    'DataPresenza', 'OraPresenza', 'PercorsoOriginaleInternal', 
                                                    'PercorsoOriginaleSenzaArt13Internal','PercorsoInternal', 
                                                    'DenominazioneAttività', 'DenominazioneAttivitaNormalizzataInternal', 
                                                    'CodicePercorso', 'TimestampPresenza']
                                                    
                            cols_report_exist = [c for c in cols_report_preferred if c in df_deleted_report.columns]
                            df_deleted_report_final = df_deleted_report[cols_report_exist]
                            
                            if 'GruppoDuplicati' in cols_report_exist and 'TimestampPresenza' in cols_report_exist:
                                df_deleted_report_final = df_deleted_report_final.sort_values(by=['GruppoDuplicati', 'TimestampPresenza'])
                                
                            report_csv_bytes = df_deleted_report_final.to_csv(index=True, index_label='OriginalIndex').encode('utf-8')
                            ts_report = datetime.now().strftime("%Y%m%d_%H%M")
                            report_filename = f"Report_Record_Eliminati_Auto_{ts_report}.csv"
                            
                            st.session_state.report_data_to_download = report_csv_bytes
                            st.session_state.report_filename_to_download = report_filename
                            
                            df_cleaned = current_df_auto.drop(index=valid_indices_to_remove_auto)
                            st.session_state.processed_df = df_cleaned
                            st.session_state.duplicates_removed = True
                            st.session_state.duplicate_detection_results = (pd.DataFrame(), [], [])
                            st.session_state.selected_indices_to_drop = []
                            
                        st.success(f"{num_valid_to_remove} record rimossi!")
                    except Exception as e:
                        st.error(f"Errore eliminazione automatica: {e}")
                        st.exception(e)
                        st.session_state.report_data_to_download = None
                        st.session_state.report_filename_to_download = None
                else:
                    st.info("Nessun record suggerito valido da rimuovere.")
                    
        with col_report_dl:
            if st.session_state.report_data_to_download is not None and st.session_state.duplicates_removed:
                st.download_button(
                    label=f"📥 Scarica Report Eliminati ({st.session_state.report_filename_to_download})",
                    data=st.session_state.report_data_to_download,
                    file_name=st.session_state.report_filename_to_download,
                    mime="text/csv",
                    key="dl_deleted_report"
                )
                
        st.markdown("---")
        st.subheader("Revisione Manuale e Rimozione Selezionata")
        st.info("Esamina i gruppi qui sotto...")
        
        selected_indices = select_duplicates_to_remove_ui(duplicates_df_display)
        st.session_state.selected_indices_to_drop = selected_indices
        
        st.divider()
        num_selected = len(st.session_state.selected_indices_to_drop)
        manual_button_disabled = st.session_state.duplicates_removed or num_selected == 0
        
        if num_selected > 0:
            st.write(f"🗑️ Selezionati manualmente: **{num_selected}** record.")
        else:
            st.write("⚠️ Nessun record selezionato manualmente.")
            
        if st.button("Rimuovi i Record Selezionati Manualmente", key="confirm_remove_manual", 
                    disabled=manual_button_disabled, 
                    help="Rimuove solo i record con 'Elimina?' spuntata."):
            selected_to_remove = st.session_state.selected_indices_to_drop
            
            if selected_to_remove:
                try:
                    current_df_manual = df_main
                    
                    if not current_df_manual.empty:
                        valid_indices_to_remove = [idx for idx in selected_to_remove if idx in current_df_manual.index]
                        num_actually_removed = len(valid_indices_to_remove)
                        
                        if num_actually_removed > 0:
                            with st.spinner(f"Rimozione di {num_actually_removed} record..."):
                                df_cleaned = current_df_manual.drop(index=valid_indices_to_remove)
                                st.session_state.processed_df = df_cleaned
                                st.session_state.duplicates_removed = True
                                st.session_state.duplicate_detection_results = (pd.DataFrame(), [], [])
                                st.session_state.selected_indices_to_drop = []
                                st.session_state.report_data_to_download = None
                                st.session_state.report_filename_to_download = None
                                
                            st.success(f"{num_actually_removed} record rimossi!")
                            st.info("Dati aggiornati. Ricaricamento interfaccia...")
                            st.rerun()
                        else:
                            st.warning("Nessuno degli indici selezionati trovato.")
                    else:
                        st.error("DataFrame non trovato.")
                except Exception as e:
                    st.error(f"Errore rimozione manuale: {e}")
                    st.exception(e)
            else:
                st.info("Nessun record selezionato.")
                
        st.divider()
        with st.expander("📄 Scarica Report Tutti i Duplicati Identificati (Prima della rimozione)"):
            duplicates_df_display_orig = st.session_state.duplicate_detection_results[0]
            
            if not duplicates_df_display_orig.empty:
                cols_show_dup_preferred = ['GruppoDuplicati', 'CodiceFiscale', 'Nome', 'Cognome', 
                                          'DataPresenza', 'OraPresenza', 'PercorsoOriginaleInternal', 
                                          'PercorsoOriginaleSenzaArt13Internal','PercorsoInternal', 
                                          'DenominazioneAttività', 'DenominazioneAttivitaNormalizzataInternal', 
                                          'CodicePercorso', 'CFU', 'SuggerisciRimuovere', 'TimestampPresenza']
                                          
                cols_show_dup_exist = [c for c in cols_show_dup_preferred if c in duplicates_df_display_orig.columns]
                
                if 'GruppoDuplicati' not in duplicates_df_display_orig.columns:
                    duplicates_df_display_orig['GruppoDuplicati'] = 0
                    
                duplicates_csv_orig = duplicates_df_display_orig[cols_show_dup_exist].to_csv(index=True, index_label='OriginalIndex').encode('utf-8')
                st.download_button(
                    label="Scarica CSV Cluster Identificati",
                    data=duplicates_csv_orig,
                    file_name="record_duplicati_cluster_identificati.csv",
                    mime="text/csv",
                    key="dl_involved_clusters_orig"
                )
            else:
                st.info("Nessun duplicato identificato.")
    else:
        if st.session_state.duplicates_removed:
            st.success("I duplicati sono già stati rimossi.")
        else:
            st.success("✅ Nessun record potenzialmente duplicato trovato.")
            

def select_duplicates_to_remove_ui(df_duplicates):
    """Crea un'interfaccia per selezionare i duplicati da rimuovere"""
    if df_duplicates.empty:
        return []
        
    selected_indices_map = {}
    
    if 'GruppoDuplicati' not in df_duplicates.columns:
        st.error("Colonna 'GruppoDuplicati' mancante.")
        return []
        
    all_groups = sorted([g for g in df_duplicates['GruppoDuplicati'].unique() if g != 0])
    if not all_groups:
        return []
        
    st.write(f"👇 **Revisiona e seleziona record da eliminare:**")
    for group_id in all_groups:
        group_df = df_duplicates[df_duplicates['GruppoDuplicati'] == group_id].copy()
        
        if len(group_df) < 2:
            continue
            
        first_row = group_df.iloc[0]
        cf = first_row.get('CodiceFiscale', 'N/A')
        date_obj = first_row.get('DataPresenza')
        date_str = date_obj.strftime('%d/%m/%Y') if pd.notna(date_obj) and hasattr(date_obj, 'strftime') else 'Data N/A'
        
        expander_label = f"Gruppo {int(group_id)}: CF {cf} - Data {date_str} ({len(group_df)} record)"
        with st.expander(expander_label, expanded=False):
            st.write(f"**CF:** {cf}, **Data:** {date_str}")
            
            group_df_edit = group_df.copy()
            group_df_edit['Elimina'] = group_df_edit['SuggerisciRimuovere']
            
            cols_to_display_editor = ['Elimina', 'OraPresenza', 'PercorsoInternal']
            cols_to_display_editor.extend([c for c in ['Nome','Cognome','DenominazioneAttività'] if c in group_df_edit.columns])
            cols_to_display_editor = [c for c in cols_to_display_editor if c in group_df_edit.columns]
            
            group_df_edit['_OriginalIndex'] = group_df_edit.index
            
            if '_OriginalIndex' in group_df_edit.columns and 'Elimina' in cols_to_display_editor:
                edited_df = st.data_editor(
                    group_df_edit[cols_to_display_editor + ['_OriginalIndex']],
                    column_config={
                        "Elimina": st.column_config.CheckboxColumn("Elimina?", default=False),
                        "OraPresenza": st.column_config.TimeColumn("Ora", format="HH:mm:ss"),
                        "PercorsoInternal": st.column_config.TextColumn("Percorso (Elab.)"),
                        "DenominazioneAttività": st.column_config.TextColumn("Attività"),
                        "Nome": st.column_config.TextColumn("Nome"),
                        "Cognome": st.column_config.TextColumn("Cognome"),
                        "_OriginalIndex": None
                    },
                    disabled=[c for c in cols_to_display_editor if c != 'Elimina'],
                    hide_index=True,
                    key=f"editor_group_{group_id}"
                )
                
                selected_in_group = edited_df[edited_df['Elimina']]['_OriginalIndex'].tolist()
                for idx in group_df_edit['_OriginalIndex']:
                    selected_indices_map[idx] = idx in selected_in_group
            else:
                st.warning(f"Dati incompleti per editor gruppo {group_id}.")
                
    return [idx for idx, selected in selected_indices_map.items() if selected]
