# Interfaccia utente per la Tab 2 (Gestione Duplicati)
import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
# Importa il modulo duplicates per la gestione dei duplicati
from modules.duplicates import detect_duplicate_records

def ensure_unique_columns(df):
    """
    Assicura che le colonne di un dataframe abbiano nomi univoci.
    Rinomina le colonne duplicate aggiungendo un suffisso numerico.
    """
    cols = df.columns.tolist()
    unique_cols = []
    col_counts = {}
    for col in cols:
        if col in col_counts:
            col_counts[col] += 1
            unique_cols.append(f"{col}_{col_counts[col]}")
        else:
            col_counts[col] = 0
            unique_cols.append(col)
    
    if cols != unique_cols:
        st.warning("Rilevate e rinominate colonne duplicate nel DataFrame.")
        df.columns = unique_cols
    return df

def render_tab2(df_main):
    """Renderizza l'interfaccia della Tab 2: Gestione Duplicati"""
    st.header("Gestione Record Potenzialmente Duplicati")
    st.markdown("Identifica cluster di timbrature ravvicinate (entro due ore) per stesso Nome, Cognome e Denominazione Attività.")

    # Inizializza lo stato se non esiste
    if 'duplicates_removed' not in st.session_state:
        st.session_state.duplicates_removed = False
    if 'duplicate_detection_results' not in st.session_state:
        # Inizializza con tuple di strutture vuote per evitare errori
        st.session_state.duplicate_detection_results = (pd.DataFrame(), [], [])
    if 'selected_indices_to_drop' not in st.session_state:
        st.session_state.selected_indices_to_drop = []
    if 'report_data_to_download' not in st.session_state:
        st.session_state.report_data_to_download = None
    if 'report_filename_to_download' not in st.session_state:
        st.session_state.report_filename_to_download = None
    if 'processed_df' not in st.session_state:
        # Inizialmente punta a df_main o è None se df_main è None/vuoto
         st.session_state.processed_df = df_main.copy() if df_main is not None and not df_main.empty else pd.DataFrame()
         
    # Usa il dataframe processato se disponibile, altrimenti quello principale
    current_df = st.session_state.processed_df if not st.session_state.processed_df.empty else df_main

    # Esegui il rilevamento solo se non sono stati rimossi duplicati E non ci sono risultati esistenti validi
    if not st.session_state.duplicates_removed and st.session_state.duplicate_detection_results[0].empty:
        if current_df is not None and not current_df.empty:
            with st.spinner("Rilevamento duplicati in corso..."):
                # Assicura che le colonne necessarie esistano
                required_dup_cols = ['TimestampPresenza', 'Nome', 'Cognome', 'DenominazioneAttività']
                if all(col in current_df.columns for col in required_dup_cols):
                     # Assicura colonne uniche prima del rilevamento
                    df_per_detect = ensure_unique_columns(current_df.copy())
                    st.session_state.duplicate_detection_results = detect_duplicate_records(df_per_detect)
                else:
                    missing_cols = [col for col in required_dup_cols if col not in current_df.columns]
                    st.warning(f"Colonne necessarie per il rilevamento duplicati ({', '.join(missing_cols)}) non trovate nel DataFrame.")
                    st.session_state.duplicate_detection_results = (pd.DataFrame(), [], [])
        else:
             # Se il dataframe è vuoto, resetta i risultati
             st.session_state.duplicate_detection_results = (pd.DataFrame(), [], [])

    # Estrai i risultati dal session state
    duplicates_df_display, involved_indices, indices_to_drop_suggested = st.session_state.duplicate_detection_results
    duplicates_found = not duplicates_df_display.empty

    if duplicates_found and not st.session_state.duplicates_removed:
        st.warning(f"Trovati **{len(duplicates_df_display)}** record potenzialmente duplicati in **{duplicates_df_display['GruppoDuplicati'].nunique()}** cluster.")
        st.markdown("---")
        
        # --- Sezione Azione Rapida ---
        st.subheader("Azione Rapida: Eliminazione Automatica Suggerita")
        
        indices_to_drop_suggested_auto = indices_to_drop_suggested # Usa l'output diretto della funzione
        current_df_auto = st.session_state.processed_df # Usa sempre il df processato per l'azione
        
        valid_indices_to_remove_auto = []
        if not current_df_auto.empty and indices_to_drop_suggested_auto:
            # Filtra gli indici suggeriti per assicurarsi che esistano ancora nel dataframe corrente
            valid_indices_to_remove_auto = [idx for idx in indices_to_drop_suggested_auto if idx in current_df_auto.index]
            
        num_valid_to_remove = len(valid_indices_to_remove_auto)
        # Disabilita se i duplicati sono già stati rimossi o non ci sono duplicati validi da rimuovere
        auto_remove_disabled = st.session_state.duplicates_removed or num_valid_to_remove == 0
        
        col_btn_auto, col_report_dl = st.columns([2, 3])
        
        with col_btn_auto:
            if st.button(f"Elimina {num_valid_to_remove} Record Suggeriti", key="auto_remove_and_report", 
                         disabled=auto_remove_disabled, 
                         help="Rimuove automaticamente i record suggeriti (tutti tranne il primo per cluster) e prepara un report CSV degli elementi eliminati."):
                
                if valid_indices_to_remove_auto:
                    try:
                        with st.spinner(f"Eliminazione di {num_valid_to_remove} record e preparazione report..."):
                            # 1. Prepara il report PRIMA di modificare il dataframe
                            df_deleted_report = current_df_auto.loc[valid_indices_to_remove_auto].copy()
                            duplicates_df_display_report = st.session_state.duplicate_detection_results[0] # Usa i risultati del rilevamento corrente
                            
                            # Aggiungi 'GruppoDuplicati' al report, se possibile
                            if not duplicates_df_display_report.empty and 'GruppoDuplicati' in duplicates_df_display_report.columns:
                                try:
                                    # Assicurati che gli indici esistano nel df dei duplicati
                                    valid_report_indices = duplicates_df_display_report.index.intersection(valid_indices_to_remove_auto)
                                    if not valid_report_indices.empty:
                                         # Assegna il GruppoDuplicati solo per gli indici validi trovati
                                         df_deleted_report['GruppoDuplicati'] = duplicates_df_display_report.loc[valid_report_indices, 'GruppoDuplicati']
                                    else:
                                         df_deleted_report['GruppoDuplicati'] = 0 # Default se non ci sono corrispondenze
                                except KeyError as ke:
                                    st.warning(f"KeyError durante l'aggiunta di GruppoDuplicati al report: {ke}. Il gruppo potrebbe mancare.")
                                    if 'GruppoDuplicati' not in df_deleted_report.columns:
                                         df_deleted_report['GruppoDuplicati'] = 0
                                except Exception as e:
                                     st.warning(f"Errore generico durante l'aggiunta di GruppoDuplicati al report: {e}. Il gruppo potrebbe mancare.")
                                     if 'GruppoDuplicati' not in df_deleted_report.columns:
                                         df_deleted_report['GruppoDuplicati'] = 0
                            else:
                                df_deleted_report['GruppoDuplicati'] = 0 # Aggiungi colonna default se non presente nei risultati
                                    
                            # Definisci colonne preferite per il report
                            cols_report_preferred = ['GruppoDuplicati', 'CodiceFiscale', 'Nome', 'Cognome', 
                                                    'DataPresenza', 'OraPresenza', 'Percorso', 
                                                    'DenominazioneAttività', 'DenominazioneAttivitaNormalizzataInternal', 
                                                    'CodicePercorso', 'TimestampPresenza']
                                                    
                            # Filtra per le colonne che esistono effettivamente nel report
                            cols_report_exist = [c for c in cols_report_preferred if c in df_deleted_report.columns]
                            df_deleted_report_final = df_deleted_report[cols_report_exist]
                            
                            # Ordina il report per Gruppo e Timestamp, se possibile
                            if 'GruppoDuplicati' in cols_report_exist and 'TimestampPresenza' in cols_report_exist:
                                df_deleted_report_final = df_deleted_report_final.sort_values(by=['GruppoDuplicati', 'TimestampPresenza'])
                            
                            # Crea il file CSV in memoria
                            report_csv_bytes = df_deleted_report_final.to_csv(index=True, index_label='OriginalIndex').encode('utf-8')
                            ts_report = datetime.now().strftime("%Y%m%d_%H%M")
                            report_filename = f"Report_Record_Eliminati_Auto_{ts_report}.csv"
                            
                            # Salva i dati del report nello stato sessione per il download
                            st.session_state.report_data_to_download = report_csv_bytes
                            st.session_state.report_filename_to_download = report_filename
                            
                            # 2. Esegui la rimozione dal dataframe principale (usando current_df_auto)
                            df_cleaned = current_df_auto.drop(index=valid_indices_to_remove_auto)
                            
                            # 3. Aggiorna lo stato della sessione
                            st.session_state.processed_df = df_cleaned # Aggiorna il dataframe processato
                            st.session_state.duplicates_removed = True # Segna che i duplicati sono stati gestiti
                            # Resetta i risultati del rilevamento perché il df è cambiato
                            st.session_state.duplicate_detection_results = (pd.DataFrame(), [], []) 
                            st.session_state.selected_indices_to_drop = [] # Resetta selezione manuale
                            
                        st.success(f"{num_valid_to_remove} record rimossi automaticamente!")
                        # Non serve rerun qui, il download button apparirà sotto e l'interfaccia si aggiornerà
                        # st.rerun() # Rerun potrebbe essere necessario se l'UI dipende da `duplicates_removed` per nascondere sezioni
                        st.rerun() # Forza aggiornamento UI
                        
                    except Exception as e:
                        st.error(f"Errore durante l'eliminazione automatica: {e}")
                        st.exception(e) # Mostra il traceback per debugging
                        # Resetta i dati del report in caso di errore
                        st.session_state.report_data_to_download = None
                        st.session_state.report_filename_to_download = None
                else:
                    st.info("Nessun record suggerito valido da rimuovere trovato nel DataFrame corrente.")
                    
        with col_report_dl:
             # Mostra il pulsante di download SOLO se ci sono dati pronti E i duplicati sono stati marcati come rimossi (dall'azione auto)
             # Questo previene che appaia dopo una rimozione manuale senza report generato
            if st.session_state.report_data_to_download is not None and st.session_state.duplicates_removed and st.session_state.report_filename_to_download:
                st.download_button(
                    label=f"📥 Scarica Report Eliminati ({st.session_state.report_filename_to_download})",
                    data=st.session_state.report_data_to_download,
                    file_name=st.session_state.report_filename_to_download,
                    mime="text/csv",
                    key="dl_deleted_report_auto" # Chiave unica per questo bottone
                )
        
        st.markdown("---")
        
        # --- Sezione Revisione Manuale ---
        st.subheader("Revisione Manuale e Rimozione Selezionata")
        st.info("Esamina i gruppi qui sotto. Puoi modificare la selezione predefinita ('Elimina?') e poi confermare la rimozione.")
        
        # Interfaccia per selezionare gli indici da rimuovere
        selected_indices = select_duplicates_to_remove_ui(duplicates_df_display)
        # Nota: `selected_indices_to_drop` nello stato sessione viene aggiornato DENTRO `select_duplicates_to_remove_ui` indirettamente tramite le chiavi dei widget
        # Lo rileggiamo qui per chiarezza e per usarlo nel bottone
        st.session_state.selected_indices_to_drop = selected_indices 

        st.divider()
        num_selected = len(selected_indices)
        st.write(f"**{num_selected}** record attualmente selezionati per la rimozione manuale.")

        # Bottone per la rimozione manuale
        # Correzione dell'indentazione: il blocco 'if st.button...' e il suo contenuto devono essere correttamente indentati
        # Il bottone va allineato con st.subheader, st.info, ecc. (livello 2 in questo contesto)
        if st.button(f"Rimuovi {num_selected} Record Selezionati Manualmente", 
                     key="manual_remove_button", 
                     disabled=not selected_indices or st.session_state.duplicates_removed): # Disabilita se niente selezionato o se già rimossi (auto o manuale)
            
            # Dentro l'azione del bottone (livello 3)
            if selected_indices: 
                try:
                    # Usa il dataframe CORRENTE processato per la rimozione
                    current_df_manual = st.session_state.processed_df 
                    
                    if not current_df_manual.empty:
                        # Verifica che gli indici selezionati esistano ANCORA nel dataframe
                        valid_indices_to_remove = [idx for idx in selected_indices if idx in current_df_manual.index]
                        num_actually_removed = len(valid_indices_to_remove)

                        if valid_indices_to_remove:
                            # Dentro la condizione di indici validi (livello 4)
                            # Uso dello spinner (livello 5)
                            with st.spinner(f"Rimozione di {num_actually_removed} record selezionati..."):
                                # Azioni di rimozione e aggiornamento stato (livello 6)
                                df_cleaned = current_df_manual.drop(index=valid_indices_to_remove)
                                st.session_state.processed_df = df_cleaned # Aggiorna il dataframe
                                st.session_state.duplicates_removed = True # Segna come rimossi
                                # Resetta i risultati del rilevamento e la selezione
                                st.session_state.duplicate_detection_results = (pd.DataFrame(), [], []) 
                                st.session_state.selected_indices_to_drop = [] 
                                # La rimozione manuale non genera un report qui, quindi resetta i dati del report
                                st.session_state.report_data_to_download = None
                                st.session_state.report_filename_to_download = None
                                
                            # Messaggi di successo/info (livello 5)
                            st.success(f"{num_actually_removed} record rimossi manualmente!")
                            st.info("Dati aggiornati. Ricaricamento dell'interfaccia per riflettere le modifiche...")
                            st.rerun() # Rerun per aggiornare l'UI e nascondere le sezioni dei duplicati
                        else:
                            # Livello 5
                            st.warning("Nessuno degli indici selezionati è stato trovato nel DataFrame corrente. Potrebbero essere già stati rimossi.")
                    else:
                        # Livello 4
                        st.error("Il DataFrame processato è vuoto. Impossibile eseguire la rimozione.")
                except Exception as e:
                    # Livello 3
                    st.error(f"Errore durante la rimozione manuale: {e}")
                    st.exception(e) # Mostra traceback
            else:
                 # Livello 3 (corrisponde a 'if selected_indices:')
                 # Questo non dovrebbe accadere se il bottone è disabilitato correttamente
                 st.info("Nessun record era selezionato per la rimozione.")
                
        st.divider()
        
        # --- Sezione Download Report Completo (prima della rimozione) ---
        with st.expander("📄 Scarica Report Tutti i Duplicati Identificati (Prima della rimozione)"):
            # Usa i risultati originali del rilevamento PRIMA di qualsiasi rimozione
            duplicates_df_display_orig = st.session_state.duplicate_detection_results[0] 
            
            if not duplicates_df_display_orig.empty:
                # Colonne preferite per il report completo dei cluster
                cols_show_dup_preferred = ['GruppoDuplicati', 'CodiceFiscale', 'Nome', 'Cognome', 
                                          'DataPresenza', 'OraPresenza', 'Percorso', 
                                          'DenominazioneAttività', 'DenominazioneAttivitaNormalizzataInternal', 
                                          'CodicePercorso', 'CFU', 'SuggerisciRimuovere', 'TimestampPresenza']
                                          
                # Filtra per le colonne esistenti nel dataframe dei duplicati
                cols_show_dup_exist = [c for c in cols_show_dup_preferred if c in duplicates_df_display_orig.columns]
                
                # Assicura che 'GruppoDuplicati' esista, anche se vuoto, per coerenza
                if 'GruppoDuplicati' not in duplicates_df_display_orig.columns:
                    st.warning("Colonna 'GruppoDuplicati' non presente nei risultati originali, verrà aggiunta come 0.")
                    duplicates_df_display_orig['GruppoDuplicati'] = 0
                    if 'GruppoDuplicati' not in cols_show_dup_exist:
                         cols_show_dup_exist.insert(0, 'GruppoDuplicati') # Aggiungi all'inizio se mancava
                    
                # Ordina per gruppo e timestamp se possibile
                if 'GruppoDuplicati' in cols_show_dup_exist and 'TimestampPresenza' in cols_show_dup_exist:
                     df_to_download = duplicates_df_display_orig[cols_show_dup_exist].sort_values(by=['GruppoDuplicati', 'TimestampPresenza'])
                else:
                     df_to_download = duplicates_df_display_orig[cols_show_dup_exist]

                # Crea CSV per il download
                duplicates_csv_orig = df_to_download.to_csv(index=True, index_label='OriginalIndex').encode('utf-8')
                ts_download = datetime.now().strftime("%Y%m%d_%H%M")
                download_filename = f"Report_Duplicati_Identificati_{ts_download}.csv"

                st.download_button(
                    label="Scarica CSV Cluster Duplicati Identificati",
                    data=duplicates_csv_orig,
                    file_name=download_filename,
                    mime="text/csv",
                    key="dl_involved_clusters_orig" # Chiave unica
                )
            else:
                st.info("Non ci sono dati sui duplicati identificati da scaricare (potrebbero essere stati già rimossi o non rilevati).")

    # --- Messaggi finali ---
    else:
        # Se siamo qui, o non sono stati trovati duplicati inizialmente, o sono stati rimossi
        if st.session_state.duplicates_removed:
            st.success("✅ I record potenzialmente duplicati sono stati gestiti e rimossi.")
            # Potresti voler offrire un modo per resettare/rianalizzare se necessario
            if st.button("Rianalizza per duplicati (se necessario)", key="reanalyze_duplicates"):
                 st.session_state.duplicates_removed = False
                 st.session_state.duplicate_detection_results = (pd.DataFrame(), [], [])
                 st.session_state.report_data_to_download = None # Pulisce anche il report scaricabile
                 st.session_state.report_filename_to_download = None
                 st.rerun()

        else:
            # Nessun duplicato trovato inizialmente
            st.success("✅ Nessun record potenzialmente duplicato trovato nei dati correnti.")

def select_duplicates_to_remove_ui(df_duplicates):
    """
    Crea un'interfaccia utente con st.data_editor per selezionare 
    i record duplicati da rimuovere all'interno di ciascun gruppo.
    Restituisce una lista di indici originali selezionati per la rimozione.
    """
    if df_duplicates.empty:
        # st.info("Nessun gruppo di duplicati da revisionare.")
        return []
        
    selected_indices_map = {} # Mappa per tenere traccia della selezione {original_index: bool_elimina}
    
    # Verifica colonna essenziale
    if 'GruppoDuplicati' not in df_duplicates.columns:
        st.error("Errore interno: Colonna 'GruppoDuplicati' mancante nel dataframe dei duplicati.")
        return []
        
    # Ottieni ID unici dei gruppi, escludendo eventuali 0 o NaN se usati come 'non gruppo'
    all_groups = sorted([g for g in df_duplicates['GruppoDuplicati'].unique() if pd.notna(g) and g != 0])
    
    if not all_groups:
        st.info("Nessun gruppo di duplicati valido trovato per la revisione.")
        return []
        
    st.write(f"👇 **Revisiona i {len(all_groups)} gruppi e seleziona i record da eliminare:**")
    
    indices_selected_overall = []

    for group_id in all_groups:
        # Filtra il dataframe per il gruppo corrente
        group_df = df_duplicates[df_duplicates['GruppoDuplicati'] == group_id].copy()
        
        # Salta gruppi con meno di 2 record, non sono duplicati in senso stretto
        if len(group_df) < 2:
            continue
            
        # Prendi informazioni dalla prima riga per l'etichetta dell'expander
        first_row = group_df.iloc[0]
        # Gestisci DataPresenza che potrebbe essere date, datetime, o NaT
        date_obj = first_row.get('DataPresenza')
        date_str = 'Data N/D'
        if pd.notna(date_obj):
             if hasattr(date_obj, 'strftime'): # Check se è date o datetime
                 date_str = date_obj.strftime('%d/%m/%Y')
             else: # Altrimenti prova a convertire
                 try:
                     date_str = pd.to_datetime(date_obj).strftime('%d/%m/%Y')
                 except:
                     date_str = str(date_obj) # Fallback a stringa

        expander_label = f"Gruppo {int(group_id)}: {first_row.get('Nome', 'N/D')} {first_row.get('Cognome', 'N/D')} - Data {date_str} ({len(group_df)} record)"
        
        with st.expander(expander_label, expanded=False):
            st.markdown(f"**Nome:** `{first_row.get('Nome', 'N/D')}` **Cognome:** `{first_row.get('Cognome', 'N/D')}` - **Data:** `{date_str}` - **Attività:** `{first_row.get('DenominazioneAttività', 'N/D')}`")
            
            # Prepara il dataframe per l'editor
            group_df_edit = group_df.copy()
            
            # La colonna 'Elimina' deve basarsi sul suggerimento iniziale ('SuggerisciRimuovere')
            if 'SuggerisciRimuovere' in group_df_edit.columns:
                 group_df_edit['Elimina'] = group_df_edit['SuggerisciRimuovere'].astype(bool)
            else:
                 st.warning(f"Colonna 'SuggerisciRimuovere' non trovata per gruppo {group_id}. Selezione 'Elimina' inizializzata a False.")
                 group_df_edit['Elimina'] = False # Default a False se manca il suggerimento
            
            # Colonne da mostrare nell'editor: Priorità a Elimina, Ora, Attività, poi Nome/Cognome se esistono
            cols_to_display_editor = ['Elimina', 'OraPresenza', 'DenominazioneAttività']
            # Aggiungi nome/cognome se presenti
            for col in ['Nome', 'Cognome']:
                 if col in group_df_edit.columns:
                     cols_to_display_editor.append(col)
                     
            # Filtra per colonne che esistono effettivamente
            cols_to_display_editor_final = [c for c in cols_to_display_editor if c in group_df_edit.columns]
            
            # Assicura unicità se ci fossero duplicati nei nomi (improbabile qui ma sicuro)
            cols_to_display_editor_final = list(dict.fromkeys(cols_to_display_editor_final))
            
            # Conserva l'indice originale per riferimento
            group_df_edit['_OriginalIndex'] = group_df_edit.index
            
            # Verifica che le colonne chiave esistano prima di chiamare data_editor
            if '_OriginalIndex' in group_df_edit.columns and 'Elimina' in cols_to_display_editor_final:
                try:
                     # Configurazione delle colonne per data_editor
                    column_config = {
                        "Elimina": st.column_config.CheckboxColumn("Elimina?", default=False, help="Seleziona per rimuovere questo record"),
                        "_OriginalIndex": None # Nasconde la colonna indice originale
                    }
                    # Aggiungi configurazioni specifiche per altre colonne comuni
                    if 'OraPresenza' in cols_to_display_editor_final:
                        column_config["OraPresenza"] = st.column_config.TimeColumn("Ora", format="HH:mm:ss")
                    if 'DenominazioneAttività' in cols_to_display_editor_final:
                         column_config["DenominazioneAttività"] = st.column_config.TextColumn("Denominazione Attività")
                    if 'Nome' in cols_to_display_editor_final:
                         column_config["Nome"] = st.column_config.TextColumn("Nome")
                    if 'Cognome' in cols_to_display_editor_final:
                         column_config["Cognome"] = st.column_config.TextColumn("Cognome")

                    # Colonne da disabilitare (tutte tranne 'Elimina')
                    disabled_cols = [c for c in cols_to_display_editor_final if c != 'Elimina']

                    # Mostra il data editor
                    edited_df = st.data_editor(
                        group_df_edit[cols_to_display_editor_final + ['_OriginalIndex']], # Include indice per recuperarlo
                        column_config=column_config,
                        disabled=disabled_cols,
                        hide_index=True, # Nasconde l'indice di default del dataframe modificato
                        key=f"editor_group_{group_id}" # Chiave unica per ogni editor di gruppo
                    )
                    
                    # Aggiorna la mappa generale con le selezioni di questo gruppo
                    # Estrai gli indici originali dove 'Elimina' è True nel dataframe modificato
                    selected_in_group = edited_df[edited_df['Elimina']]['_OriginalIndex'].tolist()
                    
                    # Aggiorna la lista generale degli indici selezionati
                    indices_selected_overall.extend(selected_in_group)

                except Exception as e:
                     st.error(f"Errore durante la creazione dell'editor per il gruppo {group_id}: {e}")
                     st.exception(e) # Log completo per debug

            else:
                st.warning(f"Dati incompleti o errati per visualizzare l'editor del gruppo {group_id}. Colonne richieste: '_OriginalIndex', 'Elimina'. Colonne disponibili: {', '.join(group_df_edit.columns)}")
                # Mostra comunque i dati in forma statica se l'editor non può essere creato
                st.dataframe(group_df[[c for c in cols_to_display_editor if c in group_df.columns]])

    # Rimuovi eventuali duplicati dagli indici selezionati (se un indice fosse aggiunto più volte)
    # e restituisci la lista univoca
    return list(dict.fromkeys(indices_selected_overall))

# --- Esempio di utilizzo (se eseguito come script principale) ---
if __name__ == '__main__':
    
    # Crea un DataFrame di esempio
    data = {
        'TimestampPresenza': pd.to_datetime([
            '2023-10-26 08:00:00', '2023-10-26 08:05:00', '2023-10-26 08:06:00', # Gruppo 1
            '2023-10-26 09:30:00',                                             # Singolo
            '2023-10-27 14:00:00', '2023-10-27 14:08:00',                       # Gruppo 2
            '2023-10-27 17:00:00', '2023-10-27 18:00:00'                        # Non duplicati (distanti)
        ]),
        'CodiceFiscale': [
            'AAA', 'AAA', 'AAA', 
            'BBB', 
            'CCC', 'CCC',
            'AAA', 'AAA'
        ],
        'Nome': ['Mario', 'Mario', 'Mario', 'Luigi', 'Anna', 'Anna', 'Mario', 'Mario'],
        'Cognome': ['Rossi', 'Rossi', 'Rossi', 'Verdi', 'Bianchi', 'Bianchi', 'Rossi', 'Rossi'],
        'DenominazioneAttività': ['Lezione A', 'Lezione A - rep', 'Lezione A - rep2', 'Lab B', 'Seminario C', 'Seminario C - rep', 'Lezione D', 'Lezione E']
    }
    sample_df = pd.DataFrame(data)
    # Aggiungi un indice non standard per testare la robustezza
    sample_df.index = [10, 20, 30, 40, 50, 60, 70, 80] 

    # Inizializza lo stato sessione se non già fatto
    if 'processed_df' not in st.session_state:
         st.session_state.processed_df = sample_df.copy() # Usa il df di esempio
    # Altre inizializzazioni necessarie potrebbero essere qui (vedi render_tab2)
    if 'duplicates_removed' not in st.session_state: st.session_state.duplicates_removed = False
    if 'duplicate_detection_results' not in st.session_state: st.session_state.duplicate_detection_results = (pd.DataFrame(), [], [])
    if 'selected_indices_to_drop' not in st.session_state: st.session_state.selected_indices_to_drop = []
    if 'report_data_to_download' not in st.session_state: st.session_state.report_data_to_download = None
    if 'report_filename_to_download' not in st.session_state: st.session_state.report_filename_to_download = None


    st.title("Test Interfaccia Tab 2 - Gestione Duplicati")
    
    # Mostra il dataframe iniziale nello stato
    st.subheader("DataFrame Attuale (in st.session_state.processed_df)")
    st.dataframe(st.session_state.processed_df)

    # Renderizza la tab
    render_tab2(st.session_state.processed_df) # Passa il df dallo stato

    st.subheader("Stato Sessione Corrente")
    st.json({
         "duplicates_removed": st.session_state.get('duplicates_removed', 'Non inizializzato'),
         "selected_indices_to_drop": st.session_state.get('selected_indices_to_drop', 'Non inizializzato'),
         "report_filename_to_download": st.session_state.get('report_filename_to_download', 'Non inizializzato'),
         "len_duplicate_detection_results[0]": len(st.session_state.get('duplicate_detection_results', ([],[],[]))[0]),
         "len_processed_df": len(st.session_state.get('processed_df', pd.DataFrame()))
    })