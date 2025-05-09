# Interfaccia utente per la Tab 3 (Calcolo Presenze ed Esportazione)
import streamlit as st
import pandas as pd
import re
from datetime import datetime, date
from io import BytesIO
from modules.attendance import calculate_attendance
from modules.utils import ensure_string_columns

# Definisco le funzioni di utilità direttamente qui per evitare problemi di importazione
def clean_sheet_name(name, used_names=None):
    """Pulisce i nomi dei fogli Excel e li converte in maiuscolo"""
    # Converte in stringa, pulisce caratteri non validi e converte in maiuscolo
    name = re.sub(r'[\\/?*\[\]:]', '', str(name)).upper()
    
    # Limita la lunghezza a 31 caratteri (limite Excel)
    name = name[:31]
    
    # Gestisce i nomi duplicati aggiungendo un contatore progressivo
    if used_names is not None:
        original_name = name
        counter = 1
        while name.lower() in used_names:
            # Assicura che ci sia spazio per il contatore nel nome
            suffix = f"_{counter}"
            name = original_name[:31-len(suffix)] + suffix
            counter += 1
    
    return name
    
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
    st.header("📊 Calcolo Presenze ed Esportazione")
    
    # Utilizzo di container per una migliore organizzazione visiva
    info_container = st.container()
    with info_container:
        col1, col2 = st.columns([3, 1])
        with col1:
            st.write("Visualizzazione presenze per studente e percorso con possibilità di esportazione dettagliata.")
        with col2:
            if st.session_state.duplicates_removed: 
                st.success("Dati depurati ✅", icon="✅")

    current_df_for_tab3 = df_main

    if not current_df_for_tab3.empty:
        required_att_cols = ['CodiceFiscale', 'DataPresenza', 'OraPresenza', 'Nome', 'Cognome', 'DenominazioneAttività']
        if not all(col in current_df_for_tab3.columns for col in required_att_cols):
            missing_cols = [col for col in required_att_cols if col not in current_df_for_tab3.columns]
            st.error(f"Impossibile procedere: colonne mancanti ({', '.join(missing_cols)})")
            attendance_df = pd.DataFrame() # Resetta per sicurezza
        else:
            # Calcolo presenze usando la visualizzazione per studente e percorso
            with st.spinner("Calcolo delle presenze in corso..."):
                group_by = "studente"
                attendance_df = calculate_attendance(current_df_for_tab3, group_by=group_by)
            
            # Se abbiamo dati validi, mostriamo i filtri in un container ben organizzato
            if not attendance_df.empty:
                filter_container = st.container()
                with filter_container:
                    st.subheader("🔍 Filtri", divider="gray")
                    st.caption("I filtri consentono di ridurre i dati visualizzati nelle tabelle sottostanti")
                    # Nota: La colonna 'DenominazioneAttività' viene rinominata in 'Percorso (Senza Art.13)' 
                    # dalla funzione calculate_attendance quando group_by è "studente"
                    
                    p_col_internal_key = 'DenominazioneAttività'
            
                # La logica per il filtraggio 
                if not attendance_df.empty:
                    if p_col_internal_key not in current_df_for_tab3.columns:
                        st.error(f"Colonna chiave interna '{p_col_internal_key}' non trovata nei dati dettagliati.")
                    else:
                        # Inizializzazione dei dataframe filtrati
                        df_filtrato_concorso_agg = attendance_df.copy()
                        df_filtrato_concorso_detail = current_df_for_tab3.copy()
                        # Nota: Il filtro per Codice Classe di concorso è stato rimosso
                        
                        # --- Filtro per Denominazione Classe di concorso ---
                        st.divider()
                        filter_denom_concorso_col1, filter_denom_concorso_col2 = st.columns([3, 1])
                        denom_concorso_sel = "Tutte"
                        
                        try:
                            # Ottieni tutte le denominazioni classi di concorso uniche dal dataframe
                            if 'Codice_classe_di_concorso_e_denominazione' in df_filtrato_concorso_agg.columns:
                                denom_concorso_list = sorted([str(d) for d in df_filtrato_concorso_agg['Codice_classe_di_concorso_e_denominazione'].unique() if pd.notna(d)])
                                
                                # Visualizza il filtro per denominazione classe di concorso
                                with filter_denom_concorso_col1:
                                    denom_concorso_sel = st.selectbox(f"🏫 Seleziona Denominazione Classe di concorso:", 
                                                                   ["Tutte"] + denom_concorso_list, 
                                                                   key="filt_denom_concorso_tab3")
                                
                                # Mostra il numero di studenti per questa denominazione classe di concorso
                                if denom_concorso_sel != "Tutte":
                                    with filter_denom_concorso_col2:
                                        studenti_per_denom = len(df_filtrato_concorso_agg[df_filtrato_concorso_agg['Codice_classe_di_concorso_e_denominazione'] == denom_concorso_sel])
                                        st.metric("Studenti nella classe", studenti_per_denom)
                            else:
                                with filter_denom_concorso_col1:
                                    st.warning("Colonna 'Codice_classe_di_concorso_e_denominazione' non presente nei dati")
                        except Exception as e:
                            st.error(f"Errore nel filtro denominazione classe di concorso: {e}")
                        
                        # Filtraggio basato sulla denominazione classe di concorso
                        df_filtrato_denom_concorso_agg = df_filtrato_concorso_agg.copy()
                        df_filtrato_denom_concorso_detail = df_filtrato_concorso_detail.copy()
                        
                        if denom_concorso_sel != "Tutte" and 'Codice_classe_di_concorso_e_denominazione' in df_filtrato_denom_concorso_agg.columns:
                            df_filtrato_denom_concorso_agg = df_filtrato_denom_concorso_agg[df_filtrato_denom_concorso_agg['Codice_classe_di_concorso_e_denominazione'] == denom_concorso_sel].copy()
                            if 'Codice_classe_di_concorso_e_denominazione' in df_filtrato_denom_concorso_detail.columns:
                                df_filtrato_denom_concorso_detail = df_filtrato_denom_concorso_detail[df_filtrato_denom_concorso_detail['Codice_classe_di_concorso_e_denominazione'] == denom_concorso_sel].copy()
                        
                        # --- Filtro per Denominazione Attività (ora gerarchico) ---
                        st.divider()
                        filter_denominazione_col1, filter_denominazione_col2 = st.columns([3, 1])
                        denominazione_sel = "Tutte"
                        
                        try:
                            # Ottieni le denominazioni attività filtrate dal dataframe dettagliato
                            if 'DenominazioneAttività' in df_filtrato_denom_concorso_detail.columns:
                                denominazione_list = sorted([str(d) for d in df_filtrato_denom_concorso_detail['DenominazioneAttività'].unique() if pd.notna(d)])
                                
                                # Visualizza il filtro per denominazione attività
                                with filter_denominazione_col1:
                                    denominazione_sel = st.selectbox(f"📝 Seleziona Denominazione Attività:", 
                                                                   ["Tutte"] + denominazione_list, 
                                                                   key="filt_denominazione_tab3")
                                
                                # Mostra il numero di record per questa denominazione
                                if denominazione_sel != "Tutte":
                                    with filter_denominazione_col2:
                                        record_per_denominazione = len(df_filtrato_denom_concorso_detail[df_filtrato_denom_concorso_detail['DenominazioneAttività'] == denominazione_sel])
                                        st.metric("Record trovati", record_per_denominazione)
                            else:
                                with filter_denominazione_col1:
                                    st.warning("Colonna 'DenominazioneAttività' non presente nei dati")
                        except Exception as e:
                            st.error(f"Errore nel filtro denominazione attività: {e}")
                            
                        # Filtraggio basato sulla denominazione attività
                        df_filtrato_denominazione_agg = df_filtrato_denom_concorso_agg.copy()
                        df_filtrato_denominazione_detail = df_filtrato_denom_concorso_detail.copy()
                        
                        if denominazione_sel != "Tutte" and 'DenominazioneAttività' in df_filtrato_denominazione_detail.columns:
                            df_filtrato_denominazione_detail = df_filtrato_denominazione_detail[df_filtrato_denominazione_detail['DenominazioneAttività'] == denominazione_sel].copy()
                            # Per i dati aggregati, filtriamo in base ai codici fiscali che hanno quella denominazione
                            if not df_filtrato_denominazione_detail.empty and 'CodiceFiscale' in df_filtrato_denominazione_detail.columns and 'CodiceFiscale' in df_filtrato_denominazione_agg.columns:
                                codici_fiscali_filtrati = df_filtrato_denominazione_detail['CodiceFiscale'].unique()
                                df_filtrato_denominazione_agg = df_filtrato_denominazione_agg[df_filtrato_denominazione_agg['CodiceFiscale'].isin(codici_fiscali_filtrati)].copy()

                        # --- Filtro Studente (sempre disponibile) ---
                        st.divider()
                        filter_studente_col1, filter_studente_col2 = st.columns([3, 1])
                        stud_sel = "Tutti gli Studenti" # Default
                        
                        # Preparazione lista studenti per il filtro
                        if not df_filtrato_denominazione_agg.empty:
                            df_filtrato_denominazione_agg['StudentIdentifier'] = df_filtrato_denominazione_agg.apply(
                                lambda row: f"{row.get('Cognome','')} {row.get('Nome','')} ({row.get('CodiceFiscale','N/A')})".strip(), axis=1
                            )
                            student_list = sorted(df_filtrato_denominazione_agg['StudentIdentifier'].unique())
                            
                            # Statistica totale studenti
                            with filter_studente_col2:
                                st.metric("Studenti disponibili", len(student_list))
                            
                            # Filtro studenti con ricerca
                            with filter_studente_col1:
                                search_placeholder = "Cerca per nome o cognome..."
                                search_term = st.text_input("🔎 Cerca studente:", placeholder=search_placeholder, key="search_student")
                                
                                if search_term:
                                    search_term = search_term.lower()
                                    filtered_student_list = [s for s in student_list if search_term in s.lower()]
                                    st.caption(f"Trovati {len(filtered_student_list)} studenti su {len(student_list)}")
                                    student_options = ["Tutti gli Studenti"] + filtered_student_list
                                else:
                                    student_options = ["Tutti gli Studenti"] + student_list
                                
                                stud_sel = st.selectbox("👤 Seleziona Studente:", student_options, key="filt_stud_tab3_v8")
                        else:
                            with filter_studente_col1:
                                st.info(f"Nessun dato aggregato trovato con i filtri applicati.")
                                df_to_display_agg = pd.DataFrame()
                                df_to_display_detail = pd.DataFrame()
                        
                        # --- Applicazione Filtri in Sequenza ---
                        st.divider()
                        st.subheader("🔍 Risultati Filtrati", divider="gray")
                        
                        # I dati già filtrati per codice classe e denominazione
                        df_to_display_agg = df_filtrato_denominazione_agg.copy()
                        df_to_display_detail = df_filtrato_denominazione_detail.copy()
                        
                        # Contatori per i filtri applicati
                        num_record_dopo_filtro_codice = len(df_to_display_agg)
                        num_record_dopo_filtro_denom = len(df_to_display_agg)
                        
                        # Applica filtro per studente se selezionato
                        if stud_sel != "Tutti gli Studenti":
                            try:
                                selected_cf = re.search(r'\((.*?)\)', stud_sel).group(1)
                                df_to_display_agg = df_to_display_agg[df_to_display_agg['CodiceFiscale'] == selected_cf].copy()
                                df_to_display_detail = df_to_display_detail[df_to_display_detail['CodiceFiscale'] == selected_cf].copy()
                            except (AttributeError, IndexError):
                                st.warning("Formato studente non riconosciuto nel filtro.")
                        
                        # Contatore record dopo filtro studente
                        num_record_dopo_filtro_stud = len(df_to_display_agg)
                        
                        # Mostriamo statistiche sui filtri applicati
                        with st.expander("📊 Statistiche Filtri", expanded=False):
                            stats_col1, stats_col2, stats_col3 = st.columns(3)
                            with stats_col1:
                                st.metric("Dopo filtro codice", num_record_dopo_filtro_codice)
                            with stats_col2:
                                st.metric("Dopo filtro denom.", num_record_dopo_filtro_denom)
                            with stats_col3:
                                st.metric("Record finali", num_record_dopo_filtro_stud)

                        try:
                            # --- Visualizzazione Tabelle ---
                            if not df_to_display_agg.empty:
                                # Aggiungi un separatore e un container per le tabelle
                                tables_container = st.container()
                                
                                with tables_container:
                                    # Statistiche riassuntive
                                    if stud_sel == "Tutti gli Studenti":
                                        st.markdown("### 📊 Statistiche")
                                        stats_col1, stats_col2, stats_col3, stats_col4 = st.columns(4)
                                        
                                        with stats_col1:
                                            avg_presenze = df_to_display_agg['Presenze'].mean()
                                            st.metric("Media Presenze", f"{avg_presenze:.1f}")
                                        
                                        with stats_col2:
                                            max_presenze = df_to_display_agg['Presenze'].max()
                                            st.metric("Presenze Max", f"{max_presenze}")
                                            
                                        with stats_col3:
                                            min_presenze = df_to_display_agg['Presenze'].min()
                                            st.metric("Presenze Min", f"{min_presenze}")
                                        
                                        with stats_col4:
                                            tot_studenti = len(df_to_display_agg)
                                            st.metric("Totale Studenti", tot_studenti)
                                    
                                    # Titolo della tabella
                                    filtri_applicati = []
                                    if denom_concorso_sel != "Tutte":
                                        filtri_applicati.append(f"Classe di concorso: {denom_concorso_sel}")
                                    if denominazione_sel != "Tutte":
                                        filtri_applicati.append(f"Attività: {denominazione_sel}")
                                    if stud_sel != "Tutti gli Studenti":
                                        filtri_applicati.append(f"Studente: {stud_sel}")
                                    
                                    if filtri_applicati:
                                        filtri_text = " | ".join(filtri_applicati)
                                        st.subheader(f"📋 Riepilogo Aggregato - {filtri_text}", divider="blue")
                                    else:
                                        st.subheader("📋 Riepilogo Aggregato - Tutti i dati", divider="blue")

                                    cols_disp_agg = ['CodiceFiscale', 'Nome', 'Cognome', 'Email', 
                                                    'Codice_classe_di_concorso_e_denominazione', 'Dipartimento', 'Matricola',
                                                    'Percorso (Senza Art.13)', 'CFU Totali', 'Presenze']
                                    cols_disp_agg_exist = [c for c in cols_disp_agg if c in df_to_display_agg.columns]
                                    sort_agg_by = ['Percorso (Senza Art.13)', 'Cognome', 'Nome']
                                    
                                    # Opzioni di visualizzazione e ordinamento
                                    visual_options_col1, visual_options_col2 = st.columns(2)
                                    
                                    with visual_options_col1:
                                        sort_options = ["Presenze (decrescente)", "Cognome e Nome", "Denominazione Attività", "Denominazione Classe di concorso"]
                                        selected_sort = st.radio("Ordinamento tabella:", sort_options, horizontal=True)
                                    
                                    with visual_options_col2:
                                        highlight_opt = st.checkbox("Evidenzia valori critici", value=True, 
                                                                  help="Evidenzia studenti con poche presenze")
                                    
                                    # Applica l'ordinamento selezionato
                                    if selected_sort == "Presenze (decrescente)" and 'Presenze' in df_to_display_agg.columns:
                                        df_to_show = df_to_display_agg[cols_disp_agg_exist].sort_values(by=['Presenze'], ascending=False)
                                    elif selected_sort == "Cognome e Nome" and 'Cognome' in df_to_display_agg.columns:
                                        df_to_show = df_to_display_agg[cols_disp_agg_exist].sort_values(by=['Cognome', 'Nome'])
                                    elif selected_sort == "Denominazione Attività" and 'Percorso (Senza Art.13)' in df_to_display_agg.columns:
                                        df_to_show = df_to_display_agg[cols_disp_agg_exist].sort_values(by=['Percorso (Senza Art.13)', 'Cognome', 'Nome'])
                                    elif selected_sort == "Denominazione Classe di concorso" and 'Codice_classe_di_concorso_e_denominazione' in df_to_display_agg.columns:
                                        df_to_show = df_to_display_agg[cols_disp_agg_exist].sort_values(by=['Codice_classe_di_concorso_e_denominazione', 'Cognome', 'Nome'])
                                    elif sort_agg_by:
                                        valid_sort_agg_by = [c for c in sort_agg_by if c in df_to_display_agg.columns]
                                        if valid_sort_agg_by:
                                            df_to_show = df_to_display_agg[cols_disp_agg_exist].sort_values(by=valid_sort_agg_by)
                                        else:
                                            df_to_show = df_to_display_agg[cols_disp_agg_exist]
                                    else:
                                        df_to_show = df_to_display_agg[cols_disp_agg_exist]                                        # Converti la colonna Matricola in stringa per evitare errori di Arrow
                                    df_to_show = ensure_string_columns(df_to_show)

                                        # Dataframe con styling condizionale
                                    if highlight_opt and 'Presenze' in df_to_show.columns:
                                        # Crea una maschera per le presenze basse (< 4)
                                        def highlight_low_attendance(val):
                                            if isinstance(val, (int, float)) and val < 4:
                                                return 'background-color: #ffcccc'
                                            return ''
                                        
                                        # Applica lo styling solo alla colonna Presenze
                                        # Sostituito applymap con map (non più deprecato)
                                        styled_df = df_to_show.style.map(
                                            highlight_low_attendance, subset=['Presenze']
                                        )
                                        st.dataframe(styled_df, use_container_width=True)
                                    else:
                                        st.dataframe(df_to_show, use_container_width=True)


                            if not df_to_display_detail.empty:
                                # Sezione per il dettaglio delle presenze
                                detail_container = st.container()
                                with detail_container:
                                    # Titolo della sezione dettaglio con filtri applicati
                                    filtri_applicati = []
                                    if denom_concorso_sel != "Tutte":
                                        filtri_applicati.append(f"Classe di concorso: {denom_concorso_sel}")
                                    if denominazione_sel != "Tutte":
                                        filtri_applicati.append(f"Attività: {denominazione_sel}")
                                    if stud_sel != "Tutti gli Studenti":
                                        filtri_applicati.append(f"Studente: {stud_sel}")
                                    
                                    if filtri_applicati:
                                        filtri_text = " | ".join(filtri_applicati)
                                        st.subheader(f"📝 Dettaglio Record Presenze - {filtri_text}", divider="blue")
                                    else:
                                        st.subheader("📝 Dettaglio Record Presenze - Tutti i dati", divider="blue")
                                    
                                    # Riepilogo record trovati e statistiche dettaglio
                                    record_count = len(df_to_display_detail)
                                    if record_count > 0:
                                        detail_stats_col1, detail_stats_col2 = st.columns(2)
                                        
                                        with detail_stats_col1:
                                            st.info(f"Trovati {record_count} record di presenza", icon="ℹ️")
                                            
                                        # Se ci sono date, mostra il periodo
                                        if 'DataPresenza' in df_to_display_detail.columns:
                                            with detail_stats_col2:
                                                try:
                                                    min_data = pd.to_datetime(df_to_display_detail['DataPresenza']).min()
                                                    max_data = pd.to_datetime(df_to_display_detail['DataPresenza']).max()
                                                    st.info(f"Periodo: dal {min_data.strftime('%d/%m/%Y')} al {max_data.strftime('%d/%m/%Y')}", icon="📅")
                                                except:
                                                    pass
    
                                    # Colonne da visualizzare
                                    cols_disp_detail = ['CodiceFiscale', 'Cognome', 'Nome', 'Email', 'DataPresenza', 'OraPresenza', 
                                                       'Codice_classe_di_concorso_e_denominazione', 'Dipartimento',
                                                       'DenominazioneAttività', 'CFU']
                                    cols_disp_detail_exist = [c for c in cols_disp_detail if c in df_to_display_detail.columns]
                                    
                                    # Ordinamento
                                    detail_sort_options = st.radio(
                                        "Ordinamento dettagli:", 
                                        ["Per Data (più recente prima)", "Per Cognome e Nome", "Per Attività", "Per Denominazione Classe di concorso"],
                                        horizontal=True
                                    )
                                    
                                    # Imposta l'ordinamento in base alla selezione
                                    if detail_sort_options == "Per Data (più recente prima)" and 'DataPresenza' in df_to_display_detail.columns:
                                        sort_by_columns = ['DataPresenza', 'OraPresenza'] 
                                        ascending = [False, False]  # Prima le date più recenti
                                    elif detail_sort_options == "Per Attività" and 'DenominazioneAttività' in df_to_display_detail.columns:
                                        sort_by_columns = ['DenominazioneAttività', 'DataPresenza']
                                        ascending = [True, True]
                                    elif detail_sort_options == "Per Denominazione Classe di concorso" and 'Codice_classe_di_concorso_e_denominazione' in df_to_display_detail.columns:
                                        sort_by_columns = ['Codice_classe_di_concorso_e_denominazione', 'Cognome', 'Nome']
                                        ascending = [True, True, True]
                                    else:  # Default: per cognome e nome
                                        sort_by_columns = ['Cognome', 'Nome']
                                        if 'DataPresenza' in df_to_display_detail.columns: 
                                            sort_by_columns.append('DataPresenza')
                                        ascending = [True] * len(sort_by_columns)
                                    
                                    # Filtra colonne valide per ordinamento
                                    valid_sort_by = [col for col in sort_by_columns if col in df_to_display_detail.columns]
                                    
                                    # Visualizza dataframe
                                    if not valid_sort_by: 
                                        df_to_show = df_to_display_detail[cols_disp_detail_exist]
                                    else: 
                                        df_to_show = df_to_display_detail[cols_disp_detail_exist].sort_values(
                                            by=valid_sort_by, 
                                            ascending=ascending[:len(valid_sort_by)]
                                        )
                                    
                                    st.dataframe(df_to_show, use_container_width=True)
                            else:
                                st.info("Nessun record dettagliato da mostrare per la selezione corrente.")
                        except Exception as e: 
                            st.error(f"Errore durante la visualizzazione: {e}")

            # --- Esportazione Excel Multi-Tab ---
            st.divider()
            
            # Utilizziamo un container per la sezione di esportazione
            export_container = st.container()
            with export_container:
                st.subheader("📤 Esportazione Dettaglio Presenze", divider="gray")
                course_col_export = 'DenominazioneAttività'  # Questa colonna esiste nei dati dettagliati
                if course_col_export not in current_df_for_tab3.columns:
                    st.error(f"Colonna chiave '{course_col_export}' non trovata.")
                else:
                    export_description_col, export_info_col = st.columns([2, 1])
                    with export_description_col:
                        st.write("Questa funzione consente di esportare i dati dettagliati in Excel o CSV.")
                    
                    with export_info_col:
                        with st.expander("Informazioni Esportazione"):
                            st.info("L'esportazione Excel crea un foglio separato per ogni denominazione attività.\nIl nome del foglio viene estratto dal codice dell'attività.")
                    
                    # Tab per i due tipi di export
                    export_tab1, export_tab2 = st.tabs(["🗂️ Export Multi-Foglio", "📊 Export CSV"])
                
                    with export_tab1:
                        st.markdown("**1️⃣ Seleziona e Ordina le Colonne:**")
                        st.caption("Seleziona le colonne e trascinale per cambiarne l'ordine")

                        all_possible_cols = current_df_for_tab3.columns.tolist()
                        internal_cols_to_exclude = ['TimestampPresenza']
                        all_exportable_cols = [col for col in all_possible_cols if col not in internal_cols_to_exclude]
                        default_cols_export_ordered = ['DataPresenza','OraPresenza','DenominazioneAttività','Cognome','Nome','Percorso','Codice_classe_di_concorso_e_denominazione','CFU']
                        default_cols_final = [col for col in default_cols_export_ordered if col in all_exportable_cols]

                        # Miglioriamo la visualizzazione dei dati di esempio
                        with st.expander("👀 Anteprima dati (primo record)"):
                            if not current_df_for_tab3.empty:
                                example_data = current_df_for_tab3.head(1)[all_exportable_cols].to_dict(orient='records')[0]
                                st.json(example_data)
                            else: 
                                st.caption("Nessun dato da mostrare.")

                        # Selezione colonne con una descrizione più chiara
                        st.caption("Le colonne selezionate verranno incluse nell'esportazione nell'ordine specificato")
                        selected_cols_export_ordered = st.multiselect(
                            "Seleziona le colonne da esportare:", 
                            options=all_exportable_cols, 
                            default=default_cols_final, 
                            key="export_cols_selector_ordered_v215"
                        )
                        
                        st.markdown("**2️⃣ Seleziona Periodo per l'Export (opzionale):**")
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
                        
                        # Aggiunta opzione per scegliere criterio di raggruppamento
                        st.markdown("**3️⃣ Seleziona Criterio di Raggruppamento:**")
                        group_by_options = ["Denominazione Attività (Default)", "Classe di Concorso", "Codice Classe di Concorso"]
                        group_by_choice = st.radio(
                            "Raggruppa per:", 
                            options=group_by_options, 
                            horizontal=True,
                            key="export_groupby_v215"
                        )
                        
                        # Pulsante per generare ed esportare file Excel
                        if st.button("📊 Genera ed Esporta File Excel", key="export_excel_ordered_v215", use_container_width=True):
                            if not selected_cols_export_ordered: 
                                st.warning("Seleziona almeno una colonna.")
                            else:
                                overall_success = True
                                sheets_written = 0
                                error_messages = []
                                used_sheet_names = set()  # Insieme per tenere traccia dei nomi foglio già usati (case-insensitive)
                                try:
                                    output = BytesIO()
                                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                                        # Determina il campo per il raggruppamento in base alla scelta dell'utente
                                        if group_by_choice == "Classe di Concorso" and 'Codice_classe_di_concorso_e_denominazione' in current_df_for_tab3.columns:
                                            grouping_col = 'Codice_classe_di_concorso_e_denominazione'
                                        elif group_by_choice == "Codice Classe di Concorso" and 'Codice_Classe_di_concorso' in current_df_for_tab3.columns:
                                            grouping_col = 'Codice_Classe_di_concorso'
                                        else:
                                            grouping_col = course_col_export
                                            
                                        unique_values = current_df_for_tab3[grouping_col].unique()
                                        unique_values = sorted([str(c) for c in unique_values if pd.notna(c)])
                                        
                                        if not unique_values: 
                                            st.error(f"Nessun valore unico trovato in '{grouping_col}'.")
                                            overall_success = False
                                        else:
                                            prog_bar = st.progress(0, text="Creazione fogli...")
                                            for i, value in enumerate(unique_values):
                                                # Estrae codice o crea un identificativo univoco per il nome foglio
                                                # Cerca prima per il nuovo formato [codice]
                                                code_match = re.search(r'^\[([-\w]+)\]', value)
                                                if code_match:
                                                    extracted_code = code_match.group(1)
                                                    sheet_name_cleaned = clean_sheet_name(extracted_code, [name.lower() for name in used_sheet_names])
                                                else:
                                                    # Fallback al metodo precedente di estrazione dalle parentesi
                                                    extracted_code = extract_code_from_parentheses(value)
                                                    if extracted_code: 
                                                        sheet_name_cleaned = clean_sheet_name(extracted_code, [name.lower() for name in used_sheet_names])
                                                    else:
                                                        # Se non è stato trovato alcun codice, crea un nome abbreviato
                                                        words = re.findall(r'\b\w+\b', value)
                                                        if words:
                                                            # Usa le prime lettere di ogni parola o le prime 4 lettere della prima parola
                                                            if len(words) > 1:
                                                                abbr = ''.join(w[0] for w in words if w)
                                                                if len(abbr) < 3 and words[0]:
                                                                    abbr = words[0][:4]
                                                            else:
                                                                abbr = words[0][:4] if words[0] else value[:4]
                                                            sheet_name_cleaned = clean_sheet_name(abbr, [name.lower() for name in used_sheet_names])
                                                        else:
                                                            # Fallback se non ci sono parole
                                                            sheet_name_cleaned = clean_sheet_name(value[:10], [name.lower() for name in used_sheet_names])
                                                        
                                                        st.caption(f"Nota: Codice non trovato per '{value}', usato identificativo: {sheet_name_cleaned}")
                                                
                                                # Aggiungi il nome foglio all'insieme dei nomi usati
                                                used_sheet_names.add(sheet_name_cleaned.lower())
                                                
                                                prog_text = f"Foglio: {sheet_name_cleaned} ({i+1}/{len(unique_values)})" 
                                                prog_bar.progress((i + 1) / len(unique_values), text=prog_text)
                                                df_sheet = current_df_for_tab3[current_df_for_tab3[grouping_col] == value].copy()
                                                
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
                                                        'DenominazioneAttivitaNormalizzataInternal': 'Attività Elaborata'
                                                    }
                                                    cols_to_rename_final = {k: v for k, v in rename_map_export.items() if k in df_sheet_export.columns}
                                                    df_sheet_export = df_sheet_export.rename(columns=cols_to_rename_final)
                                                    
                                                    # Formatta correttamente date e orari per Excel
                                                    from modules.utils import format_datetime_for_excel
                                                    df_sheet_export = format_datetime_for_excel(df_sheet_export)
                                                    
                                                    # Esporta in Excel
                                                    df_sheet_export.to_excel(writer, sheet_name=sheet_name_cleaned, index=False)
                                                    
                                                    # Applica formato alle celle
                                                    workbook = writer.book
                                                    worksheet = writer.sheets[sheet_name_cleaned]
                                                    
                                                    # Formato per le ore
                                                    time_format = workbook.add_format({'num_format': 'HH:MM'})
                                                    
                                                    # Trova colonna OraPresenza e applica il formato
                                                    if 'OraPresenza' in df_sheet_export.columns:
                                                        col_idx = df_sheet_export.columns.get_loc("OraPresenza")
                                                        worksheet.set_column(col_idx, col_idx, 10, time_format)
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
                                    
                                    # Aggiunge il criterio di raggruppamento al nome del file
                                    if group_by_choice == "Classe di Concorso":
                                        group_type = "ClasseConcorso"
                                    elif group_by_choice == "Codice Classe di Concorso":
                                        group_type = "CodiceClasseConcorso"
                                    else:
                                        group_type = "PercorsoSenzaArt13"
                                    fname = f"Report_Presenze_Dettaglio_{group_type}{period_info}_{ts}.xlsx"
                                    st.success(f"File Excel generato con {sheets_written} fogli! Raggruppamento per: {group_by_choice}")
                                    if error_messages: 
                                        st.warning("Alcuni fogli potrebbero aver avuto problemi:")
                                        [st.caption(msg) for msg in error_messages]
                                        
                                    st.download_button(
                                        label="📥 Scarica Report Excel Dettaglio",
                                        data=output,
                                        file_name=fname,
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        key="dl_excel_course_ordered_v215",
                                        use_container_width=True
                                    )
                                elif sheets_written == 0 and overall_success:
                                    # Messaggio personalizzato in base al criterio di raggruppamento
                                    if group_by_choice == "Classe di Concorso":
                                        st.warning("Nessun dato trovato per alcuna classe di concorso. File Excel non generato.")
                                    elif group_by_choice == "Codice Classe di Concorso":
                                        st.warning("Nessun dato trovato per alcun codice classe di concorso. File Excel non generato.")
                                    else:
                                        st.warning("Nessun dato trovato per alcun percorso. File Excel non generato.")
                                else:
                                    st.error("Generazione file Excel fallita o nessun foglio valido scritto.")
                                    if error_messages: 
                                        st.warning("Dettaglio errori:")
                                        [st.caption(msg) for msg in error_messages]
                
                    with export_tab2:
                        st.markdown("**Esportazione in CSV**")
                        st.caption("Il file CSV conterrà tutti i dati in un'unica tabella, filtrati per il periodo selezionato")
                        
                        # Pulsante per generare ed esportare file CSV
                        if st.button("📄 Genera ed Esporta File CSV", key="export_csv_ordered_v215", use_container_width=True):
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
                                            key="dl_csv_ordered_v215",
                                            use_container_width=True
                                        )
                
                # Footer della pagina
                if not attendance_df.empty:
                    # Footer con informazioni aggiuntive
                    st.divider()
                    footer_container = st.container()
                    with footer_container:
                        footer_col1, footer_col2 = st.columns(2)
                        
                        # Messaggio informativo sulle colonne che saranno rimosse in futuro
                        with footer_col1:
                            with st.expander("⚠️ Nota sui campi rimossi"):
                                st.warning("""
                                Le colonne relative al percorso Art.13 sono state rimosse.
                                Il database ora utilizza come chiavi: Nome, Cognome, Codice_classe_di_concorso_e_denominazione
                                """)
                                
                        # Aiuto sulla funzionalità
                        with footer_col2:
                            with st.expander("ℹ️ Informazioni su questa pagina"):
                                st.info("""
                                **Funzionalità della scheda Calcolo Presenze**
                                
                                In questa scheda puoi:
                                1. Filtrare i dati per percorso e studente
                                2. Visualizzare statistiche aggregate sulle presenze
                                3. Ordinare i dati in diversi modi
                                4. Esportare i dati in Excel o CSV
                                
                                L'esportazione consente di creare report dettagliati per ogni percorso.
                                """)
                else:
                    st.warning("Nessun dato di presenza aggregato da visualizzare. Verifica che i dati siano stati caricati correttamente.", icon="⚠️")
    else:
        st.error("Nessun dato valido caricato. Carica un file valido nella prima scheda dell'applicazione.", icon="🚫")
