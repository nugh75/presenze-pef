# streamlit_app.py
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, date, time
import os
import re
from io import BytesIO
import difflib  # Aggiunto per confrontare stringhe simili

# --- Configurazione Pagina ---
st.set_page_config(
    page_title="Gestione Presenze",
    page_icon="📊",
    layout="wide",
)

# --- Funzioni Utilità --- (Invariate)
def normalize_generic(name):
    if not isinstance(name, str): return name
    normalized = re.sub(r'\s*\(?art\.?\s*13\.?\)?.*$', '', name, flags=re.IGNORECASE)
    return normalized.strip()
    
def reposition_code_to_front(text):
    """Prende il testo, estrae il codice tra parentesi (es. (A-30)) e lo riposiziona all'inizio della stringa."""
    if not isinstance(text, str): return text
    match = re.search(r'\(([-\w]+)\)', text)  # Cerca codice alfanumerico con trattini tra parentesi
    if match:
        code = match.group(1).strip()
        # Rimuovi il codice con le parentesi dal testo originale
        cleaned_text = re.sub(r'\s*\(' + re.escape(code) + r'\)\s*', ' ', text).strip()
        # Restituisci il codice all'inizio seguito dal testo pulito
        return f"[{code}] {cleaned_text}"
    return text
    
def transform_by_codice_percorso(codice, default_name):
    if pd.isna(codice): return default_name
    codice_str = str(codice).strip();
    if len(codice_str) < 3: return default_name
    prefix = codice_str[:3]
    if prefix == '600': return "PeF60 All. 1"
    elif prefix == '300': return "PeF30 All. 2"
    elif prefix == '360': return "PeF36 All. 5"
    elif prefix == '200': return "PeF30 art. 13"
    else: return default_name
def clean_sheet_name(name):
    name = re.sub(r'[\\/?*\[\]:]', '', str(name)); return name[:31]
def extract_code_from_parentheses(text):
    if not isinstance(text, str): return None
    match = re.search(r'\((.*?)\)', text)
    if match:
        code = match.group(1).strip()
        if code: return code
    return None

# --- Funzione per caricare i CFU ---
@st.cache_data
def load_cfu_data():
    """Carica i dati dei CFU dal file 'crediti.csv'."""
    try:
        cfu_df = pd.read_csv('crediti.csv')
        if 'DenominazioneAttività' not in cfu_df.columns or 'CFU' not in cfu_df.columns:
            st.error("Il file dei CFU non contiene le colonne richieste: 'DenominazioneAttività' e 'CFU'.")
            return pd.DataFrame()
        # Normalizza i nomi delle attività per facilitare il matching
        cfu_df['DenominazioneAttivitaNormalizzata'] = cfu_df['DenominazioneAttività'].apply(lambda x: x.strip() if isinstance(x, str) else x)
        return cfu_df
    except Exception as e:
        st.error(f"Errore durante il caricamento del file dei CFU: {e}")
        return pd.DataFrame()

def match_activity_with_cfu(activity_name, cfu_data):
    """Abbina un'attività con il suo CFU dal dataset dei CFU.
    Gestisce differenze minori nei nomi delle attività, ignorando maiuscole/minuscole."""
    if not isinstance(activity_name, str) or activity_name.strip() == '' or cfu_data.empty:
        return None
    
    # Normalizza il nome dell'attività (strip e lowercase)
    normalized_activity = activity_name.strip().lower()
    
    # Confronto case-insensitive
    for idx, row in cfu_data.iterrows():
        if row['DenominazioneAttivitaNormalizzata'].lower() == normalized_activity:
            return row['CFU']
    
    # Se non trova un match esatto case-insensitive, cerca il più simile
    similarity_threshold = 0.9  # Soglia di similarità (90%)
    activities = cfu_data['DenominazioneAttivitaNormalizzata'].tolist()
    activities_lower = [act.lower() for act in activities if isinstance(act, str)]
    
    # Usa difflib per trovare il match più simile (case-insensitive)
    matches = difflib.get_close_matches(normalized_activity, activities_lower, n=1, cutoff=similarity_threshold)
    if matches:
        closest_match_lower = matches[0]
        # Trova l'indice corrispondente nell'array originale
        for idx, act in enumerate(activities):
            if isinstance(act, str) and act.lower() == closest_match_lower:
                return cfu_data.iloc[idx]['CFU']
    
    return None

# --- Funzione Caricamento Dati --- (Modificata per includere CFU)
@st.cache_data
def load_data(uploaded_file):
    if uploaded_file is None: return None
    try:
        # Carica i dati dei CFU
        cfu_data = load_cfu_data()
        if cfu_data.empty:
            st.warning("Non è stato possibile caricare i dati dei CFU. I CFU non saranno disponibili.")
        
        df = pd.read_excel(uploaded_file); original_columns = df.columns.tolist()
        required_cols = ['CodiceFiscale', 'DataPresenza', 'OraPresenza']
        if not ('percoro' in df.columns or 'DenominazionePercorso' in df.columns): st.error("Colonna Percorso ('percoro' o 'DenominazionePercorso') non trovata."); return None
        for col in required_cols:
            if col not in df.columns: st.error(f"Colonna obbligatoria '{col}' non trovata."); return None
        df['CodiceFiscale'] = df['CodiceFiscale'].astype(str).str.strip()
        percorso_col_original_name = 'percoro' if 'percoro' in df.columns else 'DenominazionePercorso'
        df['PercorsoOriginaleInternal'] = df[percorso_col_original_name]
        df['PercorsoOriginaleSenzaArt13Internal'] = df['PercorsoOriginaleInternal'].apply(normalize_generic)
        # Applica la trasformazione per spostare i codici all'inizio
        df['PercorsoOriginaleSenzaArt13Internal'] = df['PercorsoOriginaleSenzaArt13Internal'].apply(reposition_code_to_front)
        if 'CodicePercorso' in df.columns: df['PercorsoInternal'] = df.apply(lambda row: transform_by_codice_percorso(row.get('CodicePercorso'), row['PercorsoOriginaleSenzaArt13Internal']), axis=1)
        else: df['PercorsoInternal'] = df['PercorsoOriginaleSenzaArt13Internal']
        if 'DenominazioneCds' in df.columns: st.warning("'DenominazioneCds' trovata. Verrà ignorata/rimossa."); df = df.drop(columns=['DenominazioneCds'], errors='ignore')
        activity_col_norm_internal = 'DenominazioneAttivitaNormalizzataInternal'
        if 'DenominazioneAttività' in df.columns: 
            df[activity_col_norm_internal] = df['DenominazioneAttività'].apply(normalize_generic)
            
            # Aggiungi colonna CFU abbinando le attività
            if not cfu_data.empty:
                st.info("Abbinamento dei CFU alle attività in corso...")
                df['CFU'] = df['DenominazioneAttività'].apply(lambda x: match_activity_with_cfu(x, cfu_data))
                # Conta quante attività non hanno trovato un match per i CFU
                missing_cfu = df['CFU'].isna().sum()
                total_activities = len(df)
                if missing_cfu > 0:
                    st.warning(f"Non è stato possibile trovare i CFU per {missing_cfu} attività su {total_activities} ({(missing_cfu/total_activities)*100:.1f}%).")
            else:
                df['CFU'] = None
        try: df['DataPresenza'] = pd.to_datetime(df['DataPresenza'], errors='coerce').dt.date; df.loc[pd.isna(df['DataPresenza']), 'DataPresenza'] = pd.NaT
        except Exception as e: st.warning(f"Problema conversione 'DataPresenza': {e}."); df['DataPresenza'] = pd.NaT if 'DataPresenza' not in df.columns else df['DataPresenza']
        try:
            def parse_time(t):
                if isinstance(t, time): return t;
                if isinstance(t, datetime): return t.time();
                if isinstance(t, pd.Timestamp): return t.time();
                if pd.isna(t): return pd.NaT
                try: return pd.to_datetime(str(t), format='%H:%M:%S', errors='raise').time()
                except ValueError:
                    try:
                        if isinstance(t, (int, float)) and 0 <= t < 1: total_seconds = int(t * 24 * 60 * 60); hours, remainder = divmod(total_seconds, 3600); minutes, seconds = divmod(remainder, 60); return time(hours, minutes, seconds)
                        return pd.to_datetime(str(t), errors='raise').time()
                    except (ValueError, TypeError): return pd.NaT
            df['OraPresenza'] = df['OraPresenza'].apply(parse_time); df.loc[pd.isna(df['OraPresenza']), 'OraPresenza'] = pd.NaT
        except Exception as e: st.warning(f"Problema conversione 'OraPresenza': {e}."); df['OraPresenza'] = pd.NaT if 'OraPresenza' not in df.columns else df['OraPresenza']
        def combine_dt(row):
            if pd.notna(row['DataPresenza']) and isinstance(row['DataPresenza'], date) and pd.notna(row['OraPresenza']) and isinstance(row['OraPresenza'], time):
                try: return pd.Timestamp.combine(row['DataPresenza'], row['OraPresenza'])
                except Exception: return pd.NaT
            return pd.NaT
        df['TimestampPresenza'] = df.apply(combine_dt, axis=1) # TimestampPresenza è creata qui
        initial_rows = len(df); df.dropna(subset=['TimestampPresenza', 'CodiceFiscale'], inplace=True)
        df['DataPresenza'] = df['TimestampPresenza'].dt.date; df['OraPresenza'] = df['TimestampPresenza'].dt.time
        removed_rows = initial_rows - len(df)
        if removed_rows > 0: st.warning(f"Rimossi {removed_rows} record con CF, Data, Ora o Timestamp mancanti/non validi.")
        if 'Nome' not in df.columns: df['Nome'] = '';
        if 'Cognome' not in df.columns: df['Cognome'] = ''
        final_cols = ['CodiceFiscale', 'Nome', 'Cognome', 'DataPresenza', 'OraPresenza', 'PercorsoOriginaleInternal', 'PercorsoOriginaleSenzaArt13Internal', 'PercorsoInternal', 'TimestampPresenza']
        if 'DenominazioneAttività' in df.columns: final_cols.append('DenominazioneAttività')
        if activity_col_norm_internal in df.columns: final_cols.append(activity_col_norm_internal)
        if 'CodicePercorso' in original_columns and 'CodicePercorso' in df.columns: final_cols.append('CodicePercorso')
        if 'CFU' in df.columns: final_cols.append('CFU')  # Aggiungi la colonna CFU all'elenco delle colonne da mantenere
        cols_to_keep = [col for col in final_cols if col in df.columns]; df_final = df[cols_to_keep].copy()
        return df_final
    except Exception as e: st.error(f"Errore critico caricamento/elaborazione file: {e}"); st.exception(e); return None

# --- Funzione Rilevamento Duplicati --- (Invariata da v2.6)
def detect_duplicate_records(df, timestamp_col='TimestampPresenza', cf_column='CodiceFiscale', time_delta_minutes=10):
    # ... (codice invariato) ...
    if df is None or len(df) == 0: return pd.DataFrame(), [], []
    required_cols = [timestamp_col, cf_column]
    if not all(col in df.columns for col in required_cols):
        missing_but_exist = [col for col in required_cols if col in df.columns and df[col].isnull().all()];
        if len(missing_but_exist) == len(required_cols): st.warning(f"Colonne necessarie ({', '.join(required_cols)}) vuote."); return pd.DataFrame(), [], []
        missing_cols = [col for col in required_cols if col not in df.columns];
        if missing_cols: st.error(f"Colonne necessarie ({', '.join(missing_cols)}) mancanti."); return pd.DataFrame(), [], []
    df_copy = df.dropna(subset=required_cols).copy()
    if df_copy.empty: st.info("Nessun record con Timestamp e CF validi per controllo duplicati."); return pd.DataFrame(), [], []
    df_copy['OriginalIndex'] = df_copy.index; df_sorted = df_copy.sort_values(by=[cf_column, timestamp_col])
    df_sorted['time_diff'] = df_sorted.groupby(cf_column)[timestamp_col].diff(); time_threshold = timedelta(minutes=time_delta_minutes)
    df_sorted['time_diff_next'] = df_sorted.groupby(cf_column)[timestamp_col].diff(-1).abs()
    df_sorted['is_close_to_prev'] = (df_sorted['time_diff'].notna()) & (df_sorted['time_diff'] <= time_threshold)
    df_sorted['is_close_to_next'] = (df_sorted['time_diff_next'].notna()) & (df_sorted['time_diff_next'] <= time_threshold)
    df_sorted['in_cluster'] = df_sorted['is_close_to_prev'] | df_sorted['is_close_to_next']
    df_sorted['GruppoDuplicati'] = 0; current_group_id = 1; sorted_indices = df_sorted.index
    for i in range(len(sorted_indices)):
        current_idx_sorted = sorted_indices[i]
        if df_sorted.loc[current_idx_sorted, 'in_cluster']:
            current_group = df_sorted.loc[current_idx_sorted, 'GruppoDuplicati']; prev_group = 0
            if i > 0:
                prev_idx_sorted = sorted_indices[i-1]
                if df_sorted.loc[prev_idx_sorted, cf_column] == df_sorted.loc[current_idx_sorted, cf_column] and df_sorted.loc[prev_idx_sorted, 'in_cluster']: prev_group = df_sorted.loc[prev_idx_sorted, 'GruppoDuplicati']
            if current_group == 0:
                if prev_group != 0 and df_sorted.loc[current_idx_sorted, 'is_close_to_prev']: df_sorted.loc[current_idx_sorted, 'GruppoDuplicati'] = prev_group
                else: df_sorted.loc[current_idx_sorted, 'GruppoDuplicati'] = current_group_id; current_group_id += 1
            elif prev_group != 0 and current_group != prev_group and df_sorted.loc[current_idx_sorted, 'is_close_to_prev']: group_to_merge = current_group; target_group = prev_group; df_sorted.loc[df_sorted['GruppoDuplicati'] == group_to_merge, 'GruppoDuplicati'] = target_group
    involved_df_sorted = df_sorted[df_sorted['GruppoDuplicati'] != 0].copy()
    if involved_df_sorted.empty: return pd.DataFrame(), [], []
    group_mapping = pd.Series(involved_df_sorted['GruppoDuplicati'].values, index=involved_df_sorted['OriginalIndex']).to_dict()
    involved_original_indices = involved_df_sorted['OriginalIndex'].unique().tolist(); indices_to_drop_suggestion = []
    grouped = involved_df_sorted.sort_values(timestamp_col).groupby('GruppoDuplicati')
    for group_id, group_df in grouped: indices_to_drop_suggestion.extend(group_df['OriginalIndex'].iloc[1:].tolist())
    valid_involved_indices = [idx for idx in involved_original_indices if idx in df.index]
    if not valid_involved_indices: return pd.DataFrame(), [], []
    duplicates_df = df.loc[valid_involved_indices].copy()
    duplicates_df['GruppoDuplicati'] = duplicates_df.index.map(group_mapping).fillna(0).astype(int); duplicates_df['SuggerisciRimuovere'] = duplicates_df.index.isin(indices_to_drop_suggestion)
    duplicates_df = duplicates_df.sort_values(by=['GruppoDuplicati', timestamp_col]); valid_indices_to_drop = [idx for idx in indices_to_drop_suggestion if idx in df.index]
    return duplicates_df, valid_involved_indices, valid_indices_to_drop

# --- Funzione Calcolo Presenze Aggregate --- (Modificata per includere CFU)
def calculate_attendance(df, cf_column='CodiceFiscale', percorso_chiave_col='PercorsoOriginaleSenzaArt13Internal', percorso_elab_col='PercorsoInternal', original_col='PercorsoOriginaleInternal'):
    if df is None or len(df) == 0: return pd.DataFrame()
    required_cols = [cf_column, percorso_chiave_col]; optional_cols = [percorso_elab_col, original_col]
    name_cols = [col for col in ['Nome', 'Cognome'] if col in df.columns]
    if not all(col in df.columns for col in required_cols):
        missing = [col for col in required_cols if col not in df.columns]; st.error(f"Colonne chiave ({', '.join(missing)}) mancanti."); return pd.DataFrame()
    
    # Includi anche la colonna CFU se presente
    additional_cols = []
    if 'CFU' in df.columns:
        additional_cols.append('CFU')
    
    group_cols = [cf_column, percorso_chiave_col]
    first_cols = name_cols + additional_cols + [col for col in optional_cols if col in df.columns and col != percorso_chiave_col]
    attendance_counts = df.groupby(group_cols, dropna=False).size().reset_index(name='Presenze')
    
    if first_cols:
         first_info_values = df.dropna(subset=first_cols, how='all').groupby(group_cols, as_index=False)[first_cols].first()
         attendance = pd.merge(attendance_counts, first_info_values, on=group_cols, how='left')
    else: attendance = attendance_counts
    
    rename_map = {percorso_chiave_col: 'Percorso (Senza Art.13)', percorso_elab_col: 'Percorso Elaborato (Info)', original_col: 'Percorso Originale Input (Info)'}
    attendance = attendance.rename(columns={k: v for k, v in rename_map.items() if k in attendance.columns})
    
    cols_order_final = [cf_column] + [col for col in ['Nome', 'Cognome'] if col in attendance.columns] + [rename_map.get(percorso_chiave_col, percorso_chiave_col)] + [rename_map.get(col, col) for col in [percorso_elab_col, original_col] if rename_map.get(col, col) in attendance.columns]
    
    # Aggiungi CFU prima di Presenze nella visualizzazione
    if 'CFU' in attendance.columns:
        cols_order_final.append('CFU')
    
    cols_order_final.append('Presenze')
    final_cols = [c for c in cols_order_final if c in attendance.columns]; attendance = attendance[final_cols]
    
    if 'Nome' in attendance.columns: attendance['Nome'] = attendance['Nome'].fillna('')
    if 'Cognome' in attendance.columns: attendance['Cognome'] = attendance['Cognome'].fillna('')
    
    return attendance

# --- Layout Principale ---
st.title("📊 Gestione Presenze Corsi")
st.markdown("<a id='top'></a>", unsafe_allow_html=True)

# --- Sidebar --- (Invariata da v2.6)
with st.sidebar:
    st.header("Caricamento File"); uploaded_file = st.file_uploader("Carica file Excel presenze", type=['xlsx'])
    if uploaded_file:
        st.success(f"File '{uploaded_file.name}' caricato!")
        if st.button("Mostra anteprima originale"):
            try: preview = pd.read_excel(uploaded_file, nrows=10); st.write("Colonne:", preview.columns.tolist()); st.dataframe(preview)
            except Exception as e: st.error(f"Errore anteprima: {e}")
    st.divider()
    if 'processed_df' in st.session_state and st.session_state.processed_df is not None: st.markdown("[⬆️ Torna su](#top)", help="Clicca per tornare all'inizio della pagina principale")

# --- Gestione Stato Sessione --- (Invariata)
if 'duplicates_removed' not in st.session_state: st.session_state.duplicates_removed = False
if 'duplicate_detection_results' not in st.session_state: st.session_state.duplicate_detection_results = (pd.DataFrame(), [], [])
if 'selected_indices_to_drop' not in st.session_state: st.session_state.selected_indices_to_drop = []
if 'report_data_to_download' not in st.session_state: st.session_state.report_data_to_download = None
if 'report_filename_to_download' not in st.session_state: st.session_state.report_filename_to_download = None

# --- Logica caricamento e reset stato --- (Invariata da v2.6)
if uploaded_file:
    if 'current_file_name' not in st.session_state or st.session_state.current_file_name != uploaded_file.name or 'processed_df' not in st.session_state:
        with st.spinner("Caricamento ed elaborazione dati..."): st.session_state.processed_df = load_data(uploaded_file)
        if st.session_state.processed_df is not None:
            st.session_state.current_file_name = uploaded_file.name ; st.session_state.duplicates_removed = False ; st.session_state.duplicate_detection_results = (pd.DataFrame(), [], []) ; st.session_state.selected_indices_to_drop = [] ; st.session_state.report_data_to_download = None ; st.session_state.report_filename_to_download = None
            st.success("Dati caricati ed elaborati."); st.rerun()
        else:
            if 'current_file_name' in st.session_state: del st.session_state.current_file_name
            if 'processed_df' in st.session_state: del st.session_state.processed_df
            st.session_state.duplicates_removed = False ; st.session_state.duplicate_detection_results = (pd.DataFrame(), [], []) ; st.session_state.selected_indices_to_drop = [] ; st.session_state.report_data_to_download = None ; st.session_state.report_filename_to_download = None
elif 'processed_df' in st.session_state:
    del st.session_state.processed_df
    if 'current_file_name' in st.session_state: del st.session_state.current_file_name
    st.session_state.duplicates_removed = False ; st.session_state.duplicate_detection_results = (pd.DataFrame(), [], []) ; st.session_state.selected_indices_to_drop = [] ; st.session_state.report_data_to_download = None ; st.session_state.report_filename_to_download = None
    st.rerun()

df_main = st.session_state.get('processed_df', None)

# --- Tabs ---
if df_main is not None and isinstance(df_main, pd.DataFrame):
    tab1, tab2, tab3 = st.tabs(["Analisi Dati", "Gestione Duplicati", "Calcolo Presenze ed Esportazione"])

    # --- Tab 1: Analisi Dati --- (Invariata da v2.6)
    with tab1:
        st.header("Analisi Dati Caricati");
        if st.session_state.duplicates_removed: st.success("I record duplicati selezionati sono stati rimossi.")
        st.subheader("Statistiche Generali")
        metrics_data = {"Record Validi": len(df_main)};
        if 'CodiceFiscale' in df_main.columns: metrics_data["Persone Uniche (CF)"] = df_main['CodiceFiscale'].nunique()
        if 'PercorsoInternal' in df_main.columns: metrics_data["Percorsi Unici (Elab.)"] = df_main['PercorsoInternal'].nunique()
        activity_col_norm_internal = 'DenominazioneAttivitaNormalizzataInternal';
        if activity_col_norm_internal in df_main.columns: metrics_data["Attività Uniche (Norm.)"] = df_main[activity_col_norm_internal].nunique()
        num_metrics = len(metrics_data) ; cols_metrics = st.columns(num_metrics) ; i = 0
        for label, value in metrics_data.items():
            with cols_metrics[i]: st.metric(label, value) ; i += 1
        st.subheader("Dati Attuali Utilizzati per l'Analisi")
        cols_show_preferred = ['CodiceFiscale', 'Nome', 'Cognome', 'DataPresenza', 'OraPresenza', 'PercorsoOriginaleInternal', 'PercorsoOriginaleSenzaArt13Internal', 'PercorsoInternal', 'DenominazioneAttività', 'DenominazioneAttivitaNormalizzataInternal', 'CodicePercorso', 'CFU', 'TimestampPresenza']
        cols_show_exist = [col for col in cols_show_preferred if col in df_main.columns]
        st.dataframe(df_main[cols_show_exist], use_container_width=True)
        st.caption("PercorsoOriginaleInternal: Input. PercorsoOriginaleSenzaArt13Internal: Input senza Art.13 con codice riposizionato all'inizio [Codice] (usato per fogli export e filtro Tab3). PercorsoInternal: Elaborato/Trasformato. CFU: Crediti Formativi Universitari associati all'attività.")

    # --- Tab 2: Gestione Duplicati --- (Invariata da v2.6)
    with tab2:
        # ... (codice Tab 2 invariato) ...
        st.header("Gestione Record Potenzialmente Duplicati"); st.markdown("Identifica cluster di timbrature ravvicinate (≤ 10 min) per stesso CF/giorno.")
        if not st.session_state.duplicates_removed and st.session_state.duplicate_detection_results[0].empty:
             with st.spinner("Rilevamento duplicati..."):
                 df_per_detect = st.session_state.get('processed_df', pd.DataFrame());
                 if not df_per_detect.empty:
                     required_dup_cols = ['TimestampPresenza', 'CodiceFiscale']
                     if all(col in df_per_detect.columns for col in required_dup_cols): st.session_state.duplicate_detection_results = detect_duplicate_records(df_per_detect)
                     else: missing_cols = [col for col in required_dup_cols if col not in df_per_detect.columns]; st.warning(f"Colonne necessarie ({', '.join(missing_cols)}) non trovate."); st.session_state.duplicate_detection_results = (pd.DataFrame(), [], [])
                 else: st.session_state.duplicate_detection_results = (pd.DataFrame(), [], [])
        duplicates_df_display, involved_indices, indices_to_drop_suggested = st.session_state.duplicate_detection_results; duplicates_found = not duplicates_df_display.empty
        def select_duplicates_to_remove_ui(df_duplicates):
            if df_duplicates.empty: return []
            selected_indices_map = {};
            if 'GruppoDuplicati' not in df_duplicates.columns: st.error("Colonna 'GruppoDuplicati' mancante."); return []
            all_groups = sorted([g for g in df_duplicates['GruppoDuplicati'].unique() if g != 0])
            if not all_groups: return []
            st.write(f"👇 **Revisiona e seleziona record da eliminare:**")
            for group_id in all_groups:
                group_df = df_duplicates[df_duplicates['GruppoDuplicati'] == group_id].copy()
                if len(group_df) < 2: continue
                first_row = group_df.iloc[0]; cf = first_row.get('CodiceFiscale', 'N/A'); date_obj = first_row.get('DataPresenza'); date_str = date_obj.strftime('%d/%m/%Y') if pd.notna(date_obj) and hasattr(date_obj, 'strftime') else 'Data N/A'
                expander_label = f"Gruppo {int(group_id)}: CF {cf} - Data {date_str} ({len(group_df)} record)"
                with st.expander(expander_label, expanded=False):
                    st.write(f"**CF:** {cf}, **Data:** {date_str}")
                    group_df_edit = group_df.copy(); group_df_edit['Elimina'] = group_df_edit['SuggerisciRimuovere']
                    cols_to_display_editor = ['Elimina', 'OraPresenza', 'PercorsoInternal']; cols_to_display_editor.extend([c for c in ['Nome','Cognome','DenominazioneAttività'] if c in group_df_edit.columns]); cols_to_display_editor = [c for c in cols_to_display_editor if c in group_df_edit.columns]
                    group_df_edit['_OriginalIndex'] = group_df_edit.index
                    if '_OriginalIndex' in group_df_edit.columns and 'Elimina' in cols_to_display_editor:
                        edited_df = st.data_editor(group_df_edit[cols_to_display_editor + ['_OriginalIndex']], column_config={"Elimina": st.column_config.CheckboxColumn("Elimina?", default=False),"OraPresenza": st.column_config.TimeColumn("Ora", format="HH:mm:ss"),"PercorsoInternal": st.column_config.TextColumn("Percorso (Elab.)"),"DenominazioneAttività": st.column_config.TextColumn("Attività"),"Nome": st.column_config.TextColumn("Nome"),"Cognome": st.column_config.TextColumn("Cognome"),"_OriginalIndex": None}, disabled=[c for c in cols_to_display_editor if c != 'Elimina'], hide_index=True, key=f"editor_group_{group_id}" )
                        selected_in_group = edited_df[edited_df['Elimina']]['_OriginalIndex'].tolist()
                        for idx in group_df_edit['_OriginalIndex']: selected_indices_map[idx] = idx in selected_in_group
                    else: st.warning(f"Dati incompleti per editor gruppo {group_id}.")
            return [idx for idx, selected in selected_indices_map.items() if selected]
        if duplicates_found:
            st.warning(f"Trovati **{len(duplicates_df_display)}** record potenzialmente duplicati in **{duplicates_df_display['GruppoDuplicati'].nunique()}** cluster.")
            st.markdown("---"); st.subheader("Azione Rapida: Eliminazione Automatica Suggerita")
            indices_to_drop_suggested_auto = st.session_state.duplicate_detection_results[2]; current_df_auto = st.session_state.get('processed_df', pd.DataFrame()); valid_indices_to_remove_auto = []
            if not current_df_auto.empty and indices_to_drop_suggested_auto: valid_indices_to_remove_auto = [idx for idx in indices_to_drop_suggested_auto if idx in current_df_auto.index]
            num_valid_to_remove = len(valid_indices_to_remove_auto); auto_remove_disabled = st.session_state.duplicates_removed or num_valid_to_remove == 0
            col_btn_auto, col_report_dl = st.columns([2,3])
            with col_btn_auto:
                if st.button(f"Elimina {num_valid_to_remove} Record Suggeriti", key="auto_remove_and_report", disabled=auto_remove_disabled, help="Rimuove tutti tranne il primo per cluster e prepara un CSV degli eliminati."):
                   if valid_indices_to_remove_auto:
                       try:
                           with st.spinner(f"Eliminazione e preparazione report..."):
                               df_deleted_report = current_df_auto.loc[valid_indices_to_remove_auto].copy(); duplicates_df_display_report = st.session_state.duplicate_detection_results[0]
                               if not duplicates_df_display_report.empty and 'GruppoDuplicati' in duplicates_df_display_report.columns:
                                    try: valid_report_indices = duplicates_df_display_report.index.intersection(valid_indices_to_remove_auto); df_deleted_report['GruppoDuplicati'] = duplicates_df_display_report.loc[valid_report_indices, 'GruppoDuplicati'] if not valid_report_indices.empty else 0
                                    except Exception: pass
                               cols_report_preferred = ['GruppoDuplicati', 'CodiceFiscale', 'Nome', 'Cognome', 'DataPresenza', 'OraPresenza', 'PercorsoOriginaleInternal', 'PercorsoOriginaleSenzaArt13Internal','PercorsoInternal', 'DenominazioneAttività', 'DenominazioneAttivitaNormalizzataInternal', 'CodicePercorso', 'TimestampPresenza']
                               cols_report_exist = [c for c in cols_report_preferred if c in df_deleted_report.columns]; df_deleted_report_final = df_deleted_report[cols_report_exist]
                               if 'GruppoDuplicati' in cols_report_exist and 'TimestampPresenza' in cols_report_exist: df_deleted_report_final = df_deleted_report_final.sort_values(by=['GruppoDuplicati', 'TimestampPresenza'])
                               report_csv_bytes = df_deleted_report_final.to_csv(index=True, index_label='OriginalIndex').encode('utf-8'); ts_report = datetime.now().strftime("%Y%m%d_%H%M"); report_filename = f"Report_Record_Eliminati_Auto_{ts_report}.csv"
                               st.session_state.report_data_to_download = report_csv_bytes; st.session_state.report_filename_to_download = report_filename
                               df_cleaned = current_df_auto.drop(index=valid_indices_to_remove_auto); st.session_state.processed_df = df_cleaned; st.session_state.duplicates_removed = True; st.session_state.duplicate_detection_results = (pd.DataFrame(), [], []); st.session_state.selected_indices_to_drop = []
                           st.success(f"{num_valid_to_remove} record rimossi!")
                       except Exception as e: st.error(f"Errore eliminazione automatica: {e}"); st.exception(e); st.session_state.report_data_to_download = None ; st.session_state.report_filename_to_download = None
                   else: st.info("Nessun record suggerito valido da rimuovere.")
            with col_report_dl:
                 if st.session_state.report_data_to_download is not None and st.session_state.duplicates_removed:
                    st.download_button(label=f"📥 Scarica Report Eliminati ({st.session_state.report_filename_to_download})", data=st.session_state.report_data_to_download, file_name=st.session_state.report_filename_to_download, mime="text/csv", key="dl_deleted_report")
            st.markdown("---"); st.subheader("Revisione Manuale e Rimozione Selezionata"); st.info("Esamina i gruppi qui sotto...")
            selected_indices = select_duplicates_to_remove_ui(duplicates_df_display); st.session_state.selected_indices_to_drop = selected_indices
            st.divider()
            num_selected = len(st.session_state.selected_indices_to_drop); manual_button_disabled = st.session_state.duplicates_removed or num_selected == 0
            if num_selected > 0: st.write(f"🗑️ Selezionati manualmente: **{num_selected}** record.")
            else: st.write("⚠️ Nessun record selezionato manualmente.")
            if st.button("Rimuovi i Record Selezionati Manualmente", key="confirm_remove_manual", disabled=manual_button_disabled, help="Rimuove solo i record con 'Elimina?' spuntata."):
                selected_to_remove = st.session_state.selected_indices_to_drop
                if selected_to_remove:
                    try:
                        current_df_manual = st.session_state.get('processed_df', pd.DataFrame())
                        if not current_df_manual.empty:
                            valid_indices_to_remove = [idx for idx in selected_to_remove if idx in current_df_manual.index]; num_actually_removed = len(valid_indices_to_remove)
                            if num_actually_removed > 0:
                                with st.spinner(f"Rimozione di {num_actually_removed} record..."):
                                    df_cleaned = current_df_manual.drop(index=valid_indices_to_remove); st.session_state.processed_df = df_cleaned; st.session_state.duplicates_removed = True; st.session_state.duplicate_detection_results = (pd.DataFrame(), [], []); st.session_state.selected_indices_to_drop = []; st.session_state.report_data_to_download = None ; st.session_state.report_filename_to_download = None
                                st.success(f"{num_actually_removed} record rimossi!"); st.info("Dati aggiornati. Ricaricamento interfaccia...")
                                st.rerun()
                            else: st.warning("Nessuno degli indici selezionati trovato.")
                        else: st.error("DataFrame non trovato.")
                    except Exception as e: st.error(f"Errore rimozione manuale: {e}"); st.exception(e)
                else: st.info("Nessun record selezionato.")
            st.divider()
            with st.expander("📄 Scarica Report Tutti i Duplicati Identificati (Prima della rimozione)"):
                 duplicates_df_display_orig = st.session_state.duplicate_detection_results[0]
                 if not duplicates_df_display_orig.empty:
                     cols_show_dup_preferred = ['GruppoDuplicati', 'CodiceFiscale', 'Nome', 'Cognome', 'DataPresenza', 'OraPresenza', 'PercorsoOriginaleInternal', 'PercorsoOriginaleSenzaArt13Internal','PercorsoInternal', 'DenominazioneAttività', 'DenominazioneAttivitaNormalizzataInternal', 'CodicePercorso', 'CFU', 'SuggerisciRimuovere', 'TimestampPresenza']
                     cols_show_dup_exist = [c for c in cols_show_dup_preferred if c in duplicates_df_display_orig.columns]
                     if 'GruppoDuplicati' not in duplicates_df_display_orig.columns: duplicates_df_display_orig['GruppoDuplicati'] = 0
                     duplicates_csv_orig = duplicates_df_display_orig[cols_show_dup_exist].to_csv(index=True, index_label='OriginalIndex').encode('utf-8')
                     st.download_button(label="Scarica CSV Cluster Identificati", data=duplicates_csv_orig, file_name="record_duplicati_cluster_identificati.csv", mime="text/csv", key="dl_involved_clusters_orig")
                 else: st.info("Nessun duplicato identificato.")
        else:
             if st.session_state.duplicates_removed: st.success("I duplicati sono già stati rimossi.")
             else: st.success("✅ Nessun record potenzialmente duplicato trovato.")


    # --- Tab 3: Calcolo Presenze ed Esportazione (Fix Indentazione + UI Semplificata) ---
    with tab3:
        st.header("Calcolo Presenze ed Esportazione")
        st.write("Visualizza presenze per **Percorso (Senza Art.13)** e permette export dettagliato in Excel (fogli divisi per codice percorso).")
        if st.session_state.duplicates_removed: st.success("Presenze calcolate sui dati depurati.")

        current_df_for_tab3 = st.session_state.get('processed_df', pd.DataFrame())

        if not current_df_for_tab3.empty:
            required_att_cols = ['CodiceFiscale', 'PercorsoOriginaleSenzaArt13Internal', 'PercorsoInternal', 'PercorsoOriginaleInternal']
            if not all(col in current_df_for_tab3.columns for col in required_att_cols):
                 missing_cols = [col for col in required_att_cols if col not in current_df_for_tab3.columns]
                 st.error(f"Impossibile procedere: colonne mancanti ({', '.join(missing_cols)})")
                 attendance_df = pd.DataFrame() # Resetta per sicurezza
            else:
                 # Calcola dati aggregati (necessario per il selectbox)
                 attendance_df = calculate_attendance(current_df_for_tab3)

            # --- Visualizzazione Condizionale ---
            if not attendance_df.empty:
                st.subheader("Filtra Visualizzazione")
                p_col_disp_key = "Percorso (Senza Art.13)"
                p_col_internal_key = 'PercorsoOriginaleSenzaArt13Internal'

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
                        def extract_sort_key(percorso_str):
                            code_match = re.search(r'^\[([-\w]+)\]', percorso_str)
                            if code_match:
                                return code_match.group(1)  # Estrai solo il codice per ordinamento
                            return percorso_str  # Usa l'intera stringa altrimenti
                            
                        perc_list = sorted(perc_list, key=extract_sort_key)
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

                        # --- Visualizzazione Tabelle ---
                        if not df_to_display_agg.empty:
                            if perc_sel == "Tutti" or stud_sel == "Tutti gli Studenti":
                                if perc_sel == "Tutti":
                                    st.subheader("Riepilogo Aggregato per Tutti i Percorsi e Studenti")
                                else: # Percorso selezionato, ma tutti gli studenti
                                    st.subheader(f"Riepilogo Aggregato per: {perc_sel}")

                                cols_disp_agg = ['CodiceFiscale', 'Nome', 'Cognome', p_col_disp_key, 'Percorso Elaborato (Info)', 'CFU', 'Presenze']
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

                            cols_disp_detail = ['CodiceFiscale', 'Cognome', 'Nome', 'DataPresenza', 'OraPresenza', 'DenominazioneAttività', 'CFU', 'PercorsoInternal']
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

                    except Exception as e: st.error(f"Errore durante la visualizzazione: {e}")

                # --- Esportazione Excel Multi-Tab (Invariata da v2.15) ---
                # Questo blocco deve essere allineato con 'if not attendance_df.empty:'
                st.divider()
                st.subheader("Esportazione Dettaglio Presenze per Percorso (Originale senza Art.13) in Excel")
                course_col_export = 'PercorsoOriginaleSenzaArt13Internal'
                if course_col_export not in current_df_for_tab3.columns:
                    st.error(f"Colonna chiave '{course_col_export}' non trovata.")
                else:
                    st.write(f"Esporta dati **dettagliati** in Excel, un foglio per ogni **{course_col_export}** unico (nome foglio da codice tra parentesi).")
                    st.markdown("**1. Seleziona e Ordina le Colonne per l'Export:**")
                    st.caption("Usa il box qui sotto per scegliere le colonne. Puoi trascinare le colonne selezionate per cambiarne l'ordine.")

                    all_possible_cols = current_df_for_tab3.columns.tolist(); internal_cols_to_exclude = ['TimestampPresenza']; all_exportable_cols = [col for col in all_possible_cols if col not in internal_cols_to_exclude]
                    default_cols_export_ordered = ['DataPresenza','OraPresenza','DenominazioneAttività','CFU','Cognome','Nome','PercorsoInternal','PercorsoOriginaleSenzaArt13Internal']
                    default_cols_final = [col for col in default_cols_export_ordered if col in all_exportable_cols]

                    st.markdown("**Esempio Record Dati (prima riga):**")
                    if not current_df_for_tab3.empty:
                        example_data = current_df_for_tab3.head(1)[all_exportable_cols].to_dict(orient='records')[0]; st.json(example_data)
                    else: st.caption("Nessun dato da mostrare.")

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
                    
                    if st.button("Genera ed Esporta File Excel", key="export_excel_ordered_v215"):
                        if not selected_cols_export_ordered: st.warning("Seleziona almeno una colonna.")
                        else:
                            overall_success = True; sheets_written = 0; error_messages = []
                            try:
                                output = BytesIO()
                                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                                    unique_courses = current_df_for_tab3[course_col_export].unique(); unique_courses = sorted([str(c) for c in unique_courses if pd.notna(c)])
                                    if not unique_courses: st.error(f"Nessun valore unico trovato in '{course_col_export}'."); overall_success = False
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
                                            
                                            prog_text = f"Foglio: {sheet_name_cleaned} ({i+1}/{len(unique_courses)})"; 
                                            prog_bar.progress((i + 1) / len(unique_courses), text=prog_text)
                                            df_sheet = current_df_for_tab3[current_df_for_tab3[course_col_export] == course_value].copy()
                                            if df_sheet.empty: st.write(f"Info: Nessun dato per '{sheet_name_cleaned}', foglio saltato."); continue
                                            
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
                                            if not final_ordered_cols_for_sheet: st.write(f"Info: Nessuna colonna selezionata trovata per '{sheet_name_cleaned}', foglio saltato."); continue
                                            df_sheet_export = df_sheet[final_ordered_cols_for_sheet]
                                            if df_sheet_export.empty: st.write(f"Info: Nessun dato per '{sheet_name_cleaned}' dopo selezione colonne, foglio saltato."); continue
                                            try:
                                                rename_map_export = {'PercorsoInternal': 'Tipo Percorso', 'PercorsoOriginaleInternal': 'Percorso Originale Input', 'PercorsoOriginaleSenzaArt13Internal': 'Denominazione Percorso', 'DenominazioneAttivitaNormalizzataInternal': 'Attività Elaborata'}
                                                cols_to_rename_final = {k: v for k, v in rename_map_export.items() if k in df_sheet_export.columns}
                                                df_sheet_export = df_sheet_export.rename(columns=cols_to_rename_final)
                                                df_sheet_export.to_excel(writer, sheet_name=sheet_name_cleaned, index=False)
                                                sheets_written += 1
                                            except Exception as sheet_error: error_msg = f"Errore scrittura foglio '{sheet_name_cleaned}': {sheet_error}"; st.warning(error_msg); error_messages.append(error_msg); overall_success = False
                            except Exception as writer_error: st.error(f"Errore generale creazione Excel: {writer_error}"); st.exception(writer_error); overall_success = False
                            if overall_success and sheets_written > 0:
                                prog_bar.progress(1.0, text="Completato!")
                                output.seek(0); ts = datetime.now().strftime("%Y%m%d_%H%M")
                                
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
                                if error_messages: st.warning("Alcuni fogli potrebbero aver avuto problemi:"); [st.caption(msg) for msg in error_messages]
                                st.download_button(label="📥 Scarica Report Excel Dettaglio", data=output, file_name=fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_excel_course_ordered_v215")
                            elif sheets_written == 0 and overall_success: st.warning("Nessun dato trovato per alcun percorso. File Excel non generato.")
                            else:
                                st.error("Generazione file Excel fallita o nessun foglio valido scritto.")
                                if error_messages: st.warning("Dettaglio errori:"); [st.caption(msg) for msg in error_messages]

            # Questo else si riferisce a 'if not attendance_df.empty:'
            else:
                st.info("Nessun dato di presenza aggregato da visualizzare (verifica i dati o le colonne necessarie).")
        else:
             st.info("Nessun dato valido caricato.")

# Messaggio iniziale se nessun file caricato
else:
    st.info("👈 Per iniziare, carica un file Excel dalla barra laterale.")
    st.markdown("""
    ## Istruzioni d'uso
    1.  **Carica file Excel** (.xlsx) dalla sidebar.
    2.  **(Opzionale)** Vedi anteprima originale.
    3.  **Analisi Dati (Tab 1):** Controlla statistiche e dati elaborati.
    4.  **Gestione Duplicati (Tab 2):** Identifica e rimuovi timbrature ravvicinate.
    5.  **Calcolo Presenze ed Esportazione (Tab 3):**
        *   Filtra per **Percorso (Senza Art.13)** e opzionalmente per **Studente** per vedere i dettagli record per record.
        *   **Esporta in Excel:** 
            - Seleziona e ordina le colonne desiderate
            - Scegli un periodo di date per l'esportazione (opzionale)
            - Genera file **dettagliato**, un foglio per codice percorso **estratto e mostrato all'inizio [Codice]** 
            - I percorsi mostrano ora il codice all'inizio nel formato [A-30] Nome Percorso

    ### Formato File Suggerito
    *   `CodiceFiscale`, `DataPresenza`, `OraPresenza`, `DenominazionePercorso` (o `percoro`) - Obbligatori
    *   `Nome`, `Cognome`, `DenominazioneAttività`, `CodicePercorso` - Opzionali
    """)

# Footer
st.markdown("---")
st.markdown("### Gestione Presenze - Versione beta 1.1") # Aggiorna versione - aggiunto codice percorso all'inizio