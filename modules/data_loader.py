# Funzioni per il caricamento e la trasformazione dei dati
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta, date, time
import difflib
import re
# Import diretto delle funzioni necessarie senza usare l'importazione relativa
# per evitare problemi di risoluzione dei moduli

def normalize_generic(name):
    """Rimuove 'art.13' e spazi dalle stringhe"""
    if not isinstance(name, str): return name
    normalized = re.sub(r'\s*\(?art\.?\s*13\.?\)?.*$', '', name, flags=re.IGNORECASE)
    return normalized.strip()
    
def reposition_code_to_front(text):
    """Prende il testo, estrae il codice tra parentesi e lo riposiziona all'inizio della stringa."""
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
    """Trasforma il percorso in base al codice"""
    if pd.isna(codice): return default_name
    codice_str = str(codice).strip()
    if len(codice_str) < 3: return default_name
    prefix = codice_str[:3]
    if prefix == '600': return "PeF60 All. 1"
    elif prefix == '300': return "PeF30 All. 2"
    elif prefix == '360': return "PeF36 All. 5"
    elif prefix == '200': return "PeF30 art. 13"
    else: return default_name

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

@st.cache_data
@st.cache_data
def load_data(uploaded_file):
    """Carica e preprocessa i dati dal file Excel caricato"""
    if uploaded_file is None: return None
    try:
        # Carica i dati dei CFU
        cfu_data = load_cfu_data()
        if cfu_data.empty:
            st.warning("Non è stato possibile caricare i dati dei CFU. I CFU non saranno disponibili.")
        
        df = pd.read_excel(uploaded_file)
        original_columns = df.columns.tolist()
        required_cols = ['CodiceFiscale', 'DataPresenza', 'OraPresenza']
        
        if not ('percoro' in df.columns or 'DenominazionePercorso' in df.columns): 
            st.error("Colonna Percorso ('percoro' o 'DenominazionePercorso') non trovata.")
            return None
            
        for col in required_cols:
            if col not in df.columns: 
                st.error(f"Colonna obbligatoria '{col}' non trovata.")
                return None
                
        df['CodiceFiscale'] = df['CodiceFiscale'].astype(str).str.strip()
        
        # Gestione della colonna Email
        if 'recapito_ateneo' in df.columns:
            df['Email'] = df['recapito_ateneo']
            st.success("Email caricate dalla colonna 'recapito_ateneo'")
        else:
            st.warning("Colonna 'recapito_ateneo' per le email non trovata nel file. Le email non saranno disponibili.")
            df['Email'] = ''
            
        percorso_col_original_name = 'percoro' if 'percoro' in df.columns else 'DenominazionePercorso'
        df['PercorsoOriginaleInternal'] = df[percorso_col_original_name]
        df['PercorsoOriginaleSenzaArt13Internal'] = df['PercorsoOriginaleInternal'].apply(normalize_generic)
        # Applica la trasformazione per spostare i codici all'inizio
        df['PercorsoOriginaleSenzaArt13Internal'] = df['PercorsoOriginaleSenzaArt13Internal'].apply(reposition_code_to_front)
        
        if 'CodicePercorso' in df.columns: 
            df['PercorsoInternal'] = df.apply(lambda row: transform_by_codice_percorso(row.get('CodicePercorso'), row['PercorsoOriginaleSenzaArt13Internal']), axis=1)
        else: 
            df['PercorsoInternal'] = df['PercorsoOriginaleSenzaArt13Internal']
            
        if 'DenominazioneCds' in df.columns: 
            st.warning("'DenominazioneCds' trovata. Verrà ignorata/rimossa.")
            df = df.drop(columns=['DenominazioneCds'], errors='ignore')
            
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
                
        # Gestione delle date e orari
        try: 
            df['DataPresenza'] = pd.to_datetime(df['DataPresenza'], errors='coerce').dt.date
            df.loc[pd.isna(df['DataPresenza']), 'DataPresenza'] = pd.NaT
        except Exception as e: 
            st.warning(f"Problema conversione 'DataPresenza': {e}.")
            df['DataPresenza'] = pd.NaT if 'DataPresenza' not in df.columns else df['DataPresenza']
            
        try:
            def parse_time(t):
                if isinstance(t, time): return t
                if isinstance(t, datetime): return t.time()
                if isinstance(t, pd.Timestamp): return t.time()
                if pd.isna(t): return pd.NaT
                try: return pd.to_datetime(str(t), format='%H:%M:%S', errors='raise').time()
                except ValueError:
                    try:
                        if isinstance(t, (int, float)) and 0 <= t < 1: 
                            total_seconds = int(t * 24 * 60 * 60)
                            hours, remainder = divmod(total_seconds, 3600)
                            minutes, seconds = divmod(remainder, 60)
                            return time(hours, minutes, seconds)
                        return pd.to_datetime(str(t), errors='raise').time()
                    except (ValueError, TypeError): return pd.NaT
                    
            df['OraPresenza'] = df['OraPresenza'].apply(parse_time)
            df.loc[pd.isna(df['OraPresenza']), 'OraPresenza'] = pd.NaT
            
        except Exception as e: 
            st.warning(f"Problema conversione 'OraPresenza': {e}.")
            df['OraPresenza'] = pd.NaT if 'OraPresenza' not in df.columns else df['OraPresenza']
            
        def combine_dt(row):
            if pd.notna(row['DataPresenza']) and isinstance(row['DataPresenza'], date) and pd.notna(row['OraPresenza']) and isinstance(row['OraPresenza'], time):
                try: return pd.Timestamp.combine(row['DataPresenza'], row['OraPresenza'])
                except Exception: return pd.NaT
            return pd.NaT
            
        df['TimestampPresenza'] = df.apply(combine_dt, axis=1) # TimestampPresenza è creata qui
        initial_rows = len(df)
        df.dropna(subset=['TimestampPresenza', 'CodiceFiscale'], inplace=True)
        df['DataPresenza'] = df['TimestampPresenza'].dt.date
        df['OraPresenza'] = df['TimestampPresenza'].dt.time
        removed_rows = initial_rows - len(df)
        
        if removed_rows > 0: 
            st.warning(f"Rimossi {removed_rows} record con CF, Data, Ora o Timestamp mancanti/non validi.")
            
        if 'Nome' not in df.columns: df['Nome'] = ''
        if 'Cognome' not in df.columns: df['Cognome'] = ''
        
        final_cols = ['CodiceFiscale', 'Nome', 'Cognome', 'Email', 'DataPresenza', 'OraPresenza', 
                      'PercorsoOriginaleInternal', 'PercorsoOriginaleSenzaArt13Internal', 'PercorsoInternal', 'TimestampPresenza']
                      
        if 'DenominazioneAttività' in df.columns: final_cols.append('DenominazioneAttività')
        if activity_col_norm_internal in df.columns: final_cols.append(activity_col_norm_internal)
        if 'CodicePercorso' in original_columns and 'CodicePercorso' in df.columns: final_cols.append('CodicePercorso')
        if 'CFU' in df.columns: final_cols.append('CFU')  # Aggiungi la colonna CFU all'elenco delle colonne da mantenere
        
        cols_to_keep = [col for col in final_cols if col in df.columns]
        df_final = df[cols_to_keep].copy()
        return df_final
        
    except Exception as e: 
        st.error(f"Errore critico caricamento/elaborazione file: {e}")
        st.exception(e)
        return None
