# Funzioni per il caricamento e la trasformazione dei dati
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta, date, time
import difflib
import re
import os  # Aggiunto per verificare l'esistenza dei file

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
            
        # Carica i dati degli iscritti
        enrolled_students = load_enrolled_students_data()
        if enrolled_students.empty:
            st.warning("Non è stato possibile caricare i dati degli iscritti. Le informazioni aggiuntive degli studenti non saranno disponibili.")
        
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
        
        # Integra i dati degli iscritti se disponibili
        if 'enrolled_students' in locals():
            if enrolled_students.empty:
                st.error("DataFrame degli iscritti vuoto, impossibile integrare i dati.")
            else:
                with st.spinner("Integrazione dati degli iscritti in corso..."):
                    st.warning(f"Tentativo di integrazione di {enrolled_students.shape[0]} record di iscritti nei dati presenze...")
                    
                    # Stampa alcune righe di esempio degli iscritti per debug
                    st.warning("Esempio dati iscritti (prime 3 righe):")
                    if len(enrolled_students) > 0:
                        st.write(enrolled_students.head(3))
                    
                    # Eseguo l'integrazione
                    df = match_students_data(df, enrolled_students)
                    
                    # Verifica se l'integrazione è avvenuta correttamente
                    enrolled_cols = ['Percorso', 'Codice_Classe_di_concorso', 'Dipartimento']
                    integrated_cols = [col for col in enrolled_cols if col in df.columns]
                    if integrated_cols:
                        for col in integrated_cols:
                            non_null_count = df[col].notna().sum()
                            st.warning(f"Integrazione colonna '{col}': {non_null_count} record non vuoti su {len(df)}")
                    else:
                        st.error("ATTENZIONE: Nessuna colonna degli iscritti è stata integrata nei dati!")
                
        # Aggiungo le nuove colonne degli iscritti alla lista di colonne da mantenere
        new_cols = ['Percorso', 'Codice_Classe_di_concorso', 'Codice_classe_di_concorso_e_denominazione', 
                    'Dipartimento', 'LogonName', 'Matricola']
        for col in new_cols:
            if col in df.columns:
                final_cols.append(col)
                
        cols_to_keep = [col for col in final_cols if col in df.columns]
        df_final = df[cols_to_keep].copy()
        return df_final
        
    except Exception as e: 
        st.error(f"Errore critico caricamento/elaborazione file: {e}")
        st.exception(e)
        return None

@st.cache_data(ttl=1) # Disabilitiamo quasi completamente la cache per il debug
def load_enrolled_students_data():
    """Carica i dati degli studenti iscritti dal file CSV."""
    try:
        # Usiamo il path assoluto per essere sicuri
        file_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'modules', 'dati', 'iscritti_29_aprile.csv')
        if not os.path.exists(file_path):
            st.error(f"File degli iscritti non trovato: {file_path}")
            # Prova con un path relativo come fallback
            file_path = 'modules/dati/iscritti_29_aprile.csv'
            if not os.path.exists(file_path):
                st.error(f"File degli iscritti non trovato nemmeno al path relativo: {file_path}")
                return pd.DataFrame()
            
        # Stampiamo informazioni di debug
        st.warning(f"Caricamento dati iscritti da: {os.path.abspath(file_path)}")
        
        # Carica il CSV con delimitatore punto e virgola e vari encoding come fallback
        try:
            enrolled_df = pd.read_csv(file_path, delimiter=';', encoding='utf-8-sig')
        except UnicodeDecodeError:
            try:
                enrolled_df = pd.read_csv(file_path, delimiter=';', encoding='utf-8')
            except UnicodeDecodeError:
                enrolled_df = pd.read_csv(file_path, delimiter=';', encoding='latin-1')
        
        # Stampa informazioni sul dataframe caricato
        st.warning(f"File iscritti caricato: {enrolled_df.shape[0]} righe, {enrolled_df.shape[1]} colonne")
        st.warning(f"Colonne disponibili: {', '.join(enrolled_df.columns.tolist())}")
        
        # Verifico che ci siano le colonne necessarie
        required_cols = ['Cognome', 'Nome', 'CodiceFiscale']
        if not all(col in enrolled_df.columns for col in required_cols):
            missing_cols = [col for col in required_cols if col not in enrolled_df.columns]
            st.error(f"File degli iscritti: colonne richieste mancanti ({', '.join(missing_cols)})")
            return pd.DataFrame()
        
        # Pulisco i dati
        if 'CodiceFiscale' in enrolled_df.columns:
            enrolled_df['CodiceFiscale'] = enrolled_df['CodiceFiscale'].astype(str).str.strip()
        
        # Creo una colonna di identificazione per facilitare il matching
        enrolled_df['NomeCognome'] = enrolled_df['Nome'].str.lower() + ' ' + enrolled_df['Cognome'].str.lower()
        
        return enrolled_df
    except Exception as e:
        st.error(f"Errore durante il caricamento del file degli iscritti: {e}")
        return pd.DataFrame()

def match_students_data(df_presences, df_enrolled):
    """
    Integra i dati degli studenti iscritti nel dataframe delle presenze.
    
    Args:
        df_presences: DataFrame con i dati delle presenze
        df_enrolled: DataFrame con i dati degli studenti iscritti
    
    Returns:
        DataFrame con i dati integrati
    """
    try:
        if df_presences.empty:
            st.warning("DataFrame presenze vuoto, impossibile eseguire il match.")
            return df_presences
            
        if df_enrolled.empty:
            st.warning("DataFrame iscritti vuoto, impossibile eseguire il match.")
            return df_presences
        
        # Verifico che le colonne necessarie esistano
        required_presence_cols = ['Cognome', 'Nome']
        if not all(col in df_presences.columns for col in required_presence_cols):
            st.warning(f"I dati di presenza non contengono tutte le colonne necessarie per il matching: {required_presence_cols}")
            return df_presences
            
        # Verifico che le colonne degli iscritti esistano
        required_enrolled_cols = ['Cognome', 'Nome', 'CodiceFiscale', 'Percorso']
        if not all(col in df_enrolled.columns for col in required_enrolled_cols):
            missing = [col for col in required_enrolled_cols if col not in df_enrolled.columns]
            st.warning(f"I dati degli iscritti non contengono tutte le colonne necessarie: mancano {missing}")
            return df_presences
        
        # Creo una copia del dataframe per non modificare l'originale
        result_df = df_presences.copy()
    
        # Preparo colonna per il matching
        result_df['NomeCognome'] = result_df['Nome'].str.lower() + ' ' + result_df['Cognome'].str.lower()
    
        # Colonne da integrare dal file degli iscritti
        cols_to_merge = ['Percorso', 'Codice_Classe_di_concorso', 'Codice_classe_di_concorso_e_denominazione', 
                     'Dipartimento', 'LogonName', 'Matricola']
    
        # Contatori per statistiche
        matched_by_cf = 0
        matched_by_name = 0
    
        # Cerca corrispondenze in base al Codice Fiscale (metodo principale e più affidabile)
        if 'CodiceFiscale' in result_df.columns and 'CodiceFiscale' in df_enrolled.columns:
            # Creo un DataFrame con solo le righe con CF compilato
            cf_mask = result_df['CodiceFiscale'].notna() & (result_df['CodiceFiscale'] != '')
            with_cf = result_df[cf_mask]
            
            # Effettuo il merge per Codice Fiscale (lascia i valori originali se non trova match)
            if not with_cf.empty:
                merged_cf = pd.merge(
                    with_cf, 
                    df_enrolled[cols_to_merge + ['CodiceFiscale']], 
                    on='CodiceFiscale', 
                    how='left',
                    suffixes=('', '_iscritti')
                )
                
                # Aggiorno il DataFrame risultante con le righe contenenti CF
                matched_indices = []
                for idx, row in merged_cf.iterrows():
                    if idx in result_df.index:
                        match_found = False
                        for col in cols_to_merge:
                            if col in merged_cf.columns and pd.notna(row[col]):
                                result_df.loc[idx, col] = row[col]
                                match_found = True
                        if match_found:
                            matched_indices.append(idx)
                            matched_by_cf += 1
    
        # Per le righe senza corrispondenza per CF, provo con Nome e Cognome
        no_match_mask = ~result_df.index.isin(matched_indices) if 'matched_indices' in locals() else pd.Series(True, index=result_df.index)
        if no_match_mask.any():
            no_match_df = result_df[no_match_mask].copy()
            merged_name = pd.merge(
                no_match_df,
                df_enrolled[cols_to_merge + ['NomeCognome']],
                on='NomeCognome',
                how='left',
                suffixes=('', '_iscritti')
            )
            
            # Aggiorno il DataFrame risultante con le righe contenenti solo Nome/Cognome
            for idx, row in merged_name.iterrows():
                if idx in result_df.index:
                    match_found = False
                    for col in cols_to_merge:
                        if col in merged_name.columns and pd.notna(row[col]):
                            result_df.loc[idx, col] = row[col]
                            match_found = True
                    if match_found:
                        matched_by_name += 1
    
        # Elimino la colonna temporanea
        if 'NomeCognome' in result_df.columns:
            result_df = result_df.drop(columns=['NomeCognome'])
        
        # Mostro statistiche sul matching
        total_records = len(result_df)
        total_matches = matched_by_cf + matched_by_name
        
        # Verifica e mostra statistiche di integrazione dettagliate
        percorso_integrati = result_df['Percorso'].notna().sum() if 'Percorso' in result_df.columns else 0
        classe_concorso_integrati = result_df['Codice_Classe_di_concorso'].notna().sum() if 'Codice_Classe_di_concorso' in result_df.columns else 0
        
        if percorso_integrati == 0 and 'Percorso' in result_df.columns:
            st.error(f"ERRORE CRITICO: La colonna 'Percorso' esiste ma non contiene dati validi!")
        
        st.warning(f"Matching studenti: {total_matches}/{total_records} record abbinati "
                f"({matched_by_cf} tramite CF, {matched_by_name} tramite Nome/Cognome)")
        st.warning(f"Colonne integrate: Percorso={percorso_integrati}, Classe Concorso={classe_concorso_integrati}")
        
        # Verifica se ci sono record nel dataframe dei presenti che hanno corrispondenza negli iscritti
        if 'CodiceFiscale' in result_df.columns and 'CodiceFiscale' in df_enrolled.columns:
            cf_in_comune = set(result_df['CodiceFiscale'].dropna()) & set(df_enrolled['CodiceFiscale'].dropna())
            st.warning(f"Codici fiscali in comune tra presenze e iscritti: {len(cf_in_comune)}")
        
        return result_df
    except Exception as e:
        st.error(f"Errore durante il matching dei dati degli iscritti: {e}")
        st.exception(e)
        return df_presences
