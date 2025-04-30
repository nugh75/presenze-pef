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
        # Provo prima con il percorso assoluto
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        file_path = os.path.join(base_path, 'crediti.csv')
        
        # Controllo se il file esiste al path assoluto
        if not os.path.exists(file_path):
            st.warning(f"File dei CFU non trovato al path assoluto: {file_path}")
            # Provo con un path relativo come fallback
            file_path = 'crediti.csv'
            if not os.path.exists(file_path):
                st.error(f"File dei CFU non trovato nemmeno al path relativo: {file_path}")
                return pd.DataFrame()
        
        st.info(f"Caricamento dati CFU da: {os.path.abspath(file_path)}")
        
        # Carica il CSV con vari encoding come fallback
        try:
            cfu_df = pd.read_csv(file_path, encoding='utf-8-sig')
        except UnicodeDecodeError:
            try:
                cfu_df = pd.read_csv(file_path, encoding='utf-8')
            except UnicodeDecodeError:
                cfu_df = pd.read_csv(file_path, encoding='latin-1')
        
        if 'DenominazioneAttività' not in cfu_df.columns or 'CFU' not in cfu_df.columns:
            st.error(f"Il file dei CFU non contiene le colonne richieste: 'DenominazioneAttività' e 'CFU'. Colonne trovate: {', '.join(cfu_df.columns)}")
            return pd.DataFrame()
            
        # Stampa informazioni sul dataframe caricato
        st.info(f"File CFU caricato: {cfu_df.shape[0]} righe, {cfu_df.shape[1]} colonne")
        
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
        st.info("Inizializzazione caricamento dati...")
        
        # Carica i dati dei CFU
        st.info("Caricamento dati CFU...")
        cfu_data = load_cfu_data()
        if cfu_data.empty:
            st.warning("Non è stato possibile caricare i dati dei CFU. I CFU non saranno disponibili.")
        else:
            st.success(f"Dati CFU caricati con successo: {len(cfu_data)} attività trovate.")
            
        # Carica i dati degli iscritti
        st.info("Caricamento dati iscritti...")
        enrolled_students = load_enrolled_students_data()
        if enrolled_students.empty:
            st.warning("Non è stato possibile caricare i dati degli iscritti. Le informazioni aggiuntive degli studenti non saranno disponibili.")
        else:
            st.success(f"Dati iscritti caricati con successo: {len(enrolled_students)} iscritti trovati.")
        
        st.info("Caricamento file Excel...")
        df = pd.read_excel(uploaded_file)
        st.success(f"File Excel caricato con successo: {len(df)} record trovati.")
        
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
        # Nota: le colonne PercorsoOriginaleInternal, PercorsoOriginaleSenzaArt13Internal e PercorsoInternal sono state rimosse
            
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
                # Se è già un oggetto time, lo restituiamo direttamente
                if isinstance(t, time): return t
                
                # Se è un oggetto datetime o timestamp, estraiamo l'ora
                if isinstance(t, datetime): return t.time()
                if isinstance(t, pd.Timestamp): return t.time()
                
                # Gestione dei valori mancanti
                if pd.isna(t): return pd.NaT
                
                # Gestione dei formati numerici (Excel salva spesso le ore come decimali)
                if isinstance(t, (int, float)) and 0 <= t < 1: 
                    total_seconds = int(t * 24 * 60 * 60)
                    hours, remainder = divmod(total_seconds, 3600)
                    minutes, seconds = divmod(remainder, 60)
                    return time(hours, minutes, seconds)
                
                # Normalizza la stringa prima della conversione
                str_t = str(t).strip()
                
                # Se la stringa contiene già i due punti, è probabilmente già in formato orario
                if ':' in str_t:
                    # Prova i formati orari comuni
                    for fmt in ['%H:%M:%S', '%H:%M', '%I:%M:%S %p', '%I:%M %p']:
                        try:
                            return pd.to_datetime(str_t, format=fmt, errors='raise').time()
                        except ValueError:
                            continue
                
                # Se tutto fallisce, tenta una conversione generica
                try:
                    return pd.to_datetime(str_t, errors='raise').time()
                except (ValueError, TypeError):
                    return pd.NaT
                    
            # Applica la funzione di parsing migliorata
            df['OraPresenza'] = df['OraPresenza'].apply(parse_time)
            df.loc[pd.isna(df['OraPresenza']), 'OraPresenza'] = pd.NaT
            
            # Verifica la percentuale di conversioni riuscite
            ora_validi = df['OraPresenza'].notna().sum()
            if ora_validi < len(df):
                st.warning(f"Attenzione: {len(df) - ora_validi} valori di OraPresenza non sono stati convertiti correttamente.")
            
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
        
        # Standardizzazione dei campi data e ora a partire dal TimestampPresenza
        # Questo garantisce che i campi siano sempre nel formato corretto
        df['DataPresenza'] = df['TimestampPresenza'].dt.date
        df['OraPresenza'] = df['TimestampPresenza'].dt.time
        
        # Verifica che i campi siano stati convertiti correttamente
        if df['DataPresenza'].isna().any() or df['OraPresenza'].isna().any():
            st.warning("Alcuni valori di DataPresenza o OraPresenza non sono stati estratti correttamente da TimestampPresenza.")
            
        # Log informativo
        st.info("Formati dati standardizzati: DataPresenza (date) e OraPresenza (time) estratti da TimestampPresenza")
        
        removed_rows = initial_rows - len(df)
        
        if removed_rows > 0: 
            st.warning(f"Rimossi {removed_rows} record con CF, Data, Ora o Timestamp mancanti/non validi.")
            
        if 'Nome' not in df.columns: df['Nome'] = ''
        if 'Cognome' not in df.columns: df['Cognome'] = ''
        
        final_cols = ['CodiceFiscale', 'Nome', 'Cognome', 'Email', 'DataPresenza', 'OraPresenza', 'TimestampPresenza']
                      
        if 'DenominazioneAttività' in df.columns: final_cols.append('DenominazioneAttività')
        if activity_col_norm_internal in df.columns: final_cols.append(activity_col_norm_internal)
        if 'CodicePercorso' in original_columns and 'CodicePercorso' in df.columns: final_cols.append('CodicePercorso')
        if 'CFU' in df.columns: final_cols.append('CFU')  # Aggiungi la colonna CFU all'elenco delle colonne da mantenere
        
        # Integra i dati degli iscritti se disponibili
        if not enrolled_students.empty:
            with st.spinner("Integrazione dati degli iscritti in corso..."):
                st.info(f"Integrazione di {enrolled_students.shape[0]} record di iscritti nei dati presenze...")
                
                # Preparo i dati per il matching
                # Normalizzazione di Nome e Cognome per migliorare il matching
                if 'Nome' in df.columns:
                    df['Nome'] = df['Nome'].astype(str).str.strip().str.lower().str.title()
                if 'Cognome' in df.columns:
                    df['Cognome'] = df['Cognome'].astype(str).str.strip().str.lower().str.title()
                
                if 'Nome' in enrolled_students.columns:
                    enrolled_students['Nome'] = enrolled_students['Nome'].astype(str).str.strip().str.lower().str.title()
                if 'Cognome' in enrolled_students.columns:
                    enrolled_students['Cognome'] = enrolled_students['Cognome'].astype(str).str.strip().str.lower().str.title()
                
                # Eseguo l'integrazione
                df = match_students_data(df, enrolled_students)
                
                # Verifica se l'integrazione è avvenuta correttamente
                enrolled_cols = ['Percorso', 'Codice_Classe_di_concorso', 'Codice_classe_di_concorso_e_denominazione', 'Dipartimento', 'Matricola']
                integrated_cols = [col for col in enrolled_cols if col in df.columns]
                if integrated_cols:
                    for col in integrated_cols:
                        non_null_count = df[col].notna().sum()
                        st.info(f"Integrazione colonna '{col}': {non_null_count} record non vuoti su {len(df)} ({non_null_count/len(df)*100:.1f}%)")
                    
                    # Se l'integrazione è a zero o molto bassa, segnala un problema
                    if all(df[col].notna().sum() == 0 for col in integrated_cols):
                        st.error("ERRORE: Nessuna integrazione dei dati iscritti è avvenuta!")
                        
                        # Analisi delle possibili cause
                        if 'CodiceFiscale' in df.columns and 'CodiceFiscale' in enrolled_students.columns:
                            df_cf = set(df['CodiceFiscale'].astype(str).str.strip().unique())
                            enrolled_cf = set(enrolled_students['CodiceFiscale'].astype(str).str.strip().unique())
                            common_cf = df_cf.intersection(enrolled_cf)
                            
                            st.error(f"Codici fiscali comuni tra i dataset: {len(common_cf)} su {len(df_cf)} nei dati presenze e {len(enrolled_cf)} negli iscritti")
                            
                            if len(common_cf) > 0:
                                st.warning("Analisi di alcune corrispondenze (primi 3 CF in comune):")
                                for cf in list(common_cf)[:3]:
                                    st.write(f"CF: {cf}")
                                    st.write("Record presenza:", df[df['CodiceFiscale'] == cf].iloc[0])
                                    st.write("Record iscritto:", enrolled_students[enrolled_students['CodiceFiscale'] == cf].iloc[0])
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
        # Costruisco il path con varie possibilità per robustezza
        # Percorso 1: Dalla directory di root del progetto
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        file_path = os.path.join(base_path, 'modules', 'dati', 'iscritti_29_aprile.csv')
        
        # Verifica se il file esiste
        if not os.path.exists(file_path):
            st.warning(f"File degli iscritti non trovato al path: {file_path}")
            
            # Percorso 2: Considerando che __file__ potrebbe già essere in modules/
            file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'dati', 'iscritti_29_aprile.csv')
            
            if not os.path.exists(file_path):
                st.warning(f"File degli iscritti non trovato al path alternativo: {file_path}")
                
                # Percorso 3: Path relativo come ultima risorsa
                file_path = 'modules/dati/iscritti_29_aprile.csv'
                
                if not os.path.exists(file_path):
                    st.error(f"File degli iscritti non trovato in nessun percorso! Ultimo tentativo: {file_path}")
                    # Prima di rinunciare, cerca nella directory corrente
                    import glob
                    csv_files = glob.glob('**/*.csv', recursive=True)
                    st.error(f"File CSV disponibili nella directory: {csv_files}")
                    return pd.DataFrame()
        
        # Stampiamo informazioni di debug
        st.info(f"Caricamento dati iscritti da: {os.path.abspath(file_path)}")
        
        # Carica il CSV con delimitatore punto e virgola e vari encoding come fallback
        try:
            enrolled_df = pd.read_csv(file_path, delimiter=';', encoding='utf-8-sig')
        except UnicodeDecodeError:
            try:
                enrolled_df = pd.read_csv(file_path, delimiter=';', encoding='utf-8')
            except UnicodeDecodeError:
                enrolled_df = pd.read_csv(file_path, delimiter=';', encoding='latin-1')
        
        # Stampa informazioni sul dataframe caricato
        st.info(f"File iscritti caricato: {enrolled_df.shape[0]} righe, {enrolled_df.shape[1]} colonne")
        st.info(f"Colonne disponibili: {', '.join(enrolled_df.columns.tolist())}")
        
        # Verifico che ci siano le colonne necessarie
        required_cols = ['Cognome', 'Nome', 'CodiceFiscale', 'Codice_Classe_di_concorso']
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
            
        # Stampo le prime righe dei due DataFrame per debug
        st.info("DATI PRESENZE (prime 2 righe):")
        st.write(df_presences.head(2))
        
        st.info("DATI ISCRITTI (prime 2 righe):")
        st.write(df_enrolled.head(2))
        
        # Verifico che le colonne necessarie esistano
        required_presence_cols = ['Cognome', 'Nome']
        if not all(col in df_presences.columns for col in required_presence_cols):
            st.warning(f"I dati di presenza non contengono tutte le colonne necessarie per il matching: {required_presence_cols}")
            return df_presences
            
        # Verifico che le colonne degli iscritti esistano
        required_enrolled_cols = ['Cognome', 'Nome', 'CodiceFiscale', 'Codice_Classe_di_concorso']
        if not all(col in df_enrolled.columns for col in required_enrolled_cols):
            missing = [col for col in required_enrolled_cols if col not in df_enrolled.columns]
            st.warning(f"I dati degli iscritti non contengono tutte le colonne necessarie: mancano {missing}")
            # Stampo le colonne disponibili
            st.info(f"Colonne disponibili nel file iscritti: {', '.join(df_enrolled.columns)}")
            return df_presences
        
        # Creo una copia del dataframe per non modificare l'originale
        result_df = df_presences.copy()
    
        # Preparo colonna per il matching
        result_df['NomeCognome'] = result_df['Nome'].str.lower() + ' ' + result_df['Cognome'].str.lower()
    
        # Colonne da integrare dal file degli iscritti
        # Assicuriamoci di includere solo colonne che esistono nel DataFrame degli iscritti
        available_cols = [col for col in ['Percorso', 'Codice_Classe_di_concorso', 'Codice_classe_di_concorso_e_denominazione', 
                     'Dipartimento', 'LogonName', 'Matricola'] if col in df_enrolled.columns]
                     
        cols_to_merge = available_cols
    
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

@st.cache_data
def process_datetime_field(df, field_name):
    """
    Processa un campo contenente data e ora nel formato '4/29/25 18:10:26'
    e crea due nuovi campi: DataPresenza e OraPresenza
    
    Args:
        df: DataFrame contenente i dati
        field_name: Nome del campo contenente data e ora (es. "Ora di inizio")
    
    Returns:
        DataFrame con i nuovi campi DataPresenza e OraPresenza
    """
    if field_name not in df.columns:
        st.error(f"Campo {field_name} non trovato nel dataframe")
        return df
    
    result_df = df.copy()
    
    try:
        # Converti il campo in datetime
        result_df['timestamp_temp'] = pd.to_datetime(result_df[field_name], errors='coerce')
        
        # Estrai data e ora
        result_df['DataPresenza'] = result_df['timestamp_temp'].dt.date
        result_df['OraPresenza'] = result_df['timestamp_temp'].dt.time
        
        # Rimuovi il campo temporaneo
        result_df = result_df.drop(columns=['timestamp_temp'])
        
        # Conta quanti record sono stati convertiti con successo
        successful_conversions = result_df['DataPresenza'].notna().sum()
        total_records = len(result_df)
        
        if successful_conversions < total_records:
            st.warning(f"Conversione del campo {field_name}: {successful_conversions}/{total_records} record convertiti con successo")
        else:
            st.success(f"Conversione del campo {field_name} completata con successo: {successful_conversions}/{total_records} record convertiti")
            
        return result_df
    except Exception as e:
        st.error(f"Errore durante la conversione del campo {field_name}: {e}")
        return df

@st.cache_data
def load_multiple_files(uploaded_files):
    """
    Carica e preprocessa i dati da più file Excel/CSV caricati
    
    Args:
        uploaded_files: Lista di file caricati dall'utente
    
    Returns:
        DataFrame combinato con i dati di tutti i file
    """
    if not uploaded_files:
        st.error("Nessun file caricato")
        return None
        
    all_dataframes = []
    processed_files = 0
    failed_files = 0
    
    for uploaded_file in uploaded_files:
        try:
            # Determina il tipo di file
            file_ext = uploaded_file.name.split('.')[-1].lower()
            
            if file_ext == 'xlsx':
                df = pd.read_excel(uploaded_file)
            elif file_ext in ['csv', 'txt']:
                # Tenta di caricare CSV con diverse codifiche e delimitatori
                try:
                    df = pd.read_csv(uploaded_file, sep=',', encoding='utf-8-sig')
                except UnicodeDecodeError:
                    try:
                        df = pd.read_csv(uploaded_file, sep=',', encoding='utf-8')
                    except UnicodeDecodeError:
                        try:
                            df = pd.read_csv(uploaded_file, sep=';', encoding='utf-8-sig')
                        except UnicodeDecodeError:
                            df = pd.read_csv(uploaded_file, sep=';', encoding='latin-1')
            else:
                st.error(f"Formato file non supportato: {file_ext}")
                continue
                
            # Verifica schema dati per i nuovi formati
            columns = df.columns.tolist()
            
            # Identificazione e mappatura colonne
            # Formato 1: Con colonne DataPresenza e OraPresenza già presenti
            if 'DataPresenza' in columns and 'OraPresenza' in columns:
                st.info(f"File {uploaded_file.name} con formato standard riconosciuto")
            
            # Formato 2: Con colonna "Ora di inizio" che contiene data e ora insieme
            elif 'Ora di inizio' in columns:
                st.info(f"File {uploaded_file.name}: trovata colonna 'Ora di inizio', la elaboro...")
                df = process_datetime_field(df, 'Ora di inizio')
                
                # Mappatura delle altre colonne
                if 'Denominazione dell\'attività' in columns or 'Denominazione dell\'attività' in columns:
                    denominazione_col = 'Denominazione dell\'attività' if 'Denominazione dell\'attività' in columns else 'Denominazione dell\'attività'
                    df.rename(columns={denominazione_col: 'DenominazioneAttività'}, inplace=True)
                    st.info(f"Rinominata colonna '{denominazione_col}' in 'DenominazioneAttività'")
                    
                # Gestione colonna Nome con opzioni alternative
                if 'Nome (del corsista)' in columns:
                    df.rename(columns={'Nome (del corsista)': 'Nome'}, inplace=True)
                    st.info(f"Rinominata colonna 'Nome (del corsista)' in 'Nome'")
                elif 'nome2' in columns:
                    df.rename(columns={'nome2': 'Nome'}, inplace=True)
                    st.info(f"Rinominata colonna 'nome2' in 'Nome'")
                    
                if 'Cognome (del corsista)' in columns:
                    df.rename(columns={'Cognome (del corsista)': 'Cognome'}, inplace=True)
                    st.info(f"Rinominata colonna 'Cognome (del corsista)' in 'Cognome'")
                    
                if 'Tipo di percorso' in columns:
                    df.rename(columns={'Tipo di percorso': 'DenominazionePercorso'}, inplace=True)
                    st.info(f"Rinominata colonna 'Tipo di percorso' in 'DenominazionePercorso'")
                    
                if 'Posta elettronica' in columns:
                    df.rename(columns={'Posta elettronica': 'Email'}, inplace=True)
                    st.info(f"Rinominata colonna 'Posta elettronica' in 'Email'")
                    
                if 'ID' in columns:
                    df['CodiceFiscale'] = df['ID']  # Usiamo ID come sostituto del codice fiscale se non c'è altro
                    st.warning(f"Usata colonna 'ID' come sostituto per 'CodiceFiscale'")
            else:
                # Formato non riconosciuto
                st.warning(f"File {uploaded_file.name}: formato non riconosciuto, non elaborato")
                failed_files += 1
                continue
                
            # Verifica requisiti minimi
            required_cols = ['DataPresenza', 'OraPresenza']
            missing_cols = [col for col in required_cols if col not in df.columns]
            
            if missing_cols:
                st.error(f"File {uploaded_file.name}: colonne mancanti: {', '.join(missing_cols)}")
                failed_files += 1
                continue
                
            # Verifica presenza della colonna percorso (che adesso potrebbe venire da "Tipo di percorso")
            if not ('DenominazionePercorso' in df.columns or 'percoro' in df.columns):
                st.warning(f"File {uploaded_file.name}: manca una colonna per il percorso. Questo potrebbe causare problemi nell'analisi.")
                # Non blocchiamo l'elaborazione, ma avvisiamo l'utente
                
            # Aggiungi il dataframe alla lista
            all_dataframes.append(df)
            processed_files += 1
            
        except Exception as e:
            st.error(f"Errore durante il caricamento del file {uploaded_file.name}: {e}")
            failed_files += 1
            
    # Verifica se ci sono file processati con successo
    if not all_dataframes:
        st.error("Nessun file è stato processato con successo")
        return None
        
    # Combina tutti i dataframe
    try:
        combined_df = pd.concat(all_dataframes, ignore_index=True)
        st.success(f"Caricati con successo {processed_files} file, combinati {len(all_dataframes)} dataframe con un totale di {len(combined_df)} righe")
        
        if failed_files > 0:
            st.warning(f"{failed_files} file non sono stati processati a causa di errori")
            
        # Assicuriamoci che tutte le colonne richieste siano presenti
        if 'CodiceFiscale' not in combined_df.columns:
            st.error("La colonna CodiceFiscale non è presente nei dati combinati.")
            # Se non c'è il CodiceFiscale, proviamo a usare l'ID come fallback
            if 'ID' in combined_df.columns:
                combined_df['CodiceFiscale'] = combined_df['ID']
                st.warning("Usato ID come sostituto per CodiceFiscale")
                
        # Aggiungiamo colonne vuote per quelle non presenti
        for col in ['Nome', 'Cognome', 'Email']:
            if col not in combined_df.columns:
                combined_df[col] = ''
                st.warning(f"Aggiunta colonna '{col}' vuota perché non presente nei dati")
                
        # Verifica la presenza della colonna TimestampPresenza
        if 'TimestampPresenza' not in combined_df.columns:                # Creiamo il timestamp combinando data e ora
            try:
                # Normalizzazione dei tipi di dati data e ora prima di combinarli
                # Converto esplicitamente DataPresenza in oggetto date
                try:
                    combined_df['DataPresenza'] = pd.to_datetime(combined_df['DataPresenza'], errors='coerce').dt.date
                except Exception as e:
                    st.warning(f"Problema di conversione 'DataPresenza': {e}")
                
                # Parsing migliorato della colonna OraPresenza
                try:
                    def parse_time(t):
                        # Se è già un oggetto time, lo restituiamo direttamente
                        if isinstance(t, time): return t
                        
                        # Se è un oggetto datetime o timestamp, estraiamo l'ora
                        if isinstance(t, datetime): return t.time()
                        if isinstance(t, pd.Timestamp): return t.time()
                        
                        # Gestione dei valori mancanti
                        if pd.isna(t): return pd.NaT
                        
                        # Gestione dei formati numerici (Excel salva spesso le ore come decimali)
                        if isinstance(t, (int, float)) and 0 <= t < 1: 
                            total_seconds = int(t * 24 * 60 * 60)
                            hours, remainder = divmod(total_seconds, 3600)
                            minutes, seconds = divmod(remainder, 60)
                            return time(hours, minutes, seconds)
                        
                        # Normalizza la stringa prima della conversione
                        str_t = str(t).strip()
                        
                        # Se la stringa contiene già i due punti, è probabilmente già in formato orario
                        if ':' in str_t:
                            # Prova i formati orari comuni
                            for fmt in ['%H:%M:%S', '%H:%M', '%I:%M:%S %p', '%I:%M %p']:
                                try:
                                    return pd.to_datetime(str_t, format=fmt, errors='raise').time()
                                except ValueError:
                                    continue
                        
                        # Se tutto fallisce, tenta una conversione generica
                        try:
                            return pd.to_datetime(str_t, errors='raise').time()
                        except (ValueError, TypeError):
                            return pd.NaT
                    
                    # Applica la funzione di parsing migliorata
                    combined_df['OraPresenza'] = combined_df['OraPresenza'].apply(parse_time)
                except Exception as e:
                    st.warning(f"Problema di conversione 'OraPresenza': {e}")
                
                # Funzione per combinare data e ora
                def combine_dt(row):
                    if pd.notna(row['DataPresenza']) and isinstance(row['DataPresenza'], date) and pd.notna(row['OraPresenza']) and isinstance(row['OraPresenza'], time):
                        try: 
                            return pd.Timestamp.combine(row['DataPresenza'], row['OraPresenza'])
                        except Exception: 
                            return pd.NaT
                    return pd.NaT
                
                # Applica la combinazione
                combined_df['TimestampPresenza'] = combined_df.apply(combine_dt, axis=1)
                
                # Verifica la percentuale di timestamp creati con successo
                timestamp_validi = combined_df['TimestampPresenza'].notna().sum()
                if timestamp_validi < len(combined_df):
                    st.warning(f"Attenzione: creati {timestamp_validi}/{len(combined_df)} timestamp validi ({timestamp_validi/len(combined_df)*100:.1f}%)")
                else:
                    st.success("Campo TimestampPresenza creato con successo per tutti i record")
                
                # Rimuovo record senza timestamp valido
                initial_rows = len(combined_df)
                combined_df = combined_df.dropna(subset=['TimestampPresenza', 'CodiceFiscale'])
                removed_rows = initial_rows - len(combined_df)
                
                if removed_rows > 0:
                    st.warning(f"Rimossi {removed_rows} record con CF, Data, Ora o Timestamp mancanti/non validi.")
                    
                # Standardizzo i campi data e ora a partire dal TimestampPresenza
                combined_df['DataPresenza'] = combined_df['TimestampPresenza'].dt.date
                combined_df['OraPresenza'] = combined_df['TimestampPresenza'].dt.time
                
                st.info("Formati dati standardizzati: DataPresenza (date) e OraPresenza (time) estratti da TimestampPresenza")
                
            except Exception as e:
                st.error(f"Impossibile creare il campo TimestampPresenza: {e}")
                st.exception(e)
        
        # Ora processiamo il dataframe combinato con la logica standard di integrazione dati
        
        # Carica i dati dei CFU
        st.info("Caricamento dati CFU per integrazione con i file caricati...")
        cfu_data = load_cfu_data()
        
        if not cfu_data.empty:
            st.success(f"Dati CFU caricati con successo: {len(cfu_data)} attività trovate.")
            
            # Abbinamento CFU se c'è la denominazione dell'attività
            if 'DenominazioneAttività' in combined_df.columns:
                st.info("Abbinamento dei CFU alle attività in corso...")
                activity_col_norm_internal = 'DenominazioneAttivitaNormalizzataInternal'
                combined_df[activity_col_norm_internal] = combined_df['DenominazioneAttività'].apply(normalize_generic)
                combined_df['CFU'] = combined_df['DenominazioneAttività'].apply(lambda x: match_activity_with_cfu(x, cfu_data))
                
                # Conta quante attività non hanno trovato un match per i CFU
                missing_cfu = combined_df['CFU'].isna().sum()
                total_activities = len(combined_df)
                if missing_cfu > 0:
                    st.warning(f"Non è stato possibile trovare i CFU per {missing_cfu} attività su {total_activities} ({(missing_cfu/total_activities)*100:.1f}%).")
                else:
                    st.success("Abbinamento CFU completato con successo per tutte le attività.")
            else:
                st.warning("I file caricati non contengono la colonna 'DenominazioneAttività', impossibile abbinare i CFU.")
        else:
            st.warning("Non è stato possibile caricare i dati dei CFU. I CFU non saranno disponibili.")
        
        # Carica i dati degli iscritti
        st.info("Caricamento dati iscritti per integrazione con i file caricati...")
        enrolled_students = load_enrolled_students_data()
        
        if not enrolled_students.empty:
            st.success(f"Dati iscritti caricati con successo: {len(enrolled_students)} iscritti trovati.")
            
            # Integra i dati degli iscritti
            with st.spinner("Integrazione dati degli iscritti in corso..."):
                st.info(f"Integrazione di {enrolled_students.shape[0]} record di iscritti nei dati presenze...")
                
                # Preparo i dati per il matching
                # Normalizzazione di Nome e Cognome per migliorare il matching
                if 'Nome' in combined_df.columns:
                    combined_df['Nome'] = combined_df['Nome'].astype(str).str.strip().str.lower().str.title()
                if 'Cognome' in combined_df.columns:
                    combined_df['Cognome'] = combined_df['Cognome'].astype(str).str.strip().str.lower().str.title()
                
                if 'Nome' in enrolled_students.columns:
                    enrolled_students['Nome'] = enrolled_students['Nome'].astype(str).str.strip().str.lower().str.title()
                if 'Cognome' in enrolled_students.columns:
                    enrolled_students['Cognome'] = enrolled_students['Cognome'].astype(str).str.strip().str.lower().str.title()
                
                # Eseguo l'integrazione
                combined_df = match_students_data(combined_df, enrolled_students)
                
                # Verifica se l'integrazione è avvenuta correttamente
                enrolled_cols = ['Percorso', 'Codice_Classe_di_concorso', 'Codice_classe_di_concorso_e_denominazione', 'Dipartimento', 'Matricola']
                integrated_cols = [col for col in enrolled_cols if col in combined_df.columns]
                if integrated_cols:
                    for col in integrated_cols:
                        non_null_count = combined_df[col].notna().sum()
                        st.info(f"Integrazione colonna '{col}': {non_null_count} record non vuoti su {len(combined_df)} ({non_null_count/len(combined_df)*100:.1f}%)")
                    
                    # Se l'integrazione è a zero o molto bassa, segnala un problema
                    if all(combined_df[col].notna().sum() == 0 for col in integrated_cols):
                        st.error("ERRORE: Nessuna integrazione dei dati iscritti è avvenuta!")
                        
                        # Analisi delle possibili cause
                        if 'CodiceFiscale' in combined_df.columns and 'CodiceFiscale' in enrolled_students.columns:
                            df_cf = set(combined_df['CodiceFiscale'].astype(str).str.strip().unique())
                            enrolled_cf = set(enrolled_students['CodiceFiscale'].astype(str).str.strip().unique())
                            common_cf = df_cf.intersection(enrolled_cf)
                            
                            st.error(f"Codici fiscali comuni tra i dataset: {len(common_cf)} su {len(df_cf)} nei dati presenze e {len(enrolled_cf)} negli iscritti")
                            
                            if len(common_cf) > 0:
                                st.warning("Analisi di alcune corrispondenze (primi 3 CF in comune):")
                                for cf in list(common_cf)[:3]:
                                    st.write(f"CF: {cf}")
                                    st.write("Record presenza:", combined_df[combined_df['CodiceFiscale'] == cf].iloc[0])
                                    st.write("Record iscritto:", enrolled_students[enrolled_students['CodiceFiscale'] == cf].iloc[0])
                else:
                    st.error("ATTENZIONE: Nessuna colonna degli iscritti è stata integrata nei dati!")
        else:
            st.warning("Non è stato possibile caricare i dati degli iscritti. Le informazioni aggiuntive degli studenti non saranno disponibili.")
            
        st.success("Elaborazione del caricamento multiplo completata.")
        return combined_df
        
    except Exception as e:
        st.error(f"Errore durante la combinazione dei dataframe: {e}")
        st.exception(e)  # Mostra lo stack trace completo per facilitare il debug
        return None
