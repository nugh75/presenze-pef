# filepath: /mnt/git/presenze-pef/modules/duplicates.py
# Funzioni per il rilevamento e la gestione dei duplicati
import pandas as pd
import streamlit as st
from datetime import timedelta, time

def detect_duplicate_records(df, timestamp_col='TimestampPresenza', time_delta_minutes=120):
    """
    Rileva record duplicati nei dati in base a:
    - Cognome (insensibile a maiuscole/minuscole)
    - Nome (insensibile a maiuscole/minuscole) 
    - DenominazioneAttività (insensibile a maiuscole/minuscole)
    - DataPresenza (stesso giorno)
    - OraPresenza (stessa ora esatta o entro un intervallo di due ore)
    
    Due record sono considerati duplicati se:
    1. Hanno identici valori di Nome, Cognome, DenominazioneAttività, DataPresenza e OraPresenza
       (record completamente identici)
       OPPURE
    2. Hanno identici valori di Nome, Cognome, DenominazioneAttività, DataPresenza 
       e registrazioni entro un intervallo di due ore (time_delta_minutes)
    
    Note: La funzione assume che i campi DataPresenza e OraPresenza siano già stati 
    standardizzati dal modulo data_loader per garantire coerenza nei confronti.
    """
    if df is None or len(df) == 0: 
        return pd.DataFrame(), [], []
        
    required_cols = [timestamp_col, 'Nome', 'Cognome', 'DenominazioneAttività']
    if not all(col in df.columns for col in required_cols):
        missing_but_exist = [col for col in required_cols if col in df.columns and df[col].isnull().all()]
        if len(missing_but_exist) == len(required_cols): 
            st.warning(f"Colonne necessarie ({', '.join(required_cols)}) vuote.")
            return pd.DataFrame(), [], []
            
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols: 
            st.error(f"Colonne necessarie ({', '.join(missing_cols)}) mancanti.")
            return pd.DataFrame(), [], []
            
    df_copy = df.dropna(subset=required_cols).copy()
    if df_copy.empty: 
        st.info("Nessun record con Timestamp, Nome, Cognome e DenominazioneAttività validi per controllo duplicati.")
        return pd.DataFrame(), [], []
    
    # Normalizzazione delle date e orari
    # Assicurati che TimestampPresenza sia datetime
    if timestamp_col in df_copy.columns:
        df_copy[timestamp_col] = pd.to_datetime(df_copy[timestamp_col], errors='coerce')
        
        # Avvisa se ci sono valori NaT dopo la conversione
        nat_count = df_copy[timestamp_col].isna().sum()
        if nat_count > 0:
            st.warning(f"{nat_count} valori di {timestamp_col} non sono stati convertiti correttamente.")
    
    # Standardizzazione di DataPresenza e OraPresenza
    # 1. Se mancano queste colonne, derivale da TimestampPresenza
    if 'DataPresenza' not in df_copy.columns and timestamp_col in df_copy.columns:
        df_copy['DataPresenza'] = df_copy[timestamp_col].dt.date
    
    if 'OraPresenza' not in df_copy.columns and timestamp_col in df_copy.columns:
        df_copy['OraPresenza'] = df_copy[timestamp_col].dt.time
        
    # 2. Standardizza DataPresenza - assicurandosi che sia di tipo date
    if 'DataPresenza' in df_copy.columns:
        try:
            # Salva una copia del valore originale prima della standardizzazione
            df_copy['DataPresenzaOriginal'] = df_copy['DataPresenza'].astype(str)
            
            # Converti in formato standard
            if not pd.api.types.is_datetime64_dtype(df_copy['DataPresenza']):
                df_copy['DataPresenza'] = pd.to_datetime(df_copy['DataPresenza'], errors='coerce').dt.date
        except Exception as e:
            st.warning(f"Errore durante la standardizzazione di DataPresenza: {e}")
    
    # 3. Standardizza OraPresenza - assicurandosi che sia di tipo time
    if 'OraPresenza' in df_copy.columns:
        # Salva il valore originale di OraPresenza per la chiave duplicati
        df_copy['OraPresenzaOriginal'] = df_copy['OraPresenza'].astype(str)
        
        # Standardizza OraPresenza per i confronti temporali
        try:
            # Funzione helper per standardizzare i valori di ora
            def standardize_time(val):
                if pd.isna(val):
                    return None
                if isinstance(val, time):
                    return val
                # Gli altri casi sono gestiti dal data_loader, quindi qui
                # dovremmo avere già tutti i valori nel formato corretto
                return val
                
            df_copy['OraPresenza'] = df_copy['OraPresenza'].apply(standardize_time)
        except Exception as e:
            st.warning(f"Errore durante la standardizzazione di OraPresenza: {e}")
    else:
        # Se non esiste proprio OraPresenza, creiamo un campo vuoto
        df_copy['OraPresenzaOriginal'] = ''
    
    # Preparo le componenti della chiave di duplicazione
    nome_norm = df_copy['Nome'].str.lower().fillna('')
    cognome_norm = df_copy['Cognome'].str.lower().fillna('')
    attivita_norm = df_copy['DenominazioneAttività'].str.lower().fillna('')
    data_norm = df_copy['DataPresenzaOriginal'].fillna('') if 'DataPresenzaOriginal' in df_copy.columns else df_copy['DataPresenza'].astype(str).fillna('')
    ora_norm = df_copy['OraPresenzaOriginal'].fillna('')
    
    # Creo una colonna che rappresenta la combinazione delle chiavi di ricerca duplicati
    df_copy['ChiaveDuplicato'] = (
        nome_norm + '|' + 
        cognome_norm + '|' + 
        attivita_norm + '|' + 
        data_norm + '|' +
        ora_norm  # Include valore originale di OraPresenza
    )
    
    # Con questa implementazione:
    # 1. Due record con esattamente lo stesso valore in tutte le componenti (anche con formati diversi 
    #    ma uguale rappresentazione stringa) saranno considerati duplicati esatti
    # 2. Record con ore diverse ma nella stessa data/persona/attività saranno valutati come potenziali 
    #    duplicati se rientrano nell'intervallo definito da time_delta_minutes (due ore)
    
    df_copy['OriginalIndex'] = df_copy.index
    df_sorted = df_copy.sort_values(by=['ChiaveDuplicato', timestamp_col])
    
    # Aggiungo un flag per record che sono già duplicati esatti (stessa chiave duplicato completa)
    df_sorted['exact_duplicate'] = df_sorted.duplicated(subset=['ChiaveDuplicato'], keep=False)
    
    # Per record con orari diversi ma stessa chiave base, calcolo la differenza temporale
    df_sorted['time_diff'] = df_sorted.groupby('ChiaveDuplicato')[timestamp_col].diff()
    time_threshold = timedelta(minutes=time_delta_minutes)
    df_sorted['time_diff_next'] = df_sorted.groupby('ChiaveDuplicato')[timestamp_col].diff(-1).abs()
    
    df_sorted['is_close_to_prev'] = (df_sorted['time_diff'].notna()) & (df_sorted['time_diff'] <= time_threshold)
    df_sorted['is_close_to_next'] = (df_sorted['time_diff_next'].notna()) & (df_sorted['time_diff_next'] <= time_threshold)
    # Aggiungiamo exact_duplicate alla condizione in_cluster per considerare anche duplicati esatti
    df_sorted['in_cluster'] = df_sorted['is_close_to_prev'] | df_sorted['is_close_to_next'] | df_sorted['exact_duplicate']
    
    df_sorted['GruppoDuplicati'] = 0
    current_group_id = 1
    sorted_indices = df_sorted.index
    
    for i in range(len(sorted_indices)):
        current_idx_sorted = sorted_indices[i]
        if df_sorted.loc[current_idx_sorted, 'in_cluster']:
            current_group = df_sorted.loc[current_idx_sorted, 'GruppoDuplicati']
            prev_group = 0
            
            if i > 0:
                prev_idx_sorted = sorted_indices[i-1]
                if df_sorted.loc[prev_idx_sorted, 'ChiaveDuplicato'] == df_sorted.loc[current_idx_sorted, 'ChiaveDuplicato'] and df_sorted.loc[prev_idx_sorted, 'in_cluster']: 
                    prev_group = df_sorted.loc[prev_idx_sorted, 'GruppoDuplicati']
                    
            if current_group == 0:
                if prev_group != 0 and df_sorted.loc[current_idx_sorted, 'is_close_to_prev']: 
                    df_sorted.loc[current_idx_sorted, 'GruppoDuplicati'] = prev_group
                else: 
                    df_sorted.loc[current_idx_sorted, 'GruppoDuplicati'] = current_group_id
                    current_group_id += 1
            elif prev_group != 0 and current_group != prev_group and df_sorted.loc[current_idx_sorted, 'is_close_to_prev']: 
                group_to_merge = current_group
                target_group = prev_group
                df_sorted.loc[df_sorted['GruppoDuplicati'] == group_to_merge, 'GruppoDuplicati'] = target_group
                
    involved_df_sorted = df_sorted[df_sorted['GruppoDuplicati'] != 0].copy()
    if involved_df_sorted.empty: 
        return pd.DataFrame(), [], []
        
    group_mapping = pd.Series(involved_df_sorted['GruppoDuplicati'].values, index=involved_df_sorted['OriginalIndex']).to_dict()
    involved_original_indices = involved_df_sorted['OriginalIndex'].unique().tolist()
    indices_to_drop_suggestion = []
    
    grouped = involved_df_sorted.sort_values(timestamp_col).groupby('GruppoDuplicati')
    for group_id, group_df in grouped: 
        indices_to_drop_suggestion.extend(group_df['OriginalIndex'].iloc[1:].tolist())
        
    valid_involved_indices = [idx for idx in involved_original_indices if idx in df.index]
    if not valid_involved_indices: 
        return pd.DataFrame(), [], []
        
    duplicates_df = df.loc[valid_involved_indices].copy()
    duplicates_df['GruppoDuplicati'] = duplicates_df.index.map(group_mapping).fillna(0).astype(int)
    duplicates_df['SuggerisciRimuovere'] = duplicates_df.index.isin(indices_to_drop_suggestion)
    duplicates_df = duplicates_df.sort_values(by=['GruppoDuplicati', timestamp_col])
    valid_indices_to_drop = [idx for idx in indices_to_drop_suggestion if idx in df.index]
    
    return duplicates_df, valid_involved_indices, valid_indices_to_drop
