# Funzioni per il rilevamento e la gestione dei duplicati
import pandas as pd
import streamlit as st
from datetime import timedelta

def detect_duplicate_records(df, timestamp_col='TimestampPresenza', cf_column='CodiceFiscale', time_delta_minutes=10):
    """Rileva record duplicati nei dati (timbrature ravvicinate dello stesso CF)"""
    if df is None or len(df) == 0: 
        return pd.DataFrame(), [], []
        
    required_cols = [timestamp_col, cf_column]
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
        st.info("Nessun record con Timestamp e CF validi per controllo duplicati.")
        return pd.DataFrame(), [], []
        
    df_copy['OriginalIndex'] = df_copy.index
    df_sorted = df_copy.sort_values(by=[cf_column, timestamp_col])
    
    df_sorted['time_diff'] = df_sorted.groupby(cf_column)[timestamp_col].diff()
    time_threshold = timedelta(minutes=time_delta_minutes)
    df_sorted['time_diff_next'] = df_sorted.groupby(cf_column)[timestamp_col].diff(-1).abs()
    
    df_sorted['is_close_to_prev'] = (df_sorted['time_diff'].notna()) & (df_sorted['time_diff'] <= time_threshold)
    df_sorted['is_close_to_next'] = (df_sorted['time_diff_next'].notna()) & (df_sorted['time_diff_next'] <= time_threshold)
    df_sorted['in_cluster'] = df_sorted['is_close_to_prev'] | df_sorted['is_close_to_next']
    
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
                if df_sorted.loc[prev_idx_sorted, cf_column] == df_sorted.loc[current_idx_sorted, cf_column] and df_sorted.loc[prev_idx_sorted, 'in_cluster']: 
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
