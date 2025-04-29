# Funzioni per il calcolo delle presenze e delle frequenze
import pandas as pd
import streamlit as st

def calculate_attendance(df, cf_column='CodiceFiscale', percorso_chiave_col='PercorsoOriginaleSenzaArt13Internal', 
                         percorso_elab_col='PercorsoInternal', original_col='PercorsoOriginaleInternal', group_by="studente"):
    """
    Calcola le presenze aggregate in base al criterio specificato.
    
    Args:
        df: DataFrame con i dati delle presenze
        cf_column: Nome della colonna per il codice fiscale
        percorso_chiave_col: Nome della colonna per il percorso chiave (senza Art.13)
        percorso_elab_col: Nome della colonna per il percorso elaborato
        original_col: Nome della colonna per il percorso originale
        group_by: Criterio di raggruppamento ("studente", "percorso_originale", "percorso_elaborato", "lista_studenti")
    
    Returns:
        DataFrame con i dati aggregati in base al criterio specificato
    """
    if df is None or len(df) == 0: 
        return pd.DataFrame()
        
    required_cols = [cf_column]
    if group_by in ["studente", "lista_studenti"]:
        required_cols.append(percorso_chiave_col)
        
    optional_cols = [percorso_elab_col, original_col] if group_by != "percorso_elaborato" else [original_col]
    name_cols = [col for col in ['Nome', 'Cognome', 'Email'] if col in df.columns]
    
    if not all(col in df.columns for col in required_cols):
        missing = [col for col in required_cols if col not in df.columns]
        st.error(f"Impossibile procedere: colonne mancanti ({', '.join(missing)})")
        return pd.DataFrame()
    
    # Definisci i gruppi in base al criterio selezionato
    if group_by == "studente":
        # Raggruppa per studente e percorso
        group_cols = [cf_column, percorso_chiave_col]
    elif group_by == "percorso_originale":
        # Raggruppa solo per percorso originale
        group_cols = [percorso_chiave_col]
        if percorso_chiave_col not in df.columns:
            st.error(f"Impossibile procedere: colonna {percorso_chiave_col} mancante")
            return pd.DataFrame()
    elif group_by == "percorso_elaborato":
        # Raggruppa solo per percorso elaborato
        group_cols = [percorso_elab_col]
        if percorso_elab_col not in df.columns:
            st.error(f"Impossibile procedere: colonna {percorso_elab_col} mancante")
            return pd.DataFrame()
    elif group_by == "lista_studenti":
        # Lista degli studenti senza raggruppare per percorso
        group_cols = [cf_column]
    else:
        st.error(f"Criterio di raggruppamento non valido: {group_by}")
        return pd.DataFrame()
    
    # Separa i CFU dalle altre colonne per poterli sommare
    cfu_column = 'CFU'
    has_cfu = cfu_column in df.columns
    
    # Ottieni le altre colonne che saranno aggregate con first()
    if group_by == "studente" or group_by == "lista_studenti":
        # Per raggruppamento per studente, includi nome e cognome
        first_cols = name_cols + [col for col in optional_cols if col in df.columns and col not in group_cols]
    else:
        # Per raggruppamenti per percorso, non includere nome e cognome
        first_cols = [col for col in optional_cols if col in df.columns and col not in group_cols]
    
    # Calcola il conteggio delle presenze
    attendance_counts = df.groupby(group_cols, dropna=False).size().reset_index(name='Presenze')
    
    # Crea il dataframe di base con le informazioni generali
    if first_cols and (group_by == "studente" or group_by == "lista_studenti"):
        first_info_values = df.dropna(subset=first_cols, how='all').groupby(group_cols, as_index=False)[first_cols].first()
        attendance = pd.merge(attendance_counts, first_info_values, on=group_cols, how='left')
    else: 
        attendance = attendance_counts
    
    # Se abbiamo i CFU, calcoliamo la somma per ogni gruppo e li aggiungiamo al dataframe
    if has_cfu:
        # Converti i valori di CFU a numeri per assicurare che possano essere sommati
        df_cfu = df.copy()
        df_cfu[cfu_column] = pd.to_numeric(df_cfu[cfu_column], errors='coerce').fillna(0)
        
        # Calcola la somma dei CFU per ciascun gruppo
        cfu_sums = df_cfu.groupby(group_cols, dropna=False)[cfu_column].sum().reset_index()
        cfu_sums = cfu_sums.rename(columns={cfu_column: 'CFU Totali'})
        
        # Unisci la somma dei CFU con il dataframe principale
        attendance = pd.merge(attendance, cfu_sums, on=group_cols, how='left')
        
        # Rimuovi la vecchia colonna CFU se presente (che contiene solo il primo valore)
        if 'CFU' in attendance.columns:
            attendance = attendance.drop(columns=['CFU'])
    
    # Rinomina le colonne in base al raggruppamento
    if group_by == "studente" or group_by == "lista_studenti":
        rename_map = {
            percorso_chiave_col: 'Percorso (Senza Art.13)', 
            percorso_elab_col: 'Percorso Elaborato (Info)', 
            original_col: 'Percorso Originale Input (Info)'
        }
    elif group_by == "percorso_originale":
        rename_map = {
            percorso_chiave_col: 'Percorso (Senza Art.13)'
        }
    elif group_by == "percorso_elaborato":
        rename_map = {
            percorso_elab_col: 'Percorso Elaborato'
        }
    
    attendance = attendance.rename(columns={k: v for k, v in rename_map.items() if k in attendance.columns})
    
    # Ordina le colonne in base al raggruppamento
    if group_by == "studente":
        cols_order_final = [cf_column] + name_cols + [rename_map.get(percorso_chiave_col, percorso_chiave_col)] + [rename_map.get(col, col) for col in [percorso_elab_col, original_col] if col in attendance.columns]
    elif group_by == "lista_studenti":
        cols_order_final = [cf_column] + name_cols
    elif group_by == "percorso_originale":
        cols_order_final = [rename_map.get(percorso_chiave_col, percorso_chiave_col)]
    elif group_by == "percorso_elaborato":
        cols_order_final = [rename_map.get(percorso_elab_col, percorso_elab_col)]
    
    # Aggiungi CFU Totali prima di Presenze nella visualizzazione
    if 'CFU Totali' in attendance.columns:
        cols_order_final.append('CFU Totali')
    
    cols_order_final.append('Presenze')
    
    # Seleziona solo le colonne esistenti e in ordine
    final_cols = [c for c in cols_order_final if c in attendance.columns]
    attendance = attendance[final_cols]
    
    # Riempie i valori nulli in nome e cognome
    if 'Nome' in attendance.columns: attendance['Nome'] = attendance['Nome'].fillna('')
    if 'Cognome' in attendance.columns: attendance['Cognome'] = attendance['Cognome'].fillna('')
    if 'Email' in attendance.columns: attendance['Email'] = attendance['Email'].fillna('')
    
    return attendance

def calculate_lesson_attendance(df, date_filter=None, activity_filter=None, cf_column='CodiceFiscale', 
                               date_col='DataPresenza', activity_col='DenominazioneAttivitaNormalizzataInternal'):
    """
    Calcola il numero di partecipanti unici per ogni combinazione di data e attività.
    
    Args:
        df: DataFrame con i dati delle presenze
        date_filter: Data specifica da filtrare (opzionale)
        activity_filter: Attività specifica da filtrare (opzionale)
        cf_column: Nome della colonna contenente i codici fiscali
        date_col: Nome della colonna contenente le date
        activity_col: Nome della colonna contenente le attività normalizzate
    
    Returns:
        DataFrame con il conteggio dei partecipanti per combinazione data-attività
    """
    if df is None or len(df) == 0:
        return pd.DataFrame()
    
    required_cols = [cf_column, date_col, activity_col]
    if not all(col in df.columns for col in required_cols):
        missing = [col for col in required_cols if col not in df.columns]
        st.error(f"Impossibile procedere: colonne mancanti ({', '.join(missing)})")
        return pd.DataFrame()
    
    # Creazione di una copia del DataFrame per il filtraggio
    filtered_df = df.copy()
    
    # Applicazione dei filtri se specificati
    if date_filter is not None and date_filter != "Tutte le date":
        filtered_df = filtered_df[filtered_df[date_col] == date_filter]
    
    if activity_filter is not None and activity_filter != "Tutte le attività":
        filtered_df = filtered_df[filtered_df[activity_col] == activity_filter]
    
    if filtered_df.empty:
        return pd.DataFrame()
    
    # Raggruppamento per data e attività, conteggio dei CF unici
    attendance_counts = (filtered_df.groupby([date_col, activity_col], dropna=False)
                          .agg({cf_column: 'nunique'})
                          .reset_index()
                          .rename(columns={cf_column: 'Partecipanti'}))
    
    # Ordinamento per data e poi per attività
    if not attendance_counts.empty:
        attendance_counts = attendance_counts.sort_values(by=[date_col, activity_col])
    
    return attendance_counts
