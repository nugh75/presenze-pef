# Interfaccia utente per la Tab 1 (Analisi Dati)
import streamlit as st
import pandas as pd

def render_tab1(df_main):
    """Renderizza l'interfaccia della Tab 1: Analisi Dati"""
    st.header("Analisi Dati Caricati")
    if st.session_state.duplicates_removed: 
        st.success("I record duplicati selezionati sono stati rimossi.")
    
    st.subheader("Statistiche Generali")
    metrics_data = {"Record Validi": len(df_main)}
    
    if 'CodiceFiscale' in df_main.columns: 
        metrics_data["Persone Uniche (CF)"] = df_main['CodiceFiscale'].nunique()
        
    # Le metriche relative a PercorsoInternal sono state rimosse
        
    activity_col_norm_internal = 'DenominazioneAttivitaNormalizzataInternal'
    if activity_col_norm_internal in df_main.columns: 
        metrics_data["Attività Uniche (Norm.)"] = df_main[activity_col_norm_internal].nunique()
    
    num_metrics = len(metrics_data)
    cols_metrics = st.columns(num_metrics)
    
    for i, (label, value) in enumerate(metrics_data.items()):
        with cols_metrics[i]: 
            st.metric(label, value)
    
    st.subheader("Dati Attuali Utilizzati per l'Analisi")
    cols_show_preferred = ['CodiceFiscale', 'Nome', 'Cognome', 'Email', 'DataPresenza', 'OraPresenza',
                          'DenominazioneAttività', 'DenominazioneAttivitaNormalizzataInternal',
                          'Percorso', 'Codice_Classe_di_concorso', 'Codice_classe_di_concorso_e_denominazione',
                          'Dipartimento', 'LogonName', 'Matricola',
                          'CodicePercorso', 'CFU', 'TimestampPresenza']
                          
    cols_show_exist = [col for col in cols_show_preferred if col in df_main.columns]
    st.dataframe(df_main[cols_show_exist], use_container_width=True)
    
    st.caption("CFU: Crediti Formativi Universitari associati all'attività. " +
               "Percorso, Codice_Classe_di_concorso, ecc.: Dati integrati dal file degli studenti iscritti.")
