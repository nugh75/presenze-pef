# File principale dell'applicazione Gestione Presenze
import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

# Importazione dei moduli
from modules.data_loader import load_data, load_multiple_files
# Importazione diretta dai moduli tab invece che dal pacchetto ui
from modules.ui.tab1 import render_tab1
from modules.ui.tab2 import render_tab2  # Versione corretta che gestisce le colonne duplicate
from modules.ui.tab3 import render_tab3
from modules.ui.tab4 import render_tab4

# Configurazione Pagina
st.set_page_config(
    page_title="Gestione Presenze",
    page_icon="📊",
    layout="wide",
)

# Layout Principale
st.title("📊 Gestione Presenze Corsi")
st.markdown("<a id='top'></a>", unsafe_allow_html=True)

# --- Sidebar --- 
with st.sidebar:
    st.header("Caricamento File")
    upload_method = st.radio("Modalità di caricamento:", ["File singolo", "Più file contemporaneamente"], index=0)
    
    if upload_method == "File singolo":
        uploaded_file = st.file_uploader("Carica file Excel presenze", type=['xlsx'])
        if uploaded_file:
            st.success(f"File '{uploaded_file.name}' caricato!")
            if st.button("Mostra anteprima originale"):
                try: 
                    preview = pd.read_excel(uploaded_file, nrows=10)
                    st.write("Colonne:", preview.columns.tolist())
                    st.dataframe(preview)
                except Exception as e: 
                    st.error(f"Errore anteprima: {e}")
            uploaded_files = [uploaded_file]  # Per uniformità con la modalità multi-file
    else:
        uploaded_files = st.file_uploader("Carica più file contemporaneamente", 
                                         type=['xlsx', 'csv', 'txt'], 
                                         accept_multiple_files=True)
        if uploaded_files:
            file_names = [file.name for file in uploaded_files]
            st.success(f"Caricati {len(uploaded_files)} file: {', '.join(file_names)}")
            if st.button("Mostra anteprima primo file"):
                try:
                    first_file = uploaded_files[0]
                    file_ext = first_file.name.split('.')[-1].lower()
                    if file_ext == 'xlsx':
                        preview = pd.read_excel(first_file, nrows=10)
                    else:
                        preview = pd.read_csv(first_file, nrows=10)
                    st.write(f"Colonne del file {first_file.name}:", preview.columns.tolist())
                    st.dataframe(preview)
                except Exception as e:
                    st.error(f"Errore anteprima: {e}")
        else:
            uploaded_files = None
    
    st.divider()
    if 'processed_df' in st.session_state and st.session_state.processed_df is not None:
        st.markdown("[⬆️ Torna su](#top)", help="Clicca per tornare all'inizio della pagina principale")

# --- Gestione Stato Sessione ---
if 'duplicates_removed' not in st.session_state: 
    st.session_state.duplicates_removed = False
if 'duplicate_detection_results' not in st.session_state: 
    st.session_state.duplicate_detection_results = (pd.DataFrame(), [], [])
if 'selected_indices_to_drop' not in st.session_state: 
    st.session_state.selected_indices_to_drop = []
if 'report_data_to_download' not in st.session_state: 
    st.session_state.report_data_to_download = None
if 'report_filename_to_download' not in st.session_state: 
    st.session_state.report_filename_to_download = None

# --- Logica caricamento e reset stato ---
if 'uploaded_files' in locals() and uploaded_files:
    # Ottieni una lista di nomi per il controllo di caricamento
    if upload_method == "File singolo":
        current_files_name = uploaded_file.name
        need_reload = ('current_file_name' not in st.session_state or 
                       st.session_state.current_file_name != current_files_name or 
                       'processed_df' not in st.session_state)
    else:
        current_files_name = ",".join(sorted([f.name for f in uploaded_files]))
        need_reload = ('current_file_name' not in st.session_state or 
                       st.session_state.current_file_name != current_files_name or 
                       'processed_df' not in st.session_state)
    
    if need_reload:
        with st.spinner("Caricamento ed elaborazione dati..."): 
            if upload_method == "File singolo":
                st.session_state.processed_df = load_data(uploaded_file)
            else:
                st.session_state.processed_df = load_multiple_files(uploaded_files)
                
        if st.session_state.processed_df is not None:
            st.session_state.current_file_name = current_files_name
            st.session_state.duplicates_removed = False
            st.session_state.duplicate_detection_results = (pd.DataFrame(), [], [])
            st.session_state.selected_indices_to_drop = []
            st.session_state.report_data_to_download = None
            st.session_state.report_filename_to_download = None
            
            num_files = len(uploaded_files) if upload_method == "Più file contemporaneamente" else 1
            if num_files > 1:
                st.success(f"Dati caricati da {num_files} file ed elaborati con successo.")
            else:
                st.success("Dati caricati ed elaborati con successo.")
            st.rerun()
        else:
            if 'current_file_name' in st.session_state: 
                del st.session_state.current_file_name
            if 'processed_df' in st.session_state: 
                del st.session_state.processed_df
            st.session_state.duplicates_removed = False
            st.session_state.duplicate_detection_results = (pd.DataFrame(), [], [])
            st.session_state.selected_indices_to_drop = []
            st.session_state.report_data_to_download = None
            st.session_state.report_filename_to_download = None
elif 'processed_df' in st.session_state:
    del st.session_state.processed_df
    if 'current_file_name' in st.session_state: 
        del st.session_state.current_file_name
    st.session_state.duplicates_removed = False
    st.session_state.duplicate_detection_results = (pd.DataFrame(), [], [])
    st.session_state.selected_indices_to_drop = []
    st.session_state.report_data_to_download = None
    st.session_state.report_filename_to_download = None
    st.rerun()

df_main = st.session_state.get('processed_df', None)

# --- Tabs ---
if df_main is not None and isinstance(df_main, pd.DataFrame):
    tab1, tab2, tab3, tab4 = st.tabs(["Analisi Dati", "Gestione Duplicati", "Calcolo Presenze ed Esportazione", "Frequenza Lezioni"])
    
    with tab1:
        render_tab1(df_main)
    
    with tab2:
        render_tab2(df_main)
    
    with tab3:
        render_tab3(df_main)
        
    with tab4:
        render_tab4(df_main)
        
else:
    # Messaggio iniziale se nessun file caricato
    st.info("👈 Per iniziare, carica un file Excel dalla barra laterale.")
    st.markdown("""
    ## Istruzioni d'uso
    1.  **Carica file** dalla sidebar:
        * Scegli tra la modalità **File singolo** (.xlsx) o **Più file contemporaneamente** (.xlsx, .csv, .txt)
        * L'app supporta ora anche il formato con colonna "Ora di inizio" contenente data e ora
    2.  **(Opzionale)** Vedi anteprima del file originale.
    3.  **Analisi Dati (Tab 1):** Controlla statistiche e dati elaborati.
    4.  **Gestione Duplicati (Tab 2):** Identifica e rimuovi timbrature ravvicinate.
    5.  **Calcolo Presenze ed Esportazione (Tab 3):**
        *   **Seleziona il tipo di visualizzazione**:
            - **Presenze per Studente e Percorso**: Vista tradizionale con filtri per percorso e studente
            - **Riassunto per Percorso Originale**: Visualizza i totali parziali per ogni tipo di percorso originale
            - **Riassunto per Percorso Elaborato**: Visualizza i totali parziali per ogni tipo di percorso elaborato
            - **Lista Completa Studenti**: Elenco di tutti gli studenti con possibilità di ricerca e filtro
        *   **Esporta in Excel e CSV:** 
            - Seleziona e ordina le colonne desiderate (inclusa l'email)
            - Scegli un periodo di date per l'esportazione (opzionale)
            - Genera file **dettagliato**, un foglio per codice percorso **estratto e mostrato all'inizio [Codice]**
            - I percorsi mostrano ora il codice all'inizio nel formato [A-30] Nome Percorso
    6.  **Frequenza Lezioni (Tab 4):**
        *   Visualizza il numero di partecipanti unici per ogni combinazione di **Data** e **Attività**.
        *   Filtra per **Data** e/o **Attività** in modo indipendente.
        *   Consulta le **Statistiche** di frequenza (numero lezioni, media partecipanti, totale presenze).
        *   **Esporta in CSV** i dati visualizzati.

    ### Formati File Supportati
    #### Formato Standard
    *   `CodiceFiscale`, `DataPresenza`, `OraPresenza`, `DenominazionePercorso` (o `percoro`) - Obbligatori
    *   `Nome`, `Cognome`, `recapito_ateneo` (per l'email), `DenominazioneAttività`, `CodicePercorso` - Opzionali
    
    #### Nuovo Formato Supportato
    *   `ID`, `Ora di inizio` (contiene data e ora), `Denominazione dell'attività` - Obbligatori
    *   `Nome (del corsista)`, `Cognome (del corsista)`, `Tipo di percorso`, `Posta elettronica` - Importanti
    """)

# Footer
st.markdown("---")
st.markdown("### Gestione Presenze - Versione beta 1.6") # Versione aggiornata - supporto multiformato e caricamento multiplo
