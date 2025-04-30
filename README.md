# Gestione Presenze - File README

## Descrizione
Questa applicazione Streamlit permette di gestire le presenze di un corso, con particolare attenzione a:
1. Identificazione e rimozione dei record duplicati (con differenza di orario ≤ 10 minuti)
2. Calcolo delle presenze per studente, suddivise per percorso formativo

## Funzionalità Principali
- **Normalizzazione dei percorsi formativi**: L'app rimuove automaticamente la dicitura "art. 13" (e varianti) dai nomi dei percorsi per considerarli equivalenti
- **Riposizionamento codici percorso**: I codici percorso tra parentesi (es: A-30) vengono estratti e riposizionati all'inizio della stringa in formato [A-30]
- **Visualizzazione dati**: È possibile visualizzare sia i nomi originali dei percorsi che quelli normalizzati con codici in evidenza
- **Filtro per percorso**: Permette di filtrare le presenze per percorso formativo specifico
- **Identificazione duplicati**: Trova e gestisce record duplicati con tempi simili (entro 10 minuti)
- **Reportistica**: Genera report scaricabili con le presenze per percorso formativo
- **Filtro per periodo**: Permette di selezionare un intervallo di date per l'esportazione
- **Integrazione dati studenti**: Carica e integra automaticamente le informazioni aggiuntive degli studenti da un file CSV esterno
- **Integrazione dati studenti**: Carica e integra automaticamente le informazioni aggiuntive degli studenti da un file CSV esterno

## Requisiti
- Python 3.6+
- Streamlit
- Pandas
- Openpyxl (per la lettura di file Excel)
- Matplotlib (per i grafici)

## Installazione
1. Crea un ambiente virtuale (facoltativo ma consigliato):
   ```
   python -m venv venv
   source venv/bin/activate  # Linux/Mac
   venv\Scripts\activate     # Windows
   ```

2. Installa le dipendenze:
   ```
   pip install -r requirements.txt
   ```

## Utilizzo
1. Avvia l'applicazione:
   ```
   streamlit run app.py
   ```
   Oppure usa lo script:
   ```
   ./run.sh
   ```

2. Carica un file Excel contenente i dati delle presenze
3. Esplora le diverse funzionalità tramite le schede:
   - Analisi Dati: visualizza statistiche generali e i dati completi
   - Gestione Duplicati: identifica e rimuovi record duplicati
   - Calcolo Presenze per Percorso: visualizza e scarica report delle presenze per percorso formativo
   - Frequenza Lezioni: visualizza i partecipanti per data e attività
   - Frequenza Lezioni: visualizza i partecipanti per data e attività

## Integrazione dati studenti
L'applicazione può integrare dati aggiuntivi sugli studenti da un file CSV esterno:
1. Il file deve essere posizionato in `modules/dati/iscritti_29_aprile.csv`
2. I dati vengono associati automaticamente tramite Codice Fiscale o Nome/Cognome
3. Le informazioni aggiuntive (percorso, classe di concorso, dipartimento, matricola, ecc.) vengono mostrate in tutte le visualizzazioni e nei report
4. Per ulteriori dettagli consultare `docs/data_integration_guide.md`

## Formato del file Excel
Il file Excel dovrebbe contenere almeno le seguenti colonne:
- `CodiceFiscale`: identificativo univoco dello studente
- `DataPresenza`: data della presenza
- `OraPresenza`: orario della presenza
- `DenominazionePercorso`: percorso formativo associato alla presenza (l'app gestisce automaticamente varianti con "art. 13")

Colonne opzionali ma utili:
- `Nome`: nome dello studente
- `Cognome`: cognome dello studente
- `recapito_ateneo`: email dello studente (opzionale)

## Integrazione dati studenti
L'applicazione può integrare dati aggiuntivi sugli studenti da un file CSV esterno:
1. Il file deve essere posizionato in `modules/dati/iscritti_29_aprile.csv`
2. I dati vengono associati automaticamente tramite Codice Fiscale o Nome/Cognome
3. Le informazioni aggiuntive (percorso, classe di concorso, dipartimento, matricola, ecc.) vengono mostrate in tutte le visualizzazioni e nei report
4. Per ulteriori dettagli consultare `docs/data_integration_guide.md`
- `recapito_ateneo`: email dello studente (opzionale)
