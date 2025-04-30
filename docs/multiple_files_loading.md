# Guida al Caricamento di File Multipli

Questa guida spiega come funziona il caricamento e l'integrazione dei dati quando si utilizzano più file contemporaneamente.

## Formati Supportati

Il sistema supporta il caricamento di:
- File Excel (.xlsx)
- File CSV (.csv) con vari delimitatori (virgola o punto e virgola) e codifiche

## Procedura di Caricamento Multiplo

1. **Riconoscimento del Formato**:
   - Il sistema identifica automaticamente il formato di ciascun file
   - Supporta due formati principali:
     - **Formato Standard**: Con colonne DataPresenza e OraPresenza già presenti
     - **Formato Alternativo**: Con colonna "Ora di inizio" che contiene data e ora insieme

2. **Rinominazione delle Colonne**:
   - Per uniformare i dati, il sistema rinomina le colonne secondo uno schema standard
   - Mappature principali:
     - "Denominazione dell'attività" → "DenominazioneAttività"
     - "Nome (del corsista)" → "Nome"
     - "Cognome (del corsista)" → "Cognome" 
     - "Tipo di percorso" → "DenominazionePercorso"
     - "Posta elettronica" → "Email"

3. **Verifica dei Requisiti Minimi**:
   - I file devono contenere almeno le colonne "DataPresenza" e "OraPresenza"
   - Viene segnalata l'assenza della colonna "percorso" ma non blocca il caricamento

4. **Combinazione dei DataFrame**:
   - Tutti i file validi vengono uniti in un unico DataFrame
   - Vengono aggiunte colonne mancanti con valori vuoti se necessario

5. **Creazione del TimestampPresenza**:
   - Viene generato un timestamp combinando data e ora
   - Normalizzazione dei formati di data e ora per garantire coerenza

6. **Integrazione Dati Esterni**:
   - **Dati dei CFU**: Integrazione con informazioni dal file `crediti.csv`
   - **Dati degli Iscritti**: Integrazione con informazioni dal file `modules/dati/iscritti_29_aprile.csv`

## Processo di Integrazione Dati Esterni

### Integrazione dei CFU

1. Caricamento del file `crediti.csv`
2. Normalizzazione dei nomi delle attività
3. Abbinamento tra "DenominazioneAttività" dei file di presenza e "DenominazioneAttività" del file CFU
4. Supporto per abbinamenti anche con piccole differenze (case-insensitive, spazi extra, ecc.)

### Integrazione dei Dati degli Iscritti

1. Caricamento del file `modules/dati/iscritti_29_aprile.csv`
2. Normalizzazione di nomi e cognomi
3. Tentativo di abbinamento prima per codice fiscale, poi per nome e cognome
4. Integrazione delle seguenti informazioni:
   - Percorso
   - Codice_Classe_di_concorso
   - Codice_classe_di_concorso_e_denominazione
   - Dipartimento
   - Matricola

## Risoluzione dei Problemi

Se l'integrazione dei dati non funziona correttamente, verificare:
1. **Problemi di formattazione**: Controllare che i formati di data e ora siano riconosciuti
2. **Problemi di corrispondenza**: Controllare che i codici fiscali o nomi/cognomi corrispondano tra i file
3. **File mancanti o inaccessibili**: Verificare che i file `crediti.csv` e `iscritti_29_aprile.csv` siano accessibili

## Analisi Diagnostica

Il sistema fornisce informazioni dettagliate durante il processo di caricamento e integrazione:
- Percentuali di successo nelle operazioni di abbinamento
- Statistiche su record rimossi o modificati
- Analisi di corrispondenze specifiche in caso di problemi (esempi di codici fiscali comuni)
