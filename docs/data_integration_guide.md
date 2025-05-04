# Guida all'Integrazione dei Dati

Questa guida spiega il processo di integrazione dei dati esterni nel sistema di gestione delle presenze.

## Introduzione

Il sistema di gestione delle presenze può integrare due tipi di dati esterni:

1. **Dati degli Iscritti**: Informazioni sugli studenti iscritti, comprese classi di concorso, percorsi, e dati anagrafici
2. **Dati dei CFU**: Informazioni sui crediti formativi associati a ciascuna attività didattica

## File dei Dati Esterni

### File degli Iscritti: `modules/dati/iscritti_30_aprile.csv`

Questo file contiene informazioni sugli iscritti, in formato CSV con separatore ';', con le seguenti colonne principali:
- `Cognome`, `Nome`: Dati anagrafici dello studente
- `CodiceFiscale`: Identificativo univoco dello studente 
- `Codice_Classe_di_concorso`: Codice della classe di concorso
- `Percorso`: Informazione sul percorso formativo (es. PeF30, PeF60)
- `Email`, `LogonName`, `Matricola`: Altre informazioni identificative

### File dei CFU: `crediti.csv`

Questo file contiene informazioni sui crediti formativi di ciascuna attività didattica, in formato CSV con separatore ',', con le seguenti colonne:
- `DenominazioneAttività`: Nome dell'attività didattica
- `CFU`: Numero di crediti formativi assegnati all'attività

## Processo di Integrazione

Durante il caricamento dei dati delle presenze, il sistema tenta automaticamente di:

1. **Integrare i dati degli iscritti**:
   - L'abbinamento avviene principalmente tramite il codice fiscale (metodo più affidabile)
   - In mancanza di corrispondenza per CF, tenta l'abbinamento con nome e cognome
   - Le colonne integrate includono: Percorso, Codice_Classe_di_concorso, Dipartimento, Matricola

2. **Integrare i dati dei CFU**:
   - L'abbinamento viene fatto tra la denominazione dell'attività nel file presenze e nel file CFU
   - Il sistema accetta piccole differenze (case-insensitive e tolleranza di spazi)
   - Viene usato un algoritmo di "fuzzy matching" con soglia di similarità al 90%

## Risoluzione dei Problemi

Se l'integrazione dei dati non funziona correttamente:

1. **Verifica dei file**:
   - Controlla che i file `iscritti_30_aprile.csv` e `crediti.csv` esistano nei percorsi corretti
   - Verifica che i formati e i separatori dei file siano corretti

2. **Problemi di matching**:
   - Verifica che i codici fiscali siano coerenti tra i file
   - Controlla la normalizzazione di nomi e cognomi (maiuscole/minuscole, spazi)
   - Verifica che le denominazioni delle attività siano coerenti tra i file

3. **Informazioni di Debug**:
   - Il sistema fornisce messaggi informativi sul numero di record abbinati
   - Viene mostrata la percentuale di abbinamento per ciascuna colonna integrata
   - In caso di problemi gravi, viene mostrata un'analisi di alcune corrispondenze

## Note Tecniche

- Il sistema normalizza automaticamente nomi, cognomi e codici fiscali prima di tentare l'abbinamento
- Per i CFU viene utilizzato un algoritmo di fuzzy matching per gestire piccole differenze nei nomi delle attività
- L'integrazione avviene durante il caricamento dei dati e non richiede intervento manuale
