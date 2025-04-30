# Guida all'integrazione dei dati degli studenti iscritti

## Introduzione

L'applicazione Gestione Presenze supporta ora l'integrazione dei dati degli studenti iscritti dal file CSV `modules/dati/iscritti_29_aprile.csv`. Questa funzionalità consente di arricchire le informazioni sugli studenti con dati aggiuntivi come la classe di concorso, il dipartimento e altri dettagli.

## Funzionalità principali

- **Caricamento dati iscritti**: L'app carica automaticamente il file CSV degli iscritti all'avvio
- **Matching intelligente**: Gli studenti vengono riconosciuti tramite Codice Fiscale o, se non disponibile, tramite Nome e Cognome
- **Visualizzazione integrata**: I dati aggiuntivi vengono mostrati in tutte le tabelle che contengono informazioni sugli studenti
- **Esportazione completa**: Le esportazioni in CSV ed Excel includono tutte le informazioni aggiuntive

## Dati integrati

Il file `iscritti_29_aprile.csv` contiene le seguenti colonne che vengono integrate nei dati delle presenze:

1. `Percorso` - Percorso formativo dello studente
2. `Codice_Classe_di_concorso` - Codice della classe di concorso
3. `Codice_classe_di_concorso_e_denominazione` - Descrizione estesa della classe di concorso
4. `Dipartimento` - Dipartimento di afferenza
5. `LogonName` - Nome utente del sistema
6. `Matricola` - Numero di matricola dello studente

## Processo di matching

Il processo di abbinamento dei dati avviene in due passaggi:

1. **Matching per Codice Fiscale**: Prima si tenta di abbinare gli studenti tramite il codice fiscale (metodo più affidabile)
2. **Matching per Nome e Cognome**: Per gli studenti che non hanno trovato corrispondenza tramite CF, si prova con il nome e cognome

Al termine del processo, l'applicazione mostrerà statistiche sul numero di record abbinati e il metodo utilizzato.

## Utilizzo

Non è richiesta alcuna azione specifica per utilizzare questa funzionalità, poiché l'integrazione avviene automaticamente quando si carica un file delle presenze.

I dati integrati saranno visibili nelle seguenti schede dell'applicazione:
- **Analisi Dati**: Nella tabella principale dei dati
- **Calcolo Presenze ed Esportazione**: Nelle visualizzazioni per studente e nella lista completa degli studenti
- **Frequenza Lezioni**: Nella lista dei partecipanti a una lezione specifica

## Requisiti

Il file degli iscritti deve:
- Essere posizionato in `modules/dati/iscritti_29_aprile.csv`
- Usare il separatore punto e virgola (`;`)
- Contenere almeno le colonne `Nome`, `Cognome` e `CodiceFiscale`

## Risoluzione problemi

- Se il file degli iscritti non viene trovato, apparirà un messaggio di warning ma l'applicazione continuerà a funzionare senza l'integrazione dei dati
- Se il file non contiene le colonne necessarie, verrà mostrato un errore specifico
- Gli studenti per cui non è stato trovato un abbinamento manterranno solo i dati originali
