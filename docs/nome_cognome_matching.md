# Guida all'Accoppiamento Nome-Cognome

Questa guida spiega come il sistema integra i dati degli studenti iscritti basandosi esclusivamente su Nome e Cognome.

## Principio fondamentale

Il sistema ora utilizza **esclusivamente Nome e Cognome** come chiave di accoppiamento tra i record di presenza e quelli degli studenti iscritti. Questo perché:

1. Nome e Cognome sono gli unici dati certi presenti in entrambi i file
2. Il Codice Fiscale e l'Email devono essere presi dal file iscritti

## Processo di normalizzazione

Per garantire un accoppiamento efficace indipendentemente dalla formattazione, viene applicato un processo di normalizzazione:

1. **Rimozione degli spazi** all'inizio e alla fine dei nomi
2. **Conversione in minuscolo** di tutti i caratteri
3. **Normalizzazione degli accenti** e rimozione dei segni diacritici
4. **Rimozione di spazi multipli** (convertiti in spazi singoli)

Questo permette di abbinare correttamente nomi scritti in diverse forme, ad esempio:
- "MARIO ROSSI" e "mario rossi" 
- "De Andrè" e "de andre"
- "D'Angelo" e "d angelo"

## Dati integrati

Dopo l'accoppiamento, i seguenti dati vengono sempre copiati dal file iscritti:

- **Codice Fiscale**: sostituisce o integra il CF eventualmente presente nel file presenze
- **Email**: viene sempre presa dal file iscritti

Inoltre, quando disponibili, vengono integrati:
- Percorso formativo
- Codice Classe di concorso
- Denominazione completa classe di concorso
- Dipartimento
- Matricola
- Username (LogonName)

## Statistiche e diagnostica

Durante il processo di accoppiamento, il sistema:
1. Conta quanti record sono stati abbinati con successo
2. Verifica la percentuale di completezza per ciascuna colonna integrata
3. In caso di problemi, mostra esempi di record per facilitare l'analisi

## Risoluzione problemi

Se l'accoppiamento non funziona correttamente:

1. **Verificare la presenza dei dati**:
   - Controllare che i campi Nome e Cognome esistano in entrambi i file
   - Assicurarsi che non ci siano valori mancanti o anomali

2. **Verificare la formattazione**:
   - Controllare se ci sono formati particolari (es. cognome e nome invertiti)
   - Verificare la presenza di caratteri speciali che potrebbero influire sulla normalizzazione

3. **Verificare la corrispondenza**:
   - Controllare manualmente alcuni record di esempio per verificare che i nomi e cognomi corrispondano effettivamente

## Limitazioni

Questo approccio presenta alcune limitazioni:
- In caso di omonimia (più persone con lo stesso nome e cognome), verrà utilizzato solo il primo record trovato
- Le differenze di formattazione troppo estreme potrebbero compromettere l'accoppiamento
