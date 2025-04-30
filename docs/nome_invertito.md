# Guida all'Accoppiamento Nome-Cognome con Rilevamento Inversione

Questo documento spiega come il sistema gestisce il caso in cui nome e cognome siano invertiti nei dati di input.

## Problema rilevato

È stato riscontrato che in alcuni file di input, i campi Nome e Cognome potrebbero essere invertiti rispetto al file degli iscritti. 
Questo causa problemi nell'accoppiamento dei dati, poiché:

- Nel file degli iscritti: la persona è registrata come `Nome=Mario, Cognome=Rossi`
- Nel file delle presenze: la stessa persona potrebbe essere registrata come `Nome=Rossi, Cognome=Mario`

## Soluzione implementata

Il sistema ora implementa un algoritmo che:

1. **Rileva automaticamente** se i campi potrebbero essere invertiti, verificando con un test preliminare su un campione di record
2. **Prova entrambe le combinazioni** di nome e cognome quando necessario
3. **Visualizza avvisi** quando rileva che i campi sono probabilmente invertiti

## Come funziona

1. Il sistema effettua un test preliminare sui primi record per verificare se ci sono più match quando nome e cognome sono invertiti
2. Se vengono rilevati match con campi invertiti, viene attivata la modalità di "doppio matching"
3. Per ogni record, il sistema prova prima il matching normale, e se non trova corrispondenza, prova invertendo i campi

## Verifica dell'inversione

Quando il sistema rileva un'inversione dei campi, mostrerà questi messaggi:

```
Rilevati N possibili match con nome e cognome invertiti. Proverò entrambe le combinazioni.
```

E quando trova una corrispondenza con campi invertiti:

```
Match trovato invertendo nome e cognome per: [Nome] [Cognome]
```

## Cosa fare se si verifica questo problema

1. **Verificare i dati di input**: Controllare come sono etichettate le colonne nel file originale
2. **Rinominare le colonne** durante il caricamento: Se possibile, rinominare le colonne in modo che corrispondano correttamente
3. **Controllare i risultati**: Verificare che l'integrazione dei dati sia avvenuta correttamente

Il sistema ora dovrebbe gestire automaticamente il problema, ma è sempre consigliabile correggere i dati alla fonte per evitare confusione.
