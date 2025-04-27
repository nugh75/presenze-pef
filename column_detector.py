# Programma per identificare i nomi delle colonne nel file Excel
import pandas as pd
import os

def main():
    try:
        file_path = 'Presenze - 2025_04_24.xlsx'
        if not os.path.exists(file_path):
            print(f"File non trovato: {file_path}")
            print("Directory corrente:", os.getcwd())
            print("File nella directory corrente:", os.listdir())
            return
            
        # Legge il file Excel e stampa i nomi delle colonne
        print(f"Lettura del file: {file_path}")
        df = pd.read_excel(file_path)
        print("Nomi delle colonne:")
        print(df.columns.tolist())
        
        # Stampa informazioni aggiuntive
        print(f"Numero totale di record: {len(df)}")
        print("Prime 3 righe:")
        print(df.head(3))
        
    except Exception as e:
        print(f"Errore: {str(e)}")

if __name__ == "__main__":
    main()
