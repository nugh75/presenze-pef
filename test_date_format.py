"""
Test di formattazione date per Excel.
Questo script testa la funzione format_datetime_for_excel
per verificare che formatti correttamente le date e le ore 
per l'esportazione in Excel.
"""
import pandas as pd
import numpy as np
from datetime import datetime, date, time
from io import BytesIO
from modules.utils import format_datetime_for_excel

def main():
    # Crea DataFrame di test con vari formati di data e ora
    data = {
        'DataPresenza': [
            date(2023, 5, 10),
            date(2023, 6, 15),
            pd.NaT,
            date(2023, 7, 20)
        ],
        'OraPresenza': [
            time(9, 30, 0),
            time(14, 15, 0),
            time(11, 0, 0),
            pd.NaT
        ],
        'Nome': ['Mario', 'Luigi', 'Giovanni', 'Antonio'],
        'Cognome': ['Rossi', 'Bianchi', 'Verdi', 'Neri']
    }

    df = pd.DataFrame(data)
    
    print("DataFrame originale:")
    print(df)
    print("\nTipi di dato:")
    print(df.dtypes)
    
    # Applica la funzione di formattazione
    df_formatted = format_datetime_for_excel(df)
    
    print("\nDataFrame formattato:")
    print(df_formatted)
    print("\nTipi di dato formattati:")
    print(df_formatted.dtypes)
    
    # Esporta in Excel per verifica
    with pd.ExcelWriter('test_date_format.xlsx', engine='xlsxwriter') as writer:
        df_formatted.to_excel(writer, sheet_name='Test', index=False)
    
    print("\nFile Excel salvato come 'test_date_format.xlsx'")
    print("Verifica che le date e le ore appaiano correttamente in Excel.")

if __name__ == "__main__":
    main()
