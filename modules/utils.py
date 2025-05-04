# Funzioni di utilità per l'applicazione Gestione Presenze
import re
import pandas as pd
import numpy as np
from datetime import datetime, date, time
import unicodedata

def normalize_name_advanced(name):
    """
    Normalizza un nome rimuovendo accenti, apostrofi e altri caratteri speciali.
    Standardizza le variazioni comuni per migliorare il matching dei nomi degli studenti.
    
    Args:
        name: Nome o cognome da normalizzare
        
    Returns:
        Stringa normalizzata in formato lowercase senza accenti o caratteri speciali
    """
    if not isinstance(name, str) or not name.strip():
        return ""
    
    # Converti in minuscolo
    name = name.lower().strip()
    
    # Rimuovi accenti (decomposizione NFD e rimozione dei caratteri combinanti)
    name = ''.join(c for c in unicodedata.normalize('NFD', name) if not unicodedata.combining(c))
    
    # Gestisci apostrofi e caratteri speciali
    # - Rimuovi apostrofi e caratteri speciali (mantieni solo lettere e spazi)
    name = re.sub(r'[^\w\s]', '', name)
    
    # Standardizza spazi multipli
    name = ' '.join(name.split())
    
    # Gestisci casi particolari comuni
    replacements = {
        # Variazioni comuni
        'maria': 'maria',
        'anna': 'anna',
        'giovanni': 'giovanni',
        'giuseppe': 'giuseppe',
        'angelo': 'angelo',
        'deangelo': 'de angelo',  # Gestione spazi in nomi composti
        'de angelo': 'de angelo',
        'dell': 'dell',           # Prefissi comuni
        'della': 'della',
        'dello': 'dello',
        'dal': 'dal',
        'dalla': 'dalla',
        'del': 'del',
    }
    
    # Applica le sostituzioni per standardizzare i nomi composti comuni
    for key, value in replacements.items():
        # Sostituzione solo se è una parola completa
        name = re.sub(r'\b' + key + r'\b', value, name)
    
    return name

def normalize_generic(name):
    """Rimuove 'art.13' e spazi dalle stringhe"""
    if not isinstance(name, str): return name
    normalized = re.sub(r'\s*\(?art\.?\s*13\.?\)?.*$', '', name, flags=re.IGNORECASE)
    return normalized.strip()
    
def reposition_code_to_front(text):
    """Prende il testo, estrae il codice tra parentesi (es. (A-30)) e lo riposiziona all'inizio della stringa."""
    if not isinstance(text, str): return text
    match = re.search(r'\(([-\w]+)\)', text)  # Cerca codice alfanumerico con trattini tra parentesi
    if match:
        code = match.group(1).strip()
        # Rimuovi il codice con le parentesi dal testo originale
        cleaned_text = re.sub(r'\s*\(' + re.escape(code) + r'\)\s*', ' ', text).strip()
        # Restituisci il codice all'inizio seguito dal testo pulito
        return f"[{code}] {cleaned_text}"
    return text
    
def transform_by_codice_percorso(codice, default_name):
    """Trasforma il percorso in base al codice"""
    if pd.isna(codice): return default_name
    codice_str = str(codice).strip()
    if len(codice_str) < 3: return default_name
    prefix = codice_str[:3]
    if prefix == '600': return "PeF60 All. 1"
    elif prefix == '300': return "PeF30 All. 2"
    elif prefix == '360': return "PeF36 All. 5"
    elif prefix == '200': return "PeF30 art. 13"
    else: return default_name
    
def clean_sheet_name(name):
    """Pulisce i nomi dei fogli Excel"""
    name = re.sub(r'[\\/?*\[\]:]', '', str(name))
    return name[:31]
    
def extract_code_from_parentheses(text):
    """Estrae codici tra parentesi"""
    if not isinstance(text, str): return None
    match = re.search(r'\((.*?)\)', text)
    if match:
        code = match.group(1).strip()
        if code: return code
    return None

def extract_sort_key(percorso_str):
    """Estrai chiavi di ordinamento dai percorsi"""
    code_match = re.search(r'^\[([-\w]+)\]', str(percorso_str))
    if code_match:
        return code_match.group(1)
    return str(percorso_str)
    
def format_datetime_for_excel(df):
    """
    Formatta correttamente le colonne di data e ora per l'esportazione Excel.
    
    Args:
        df: DataFrame pandas da preparare per l'export in Excel
        
    Returns:
        DataFrame con date formattate per Excel
    """
    df_export = df.copy()
    
    # Lista delle possibili colonne contenenti date/orari
    date_columns = ['DataPresenza']
    time_columns = ['OraPresenza']
    
    try:
        # Formatta le colonne di data (mantenendo il tipo datetime)
        for col in date_columns:
            if col in df_export.columns:
                # Se la colonna contiene oggetti date, convertili in datetime
                if pd.api.types.is_datetime64_dtype(df_export[col]) or df_export[col].apply(lambda x: isinstance(x, date) if not pd.isna(x) else False).any():
                    # Teniamo il formato datetime che Excel riconoscerà correttamente
                    df_export[col] = pd.to_datetime(df_export[col], errors='coerce')
        
        # Formatta le colonne di ora come stringa in formato "HH:MM"
        for col in time_columns:
            if col in df_export.columns:
                # Convertiamo in formato stringa "HH:MM" che sarà leggibile in Excel
                df_export[col] = df_export[col].apply(
                    lambda x: x.strftime('%H:%M') 
                    if isinstance(x, time) else (
                        x.strftime('%H:%M')
                        if isinstance(x, datetime) or isinstance(x, pd.Timestamp) 
                        else x
                    ) if not pd.isna(x) else ""
                )
    
    except Exception as e:
        # In caso di errore nella formattazione, restituiamo il dataframe originale
        print(f"Errore nella formattazione delle date: {e}")
        
    return df_export
