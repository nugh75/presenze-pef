# Funzioni di utilità per l'applicazione Gestione Presenze
import re
import pandas as pd

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
