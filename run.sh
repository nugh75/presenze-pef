#!/bin/bash

# Script per avviare l'applicazione Streamlit di gestione presenze

echo "Avvio dell'applicazione Gestione Presenze..."
streamlit run app.py --server.maxUploadSize=1024
