# Pacchetto modules.ui per le interfacce utente dell'applicazione Gestione Presenze

from .tab1 import render_tab1
# Importo la funzione render_tab2 usando un metodo alternativo
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from .tab2 import render_tab2
from .tab3 import render_tab3
from .tab4 import render_tab4
