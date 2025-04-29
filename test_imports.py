import sys
print("Path di sistema:", sys.path)

try:
    import modules
    print("Importato: modules")
    
    import modules.utils
    print("Importato: modules.utils")
    
    import modules.data_loader
    print("Importato: modules.data_loader")
    
    from modules.data_loader import load_data
    print("Importato: load_data")
    
    import modules.ui
    print("Importato: modules.ui")
    
    try:
        from modules.ui import render_tab1
        print("Importato: render_tab1")
    except ImportError as e:
        print("Errore importando render_tab1:", e)
        
        try:
            import modules.ui.tab1
            print("Importato: modules.ui.tab1")
            print("Ha render_tab1?", hasattr(modules.ui.tab1, "render_tab1"))
            if hasattr(modules.ui.tab1, "render_tab1"):
                render_tab1 = modules.ui.tab1.render_tab1
                print("Assegnato manualmente: render_tab1")
        except ImportError as e2:
            print("Errore importando modules.ui.tab1:", e2)
        
except ImportError as e:
    print("Errore importando moduli:", e)
