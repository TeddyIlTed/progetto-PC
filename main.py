import os
import sys
import tkinter as tk
from ui import GestoreFoglioDiMarciaApp
import logging

def setup_logging():
    # Configura il logging per registrare i messaggi di errore e di informazione
    logging.basicConfig(filename='app.log', filemode='a', 
                        format='%(asctime)s - %(levelname)s - %(message)s', 
                        level=logging.INFO)

def main():
    setup_logging()
    logging.info("Avvio dell'applicazione")

    # Determina il percorso della cartella dell'eseguibile
    try:
        if getattr(sys, 'frozen', False):
            application_path = os.path.dirname(sys.executable)
        else:
            application_path = os.path.dirname(os.path.abspath(__file__))
        logging.info(f"Percorso della cartella dell'eseguibile: {application_path}")
    except Exception as e:
        logging.error(f"Errore nel determinare il percorso dell'applicazione: {e}")
        sys.exit("Errore nel determinare il percorso dell'applicazione.")

    try:
        root = tk.Tk()
        app = GestoreFoglioDiMarciaApp(root)
        app.run()
    except Exception as e:
        logging.error(f"Errore nell'esecuzione dell'applicazione: {e}")
        sys.exit("Errore nell'esecuzione dell'applicazione.")

if __name__ == "__main__":
    main()
