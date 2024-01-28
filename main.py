import os
import sys
import tkinter as tk
from ui import GestoreFoglioDiMarciaApp

def main():
    # Determina il percorso della cartella dell'eseguibile
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    else:
        application_path = os.path.dirname(os.path.abspath(__file__))

    # Stampa i percorsi completi dei file
    print("Percorso della cartella dell'eseguibile:", application_path)
    print("Percorso del file Excel:", os.path.join(application_path, "foglio_di_marcia_completo.xlsx"))
    print("Percorso del file .txt:", os.path.join(application_path, "numero_servizio.txt"))

    root = tk.Tk()
    app = GestoreFoglioDiMarciaApp(root)
    app.run()

if __name__ == "__main__":
    main()
