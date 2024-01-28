import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
from tkcalendar import DateEntry
from excel_manager import ExcelManager
from service_number import ServiceNumberManager
import logging

class GestoreFoglioDiMarciaApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Gestore Foglio di Marcia")
        self.excel_manager = ExcelManager()
        self.service_number_manager = ServiceNumberManager()
        self.create_interface()

    def create_interface(self):
        # Creazione dei campi di input
        self.inputs = {}
        labels = [
            "Data (formato DD/MM/YY)", 
            "Mezzo", 
            "Orario uscita (formato HH:MM)", 
            "Orario rientro (formato HH:MM)", 
            "Km finali", 
            "Matricola autista", 
            "Matricola CS",
            *[f"Matricola volontario {i}" for i in range(1, 9)],
            "Litri rifornimento", 
            "Motivo uscita"
        ]

        for idx, text in enumerate(labels):
            tk.Label(self.master, text=text).grid(row=idx, column=0, sticky="w")
            if text == "Mezzo":
                self.inputs[text] = tk.StringVar(self.master)
                dropdown = tk.OptionMenu(self.master, self.inputs[text], *self.excel_manager.mezzi)
                dropdown.grid(row=idx, column=1, sticky="ew")
                dropdown.config(width=20)
            elif text == "Data (formato DD/MM/YY)":
                self.inputs[text] = DateEntry(self.master, width=17, background='darkblue',
                                              foreground='white', borderwidth=2, date_pattern='dd/mm/y')
                self.inputs[text].grid(row=idx, column=1, sticky="ew")
            else:
                self.inputs[text] = tk.Entry(self.master)
                self.inputs[text].grid(row=idx, column=1, sticky="ew")

        # Bottoni
        tk.Button(self.master, text="Salva", command=self.save_data).grid(row=len(labels), column=0, sticky="ew")
        tk.Button(self.master, text="Chiudi", command=self.master.quit).grid(row=len(labels), column=1, sticky="ew")

        self.master.grid_columnconfigure(1, weight=1)

    def save_data(self):
        # Validazione e salvataggio dei dati
        try:
            data = {label: self.inputs[label].get() for label in self.inputs}
            data['Numero servizio'] = self.service_number_manager.service_number
            self.validate_data(data)
            self.excel_manager.add_data_to_sheet(data)
            self.service_number_manager.increment_and_save()
            self.reset_fields()
            messagebox.showinfo("Successo", "I dati sono stati salvati con successo.")
        except Exception as e:
            logging.error(f"Errore nel salvataggio dei dati: {e}")
            messagebox.showerror("Errore", "Si è verificato un errore durante il salvataggio dei dati.")

    def validate_data(self, data):
        # Qui puoi aggiungere ulteriori controlli di validazione
        if not data["Data (formato DD/MM/YY)"]:
            raise ValueError("Il campo 'Data' è obbligatorio.")
        if not data["Mezzo"]:
            raise ValueError("Selezionare un mezzo.")

    def reset_fields(self):
        for label, input_field in self.inputs.items():
            if isinstance(input_field, tk.Entry):
                input_field.delete(0, tk.END)
            elif isinstance(input_field, tk.StringVar) and label == "Mezzo":
                input_field.set('')

    def run(self):
        self.master.mainloop()

# Codice per avviare l'applicazione
if __name__ == "__main__":
    root = tk.Tk()
    app = GestoreFoglioDiMarciaApp(root)
    app.run()
