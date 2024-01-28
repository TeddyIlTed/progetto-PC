import tkinter as tk
from tkinter import messagebox
from excel_manager import ExcelManager
from service_number import ServiceNumberManager
from tkinter import filedialog
import os 
class GestoreFoglioDiMarciaApp:
    def __init__(self, master):
        self.master = master
        self.excel_manager = ExcelManager()
        self.service_number_manager = ServiceNumberManager()
        self.create_interface()

    def create_interface(self):
        self.master.title("Gestore Foglio di Marcia")

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
            else:
                self.inputs[text] = tk.Entry(self.master)
                self.inputs[text].grid(row=idx, column=1, sticky="ew")

        # Bottoni
        tk.Button(self.master, text="Salva", command=self.save_data).grid(row=len(labels), column=0, sticky="ew")
        tk.Button(self.master, text="Chiudi", command=self.master.quit).grid(row=len(labels), column=1, sticky="ew")

        self.master.grid_columnconfigure(1, weight=1)

    def ask_excel_save_location(self):
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Salva come"
        )
        if save_path:
            self.excel_manager.set_file_path(save_path)
        else:
            messagebox.showerror("Errore", "È necessario scegliere un percorso per salvare il file Excel.")

    def save_data(self):
        data = {label: self.inputs[label].get() for label in self.inputs}
        data['Numero servizio'] = self.service_number_manager.service_number

        if not data["Data (formato DD/MM/YY)"]:
            messagebox.showerror("Errore", "Il campo 'Data' è obbligatorio.")
            return
        if not data["Mezzo"]:
            messagebox.showerror("Errore", "Selezionare un mezzo.")
            return

        if not os.path.exists(self.excel_manager.file_path):
            save_path = self.ask_excel_save_location()
            if save_path:
                self.excel_manager.file_path = save_path
            else:
                messagebox.showerror("Errore", "È necessario scegliere un percorso per salvare il file Excel.")
                return

        self.excel_manager.add_data_to_sheet(data)
        self.service_number_manager.increment_and_save()
        self.reset_fields()
        messagebox.showinfo("Successo", "I dati sono stati salvati con successo.")



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
