from datetime import datetime
import os
import logging

class ServiceNumberManager:
    def __init__(self):
        self.current_year = datetime.now().year
        self.file_path = f"numero_servizio_{self.current_year}.txt"
        self.service_number = self.load_or_initialize_number()

    def load_or_initialize_number(self):
        try:
            if os.path.exists(self.file_path):
                with open(self.file_path, "r") as file:
                    return int(file.read().strip())
            else:
                return self.reset_number()
        except Exception as e:
            logging.error(f"Errore durante il caricamento del numero di servizio: {e}")
            return self.reset_number()

    def reset_number(self):
        logging.info("Reset del numero di servizio per il nuovo anno")
        return 1

    def increment_and_save(self):
        self.reset_number_if_new_year()
        self.service_number += 1
        try:
            with open(self.file_path, "w") as file:
                file.write(str(self.service_number))
        except Exception as e:
            logging.error(f"Errore durante il salvataggio del nuovo numero di servizio: {e}")

    def reset_number_if_new_year(self):
        if datetime.now().year != self.current_year:
            self.current_year = datetime.now().year
            self.service_number = self.reset_number()
            self.file_path = f"numero_servizio_{self.current_year}.txt"
