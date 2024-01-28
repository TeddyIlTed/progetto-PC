from datetime import datetime
import os 
class ServiceNumberManager:
    def __init__(self):
        self.current_year = datetime.now().year
        self.service_number = self.load_or_initialize_number()
        self.file_path = os.path.join("numero_servizio.txt")


    def load_or_initialize_number(self):
        file_path = f"numero_servizio_{self.current_year}.txt"
        try:
            with open(file_path, "r") as file:
                return int(file.read().strip())
        except FileNotFoundError:
            return 1

    def increment_and_save(self):
        self.service_number += 1
        with open(f"numero_servizio_{self.current_year}.txt", "w") as file:
            file.write(str(self.service_number))

    def reset_number_if_new_year(self):
        if datetime.now().year != self.current_year:
            self.current_year = datetime.now().year
            self.service_number = 1
            self.increment_and_save()
