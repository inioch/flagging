import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import warnings
from datetime  import datetime 
import locale


class App:

    def __init__(self,root):
        self.root = root
        self.root.title("Generator plomby")
        self.root.geometry("500x500")

        self.seal = ""  # Inicjalizacja atrybutu instancji
        self.products = ""
        self.file_path = ""
# sprawdzenie dnia tygodnia
        self.dzis = datetime.today().weekday()
        locale.setlocale(locale.LC_ALL, 'pl_PL.UTF-8')


        self.name_of_day = datetime.today().strftime('%A')


# wczytywanie pliku
        self.btn_load = tk.Button(root, text="Wybierz plik Excel", command=self.select_data)
        self.btn_load.pack(pady=10)
# pokazanie wybranego pliku
        self.seal_label = tk.Label(root, text="Wybrana plomba:")
        self.seal_label.pack(pady=5)

        self.seal_number = tk.Entry(root,width=50, fg="blue")
        self.seal_number.pack(pady=5)
# czy są baterie?

        self.type_of_batteries = tk.IntVar()
        self.type_of_batteries.set(1)

        self.label_bat = tk.Label(root, text="Załadowane baterie?")
        self.label_bat.pack()
        self.radio_bat = tk.Radiobutton(root, text="Brak", variable=self.type_of_batteries, value=1)
        self.radio_bat.pack()
        self.radio_bat2 = tk.Radiobutton(root, text="LIT-ION", variable=self.type_of_batteries, value=2)
        self.radio_bat2.pack()
        self.radio_bat3 = tk.Radiobutton(root, text="LIT-MET", variable=self.type_of_batteries, value=3)
        self.radio_bat3.pack()
        
# typ auta
        self.car_label = tk.Label(root, text="Wybierz typ auta:")
        self.car_label.pack(pady=5)

        self.car_type = tk.IntVar()

        self.r1 = tk.Radiobutton(root, text="COY", variable=self.car_type, value= 1)
        self.r1.pack()
        self.r2 = tk.Radiobutton(root, text="NCY", variable=self.car_type, value= 2)
        self.r2.pack()
        self.r3 = tk.Radiobutton(root, text="CNY", variable=self.car_type, value= 3)
        self.r3.pack()

# dodanie opbsugi sobota
        self.saturday = tk.IntVar()

        self.saturday_label = tk.Label(root, text="Czy są paczki na sobotę?")
        self.saturday_label.pack()
        self.checkbox_saturday = tk.Checkbutton(root, text="Tak", variable=self.saturday, onvalue=True, offvalue=False)
        self.checkbox_saturday.pack()
   
# wyliczenie plomby

        self.result_btn = tk.Button(root, text="Stwórz plombe", command=self.check_if_data_available)
        self.result_btn.pack(pady=10)

        self.result_label = tk.Label(root, text="Wygenerowany seal:")
        self.result_label.pack(pady=5)

        self.result_text = tk.Entry(root,width=50)
        self.result_text.pack(pady=5)


        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
        
    def toggle_batteries(self):
        if self.is_batteries == False:
            self.is_batteries = True
        else:    
            self.is_batteries = False


    def select_data(self):
        self.file_path = filedialog.askopenfilename(title="Wybierz plik Excel", filetypes=[("Pliki Excel", "*.xlsx")])
        if not self.file_path:
            return

        self.seal_number.delete(0, tk.END)
        self.seal_number.insert(0, self.file_path)

        try:
            df = pd.read_excel(self.file_path, engine='openpyxl')
            self.products = df["Product"].astype(str).unique()  # Konwersja na string + usunięcie duplikatów
        except FileNotFoundError:
            messagebox.showerror("Błąd", " Plik Excel nie został znaleziony.")
        except KeyError:
            messagebox.showerror("Błąd", " Kolumna 'Product' nie istnieje w pliku Excel. Napewno wybrałeś odpowiedni plik?")
        except Exception as e:
            messagebox.showerror("Błąd", f"Wystąpił nieoczekiwany błąd: {e}")

    def check_if_data_available(self):
        if self.file_path:
            self.create_seal()
        else:
            messagebox.showerror("Błąd", "Nie wybrano pliku.")
    def create_seal(self):
        seal_parts = []

        if "Q" in self.products:
            seal_parts.append("*ICE")
        if self.type_of_batteries.get() == 2:
            seal_parts.append("RLI")
        if self.type_of_batteries.get() == 3:
            seal_parts.append("RLM")
        if self.saturday.get() == True:
            seal_parts.append("DD6")

        if "W" in self.products and "I" in self.products:
            seal_parts.append("DDI")
        elif "I" in self.products:
            seal_parts.append("ESI")
        elif "W" in self.products:
            seal_parts.append("ESU")
        if "P" in self.products and "U" in self.products and "C" in self.products:
            seal_parts.append("TMX")
        elif "C" in self.products:
            seal_parts.append("CMX")
        elif "Q" in self.products:
            seal_parts.append("WMX")
        elif "P" in self.products and "U" in self.products:
            seal_parts.append("MIP")
        elif "P" in self.products and "C" in self.products or "P" in self.products and "Q" in self.products:
            seal_parts.append("TMX")
        elif "U" in self.products and "C" in self.products or "U" in self.products and "Q" in self.products:
            seal_parts.append("TMX")
        elif "P" in self.products:
            seal_parts.append("WPX")
        elif "U" in self.products:
            seal_parts.append("ECX")


        match self.car_type.get():
            case 1:
                seal_parts.append("COY")
            case 2:
                seal_parts.append("NCY")
            case 3:
                seal_parts.append("CNY")
        seal_parts.append("ORGKRK")
        self.seal = "".join(seal_parts)
        if self.car_type.get() != 0:
            if len(self.seal) > 29:
                self.seal = self.seal[:-6]
            self.result_text.delete(0, tk.END)
            self.result_text.insert(0, self.seal)
            if self.type_of_batteries.get() == 1:
                messagebox.showwarning("Nie wybrano baterii"," Zaleca sie wybranie baterii. Jesli są załadowane.")
            if self.saturday.get() == False and self.dzis in(3,4) :
                messagebox.showwarning("Sobota?",f"Dzisiaj {self.name_of_day}. Sprawdź czy nie są załadowane paczki na sobotę!")
        else:
            messagebox.showerror("Błąd", "Nie wybrano typu auta.")
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()

