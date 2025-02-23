import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

class App:

    def __init__(self,root):
        self.root = root
        self.root.title("Generator plomby")
        self.root.geometry("500x400")

        self.seal = ""  # Inicjalizacja atrybutu instancji
        self.products = ""
        self.file_path = ""
# wczytywanie pliku
        self.btn_load = tk.Button(root, text="Wybierz plik Excel", command=self.select_data)
        self.btn_load.pack(pady=10)


# pokazanie wybranego pliku
        self.seal_label = tk.Label(root, text="Wybrana plomba:")
        self.seal_label.pack(pady=5)

        self.seal_number = tk.Entry(root,width=50, fg="blue")
        self.seal_number.pack(pady=5)
# typ auta

        self.car_type = tk.IntVar()

        self.r1 = tk.Radiobutton(root, text="COY", variable=self.car_type, value= 1, command=self.toggle)
        self.r1.pack()
        self.r2 = tk.Radiobutton(root, text="NCY", variable=self.car_type, value= 2, command=self.toggle)
        self.r2.pack()
        self.r3 = tk.Radiobutton(root, text="CNY", variable=self.car_type, value= 3, command=self.toggle)
        self.r3.pack()

   
# wyliczenie plomby

        self.result_btn = tk.Button(root, text="Stwórz plombe", command=self.check_if_data_available)
        self.result_btn.pack(pady=10)

        self.result_label = tk.Label(root, text="Wygenerowany seal:")
        self.result_label.pack(pady=5)

        self.result_text = tk.Entry(root,width=50)
        self.result_text.pack(pady=5)
        
    def toggle(self):
        print(f"Stan car_type: {self.car_type.get()}")


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
            seal_parts.append("*ICEKTWGTU")
        else:
            seal_parts.append("KTWGTU")
        if "P" in self.products:
            if "U" in self.products:
                if "C" in self.products:
                    seal_parts.append("TMX")
                else:
                    seal_parts.append("MIP")
            else:
                seal_parts.append("WPX")
        else:
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
            self.result_text.delete(0, tk.END)
            self.result_text.insert(0, self.seal)
        else:
            messagebox.showerror("Błąd", "Nie wybrano typu auta.")
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()

