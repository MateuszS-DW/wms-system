import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import csv

from wms_DB import WMSDatabase

try:
    import openpyxl
    from openpyxl.styles import Font
    OPENPYXL_DOSTEPNY = True
except ImportError:
    OPENPYXL_DOSTEPNY = False

CONNECTION_STRING = (
    'DRIVER={ODBC Driver 17 for SQL Server};'
    'SERVER=MATEUSZ_EDU\\SQLEXPRESS;'
    'DATABASE=TestDB;'
    'Trusted_Connection=yes;'
    'TrustServerCertificate=yes;'
)

class WMSApp:
    def __init__(self, root):
        self.root = root
        self.db = WMSDatabase(CONNECTION_STRING)

        root.title("System Zarządzania Magazynem (WMS)")
        root.geometry("850x500")

        self._buduj_pasek_przyciskow()
        self._buduj_tabele()

        self.pokaz_stan_magazynu()

    # ------- BUDOWA INTERFEJSU------------

    def _buduj_pasek_przyciskow(self):
        pasek = tk.Frame(self.root, pady=8)
        pasek.pack(side="top", fill="x")

        tk.Button(pasek, text="Stan magazynowy", width=16,
                  command=self.pokaz_stan_magazynu).pack(side="left", padx=5)
        tk.Button(pasek, text="Historia ruchów", width=16,
                  command=self.pokaz_historie).pack(side="left", padx=5)
        tk.Button(pasek, text="Przyjęcie (PZ)", width=16,
                  command=self.otworz_okno_pz).pack(side="left", padx=5)
        tk.Button(pasek, text="Wydanie (WZ)", width=16,
                  command=self.otworz_okno_wz).pack(side="left", padx=5)
        tk.Button(pasek, text="Eksport raportu", width=16,
                  command=self.eksport_raportu).pack(side="left", padx=5)
        
    def _buduj_tabele(self):
        self.tree = ttk.Treeview(self.root, show="headings")
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)

    def _ustaw_kolumny(self, kolumny):
        """Czyści tabelę i ustawia nowe nagłówki kolumn."""
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = kolumny
        for k in kolumny:
            self.tree.heading(k, text=k)
            self.tree.column(k, width=130, anchor="w")

    # --------- WIDOK DANYCH ----------------

    def pokaz_stan_magazynu(self):
        self._ustaw_kolumny(["SKU", "Nazwa", "Ilość", "Cena"])
        for sku, nazwa, ilosc, cena in self.db.pobierz_stan_magazynu():
            self.tree.insert("", "end", values=(sku, nazwa, ilosc, f"{cena:.2f}"))

    def pokaz_historie(self):
        self._ustaw_kolumny(["ID", "Data", "Typ", "SKU", "Nazwa", "Ilość"])
        for r_id, sku, nazwa, typ, ilosc, data in self.db.pobierz_historie():
            data_str = data.strftime("%Y-%m-%d %H:%M:%S")
            self.tree.insert("", "end", values=(r_id, data_str, typ, sku, nazwa, ilosc))

    # ----- Okno przyjęcia (PZ)-------------------

    def otworz_okno_pz(self):
        okno = tk.Toplevel(self.root)
        okno.title("Przyjęcie towaru (PZ)")
        okno.geometry("350x260")
        okno.grab_set()        # blokuje klikanie w główne okno, dopóki to jest otwarte

        wszystkie_sku = [row[0] for row in self.db.pobierz_stan_magazynu()]

        tk.Label(okno, text="SKU:").pack(pady=(15,0))
        combo_sku = ttk.Combobox(okno, values=wszystkie_sku)
        combo_sku.pack()
        self._dodaj_filtrowanie(combo_sku, wszystkie_sku)

        tk.Label(okno, text="Ilość").pack(pady=(10,0))
        entry_ilosc = tk.Entry(okno)
        entry_ilosc.pack()

        tk.Label(okno, text="Nazwa (tylko dla nowego SKU):").pack(pady=(10,0))
        entry_nazwa = tk.Entry(okno)
        entry_nazwa.pack()

        tk.Label(okno, text="Cena (tylko dla nowego SKU):").pack(pady=(10,0))
        entry_cena = tk.Entry(okno)
        entry_cena.pack()

        def zatwierdz():
            sku = combo_sku.get().strip()
            if not sku:
                messagebox.showwarning("Błąd", "Podaj SKU.")
                return

            ilosc_str = entry_ilosc.get().strip()
            if not ilosc_str.isdigit() or int(ilosc_str) <=0:
                messagebox.showwarning("Błąd", "Ilość musi być dodatnią liczbą całkowitą.")
                return
            ilosc = int(ilosc_str)

            nazwa = None
            cena = None
            if not self.db.sku_istnieje(sku):
                nazwa = entry_nazwa.get().strip()
                cena_str = entry_cena.get().strip().replace(",", ".")
                if not nazwa or not cena_str:
                    messagebox.showwarning("Błąd", "Nowy SKU wymaga podania nazwy i ceny.")
                    return
                try:
                    cena = float(cena_str)
                except ValueError:
                    messagebox.showwarning("Błąd", "Nieprawidłowa cena.")
                    return

            sukces, komunikat = self.db.dodaj_pz(sku, ilosc, nazwa, cena)
            if sukces:
                messagebox.showinfo("Sukces", komunikat)
                okno.destroy()
                self.pokaz_stan_magazynu()
            else:
                messagebox.showerror("Błąd", komunikat)

        tk.Button(okno, text="Zatwierdź", command=zatwierdz).pack(pady=15)

    # ------------ Okno wydania (WZ) -------------

    def otworz_okno_wz(self):
        okno = tk.Toplevel(self.root)
        okno.title("Wydanie towaru (WZ)")
        okno.geometry("350x180")
        okno.grab_set() 

        wszystkie_sku = [row[0] for row in self.db.pobierz_stan_magazynu()]

        tk.Label(okno, text="SKU:").pack(pady=(15, 0))
        combo_sku = ttk.Combobox(okno, values=wszystkie_sku)
        combo_sku.pack()
        self._dodaj_filtrowanie(combo_sku, wszystkie_sku)

        tk.Label(okno, text="Ilość do wydania:").pack(pady=(10, 0))
        entry_ilosc = tk.Entry(okno)
        entry_ilosc.pack()

        def zatwierdz():
            sku = combo_sku.get().strip()
            ilosc_str = entry_ilosc.get().strip()

            if not sku:
                messagebox.showwarning("Błąd", "Podaj SKU.")
                return
            if not ilosc_str.isdigit() or int(ilosc_str) <= 0:
                messagebox.showwarning("Błąd", "Ilość musi być dodatnią liczbą całkowitą.")
                return

            sukces, komunikat = self.db.dodaj_wz(sku, int(ilosc_str))
            if sukces:
                messagebox.showinfo("Sukces", komunikat)
                okno.destroy()
                self.pokaz_stan_magazynu()
            else:
                messagebox.showerror("Błąd", komunikat)

        tk.Button(okno, text="Zatwierdź", command=zatwierdz).pack(pady=15)

    # --------------- Pomocnicze: Podpowiedzi SKU w Combobox -----------------

    def _dodaj_filtrowanie(self, combo, wszystkie_wartosci):
        """Filtruje listę rozwijaną combobox w miarę pisania."""
        def przy_wpisywaniu(event):
            wpisany_tekst = combo.get().upper()
            if wpisany_tekst == "":
                combo["values"] = wszystkie_wartosci
            else:
                dopasowane = [s for s in wszystkie_wartosci if s.upper().startswith(wpisany_tekst)]
                combo["values"] = dopasowane
        combo.bind("<KeyRelease>", przy_wpisywaniu)

    # ---------- EKSPORT ----------

    def eksport_raportu(self):
        wiersze = self.db.pobierz_stan_magazynu()
        if not wiersze:
            messagebox.showinfo("Eksport", "Brak danych do eksportu.")
            return

        typy_plikow = [("Plik CSV", "*.csv")]
        if OPENPYXL_DOSTEPNY:
            typy_plikow.append(("Plik Excel", "*.xlsx"))

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        sciezka = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=typy_plikow,
            initialfile=f"raport_magazyn_{timestamp}"
        )
        if not sciezka:
            return  # użytkownik anulował okno zapisu

        if sciezka.endswith(".xlsx"):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Stan magazynowy"
            ws.append(["SKU", "Nazwa", "Ilosc", "Cena"])
            for kom in ws[1]:
                kom.font = Font(bold=True)
            for sku, nazwa, ilosc, cena in wiersze:
                ws.append([sku, nazwa, ilosc, float(cena)])
            wb.save(sciezka)
        else:
            with open(sciezka, mode="w", newline="", encoding="utf-8-sig") as plik:
                writer = csv.writer(plik, delimiter=";")
                writer.writerow(["SKU", "Nazwa", "Ilosc", "Cena"])
                for sku, nazwa, ilosc, cena in wiersze:
                    writer.writerow([sku, nazwa, ilosc, cena])

        messagebox.showinfo("Eksport", f"Zapisano raport:\n{sciezka}")


if __name__ == "__main__":
    root = tk.Tk()
    app = WMSApp(root)
    root.mainloop()


    