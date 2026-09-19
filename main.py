import csv
from datetime import datetime
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

def pokaz_stan_magazynu(db):
    wiersze = db.pobierz_stan_magazynu()

    print("\n" + "=" * 60)
    print(" AKTUALNY STAN MAGAZYNOWY")
    print("=" * 60)
    for sku, nazwa, ilosc, cena in wiersze:
        print(f"SKU: {sku:<10} | Nazwa: {nazwa:<25} | Stan: {ilosc:>4} szt. | Cena: {cena:>6.2f} PLN")
    print("=" * 60 + "\n")


def pokaz_historie_ruchow(db):
    wiersze = db.pobierz_historie()

    print("\n" + "=" * 75)
    print(" HISTORIA RUCHÓW MAGAZYNOWYCH")
    print("=" * 75)
    for r_id, sku, nazwa, typ, ilosc, data in wiersze:
        data_str = data.strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{r_id:>3}] {data_str} | Typ: {typ:<3} | SKU: {sku:<10} | Ilosc: {ilosc:>4} szt. | {nazwa}")
    print("=" * 75 + "\n")

def wybierz_sku(db, wymagaj_istniejacego):
    """Pyta o SKU. Jeśli nie ma dokładnego dopasowania, proponuje podobne pozycje.
    wymagaj_istniejacego=True -> SKU musi istnieć w bazie (użycie: WZ)
    wymagaj_istniejacego=False -> brak dopasowań oznacza nowy produkt (użycie: PZ)
    """
    while True:
        sku = input("Podaj SKU: ").strip()
        if not sku:
            print("[!] SKU nie może być puste.\n")
            continue

        if db.sku_istnieje(sku):
            return sku

        dopasowania = db.szukaj_po_prefiksie(sku)
        if dopasowania:
            print(f"\nNie znaleziono dokładnego SKU '{sku}'. Podobne pozycje:")
            for i, (d_sku, d_nazwa, d_ilosc) in enumerate(dopasowania, start=1):
                print(f"  {i}. {d_sku:<10} | {d_nazwa:<25} | Stan: {d_ilosc} szt.")
            print("  0. Wpisz SKU ponownie")

            wybor = input("Wybierz numer pozycji: ").strip()
            if wybor == "0":
                print()
                continue
            if wybor.isdigit() and 1 <= int(wybor) <= len(dopasowania):
                return dopasowania[int(wybor) - 1][0]
            print("[!] Nieprawidłowy wybór.\n")
            continue

        if wymagaj_istniejacego:
            print(f"[!] Brak towaru o SKU zaczynającym się na '{sku}'.\n")
            return None
        else:
            return sku  # traktujemy jako nowy produkt

def przyjecie_pz(db):
    print("\n--- PRZYJĘCIE TOWARU (PZ) ---")
    sku = wybierz_sku(db, wymagaj_istniejacego=False)

    ilosc_str = input("Podaj ilość do przyjęcia: ").strip()
    if not ilosc_str.isdigit() or int(ilosc_str) <= 0:
        print("[!] Ilość musi być dodatnią liczbą całkowitą.\n")
        return
    ilosc = int(ilosc_str)

    nazwa = None
    cena = None

    if not db.sku_istnieje(sku):
        nazwa = input("Nowy SKU - podaj nazwę produktu: ").strip()
        cena_str = input("Podaj cenę (np. 12.50): ").strip().replace(",", ".")
        try:
            cena = float(cena_str)
        except ValueError:
            print("[!] Nieprawidłowa cena. Anulowano operację.\n")
            return

    sukces, komunikat = db.dodaj_pz(sku, ilosc, nazwa, cena)
    prefiks = "-->" if sukces else "--> [BŁĄD]"
    print(f"{prefiks} {komunikat}\n")

def wydanie_wz(db):
    print("\n--- WYDANIE TOWARU (WZ) ---")
    sku = wybierz_sku(db, wymagaj_istniejacego=True)
    if sku is None:
        return

    ilosc_str = input("Podaj ilość do wydania: ").strip()
    if not ilosc_str.isdigit() or int(ilosc_str) <= 0:
        print("[!] Ilość musi być dodatnią liczbą całkowitą.\n")
        return   
    ilosc = int(ilosc_str)

    sukces, komunikat = db.dodaj_wz(sku, ilosc)
    prefiks = "-->" if sukces else "--> [BŁĄD]"
    print(f"{prefiks} {komunikat}\n")

     

def eksport_raportu(db):
    print("\n--- EKSPORT RAPORTU STANU MAGAZYNOWEGO ---")
    print("1. CSV")
    print("2. Excel (.xlsx)")
    format_wybor = input("Wybierz format (1-2): ").strip()

    if format_wybor not in ("1", "2"):
        print("[!] Nieprawidłowy wybór formatu.\n")
        return

    if format_wybor == "2" and not OPENPYXL_DOSTEPNY:
        print("[!] Biblioteka 'openpyxl' nie jest zainstalowana.")
        print("    Zainstaluj ją poleceniem: pip install openpyxl\n")
        return

    wiersze = db.pobierz_stan_magazynu()

    if not wiersze:
        print("[!] Brak danych do eksportu.\n")
        return

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    if format_wybor == "1":
        nazwa_pliku = f"raport_magazyn_{timestamp}.csv"
        with open(nazwa_pliku, mode="w", newline="", encoding="utf-8-sig") as plik:
            writer = csv.writer(plik, delimiter=";")
            writer.writerow(["SKU", "Nazwa", "Ilosc", "Cena"])
            for sku, nazwa, ilosc, cena in wiersze:
                writer.writerow([sku, nazwa, ilosc, cena])
        print(f"--> Wyeksportowano raport do pliku: {nazwa_pliku}\n")

    else:  # Excel
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Stan magazynowy"

        ws.append(["SKU", "Nazwa", "Ilosc", "Cena"])
        for kom in ws[1]:
            kom.font = Font(bold=True)

        for sku, nazwa, ilosc, cena in wiersze:
            ws.append([sku, nazwa, ilosc, float(cena)])

        szerokosci = [12, 30, 10, 10]
        for i, szer in enumerate(szerokosci, start=1):
            ws.column_dimensions[chr(64 + i)].width = szer

        nazwa_pliku = f"raport_magazyn_{timestamp}.xlsx"
        wb.save(nazwa_pliku)
        print(f"--> Wyeksportowano raport do pliku: {nazwa_pliku}\n")


def menu():
    db = WMSDatabase(CONNECTION_STRING)

    while True:
        print("--- SYSTEM MANAGEMENT WMS ---")
        print("1. Wyświetl stan magazynowy")
        print("2. Wyświetl historię ruchów (PZ/WZ)")
        print("3. Przyjęcie towaru (PZ)")
        print("4. Wydanie towaru (WZ)")
        print("5. Eksport raportu (CSV/Excel)")
        print("0. Wyjście z programu")

        wybor = input("\nWybierz opcję (0-5): ").strip()

        if wybor == "1":
            pokaz_stan_magazynu(db)
        elif wybor == "2":
            pokaz_historie_ruchow(db)
        elif wybor == "3":
            przyjecie_pz(db)
        elif wybor == "4":
            wydanie_wz(db)
        elif wybor == "5":
            eksport_raportu(db)
        elif wybor == "0":
            print("\nZamknięcie systemu WMS. Do widzenia!")
            break
        else:
            print("\n[!] Niepoprawna opcja. Spróbuj ponownie.\n")


if __name__ == "__main__":
    menu()