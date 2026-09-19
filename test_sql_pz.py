import pyodbc

conn = pyodbc.connect(
    'DRIVER={ODBC Driver 17 for SQL Server};'
    'SERVER=MATEUSZ_EDU\\SQLEXPRESS;'
    'DATABASE=TestDB;'
    'Trusted_Connection=yes;'
    'TrustServerCertificate=yes;'
)

cursor = conn.cursor()

try:
    # Dane do przyjęcia PZ
    sku = "KART-003"
    nazwa = "Karton Gabaryt L"
    ilosc_pz = 100
    cena = 8.50

    # 1. Sprawdzamy, czy SKU już istnieje w bazie
    cursor.execute("SELECT ID FROM Produkty WHERE SKU = ?", (sku,))
    row = cursor.fetchone()

    if row:
        # SCENARIUSZ A: Produkt istnieje -> Zwiększamy stan
        produkt_id = row[0]
        cursor.execute(
            "UPDATE Produkty SET Ilosc = Ilosc + ? WHERE ID = ?",
            (ilosc_pz, produkt_id)
        )
        print(f"--> Zwiększono stan istniejącego produktu {sku} o {ilosc_pz} szt.")
    else:
        # SCENARIUSZ B: Nowy produkt -> wstawiamy do tabeli Produkty
        # OUTPUT INSERTED.ID zwraca ID nowego wiersza w TYM SAMYM poleceniu
        sql_insert_prod = """
            INSERT INTO Produkty (SKU, Nazwa, Ilosc, Cena)
            OUTPUT INSERTED.ID
            VALUES (?, ?, ?, ?);
        """
        cursor.execute(sql_insert_prod, (sku, nazwa, ilosc_pz, cena))
        produkt_id = cursor.fetchone()[0]  # <-- ID przychodzi od razu z INSERT-a
        print(f"--> Dodano nowy produkt do bazy: {nazwa} ({sku})")

    # 2. Rejestracja ruchu w historii (dla obu scenariuszy)
    sql_log = """
        INSERT INTO RuchyMagazynowe (ProduktID, TypRuchu, Ilosc)
        VALUES (?, 'PZ', ?);
    """
    cursor.execute(sql_log, (produkt_id, ilosc_pz))

    conn.commit()  # Trwałe zapisanie zmian w bazie
    print(f"--> [PZ SUCCESS] Zarejestrowano przyjęcie {ilosc_pz} szt. w historii.")

except pyodbc.Error as e:
    conn.rollback()
    print(f"Błąd bazy danych: {e}")

finally:
    cursor.close()
    conn.close()