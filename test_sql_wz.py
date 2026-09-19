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
    sku_do_wydania = "KART-002"
    ile_wydac = 100

    # 1. Pobieramy ID produktu na podstawie SKU
    cursor.execute("SELECT ID FROM Produkty WHERE SKU = ?", (sku_do_wydania,))
    row = cursor.fetchone()

    if not row:
        print(f"--> [BŁĄD WZ] Nie znaleziono towaru o SKU: {sku_do_wydania}")
    else:
        produkt_id = row[0]

        # 2. Aktualizacja stanu w tabeli Produkty (z blokadą ujemnego stanu)
        sql_update = """
            UPDATE Produkty
            SET Ilosc = Ilosc - ?
            WHERE ID = ? AND Ilosc >= ?;
        """
        cursor.execute(sql_update, (ile_wydac, produkt_id, ile_wydac))

        if cursor.rowcount > 0:
            # 3. Dodanie wpisu do historii (RuchyMagazynowe)
            sql_log = """
                INSERT INTO RuchyMagazynowe (ProduktID, TypRuchu, Ilosc)
                VALUES (?, 'WZ', ?);
            """
            cursor.execute(sql_log, (produkt_id, ile_wydac))

            # Zatwierdzamy OBA kroki naraz!
            conn.commit()
            print(f"--> [WZ] Wydano {ile_wydac} szt. dla SKU: {sku_do_wydania}. Zarejestrowano ruch w historii.")
        else:
            conn.rollback()
            print(f"--> [BŁĄD WZ] Zbyt mała ilość na stanie dla SKU: {sku_do_wydania}")

except pyodbc.Error as e:
    conn.rollback()
    print(f"Błąd bazy danych: {e}")

finally:
    cursor.close()
    conn.close()