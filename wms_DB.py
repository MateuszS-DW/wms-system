import pyodbc


class WMSDatabase:
    def __init__(self, connection_string):
        self.conn_str = connection_string

    def _get_connection(self):
        return pyodbc.connect(self.conn_str)

    def pobierz_stan_magazynu(self):
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT SKU, Nazwa, Ilosc, Cena FROM Produkty")
            wynik = cursor.fetchall()
            cursor.close()
            return wynik

    def pobierz_historie(self):
        with self._get_connection() as conn:
            cursor = conn.cursor()
            sql = """
                SELECT R.ID, P.SKU, P.Nazwa, R.TypRuchu, R.Ilosc,R.DataOperacji
                FROM RuchyMagazynowe R
                JOIN Produkty P ON R.ProduktID = P.ID
                ORDER BY R.DataOperacji DESC;
            """
            cursor.execute(sql)
            wynik = cursor.fetchall()
            cursor.close()
            return wynik

    def sku_istnieje(self, sku):
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT ID FROM Produkty WHERE SKU = ?", (sku,))
            wynik = cursor.fetchone()
            cursor.close()
            return wynik is not None

    def szukaj_po_prefiksie(self,prefiks):
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT SKU, Nazwa, Ilosc FROM Produkty WHERE SKU LIKE ? ORDER BY SKU",
                (prefiks + '%')
            )
            wynik = cursor.fetchall()
            cursor.close()
            return wynik

    def dodaj_pz(self, sku, ilosc, nazwa=None, cena=None):
        conn = self._get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT ID FROM Produkty WHERE SKU = ?", (sku, ))
            row = cursor.fetchone()

            if row:
                produkt_id = row[0]
                cursor.execute(
                    "UPDATE Produkty SET Ilosc = Ilosc + ? WHERE ID = ?",
                    (ilosc, produkt_id)
                )
            else:
                if not nazwa or cena is None:
                    raise ValueError("Dla nowego SKU wymagana jest nazwa i cena")
                sql_insert = """
                    INSERT INTO Produkty (SKU, Nazwa, Ilosc, Cena)
                    OUTPUT INSERTED.ID
                    VALUES (?, ?, ?, ?);
                """
                cursor.execute(sql_insert, (sku, nazwa, ilosc, cena))
                produkt_id = cursor.fetchone()[0]

            cursor.execute(
                "INSERT INTO RuchyMagazynowe (ProduktID, TypRuchu, Ilosc) VALUES (?, 'PZ', ?)",
                (produkt_id, ilosc)
            )
            conn.commit()
            return True, "Przyjęcie PZ wykonane pomyslnie."

        except Exception as e:
            conn.rollback()
            return False, f"Błąd transakcji PZ: {e}"

        finally:
            cursor.close()
            conn.close()

    def dodaj_wz(self, sku, ilosc):
        conn = self._get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT ID FROM Produkty WHERE SKU = ?", (sku, ))
            row = cursor.fetchone()

            if not row:
                return False, f"Brak towaru o podanym SKU: {sku}"

            produkt_id = row[0]
            sql_update = "UPDATE Produkty SET Ilosc = Ilosc - ? WHERE ID = ? AND Ilosc >= ?"
            cursor.execute(sql_update, (ilosc, produkt_id, ilosc))

            if cursor.rowcount > 0:
                cursor.execute(
                    "INSERT INTO RuchyMagazynowe (ProduktID, TypRuchu, Ilosc) VALUES (?, 'WZ', ?)",
                    (produkt_id, ilosc)
                )
                conn.commit()
                return True, f"Pomyślnie wydano {ilosc} szt. SKU: {sku}"
            else:
                conn.rollback()
                return False, "Odrzucono: Niewystarczający stan magazynowy!"

        except Exception as e:
            conn.rollback()
            return False, f"Błąd transakcji WZ: {e}"

        finally:
            cursor.close()
            conn.close()

    