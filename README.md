# WMS System

Prosty system zarządzania magazynem (Warehouse Management System) napisany w Pythonie, korzystający z bazy danych Microsoft SQL Server.

## Funkcjonalności

- 📦 Podgląd aktualnego stanu magazynowego
- 📜 Historia wszystkich ruchów magazynowych (przyjęcia i wydania)
- ➕ Przyjęcie towaru (PZ) — zarówno zwiększenie stanu istniejącego produktu, jak i dodanie zupełnie nowej pozycji
- ➖ Wydanie towaru (WZ) — z blokadą przed zejściem stanu magazynowego poniżej zera
- 🔍 Podpowiedzi SKU — jeśli wpisany kod nie pasuje dokładnie do żadnego produktu, system proponuje podobne pozycje na podstawie prefiksu
- 📊 Eksport raportu stanu magazynowego do pliku CSV lub Excel (.xlsx)

## Technologie

- **Python 3.13**
- **Microsoft SQL Server** (SQL Server Express)
- **pyodbc** — połączenie z bazą danych
- **openpyxl** — generowanie plików Excel

## Struktura projektu

```
├── main.py          # Interfejs użytkownika (menu, obsługa konsoli)
├── wms_DB.py         # Warstwa dostępu do bazy danych (klasa WMSDatabase)
└── .gitignore
```

Projekt jest podzielony na dwie warstwy:
- **`wms_DB.py`** zawiera klasę `WMSDatabase`, odpowiedzialną wyłącznie za komunikację z bazą danych (zapytania SQL, transakcje, obsługa błędów)
- **`main.py`** odpowiada za interakcję z użytkownikiem i wywołuje metody klasy `WMSDatabase`

Dzięki temu logika bazodanowa jest niezależna od interfejsu — w przyszłości można ją łatwo podłączyć np. pod interfejs graficzny lub API, bez zmiany kodu obsługującego bazę.

## Struktura bazy danych

Projekt korzysta z dwóch głównych tabel:

**Produkty**
| Kolumna | Typ |
|---|---|
| ID | INT (klucz główny, IDENTITY) |
| SKU | VARCHAR |
| Nazwa | NVARCHAR |
| Ilosc | INT |
| Cena | DECIMAL |

**RuchyMagazynowe**
| Kolumna | Typ |
|---|---|
| ID | INT (klucz główny, IDENTITY) |
| ProduktID | INT (klucz obcy → Produkty.ID) |
| TypRuchu | VARCHAR ('PZ' lub 'WZ') |
| Ilosc | INT |
| DataOperacji | DATETIME |

## Uruchomienie

1. Zainstaluj wymagane biblioteki:
```bash
pip install pyodbc openpyxl
```

2. Skonfiguruj connection string w `main.py` zgodnie ze swoim serwerem SQL Server.

3. Uruchom program:
```bash
python main.py
```

## Czego się nauczyłem przy tym projekcie

- Pracy z bazą danych SQL Server z poziomu Pythona (pyodbc)
- Zapytań parametryzowanych jako ochrony przed SQL Injection
- Transakcji bazodanowych (commit/rollback) i dbania o spójność danych
- Podstaw programowania obiektowego (klasy, hermetyzacja logiki bazodanowej)
- Eksportu danych do CSV i Excel
- Podstaw pracy z Git i GitHub
