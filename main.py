import pdfplumber
import re
import pandas as pd
import os

# Katalog der Standorte mit präzisen Adressen
STANDORTE = {
    "Florianstr. 15": "Horb am Neckar",
    "Coblitzallee 1-9": "Mannheim",
    "Erzbergstr. 121": "Karlsruhe",
    "Lohrtalweg 14": "Mosbach",
    "Schloß 2": "Bad Mergentheim",
    "Hangstr. 46-50": "Lörrach",
    "Friedrichstr. 14": "Stuttgart (Präsidium)",
    "Herdweg 21": "Stuttgart (DHBW)",
    "Friedrich-Ebert-Str. 30": "Villingen-Schwenningen",
    "Marienstr. 20": "Heidenheim",
    "Marienplatz 2": "Ravensburg",
    "Fallenbrunnen 2": "Friedrichshafen",
    "Bildungscampus 4": "Heilbronn (DHBW)",
    "Bildungscampus 23": "Heilbronn (CAS)"
}

def extract_data(pdf_path):
    """
    Öffnet die PDF, extrahiert das Datum des Anschreibens und die Tabelleneinträge,
    und gibt ein pandas DataFrame mit den Spalten Position, Datum des Anschreibens, Betrag zurück.
    """
    try:
        # PDF einlesen
        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() + "\n"

        # Datum des Anschreibens extrahieren (z.B. '05.09.2024')
        date_match = re.search(r'\b(\d{2}\.\d{2}\.\d{4})\b', text)
        if not date_match:
            print(f"Warnung: Datum des Anschreibens nicht gefunden in {pdf_path}")
            letter_date = None
        else:
            letter_date = date_match.group(1)

        # Tabelleneinträge extrahieren (Position und Betrag)
        pattern_tabelle = r'\n([^\n€]+?)\s*(?:\d+\)|\*+|†)?\s+(-?\d{1,3}(?:\.\d{3})*,\d{2})\s*€'
        matches = re.findall(pattern_tabelle, text)

        # Standort aus dem Katalog finden
        standort = None
        for adresse, ort in STANDORTE.items():
            # Flexible Suche, um verschiedene Formatierungen zu berücksichtigen
            pattern = r'\b' + re.escape(adresse.split()[0]) + r'(?:\s*\d*-?\d*)?'
            if re.search(pattern, text):
                # Genauere Überprüfung für Adressen mit Hausnummern
                if len(adresse.split()) > 1:
                    hausnummer = adresse.split()[1]
                    if hausnummer in text or re.search(r'\b' + re.escape(adresse) + r'\b', text):
                        standort = ort
                        break
                else:
                    standort = ort
                    break

        # Daten aufbereiten
        rows = []
        for position, amount_str in matches:
            pos_clean = position.strip()
            # Betrag in float umwandeln (Komma zu Punkt, Tausendertrennzeichen entfernen)
            amount = float(amount_str.replace('.', '').replace(',', '.'))
            rows.append({
                'Position': pos_clean,
                'Standort': standort if standort else '',
                'Datum des Anschreibens': letter_date,
                'Betrag (€)': amount,
                'Quelldatei': os.path.basename(pdf_path)  # Nur Dateiname ohne Pfad
            })

        return rows

    except Exception as e:
        print(f"Fehler bei der Verarbeitung von {pdf_path}: {str(e)}")
        return []

def process_all_pdfs(folder_path):
    """
    Verarbeitet alle PDFs in einem Ordner und gibt ein kombiniertes DataFrame zurück.
    """
    all_rows = []

    # Alle PDF-Dateien im Ordner finden
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]

    if not pdf_files:
        print(f"Keine PDF-Dateien im Ordner {folder_path} gefunden.")
        return pd.DataFrame()

    # Jede PDF-Datei verarbeiten
    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        print(f"Verarbeite {pdf_file}...")

        rows = extract_data(pdf_path)
        all_rows.extend(rows)

    # Kombiniertes DataFrame erstellen
    if all_rows:
        df = pd.DataFrame(all_rows)
        return df
    else:
        print("Keine Daten gefunden.")
        return pd.DataFrame()

if __name__ == "__main__":
    folder_path = ".data"  # Aktueller Ordner, kann angepasst werden

    # Alle PDFs verarbeiten und kombiniertes DataFrame erstellen
    df = process_all_pdfs(folder_path)

    if not df.empty:
        print("\nGesamtes DataFrame:")
        print(df)

        # Optional: DataFrame in CSV-Datei speichern
        df.to_csv("extracted_data.csv", index=False)
        print("\nDaten wurden in 'extracted_data.csv' gespeichert.")