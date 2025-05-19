import pdfplumber
import re
import pandas as pd
import os

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
        pattern_ort = r'\b\d{5}\s+(?!Fellbach)([A-ZÄÖÜß][a-zäöüß]+(?:[-\s]+(?:am|im|an|der|Bad|Sankt|St\.|[A-ZÄÖÜß])[a-zäöüß]+)*?)(?=\n|$)'
        matches_ort = re.findall(pattern_ort, text)

        # Daten aufbereiten
        rows = []
        for position, amount_str in matches:
            pos_clean = position.strip()
            # Betrag in float umwandeln (Komma zu Punkt, Tausendertrennzeichen entfernen)
            amount = float(amount_str.replace('.', '').replace(',', '.'))
            rows.append({
                'Position': pos_clean,
                'Standort': matches_ort[0] if matches_ort else '',
                'Datum des Anschreibens': letter_date,
                'Betrag (€)': amount,
                'Quelldatei': pdf_path  # Neue Spalte für den Dateinamen
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