import pdfplumber
import re
import pandas as pd
import os

def extract_data(pdf_path):
    """
    Öffnet die PDF, extrahiert das Datum des Anschreibens und die Tabelleneinträge,
    und gibt ein pandas DataFrame mit den Spalten Position, Datum des Anschreibens, Betrag zurück.
    """
    # PDF einlesen
    with pdfplumber.open(pdf_path) as pdf:
        text = ""
        for page in pdf.pages:
            text += page.extract_text() + "\n"

    # Datum des Anschreibens extrahieren (z.B. '05.09.2024')
    date_match = re.search(r'\b(\d{2}\.\d{2}\.\d{4})\b', text)
    if not date_match:
        raise ValueError("Datum des Anschreibens nicht gefunden.")
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
            'Betrag (€)': amount
        })

    # DataFrame erstellen
    df = pd.DataFrame(rows)
    return df

if __name__ == "__main__":

    pdf_path = "test.pdf"  # Pfad zur PDF-Datei
    df = extract_data(pdf_path)
    print(df)
