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
    Öffnet deine einzelne PDF, extrahiert das Datum des Anschreibens und die Tabelleneinträge,
    und gibt ein pandas DataFrame mit den Spalten Position, Datum des Anschreibens, Betrag zurück.
    """
    try:
        # PDF einlesen
        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() + "\n"
            print(f"\n=== Verarbeite: {os.path.basename(pdf_path)} ===")

        # Datum des Anschreibens extrahieren (z.B. '05.09.2024')
        date_match = re.search(r'\b(\d{2}\.\d{2}\.\d{4})\b', text)
        if not date_match:
            print(f"Warnung: Datum des Anschreibens nicht gefunden in {pdf_path}")
            letter_date = None
        else:
            letter_date = date_match.group(1)
            print(f"Datum gefunden: {letter_date}")

        # VERBESSERTE Tabelleneinträge-Extraktion
        # Wichtig: Nur Beträge mit LEERZEICHEN vor dem € sind Tabellenwerte!
        # Pattern sucht nach: Position + Betrag mit Leerzeichen + € am Zeilenende
        pattern_tabelle = r'^(.+?)\s+(-?\d{1,3}(?:\.\d{3})*,\d{2})\s+€\s*$'

        rows = []

        # Verarbeite Zeilen einzeln
        lines = text.split('\n')
        for line in lines:
            line = line.strip()
            if not line:
                continue

            # Suche nach Tabelleneinträgen
            match = re.search(pattern_tabelle, line, re.MULTILINE)
            if match:
                position = match.group(1).strip()
                # Bereinige Fußnoten (z.B. "1)")
                position = re.sub(r'\s*\d+\)\s*$', '', position)
                amount_str = match.group(2)

                # Konvertiere Betrag zu float
                amount = float(amount_str.replace('.', '').replace(',', '.'))

                rows.append({
                    'Position': position,
                    'Betrag (€)': amount
                })
                print(f"  ✓ Gefunden: {position} | {amount_str} €")

        # Standort aus dem Katalog finden
        standort = None
        for adresse, ort in STANDORTE.items():
            # Flexible Suche
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

        print(f"Standort: {standort if standort else 'Nicht gefunden'}")

        # Verwendungszweck finden
        pattern_verwendungszweck = r'\d{13}\s+Erst\.\:\s+\d{2}/\d{4}\s*[A-Z]\s+\+\s+\d{2}/\d{4}\s*[A-Z]'
        verwendungszweck_match = re.search(pattern_verwendungszweck, text)
        verwendungszweck = verwendungszweck_match.group(0) if verwendungszweck_match else ''

        # Füge Metadaten zu allen Zeilen hinzu
        for row in rows:
            row['Standort'] = standort if standort else ''
            row['Datum des Anschreibens'] = letter_date
            row['Quelldatei'] = os.path.basename(pdf_path)
            row['Abrechnungsstelle'] = str(os.path.basename(pdf_path)).split()[0]
            row['Verwendungszweck'] = verwendungszweck

        print(f"\nGefundene Positionen: {len(rows)}")

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
        rows = extract_data(pdf_path)
        all_rows.extend(rows)

    # Kombiniertes DataFrame erstellen
    if all_rows:
        df = pd.DataFrame(all_rows)
        # Sortiere nach Quelldatei
        df = df.sort_values(['Quelldatei', 'Position'])
        return df
    else:
        print("Keine Daten gefunden.")
        return pd.DataFrame()

if __name__ == "__main__":
    folder_path = ".data"  # Aktueller Ordner, kann angepasst werden

    # Alle PDFs verarbeiten und kombiniertes DataFrame erstellen
    df = process_all_pdfs(folder_path)

    if not df.empty:
        print("\n" + "="*80)
        print("EXTRAHIERTE TABELLENPOSITIONEN:")
        print("="*80)

        # Zeige die Daten schön formatiert
        pd.set_option('display.max_columns', None)
        pd.set_option('display.width', None)
        pd.set_option('display.max_colwidth', 50)

        print(df[['Position', 'Betrag (€)', 'Standort', 'Datum des Anschreibens']].to_string(index=False))

        # Zusammenfassung
        print("\n" + "="*80)
        print("ZUSAMMENFASSUNG:")
        print("="*80)
        print(f"Anzahl Positionen: {len(df)}")

        # Prüfe ob die Summe mit dem "verbleibenden Betrag" übereinstimmt
        verbleibend = df[df['Position'].str.contains('verbleibender Betrag', case=False)]['Betrag (€)'].values
        if len(verbleibend) > 0:
            print(f"\nVerbleibender Betrag laut Dokument: {verbleibend[0]:,.2f} €")

        # CSV und Excel Export
        df.to_csv("extracted_data.csv", index=False)
        print("\nDaten wurden in 'extracted_data.csv' gespeichert.")

        with pd.ExcelWriter("extracted_data.xlsx", engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Tabellenpositionen', index=False)

            # Füge Zusammenfassung hinzu
            summary_df = pd.DataFrame({
                'Metrik': ['Anzahl Positionen', 'Gesamtsumme'],
                'Wert': [len(df), f"{total:,.2f} €"]
            })
            summary_df.to_excel(writer, sheet_name='Zusammenfassung', index=False)

        print("Daten wurden auch in 'extracted_data.xlsx' gespeichert.")
