import os
import subprocess
from pathlib import Path
from pypdf import PdfReader, PdfWriter


def extract_pdf_pages(input_pdf_path: str, output_pdf_path: str, start_page: int, end_page: int):
    """
    Schritt 1a: Schneidet einen bestimmten Seitenbereich aus einem PDF heraus.
    Nutzt 1-basierte Seitenzahlen (wie man sie im PDF-Reader sieht).
    """
    print(f"--- Schritt 1a: Extrahiere Seiten {start_page} bis {end_page} ---")
    
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()
    
    # Gesamtseiten prüfen
    total_pages = len(reader.pages)
    if start_page < 1 or end_page > total_pages or start_page > end_page:
        raise ValueError(f"Ungültiger Seitenbereich. Das PDF hat {total_pages} Seiten.")
    
    # pypdf nutzt 0-basierte Indizes, daher (start_page - 1)
    for page_num in range(start_page - 1, end_page):
        writer.add_page(reader.pages[page_num])
        
    with open(output_pdf_path, "wb") as output_file:
        writer.write(output_file)
        
    print(f"Erfolgreich extrahiert: {output_pdf_path}\n")


def run_marker_ocr(input_pdf_path: str, output_dir: str):
    """
    Schritt 1b: Ruft das installierte 'marker'-Tool über die Konsole auf,
    um das extrahierte PDF in sauberes Markdown umzuwandeln.
    """
    print(f"--- Schritt 1b: Starte Marker-OCR für {input_pdf_path} ---")
    
    # Sicherstellen, dass der Output-Ordner existiert
    os.makedirs(output_dir, exist_ok=True)
    
    # Befehlsaufruf für marker (analog zu deinem Befehl: marker_single)
    # Wir nutzen subprocess, um den CLI-Befehl direkt auszuführen
    command = [
        "marker_single",
        str(input_pdf_path),
        "--output_dir", str(output_dir)
    ]
    
    try:
        # Ausführen des Befehls und Live-Ausgabe im Terminal anzeigen
        result = subprocess.run(command, check=True, text=True)
        print(f"\nMarker erfolgreich ausgeführt. Output in: {output_dir}\n")
    except subprocess.CalledProcessError as e:
        print(f"\nFehler beim Ausführen von Marker: {e}")
        raise
    except FileNotFoundError:
        print("\nFehler: Der Befehl 'marker_single' wurde nicht gefunden.")
        print("Stelle sicher, dass marker in deiner virtuellen Umgebung (venv) aktiviert und installiert ist.")
        raise


if __name__ == "__main__":
    # --- KONFIGURATION FÜR DEN ERSTEN TESTUNGLAUF ---
    
    # 1. Pfade definieren (Passe diese an deine Dateien an)
    QUELL_PDF = "meineQuelle.pdf" 
    GEKÜRZTES_PDF = "temp_extrahiert.pdf"
    MARKER_OUTPUT_ORDNER = "workspace/output"
    
    # 2. Welche Kapitel/Seiten möchtest du bearbeiten? (1-basiert)
    START_SEITE = 5
    END_SEITE = 12
    
    try:
        # Falls das Quell-PDF existiert, starten wir die Pipeline
        if os.path.exists(QUELL_PDF):
            
            # A) PDF zuschneiden
            extract_pdf_pages(QUELL_PDF, GEKÜRZTES_PDF, START_SEITE, END_SEITE)
            
            # B) Text mit Marker extrahieren
            run_marker_ocr(GEKÜRZTES_PDF, MARKER_OUTPUT_ORDNER)
            
            print("=== Schritt 1 erfolgreich abgeschlossen! ===")
            print(f"Deine Markdown-Datei liegt jetzt im Ordner: {MARKER_OUTPUT_ORDNER}")
            
        else:
            print(f"Bitte lege eine Testdatei namens '{QUELL_PDF}' in diesen Ordner oder passe den Pfad im Skript an.")
            
    except Exception as e:
        print(f"Pipeline abgebrochen wegen: {e}")