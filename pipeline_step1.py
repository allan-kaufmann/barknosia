import os
import subprocess
from pathlib import Path
from pypdf import PdfReader, PdfWriter


def extract_pdf_pages(input_pdf_path: str, output_pdf_path: str, start_page: int = None, end_page: int = None):
    """
    Schritt 1a: Schneidet optional einen Seitenbereich aus.
    Wenn start_page oder end_page None sind, wird das komplette PDF genutzt.
    """
    # Wenn keine Seiten angegeben sind, überspringen wir das Schneiden
    if start_page is None and end_page is None:
        print("--- Schritt 1a: Gesamtes PDF wird verwendet (keine Seitenbegrenzung) ---")
        return input_pdf_path

    print(f"--- Schritt 1a: Extrahiere Seiten {start_page} bis {end_page} ---")
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()
    total_pages = len(reader.pages)
    
    # Standardwerte setzen, falls nur eins von beiden angegeben wurde
    s_page = start_page if start_page is not None else 1
    e_page = end_page if end_page is not None else total_pages
    
    if s_page < 1 or e_page > total_pages or s_page > e_page:
        raise ValueError(f"Ungültiger Seitenbereich. Das PDF hat {total_pages} Seiten.")
    
    for page_num in range(s_page - 1, e_page):
        writer.add_page(reader.pages[page_num])
        
    with open(output_pdf_path, "wb") as output_file:
        writer.write(output_file)
        
    print(f"Teil-PDF erfolgreich extrahiert: {output_pdf_path}\n")
    return output_pdf_path


def run_marker_ocr(input_pdf_path: str, output_dir: str) -> str:
    """
    Schritt 1b: Ruft das 'marker'-Tool über die Konsole auf.
    Gibt den Pfad zur generierten .md-Datei zurück.
    """
    print(f"--- Schritt 1b: Starte Marker-OCR für {input_pdf_path} ---") [cite: 2]
    os.makedirs(output_dir, exist_ok=True)
    
    # marker_single Befehl aufbauen
    command = [
        "marker_single",
        str(input_pdf_path),
        "--output_dir", str(output_dir)
    ] [cite: 2]
    
    try:
        subprocess.run(command, check=True, text=True)
        print(f"Marker erfolgreich ausgeführt.\n") [cite: 2]
        
        # Marker erstellt einen Unterordner mit dem Namen der PDF-Datei
        # Wir suchen nach der erzeugten .md Datei in diesem Output-Verzeichnis
        pdf_stem = Path(input_pdf_path).stem
        expected_md_path = Path(output_dir) / pdf_stem / f"{pdf_stem}.md"
        
        if expected_md_path.exists():
            return str(expected_md_path)
        else:
            # Fallback: Falls marker den Ordner anders benannt hat, suchen wir die .md Datei
            found_md_files = list(Path(output_dir).glob("**/*.md"))
            if found_md_files:
                return str(found_md_files[0])
            raise FileNotFoundError("Marker hat den Prozess beendet, aber es wurde keine .md Datei gefunden.")
            
    except subprocess.CalledProcessError as e:
        print(f"\nFehler beim Ausführen von Marker: {e}")
        raise
    except FileNotFoundError:
        print("\nFehler: Befehl 'marker_single' nicht gefunden. Ist deine venv aktiv?")
        raise


if __name__ == "__main__":
    # --- KONFIGURATION ---
    QUELL_PDF = "meineQuelle.pdf"  # Deine Testdatei hier eintragen
    TEMPORÄRES_PDF = "temp_verarbeitung.pdf"
    MARKER_OUTPUT_ORDNER = "workspace/output"
    
    # SEITEN-BEGRENZUNG (Lass beide auf None für das komplette PDF!)
    START_SEITE = None
    END_SEITE = None
    
    try:
        if os.path.exists(QUELL_PDF):
            # 1. PDF-Vorbereitung (entweder komplett oder geschnitten)
            pdf_zu_verarbeiten = extract_pdf_pages(QUELL_PDF, TEMPORÄRES_PDF, START_SEITE, END_SEITE) [cite: 1]
            
            # 2. Marker OCR laufen lassen
            markdown_datei_pfad = run_marker_ocr(pdf_zu_verarbeiten, MARKER_OUTPUT_ORDNER) [cite: 2]
            
            # 3. Kontroll-Check: Haben wir den Text erfolgreich im Skript?
            with open(markdown_datei_pfad, "r", encoding="utf-8") as f:
                extrahierter_text = f.read()
                
            print("=== Schritt 1 erfolgreich abgeschlossen! ===")
            print(f"Datei generiert: {markdown_datei_pfad}")
            print(f"Gelesene Zeichenlänge für die nächsten Schritte: {len(extrahierter_text)} Zeichen.")
            
            # Aufräumen: Wenn wir das komplette PDF genutzt haben, brauchen wir die temp-Datei nicht
            if START_SEITE is None and END_SEITE is None and os.path.exists(TEMPORÄRES_PDF):
                os.remove(TEMPORÄRES_PDF)
                
        else:
            print(f"Bitte lege eine Testdatei namens '{QUELL_PDF}' bereit.")
            
    except Exception as e:
        print(f"Pipeline abgebrochen wegen: {e}")