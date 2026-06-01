import os
import argparse
import subprocess
from pathlib import Path
from pypdf import PdfReader, PdfWriter
from dotenv import load_dotenv
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Pt, RGBColor

# Lädt die Umgebungsvariablen aus der .env-Datei
load_dotenv()

# Gemini API-Client initialisieren
client = genai.Client()

def extract_pdf_pages(input_pdf_path: str, output_pdf_path: str, start_page: int = None, end_page: int = None):
    """Schritt 1a: Schneidet optional einen Seitenbereich aus."""
    if start_page is None and end_page is None:
        print("--- Schritt 1a: Gesamtes PDF wird verwendet (keine Seitenbegrenzung) ---")
        return input_pdf_path

    print(f"--- Schritt 1a: Extrahiere Seiten {start_page} bis {end_page} ---")
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()
    total_pages = len(reader.pages)
    
    s_page = start_page if start_page is not None else 1
    e_page = end_page if end_page is not None else total_pages
    
    if s_page < 1 or e_page > total_pages or s_page > e_page:
        raise ValueError(f"Ungültiger Seitenbereich. Das PDF hat {total_pages} Seiten.")
    
    for page_num in range(s_page - 1, e_page):
        writer.add_page(reader.pages[page_num])
        
    with open(output_pdf_path, "wb") as output_file:
        writer.write(output_file)
    return output_pdf_path


def run_marker_ocr(input_pdf_path: str, output_dir: str) -> str:
    """Schritt 1b: Ruft das 'marker'-Tool über den absoluten Systempfad auf."""
    print(f"--- Schritt 1b: Starte Marker-OCR für {input_pdf_path} ---")
    os.makedirs(output_dir, exist_ok=True)
    
    try:
        marker_path = subprocess.check_output(["which", "marker_single"], text=True).strip()
        print(f"Marker-Pfad gefunden: {marker_path}")
    except subprocess.CalledProcessError:
        marker_path = "marker_single"
        print("Warnung: Konnte absoluten Pfad für 'marker_single' nicht ermitteln. Nutze Standard-Aufruf.")

    command = [marker_path, str(input_pdf_path), "--output_dir", str(output_dir)]
    
    try:
        cmd_string = " ".join(command)
        subprocess.run(cmd_string, check=True, text=True, stdout=subprocess.DEVNULL, shell=True)
        print(f"Marker erfolgreich ausgeführt.\n")
        
        pdf_stem = Path(input_pdf_path).stem
        expected_md_path = Path(output_dir) / pdf_stem / f"{pdf_stem}.md"
        
        if expected_md_path.exists():
            return str(expected_md_path)
        else:
            found_md_files = list(Path(output_dir).glob("**/*.md"))
            if found_md_files:
                return str(found_md_files[0])
            raise FileNotFoundError("Keine .md Datei von Marker gefunden.")
    except Exception as e:
        print(f"Fehler bei Marker OCR: {e}")
        raise


def check_if_english(text: str) -> bool:
    """Prüft per schnellem KI-Aufruf via gemini-2.5-flash, ob der Text englisch ist."""
    print("--- Schritt 2a: Prüfe Sprache des Dokuments ---")
    
    leseprobe = text[:2000]
    prompt = (
        "Antworte mit exakt einem Wort, entweder 'YES' oder 'NO'. "
        "Ist der folgende Text hauptsächlich in englischer Sprache verfasst?\n\n"
        f"Text:\n{leseprobe}"
    )
    
    try:
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
            config=types.GenerateContentConfig(
                temperature=0.0,
                max_output_tokens=5
            )
        )
        ergebnis = response.text.strip().upper()
        return "YES" in ergebnis
    except Exception as e:
        print(f"Sprachprüfung fehlgeschlagen ({e}), weiche standardmäßig auf Übersetzung aus.")
        return True


def translate_text(text: str) -> str:
    """Übersetzt den Text mit gemini-2.5-pro nach den strengen Regeln aus 01_b_text_uebersetzen.txt"""
    print("--- Schritt 2b: Übersetze englischen Text ins Deutsche (via Gemini 2.5 Pro) ---")
    
    system_prompt = (
        "Übersetze den folgenden englischen wissenschaftlichen Text originalgetreu ins Deutsche.\n\n"
        "Ziel:\n"
        "Eine vollständige, sinntreue Übersetzung, keine Zusammenfassung[cite: 5].\n\n"
        "Strenge Regeln:\n"
        "- Nichts auslassen[cite: 5].\n"
        "- Nichts ergänzen[cite: 6].\n"
        "- Nichts interpretieren[cite: 6].\n"
        "- Keine Inhalte glätten, kürzen oder zusammenfassen[cite: 6].\n"
        "- Fachbegriffe konsistent übersetzen[cite: 6].\n"
        "- Überschriften, Absatzstruktur, Listen und Tabellenstruktur beibehalten.\n"
        "- Zitate, Autorennamen, Jahreszahlen, Variablennamen, Skalen, Hypothesen und statistische Angaben exakt erhalten.\n"
        "- Unklare oder beschädigte Stellen mit [UNKLAR: Originalstelle] markieren, nicht erraten[cite: 8].\n"
        "- Bildverweise, Tabellenverweise und Abbildungsbeschriftungen erhalten[cite: 8].\n"
        "- Markdown-Struktur beibehalten[cite: 8].\n\n"
        "Ausgabeformat:\n"
        "1. Nur die deutsche Übersetzung[cite: 9].\n"
        "2. Danach eine kurze Kontrollliste:\n"
        "   - Anzahl erkannter Absätze im Original [cite: 9]\n"
        "   - Anzahl übersetzten Absätze [cite: 9]\n"
        "   - Hinweise auf unklare Stellen [cite: 9]\n"
        "   - Hinweise auf mögliche fehlende Tabellen/Bildinhalte [cite: 9]"
    )
    
    try:
        response = client.models.generate_content(
            model='gemini-2.5-pro',
            contents=f"Text:\n{text}",
            config=types.GenerateContentConfig(
                system_instruction=system_prompt,
                temperature=0.1,
            )
        )
        return response.text
    except Exception as e:
        print(f"Fehler bei der Gemini-Übersetzung: {e}")
        raise


def create_word_document(markdown_text: str, output_docx_path: str, hide_original: bool = True):
    """
    Schritt 3: Konvertiert den strukturierten Markdown-Text in ein Word-Dokument.
    Falls hide_original=True, wird der Quelltext in einem sehr hellen Grau formatiert,
    damit er visuell in den Hintergrund rückt (ausgeblendet ist).
    """
    print(f"--- Schritt 3: Generiere strukturiertes Word-Dokument -> {output_docx_path} ---")
    
    doc = Document()
    
    # Standard-Styles einrichten
    style_normal = doc.styles['Normal']
    font = style_normal.font
    font.name = 'Arial'
    font.size = Pt(11)
    
    lines = markdown_text.split('\n')
    
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
            
        # Überschriften erkennen und Formatierung matchen [cite: 7, 27]
        if stripped.startswith('# '):
            p = doc.add_heading(stripped[2:], level=1)
            p.style.font.color.rgb = RGBColor(0x00, 0x33, 0x66) # Dunkelblau für Struktur
        elif stripped.startswith('## '):
            p = doc.add_heading(stripped[3:], level=2)
            p.style.font.color.rgb = RGBColor(0x00, 0x44, 0x88)
        elif stripped.startswith('### '):
            p = doc.add_heading(stripped[4:], level=3)
        else:
            # Normaler Textabsatz 
            p = doc.add_paragraph()
            run = p.add_run(line)
            
            # Falls ausgeblendet gewünscht, färben wir das Original hellgrau 
            if hide_original:
                run.font.color.rgb = RGBColor(0xBB, 0xBB, 0xBB)
                run.font.size = Pt(9.5)

    doc.save(output_docx_path)
    print(f"Word-Dokument erfolgreich erstellt.\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="KI-gestützte PDF-Übersetzungs- und Zusammenfassungs-Pipeline")
    parser.add_argument("pdf_path", type=str, help="Pfad zur Quell-PDF-Datei")
    parser.add_argument("--start", type=int, default=None, help="Startseite (optional)")
    parser.add_argument("--end", type=int, default=None, help="Endseite (optional)")
    
    args = parser.parse_args()
    
    TEMPORÄRES_PDF = "temp_verarbeitung.pdf"
    MARKER_OUTPUT_ORDNER = "workspace/output"
    
    try:
        if not os.getenv("GEMINI_API_KEY"):
            raise ValueError("Kein GEMINI_API_KEY in der .env-Datei gefunden!")

        if os.path.exists(args.pdf_path):
            # 1. PDF-Vorbereitung und Marker OCR
            pdf_zu_verarbeiten = extract_pdf_pages(args.pdf_path, TEMPORÄRES_PDF, args.start, args.end)
            markdown_datei_pfad = run_marker_ocr(pdf_zu_verarbeiten, MARKER_OUTPUT_ORDNER)
            
            with open(markdown_datei_pfad, "r", encoding="utf-8") as f:
                aktueller_text = f.read()
            
            # Temp-Datei aufräumen
            if args.start is None and args.end is None and os.path.exists(TEMPORÄRES_PDF):
                os.remove(TEMPORÄRES_PDF)

            # 2. Sprache prüfen & ggfls. übersetzen
            is_english = check_if_english(aktueller_text)
            
            if is_english:
                print("Text ist Englisch. Starte Übersetzung...")
                aktueller_text = translate_text(aktueller_text)
                
                # Speichere die rohe Übersetzung als Backup
                output_md_pfad = Path(markdown_datei_pfad).parent / "de_uebersetzung.md"
                with open(output_md_pfad, "w", encoding="utf-8") as f:
                    f.write(aktueller_text)
            else:
                print("Text ist bereits Deutsch. Keine Übersetzung notwendig.")
            
            # 3. Word-Dokument erstellen (Originaltext/Übersetzung wird hellgrau) 
            pdf_stem = Path(args.pdf_path).stem
            output_docx = Path(MARKER_OUTPUT_ORDNER) / pdf_stem / f"{pdf_stem}_studienbasis.docx"
            
            create_word_document(aktueller_text, str(output_docx), hide_original=True)
            
            print(f"=== Pipeline-Etappe erfolgreich! ===")
            print(f"Bereit für den Zusammenfassungs-Schritt. Word-Basis liegt in: {output_docx}")
                
        else:
            print(f"Fehler: Die Datei '{args.pdf_path}' wurde nicht gefunden.")
            
    except Exception as e:
        print(f"Pipeline abgebrochen wegen: {e}")