import os
import argparse
import subprocess
from pathlib import Path
from pypdf import PdfReader, PdfWriter
from dotenv import load_dotenv
from google import genai
from google.genai import types

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
    
    # Ermittle den absoluten Pfad von marker_single im System
    try:
        marker_path = subprocess.check_output(["which", "marker_single"], text=True).strip()
        print(f"Marker-Pfad gefunden: {marker_path}")
    except subprocess.CalledProcessError:
        marker_path = "marker_single"
        print("Warnung: Konnte absoluten Pfad für 'marker_single' nicht ermitteln. Nutze Standard-Aufruf.")

    # Befehl aufbauen
    command = [marker_path, str(input_pdf_path), "--output_dir", str(output_dir)]
    
    try:
        # shell=True und String-Aufruf für korrekte Rechte- und Pfadauflösung unter WSL/Linux
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
        "Eine vollständige, sinntreue Übersetzung, keine Zusammenfassung.\n\n"
        "Strenge Regeln:\n"
        "- Nichts auslassen.\n"
        "- Nichts ergänzen.\n"
        "- Nichts interpretieren.\n"
        "- Keine Inhalte glätten, kürzen oder zusammenfassen.\n"
        "- Fachbegriffe konsistent übersetzen.\n"
        "- Überschriften, Absatzstruktur, Listen und Tabellenstruktur beibehalten.\n"
        "- Zitate, Autorennamen, Jahreszahlen, Variablennamen, Skalen, Hypothesen und statistische Angaben exakt erhalten.\n"
        "- Unklare oder beschädigte Stellen mit [UNKLAR: Originalstelle] markieren, nicht erraten.\n"
        "- Bildverweise, Tabellenverweise und Abbildungsbeschriftungen erhalten.\n"
        "- Markdown-Struktur beibehalten.\n\n"
        "Ausgabeformat:\n"
        "1. Nur die deutsche Übersetzung.\n"
        "2. Danach eine kurze Kontrollliste:\n"
        "   - Anzahl erkannter Absätze im Original\n"
        "   - Anzahl übersetzter Absätze\n"
        "   - Hinweise auf unklare Stellen\n"
        "   - Hinweise auf mögliche fehlende Tabellen/Bildinhalte"
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
                original_text = f.read()
            
            # Temp-Datei aufräumen
            if args.start is None and args.end is None and os.path.exists(TEMPORÄRES_PDF):
                os.remove(TEMPORÄRES_PDF)

            # 2. Sprache prüfen & ggfls. übersetzen
            is_english = check_if_english(original_text)
            
            if is_english:
                print("Text ist Englisch. Starte Übersetzung...")
                uebersetzter_text = translate_text(original_text)
                
                output_md_pfad = Path(markdown_datei_pfad).parent / "de_uebersetzung.md"
                with open(output_md_pfad, "w", encoding="utf-8") as f:
                    f.write(uebersetzter_text)
                    
                print(f"\n=== Schritt 1 & 2 erfolgreich abgeschlossen! ===")
                print(f"Übersetzung gespeichert unter: {output_md_pfad}")
            else:
                print("Text ist bereits Deutsch. Keine Übersetzung notwendig.")
                
        else:
            print(f"Fehler: Die Datei '{args.pdf_path}' wurde nicht gefunden.")
            
    except Exception as e:
        print(f"Pipeline abgebrochen wegen: {e}")