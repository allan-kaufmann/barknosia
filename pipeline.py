import os
import argparse
import subprocess
import sys
from pathlib import Path
from dotenv import load_dotenv
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Pt, RGBColor

# Lädt die Umgebungsvariablen aus der .env-Datei
load_dotenv()

# Gemini API-Client initialisieren
client = genai.Client()

def run_marker_ocr(input_pdf_path: str, output_dir: str) -> str:
    """
    Schritt 1: Ruft das 'marker'-Modul direkt über den Python-Interpreter auf.
    Das verhindert den 'Permission denied' Fehler unter WSL/Linux vollständig.
    """
    print(f"--- Schritt 1: Starte Marker-OCR für {input_pdf_path} ---")
    os.makedirs(output_dir, exist_ok=True)
    
    # Wir nehmen den exakten Python-Interpreter deiner aktiven virtuellen Umgebung
    python_executable = sys.executable
    
    # Befehl wird als Modulaufruf gestartet (entspricht: python -m marker.cli.convert_single ...)
    command = [
        python_executable, 
        "-m", "marker.cli.convert_single", 
        str(input_pdf_path), 
        "--output_dir", str(output_dir)
    ]
    
    print(f"Führe OCR aus (das kann einen Moment dauern)...")
    try:
        # Führt den Befehl aus und leitet Ausgaben um, damit das Terminal übersichtlich bleibt
        subprocess.run(command, check=True, text=True, stdout=subprocess.DEVNULL)
        print(f"Marker erfolgreich ausgeführt.\n")
        
        # Den Pfad der generierten .md-Datei ermitteln
        pdf_stem = Path(input_pdf_path).stem
        expected_md_path = Path(output_dir) / pdf_stem / f"{pdf_stem}.md"
        
        if expected_md_path.exists():
            return str(expected_md_path)
        else:
            # Fallback: Falls marker den Ordnernamen leicht abgewandelt hat
            found_md_files = list(Path(output_dir).glob("**/*.md"))
            if found_md_files:
                return str(found_md_files[0])
            raise FileNotFoundError("Marker hat den Prozess beendet, aber es wurde keine .md-Datei gefunden.")
            
    except subprocess.CalledProcessError as e:
        print(f"\n[FEHLER] Marker-OCR fehlgeschlagen. Vergewissere dich, dass 'marker' in deiner .venv installiert ist.")
        raise e


def check_if_english(text: str) -> bool:
    """Schritt 2a: Prüfe per schnellem KI-Aufruf via gemini-2.5-flash, ob der Text englisch ist."""
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
    """Schritt 2b: Übersetzt den Text mit gemini-2.5-pro nach deinen strengen Regeln."""
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


def create_word_document(markdown_text: str, output_docx_path: str, hide_original: bool = True):
    """Schritt 3: Konvertiert den strukturierten Markdown-Text in ein Word-Dokument."""
    print(f"--- Schritt 3: Generiere strukturiertes Word-Dokument -> {output_docx_path} ---")
    
    doc = Document()
    
    style_normal = doc.styles['Normal']
    font = style_normal.font
    font.name = 'Arial'
    font.size = Pt(11)
    
    lines = markdown_text.split('\n')
    
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
            
        if stripped.startswith('# '):
            p = doc.add_heading(stripped[2:], level=1)
            p.style.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
        elif stripped.startswith('## '):
            p = doc.add_heading(stripped[3:], level=2)
            p.style.font.color.rgb = RGBColor(0x00, 0x44, 0x88)
        elif stripped.startswith('### '):
            p = doc.add_heading(stripped[4:], level=3)
        else:
            p = doc.add_paragraph()
            run = p.add_run(line)
            
            if hide_original:
                run.font.color.rgb = RGBColor(0xBB, 0xBB, 0xBB)
                run.font.size = Pt(9.5)

    doc.save(output_docx_path)
    print(f"Word-Dokument erfolgreich erstellt.\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Vollautomatische PDF-Übersetzungs- und Formatierungs-Pipeline")
    
    # Nimmt jetzt wie gewünscht wieder direkt den Pfad zur .pdf Datei an
    parser.add_argument("pdf_path", type=str, help="Pfad zur Quell-PDF-Datei")
    
    args = parser.parse_args()
    
    MARKER_OUTPUT_ORDNER = "workspace/output"
    
    try:
        if not os.getenv("GEMINI_API_KEY"):
            raise ValueError("Kein GEMINI_API_KEY in der .env-Datei gefunden!")

        if os.path.exists(args.pdf_path):
            # 1. Marker OCR direkt aus dem Skript starten
            markdown_datei_pfad = run_marker_ocr(args.pdf_path, MARKER_OUTPUT_ORDNER)
            
            # Die von Marker neu erzeugte .md Datei einlesen
            with open(markdown_datei_pfad, "r", encoding="utf-8") as f:
                aktueller_text = f.read()
            
            # 2. Sprache prüfen & ggfls. übersetzen
            is_english = check_if_english(aktueller_text)
            
            if is_english:
                print("Text ist Englisch. Starte Übersetzung...")
                aktueller_text = translate_text(aktueller_text)
                
                # Speichere die rohe Übersetzung als Backup ab
                output_md_pfad = Path(markdown_datei_pfad).parent / "de_uebersetzung.md"
                with open(output_md_pfad, "w", encoding="utf-8") as f:
                    f.write(aktueller_text)
                print(f"Übersetztes Markdown gesichert unter: {output_md_pfad}")
            else:
                print("Text ist bereits Deutsch. Keine Übersetzung notwendig.")
            
            # 3. Word-Dokument erstellen
            pdf_stem = Path(args.pdf_path).stem
            output_docx = Path(MARKER_OUTPUT_ORDNER) / pdf_stem / f"{pdf_stem}_studienbasis.docx"
            
            create_word_document(aktueller_text, str(output_docx), hide_original=True)
            
            print(f"=== Pipeline-Etappe erfolgreich! ===")
            print(f"Word-Basis liegt bereit in: {output_docx}")
                
        else:
            print(f"Fehler: Die Datei '{args.pdf_path}' wurde nicht gefunden.")
            
    except Exception as e:
        print(f"Pipeline abgebrochen wegen: {e}")