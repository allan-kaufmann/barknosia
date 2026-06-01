import os
import argparse
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
        "- Nichts auslassen. [cite: 5]\n"
        "- Nichts ergänzen. [cite: 6]\n"
        "- Nichts interpretieren. [cite: 6]\n"
        "- Keine Inhalte glätten, kürzen oder zusammenfassen. [cite: 6]\n"
        "- Fachbegriffe konsistent übersetzen. [cite: 6]\n"
        "- Überschriften, Absatzstruktur, Listen und Tabellenstruktur beibehalten. [cite: 7]\n"
        "- Zitate, Autorennamen, Jahreszahlen, Variablennamen, Skalen, Hypothesen und statistische Angaben exakt erhalten. [cite: 7]\n"
        "- Unklare oder beschädigte Stellen mit [UNKLAR: Originalstelle] markieren, nicht erraten. [cite: 8]\n"
        "- Bildverweise, Tabellenverweise und Abbildungsbeschriftungen erhalten. [cite: 8]\n"
        "- Markdown-Struktur beibehalten. [cite: 8]\n\n"
        "Ausgabeformat:\n"
        "1. Nur die deutsche Übersetzung. [cite: 9]\n"
        "2. Danach eine kurze Kontrollliste:\n"
        "   - Anzahl erkannter Absätze im Original [cite: 9]\n"
        "   - Anzahl übersetzter Absätze [cite: 9]\n"
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
    parser = argparse.ArgumentParser(description="KI-gestützte Übersetzungs- und Zusammenfassungs-Pipeline ab Markdown")
    parser.add_argument("md_path", type=str, help="Pfad zur Quell-Markdown-Datei (.md)")
    
    args = parser.parse_args()
    
    try:
        if not os.getenv("GEMINI_API_KEY"):
            raise ValueError("Kein GEMINI_API_KEY in der .env-Datei gefunden!")

        # Sicherheits-Check vorab: Wurde fälschlicherweise ein PDF übergeben?
        if args.md_path.lower().endswith('.pdf'):
            print("\n[FEHLER] Du hast dem Skript eine .pdf-Datei übergeben.")
            print("Dieses Skript benötigt die bereits von Marker extrahierte .md-Datei!")
            print("Beispiel für den richtigen Aufruf:")
            print("python pipeline.py workspace/output/Sonnentag-.../Sonnentag-....md\n")
            exit(1)

        if os.path.exists(args.md_path):
            # 1. Existierende Markdown-Datei einlesen
            print(f"--- Schritt 1: Lese extrahierte Markdown-Datei {args.md_path} ---")
            with open(args.md_path, "r", encoding="utf-8") as f:
                aktueller_text = f.read()
            
            # 2. Sprache prüfen & ggfls. übersetzen
            is_english = check_if_english(aktueller_text)
            
            if is_english:
                print("Text ist Englisch. Starte Übersetzung...")
                aktueller_text = translate_text(aktueller_text)
                
                # Speichere die rohe Übersetzung als Backup ab
                output_md_pfad = Path(args.md_path).parent / "de_uebersetzung.md"
                with open(output_md_pfad, "w", encoding="utf-8") as f:
                    f.write(aktueller_text)
                print(f"Übersetztes Markdown gesichert unter: {output_md_pfad}")
            else:
                print("Text ist bereits Deutsch. Keine Übersetzung notwendig.")
            
            # 3. Word-Dokument erstellen
            md_path_obj = Path(args.md_path)
            output_docx = md_path_obj.parent / f"{md_path_obj.stem}_studienbasis.docx"
            
            create_word_document(aktueller_text, str(output_docx), hide_original=True)
            
            print(f"=== Pipeline-Etappe erfolgreich! ===")
            print(f"Word-Basis liegt bereit in: {output_docx}")
                
        else:
            print(f"Fehler: Die Datei '{args.md_path}' wurde nicht gefunden.")
            
    except Exception as e:
        print(f"Pipeline abgebrochen wegen: {e}")