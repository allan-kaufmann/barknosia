import os
import argparse
import subprocess
import sys
import time
import re
from pathlib import Path
from dotenv import load_dotenv
from google import genai
from google.genai import types
from google.genai.errors import APIError
from docx import Document
from docx.shared import Pt, RGBColor

# Lädt die Umgebungsvariablen aus der .env-Datei
load_dotenv()

# Gemini API-Client mit der offiziellen Bibliothek initialisieren
client = genai.Client()

def call_gemini_with_retry(model_name: str, contents, config, max_retries: int = 5, delay: int = 5):
    """Hilfsfunktion: Ruft Gemini auf und wiederholt den Versuch bei Serverüberlastung (503)."""
    for attempt in range(1, max_retries + 1):
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=contents,
                config=config
            )
            return response
        except APIError as e:
            if e.code in [503, 429] and attempt < max_retries:
                print(f"      [Server ausgelastet] Fehler {e.code}. Warte {delay} Sekunden (Versuch {attempt}/{max_retries})...")
                time.sleep(delay)
                delay *= 2
            else:
                raise e
    raise APIError("Maximale Anzahl an Wiederholungsversuchen erreicht.")


def run_marker_ocr(input_pdf_path: str, output_dir: str) -> str:
    """Schritt 1: Konvertiert das PDF über den nativen Systemaufruf in Markdown."""
    print(f"--- Schritt 1: Starte Marker-OCR für {input_pdf_path} ---")
    os.makedirs(output_dir, exist_ok=True)
    
    abs_pdf_path = os.path.abspath(input_pdf_path)
    abs_output_dir = os.path.abspath(output_dir)
    
    if not os.path.exists(abs_pdf_path):
        raise FileNotFoundError(f"Die PDF-Datei wurde unter '{abs_pdf_path}' nicht gefunden.")
    
    env = os.environ.copy()
    inner_command = f"marker_single \"{abs_pdf_path}\" --output_dir \"{abs_output_dir}\""
    command = ["/bin/bash", "-i", "-c", inner_command]
    
    print("Führe OCR via Benutzer-Shell aus (das kann einen Moment dauern)...")
    try:
        subprocess.run(command, check=True, shell=True, env=env, stdout=subprocess.DEVNULL, stderr=sys.stderr)
        print("Marker erfolgreich ausgeführt.\n")
        
    except subprocess.CalledProcessError:
        print("\nStandard-CLI-Aufruf fehlgeschlagen. Versuche Modul-Fallback via Python-Interpreter...")
        fallback_command = [
            sys.executable, 
            "-m", "marker.scripts.convert_single", 
            str(abs_pdf_path), 
            "--output_dir", str(abs_output_dir)
        ]
        try:
            subprocess.run(fallback_command, check=True, env=env, stdout=subprocess.DEVNULL, stderr=sys.stderr)
            print("Marker über Modul-Fallback erfolgreich ausgeführt.\n")
        except subprocess.CalledProcessError as e:
            print("\n[FEHLER] Beide Marker-Aufrufe sind fehlgeschlagen.")
            raise e

    pdf_stem = Path(input_pdf_path).stem
    expected_md_path = Path(abs_output_dir) / pdf_stem / f"{pdf_stem}.md"
    
    if expected_md_path.exists():
        return str(expected_md_path)
    else:
        found_md_files = list(Path(abs_output_dir).glob("**/*.md"))
        if found_md_files:
            return str(found_md_files[0])
        raise FileNotFoundError("Marker hat den Prozess beendet, aber es wurde keine .md-Datei gefunden.")


def check_if_english(text: str) -> bool:
    """Schritt 2a: Prüfe, ob der Text englisch ist."""
    print("--- Schritt 2a: Prüfe Sprache des Dokuments ---")
    leseprobe = text[:2000]
    prompt = (
        "Antworte mit exakt einem Wort, entweder 'YES' oder 'NO'. "
        "Ist der folgende Text hauptsächlich in englischer Sprache verfasst?\n\n"
        f"Text:\n{leseprobe}"
    )
    try:
        response = call_gemini_with_retry(
            model_name='gemini-2.5-flash',
            contents=prompt,
            config=types.GenerateContentConfig(temperature=0.0, max_output_tokens=5)
        )
        return "YES" in response.text.strip().upper()
    except Exception as e:
        print(f"Sprachprüfung fehlgeschlagen ({e}), weiche auf Übersetzung aus.")
        return True


def split_text_by_headings(text: str, max_chars: int = 15000) -> list:
    """Hilfsfunktion: Splittet Markdown-Text an Überschriften in logische Abschnitte."""
    chunks = []
    current_chunk = []
    current_length = 0
    
    for line in text.split('\n'):
        if (line.startswith('# ') or line.startswith('## ')) and current_length > max_chars:
            chunks.append('\n'.join(current_chunk))
            current_chunk = []
            current_length = 0
            
        current_chunk.append(line)
        current_length += len(line)
        
    if current_chunk:
        chunks.append('\n'.join(current_chunk))
        
    return chunks


def translate_text(text: str) -> str:
    """Schritt 2b: Übersetzt den Text abschnittsweise nach den strengen Regeln aus 01_b_text_uebersetzen.txt."""
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
    
    chunks = split_text_by_headings(text)
    translated_chunks = []
    
    for i, chunk in enumerate(chunks, 1):
        print(f"   -> Übersetze Abschnitt {i} von {len(chunks)}...")
        try:
            response = call_gemini_with_retry(
                model_name='gemini-2.5-pro',
                contents=f"Text:\n{chunk}",
                config=types.GenerateContentConfig(system_instruction=system_prompt, temperature=0.1)
            )
            translated_chunks.append(response.text)
            time.sleep(2)
        except Exception as e:
            print(f"Fehler bei der Übersetzung von Abschnitt {i}: {e}")
            raise
            
    return '\n\n'.join(translated_chunks)


def generate_summary(text: str) -> str:
    """Schritt 4: Erstellt eine lernorientierte Zusammenfassung nach 02_prompts-zusammenfassung.txt."""
    print("--- Schritt 4: Erstelle lernorientierte Zusammenfassung (via Gemini 2.5 Pro) ---")
    
    prompt = (
        "Erstelle eine lernorientierte Zusammenfassung zum nachfolgenden Text, der nach 'Inhalt:' kommt.\n\n"
        "Anforderungen:\n"
        "Alle zentralen Konzepte enthalten\n"
        "Keine Beispiele entfernen, wenn sie zum Verständnis nötig sind\n"
        "Definitionen vollständig übernehmen\n"
        "Studienergebnisse erhalten\n"
        "Keine neuen Informationen ergänzen\n"
        "Struktur des Originals beibehalten (wichtig! Auch alle Unterkapitel, es darf keines fehlen! Die Gliederungsstruktur muss 100% erhalten bleiben)\n"
        "Möglichst kurz und stichpunktartig. Maximal 40 % der ursprünglichen Länge (wichtig!)\n"
        "Es darf aber nicht zu kurz sein, es muss alles vorhanden sein was in Prüfungsfragen dramkommen könnte (sehr wichtig!)\n"
        "Berücksichtige Abbildungen im Text und erläutere diese kurz.\n\n"
        "Prüfe:\n"
        "Welche Informationen aus dem Original in der Zusammenfassung fehlen\n"
        "Welche Definitionen verloren gingen\n"
        "Welche Einschränkungen oder Bedingungen fehlen\n\n"
        f"Inhalt:\n{text}"
    )
    try:
        response = call_gemini_with_retry(
            model_name='gemini-2.5-pro',
            contents=prompt,
            config=types.GenerateContentConfig(temperature=0.2)
        )
        return response.text
    except Exception as e:
        print(f"Fehler bei der Generierung der Zusammenfassung: {e}")
        raise


def verify_with_questions(summary_text: str, questions_path: str) -> str:
    """Schritt 5: Qualitätssicherung der Zusammenfassung anhand der Fragen aus 03_prompt_Fragen.txt."""
    print(f"--- Schritt 5: Qualitätssicherung via Leitfragen aus {questions_path} ---")
    
    with open(questions_path, "r", encoding="utf-8") as f:
        questions = f.read()
        
    prompt = (
        "Rolle:\nDu bist Lerncoach und Prüfer für Wirtschaftspsychologie.\n\n"
        "Aufgabe:\nBeantworte die leseleitenden Fragen ultrakompakt und mit exakt einem Unterkapitelverweis. "
        "Nutze ausschließlich den hochgeladenen Text als Wissensbasis.\n\n"
        "Bevor du antwortest:\n"
        "Schritt 1: Suche die relevanten Stellen im Dokument.\n"
        "Schritt 2: Liste die Textstellen stichpunktartig auf.\n"
        "Schritt 3: Erst danach beantworte die Frage.\n\n"
        "Wenn keine passende Stelle existiert:\n'Im Dokument nicht enthalten'. Nicht raten.\n\n"
        "Antwortregeln:\n"
        "- Maximal 3 Sätze pro Frage.\n"
        "- Keine Einleitung.\n"
        "- Keine Wiederholung der Frage.\n"
        "- Keine ausführlichen Erklärungen.\n"
        "- Nur prüfungsrelevante Kernaussage.\n"
        "- Wenn Zahlen/Studienwerte relevant sind: nennen.\n"
        "- Wenn die Antwort im Dokument nicht eindeutig steht: „Im Dokument nicht eindeutig beantwortbar.“\n\n"
        "Quellenregeln:\n"
        "- Verweise immer auf die genaueste vorhandene Überschrift.\n"
        "- Nicht nur „Kapitel 6.4“, sondern z. B. „6.4.1.2 Eine umfassende Übersicht“.\n"
        "- Wenn mehrere Unterkapitel nötig sind, maximal 3 nennen.\n"
        "- Zusätzlich 1–3 Schlüsselbegriffe aus der Textstelle nennen.\n"
        "- Keine groben Kapitelverweise, wenn Unterkapitel vorhanden sind.\n\n"
        "Ausgabeformat pro Frage:\n"
        "Frage X\n"
        "Antwort: [max. 3 Sätze]\n"
        "Textgrundlage: [genaues Unterkapitel]\n"
        "Schlüsselbegriffe: [1–3 Begriffe]\n"
        "Abdeckung: vollständig / teilweise / nicht enthalten\n\n"
        f"Wissensbasis (Zusammenfassung):\n{summary_text}\n\n"
        f"Fragen:\n{questions}"
    )
    try:
        response = call_gemini_with_retry(
            model_name='gemini-2.5-pro',
            contents=prompt,
            config=types.GenerateContentConfig(temperature=0.1)
        )
        return response.text
    except Exception as e:
        print(f"Fehler bei der Qualitätssicherung: {e}")
        raise


def add_formatted_text(paragraph, text, default_color=None):
    """Hilfsfunktion: Parse Markdown-Fettungen (**text**) und füge sie als Word-Runs hinzu."""
    parts = re.split(r'(\*\*.*?\*\*)', text)
    for part in parts:
        if part.startswith('**') and part.endswith('**'):
            run = paragraph.add_run(part[2:-2])
            run.bold = True
        else:
            run = paragraph.add_run(part)
        
        if default_color:
            run.font.color.rgb = default_color


def process_markdown_to_docx(doc, block_text, hide_text=False):
    """Interpretiert Markdown-Zeilen und fügt sie sauber formatiert dem Word-Dokument hinzu."""
    color_map = {
        'heading1': RGBColor(0x00, 0x33, 0x66), # Dunkelblau für Strukturen
        'heading2': RGBColor(0x00, 0x44, 0x88),
        'heading3': RGBColor(0x33, 0x66, 0x99),
        'hidden_text': RGBColor(0xCC, 0xCC, 0xCC), # Diskretes Hellgrau
        'hidden_heading1': RGBColor(0x99, 0x99, 0x99),
        'hidden_heading2': RGBColor(0xAA, 0xAA, 0xAA)
    }

    for line in block_text.split('\n'):
        stripped = line.strip()
        if not stripped:
            continue

        # 1. Überschriften übersetzen
        if stripped.startswith('# '):
            p = doc.add_heading(level=1)
            color = color_map['hidden_heading1'] if hide_text else color_map['heading1']
            add_formatted_text(p, stripped[2:], default_color=color)
        elif stripped.startswith('## '):
            p = doc.add_heading(level=2)
            color = color_map['hidden_heading2'] if hide_text else color_map['heading2']
            add_formatted_text(p, stripped[3:], default_color=color)
        elif stripped.startswith('### '):
            p = doc.add_heading(level=3)
            color = color_map['hidden_heading2'] if hide_text else color_map['heading3']
            add_formatted_text(p, stripped[4:], default_color=color)
            
        # 2. Aufzählungspunkte übersetzen
        elif stripped.startswith('* ') or stripped.startswith('- '):
            p = doc.add_paragraph(style='List Bullet')
            color = color_map['hidden_text'] if hide_text else None
            add_formatted_text(p, stripped[2:], default_color=color)
            if hide_text:
                p.style.font.size = Pt(9.5)
                
        # 3. Normaler Fließtext
        else:
            p = doc.add_paragraph()
            color = color_map['hidden_text'] if hide_text else None
            add_formatted_text(p, line, default_color=color)
            if hide_text:
                p.style.font.size = Pt(9.5)


def build_final_word_document(translated_text: str, summary_text: str, qa_text: str, output_path: str):
    """Schritt 3 erweitert: Erstellt das finale Word-Dokument ohne MD-Überreste."""
    print(f"--- Schritt 3/Final: Erstelle formatiertes Word-Dokument -> {output_path} ---")
    doc = Document()
    
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)
    
    # 1. QA / Quizfragen-Prüfung ganz oben anheften
    doc.add_heading("Qualitätsprüfung & Leitfragen-Abdeckung", level=1)
    process_markdown_to_docx(doc, qa_text, hide_text=False)
            
    doc.add_page_break()
    
    # 2. Lernorientierte Zusammenfassung
    doc.add_heading("Lernorientierte Zusammenfassung", level=1)
    process_markdown_to_docx(doc, summary_text, hide_text=False)
            
    doc.add_page_break()
    
    # 3. Übersetzter Originaltext (visuell "ausgeblendet" in Hellgrau)
    doc.add_heading("Vollständige Textgrundlage (Original/Übersetzung)", level=1)
    process_markdown_to_docx(doc, translated_text, hide_text=True)
            
    doc.save(output_path)
    print("Word-Dokument erfolgreich strukturiert und bereinigt.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="End-to-End PDF Translation & Learning Pipeline")
    parser.add_argument("pdf_path", type=str, help="Pfad zur Quell-PDF-Datei")
    parser.add_argument("--questions", type=str, default=None, help="Pfad zu den leseleitenden Fragen (optional)")
    
    args = parser.parse_args()
    OUTPUT_BASE = "workspace/output"
    
    try:
        if not os.getenv("GEMINI_API_KEY"):
            raise ValueError("GEMINI_API_KEY fehlt in der .env-Datei!")
            
        if not os.path.exists(args.pdf_path):
            raise FileNotFoundError(f"Die Datei {args.pdf_path} wurde nicht gefunden.")
            
        # 1. OCR ausführen
        md_file_path = run_marker_ocr(args.pdf_path, OUTPUT_BASE)
        
        with open(md_file_path, "r", encoding="utf-8") as f:
            working_text = f.read()
            
        # 2. Sprache prüfen & Übersetzen
        if check_if_english(working_text):
            print("Text ist Englisch. Starte Übersetzung...")
            working_text = translate_text(working_text)
            
            backup_md = Path(md_file_path).parent / "de_uebersetzung.md"
            with open(backup_md, "w", encoding="utf-8") as f:
                f.write(working_text)
        else:
            print("Text ist bereits Deutsch. Keine Übersetzung notwendig.")
            
        # 4. Zusammenfassung generieren
        summary_result = generate_summary(working_text)
        
        # 5. Optionale Qualitätssicherung über Fragen
        qa_result = "Keine Leitfragen zur Prüfung übergeben."
        if args.questions:
            if os.getenv("GEMINI_API_KEY") and os.path.exists(args.questions):
                qa_result = verify_with_questions(summary_result, args.questions)
            else:
                print(f"Warnung: Fragen-Datei '{args.questions}' nicht gefunden. Überspringe QS.")
                
        # 3. Word-Dokument zusammensetzen
        pdf_stem = Path(args.pdf_path).stem
        final_docx_path = Path(OUTPUT_BASE) / pdf_stem / f"{pdf_stem}_Lernskript.docx"
        
        build_final_word_document(working_text, summary_result, qa_result, str(final_docx_path))
        
        print(f"\n=== PIPELINE ERFOLGREICH BEENDET ===")
        print(f"Dein fertiges Dokument liegt hier: {final_docx_path}")
        
    except Exception as e:
        print(f"\nPipeline abgebrochen wegen: {e}")