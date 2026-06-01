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
    """Schritt 1: Konvertiert das PDF über den nativen Systemaufruf in Markdown."""
    print(f"--- Schritt 1: Starte Marker-OCR für {input_pdf_path} ---")
    os.makedirs(output_dir, exist_ok=True)
    
    # Nutzt exakt die Umgebungsvariablen deiner aktiven .venv
    env = os.environ.copy()
    
    # Wir übergeben den Befehl exakt so, wie er auf der CLI läuft
    command = f"marker_single \"{input_pdf_path}\" --output_dir \"{output_dir}\""[cite: 1]
    
    print("Führe OCR via System-Shell aus (das kann einen Moment dauern)...")
    try:
        # shell=True stellt sicher, dass WSL den Befehl 'marker_single' in der venv findet
        subprocess.run(command, check=True, shell=True, env=env, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        print("Marker erfolgreich ausgeführt.\n")
        
        pdf_stem = Path(input_pdf_path).stem
        expected_md_path = Path(output_dir) / pdf_stem / f"{pdf_stem}.md"
        
        if expected_md_path.exists():
            return str(expected_md_path)
        else:
            found_md_files = list(Path(output_dir).glob("**/*.md"))
            if found_md_files:
                return str(found_md_files[0])
            raise FileNotFoundError("Marker hat den Prozess beendet, aber es wurde keine .md-Datei gefunden.")
            
    except subprocess.CalledProcessError as e:
        print("\n[FEHLER] Marker-OCR fehlgeschlagen. Bitte prüfe, ob 'marker_single' manuell im Terminal startet.")
        raise e


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
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
            config=types.GenerateContentConfig(temperature=0.0, max_output_tokens=5)
        )
        return "YES" in response.text.strip().upper()
    except Exception as e:
        print(f"Sprachprüfung fehlgeschlagen ({e}), weiche auf Übersetzung aus.")
        return True


def translate_text(text: str) -> str:
    """Schritt 2b: Übersetzt den Text nach den strengen Regeln aus 01_b_text_uebersetzen.txt"""
    print("--- Schritt 2b: Übersetze englischen Text ins Deutsche (via Gemini 2.5 Pro) ---")[cite: 1]
    
    system_prompt = (
        "Übersetze den folgenden englischen wissenschaftlichen Text originalgetreu ins Deutsche.[cite: 2]\n\n"
        "Ziel:\n"
        "Eine vollständige, sinntreue Übersetzung, keine Zusammenfassung.[cite: 2]\n\n"
        "Strenge Regeln:\n"
        "- Nichts auslassen.[cite: 2]\n"
        "- Nichts ergänzen.[cite: 2]\n"
        "- Nichts interpretieren.[cite: 2]\n"
        "- Keine Inhalte glätten, kürzen oder zusammenfassen.[cite: 2]\n"
        "- Fachbegriffe konsistent übersetzen.[cite: 2]\n"
        "- Überschriften, Absatzstruktur, Listen und Tabellenstruktur beibehalten.[cite: 2]\n"
        "- Zitate, Autorennamen, Jahreszahlen, Variablennamen, Skalen, Hypothesen und statistische Angaben exakt erhalten.[cite: 2]\n"
        "- Unklare oder beschädigte Stellen mit [UNKLAR: Originalstelle] markieren, nicht erraten.[cite: 2]\n"
        "- Bildverweise, Tabellenverweise und Abbildungsbeschriftungen erhalten.[cite: 2]\n"
        "- Markdown-Struktur beibehalten.[cite: 2]\n\n"
        "Ausgabeformat:\n"
        "1. Nur die deutsche Übersetzung.[cite: 2]\n"
        "2. Danach eine kurze Kontrollliste:\n"
        "   - Anzahl erkannter Absätze im Original[cite: 2]\n"
        "   - Anzahl übersetzter Absätze[cite: 2]\n"
        "   - Hinweise auf unklare Stellen[cite: 2]\n"
        "   - Hinweise auf mögliche fehlende Tabellen/Bildinhalte[cite: 2]"
    )
    try:
        response = client.models.generate_content(
            model='gemini-2.5-pro',
            contents=f"Text:\n{text}",
            config=types.GenerateContentConfig(system_instruction=system_prompt, temperature=0.1)
        )
        return response.text
    except Exception as e:
        print(f"Fehler bei der Übersetzung: {e}")
        raise


def generate_summary(text: str) -> str:
    """Schritt 4: Erstellt eine lernorientierte Zusammenfassung nach 02_prompts-zusammenfassung.txt"""
    print("--- Schritt 4: Erstelle lernorientierte Zusammenfassung (via Gemini 2.5 Pro) ---")
    
    prompt = (
        "Erstelle eine lernorientierte Zusammenfassung zum nachfolgenden Text, der nach 'Inhalt:' kommt.[cite: 3]\n\n"
        "Anforderungen:\n"
        "Alle zentralen Konzepte enthalten[cite: 3]\n"
        "Keine Beispiele entfernen, wenn sie zum Verständnis nötig sind[cite: 3]\n"
        "Definitionen vollständig übernehmen[cite: 3]\n"
        "Studienergebnisse erhalten[cite: 3]\n"
        "Keine neuen Informationen ergänzen[cite: 3]\n"
        "Struktur des Originals beibehalten (wichtig! Auch alle Unterkapitel, es darf keines fehlen! Die Gliederungsstruktur muss 100% erhalten bleiben)[cite: 3]\n"
        "Möglichst kurz und stichpunktartig. Maximal 40 % der ursprünglichen Länge (wichtig!)[cite: 3]\n"
        "Es darf aber nicht zu kurz sein, es muss alles vorhanden sein was in Prüfungsfragen dramkommen könnte (sehr wichtig!)[cite: 3]\n"
        "Berücksichtige Abbildungen im Text und erläutere diese kurz.[cite: 3]\n\n"
        "Prüfe:\n"
        "Welche Informationen aus dem Original in der Zusammenfassung fehlen[cite: 3]\n"
        "Welche Definitionen verloren gingen[cite: 3]\n"
        "Welche Einschränkungen oder Bedingungen fehlen[cite: 3]\n\n"
        f"Inhalt:\n{text}"[cite: 3]
    )
    try:
        response = client.models.generate_content(
            model='gemini-2.5-pro',
            contents=prompt,
            config=types.GenerateContentConfig(temperature=0.2)
        )
        return response.text
    except Exception as e:
        print(f"Fehler bei der Generierung der Zusammenfassung: {e}")
        raise


def verify_with_questions(summary_text: str, questions_path: str) -> str:
    """Schritt 5: Qualitätssicherung der Zusammenfassung anhand der Fragen aus 03_prompt_Fragen.txt"""
    print(f"--- Schritt 5: Qualitätssicherung via Leitfragen aus {questions_path} ---")
    
    with open(questions_path, "r", encoding="utf-8") as f:
        questions = f.read()
        
    prompt = (
        "Rolle:\nDu bist Lerncoach und Prüfer für Wirtschaftspsychologie.\n\n"
        "Aufgabe:\nBeantworte die leseleitenden Fragen ultrakompakt und mit exakt einem Unterkapitelverweis. "
        "Nutze ausschließlich den hochgeladenen Text als Wissensbasis.[cite: 4]\n\n"
        "Bevor du antwortest:\n"
        "Schritt 1: Suche die relevanten Stellen im Dokument.[cite: 4]\n"
        "Schritt 2: Liste die Textstellen stichpunktartig auf.[cite: 4]\n"
        "Schritt 3: Erst danach beantworte die Frage.[cite: 4]\n\n"
        "Wenn keine passende Stelle existiert:\n'Im Dokument nicht enthalten'. Nicht raten.[cite: 4]\n\n"
        "Antwortregeln:\n"
        "- Maximal 3 Sätze pro Frage.[cite: 4]\n"
        "- Keine Einleitung.[cite: 4]\n"
        "- Keine Wiederholung der Frage.[cite: 4]\n"
        "- Keine ausführlichen Erklärungen.[cite: 4]\n"
        "- Nur prüfungsrelevante Kernaussage.[cite: 4]\n"
        "- Wenn Zahlen/Studienwerte relevant sind: nennen.[cite: 4]\n"
        "- Wenn die Antwort im Dokument nicht eindeutig steht: „Im Dokument nicht eindeutig beantwortbar.“[cite: 4]\n\n"
        "Quellenregeln:\n"
        "- Verweise immer auf die genaueste vorhandene Überschrift.[cite: 4]\n"
        "- Nicht nur „Kapitel 6.4“, sondern z. B. „6.4.1.2 Eine umfassende Übersicht“.[cite: 4]\n"
        "- Wenn mehrere Unterkapitel nötig sind, maximal 3 nennen.[cite: 4]\n"
        "- Zusätzlich 1–3 Schlüsselbegriffe aus der Textstelle nennen.[cite: 4]\n"
        "- Keine groben Kapitelverweise, wenn Unterkapitel vorhanden sind.[cite: 4]\n\n"
        "Ausgabeformat pro Frage:\n"
        "Frage X\n"
        "Antwort: [max. 3 Sätze]\n"
        "Textgrundlage: [genaues Unterkapitel]\n"
        "Schlüsselbegriffe: [1–3 Begriffe]\n"
        "Abdeckung: vollständig / teilweise / nicht enthalten[cite: 4]\n\n"
        f"Wissensbasis (Zusammenfassung):\n{summary_text}\n\n"
        f"Fragen:\n{questions}"[cite: 4]
    )
    try:
        response = client.models.generate_content(
            model='gemini-2.5-pro',
            contents=prompt,
            config=types.GenerateContentConfig(temperature=0.1)
        )
        return response.text
    except Exception as e:
        print(f"Fehler bei der Qualitätssicherung: {e}")
        raise


def build_final_word_document(translated_text: str, summary_text: str, qa_text: str, output_path: str):
    """Schritt 3 erweitert: Erstellt das Word-Dokument mit ausgeblendetem Original, Zusammenfassung und QA."""
    print(f"--- Schritt 3/Final: Erstelle finalisiertes Word-Dokument -> {output_path} ---")
    doc = Document()
    
    # Globaler Style
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)
    
    # 1. QA / Quizfragen-Prüfung ganz oben als Qualitäts-Zertifikat anheften
    doc.add_heading("Qualitätsprüfung & Leitfragen-Abdeckung", level=1)
    for line in qa_text.split("\n"):
        if line.strip():
            doc.add_paragraph(line)
            
    doc.add_page_break()
    
    # 2. Lernorientierte Zusammenfassung
    doc.add_heading("Lernorientierte Zusammenfassung", level=1)[cite: 3]
    for line in summary_text.split("\n"):
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.startswith('# '):
            doc.add_heading(stripped[2:], level=2)
        elif stripped.startswith('## '):
            doc.add_heading(stripped[3:], level=3)
        else:
            doc.add_paragraph(line)
            
    doc.add_page_break()
    
    # 3. Übersetzter Originaltext (visuell "ausgeblendet" in Hellgrau)
    doc.add_heading("Vollständige Textgrundlage (Original/Übersetzung)", level=1)
    for line in translated_text.split("\n"):
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.startswith('# '):
            p = doc.add_heading(stripped[2:], level=2)
            p.style.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
        elif stripped.startswith('## '):
            p = doc.add_heading(stripped[3:], level=3)
            p.style.font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)
        else:
            p = doc.add_paragraph()
            run = p.add_run(line)
            run.font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC) # Sehr helles Grau zum "Ausblenden"
            run.font.size = Pt(9.5)
            
    doc.save(output_path)
    print("Word-Dokument erfolgreich finalisiert.")


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
            
        # 1. OCR mit Marker über Shell-Ebenen-Passthrough
        md_file_path = run_marker_ocr(args.pdf_path, OUTPUT_BASE)
        
        with open(md_file_path, "r", encoding="utf-8") as f:
            working_text = f.read()
            
        # 2. Sprache prüfen & Übersetzen[cite: 1]
        if check_if_english(working_text):
            print("Text ist Englisch. Starte Übersetzung...")[cite: 1]
            working_text = translate_text(working_text)[cite: 1]
            
            # Backup-MD sichern
            backup_md = Path(md_file_path).parent / "de_uebersetzung.md"
            with open(backup_md, "w", encoding="utf-8") as f:
                f.write(working_text)
        else:
            print("Text ist bereits Deutsch. Keine Übersetzung notwendig.")[cite: 1]
            
        # 4. Zusammenfassung generieren
        summary_result = generate_summary(working_text)
        
        # 5. Optionale Qualitätssicherung über Fragen
        qa_result = "Keine Leitfragen zur Prüfung übergeben."
        if args.questions:
            if os.path.exists(args.questions):
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