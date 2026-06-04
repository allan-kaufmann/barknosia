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
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Heading-Farben aus MM-Skript MM_36633_AuG_SS2026
MM_HEADING_COLORS = {
    1: RGBColor(0xC4, 0x9A, 0x00),  # H1: dunkles Gold
    2: RGBColor(0xFF, 0xCA, 0x08),  # H2: helles Gold
    3: RGBColor(0xFF, 0xCA, 0x08),  # H3: helles Gold
    4: RGBColor(0x82, 0x66, 0x00),  # H4: dunkles Amber
    5: RGBColor(0x82, 0x66, 0x00),  # H5: dunkles Amber
    6: RGBColor(0x82, 0x66, 0x00),  # H6: dunkles Amber
    7: RGBColor(0x82, 0x66, 0x00),  # H7: dunkles Amber
    8: RGBColor(0x27, 0x27, 0x27),  # H8: fast Schwarz
    9: RGBColor(0x27, 0x27, 0x27),  # H9: fast Schwarz
}

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


def load_or_run(path: Path, generator_fn, label: str) -> str:
    """Resume-Hilfsfunktion: Lädt Ergebnis aus Datei wenn vorhanden, sonst generator_fn() ausführen und speichern."""
    if path.exists():
        print(f"[SKIP] {label} – bereits vorhanden: {path}")
        return path.read_text(encoding="utf-8")
    print(f"[RUN]  {label} – starte Berechnung...")
    result = generator_fn()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(result, encoding="utf-8")
    print(f"       Ergebnis gespeichert: {path}")
    return result


def run_marker_ocr(input_pdf_path: str, output_dir: str) -> str:
    """Schritt 1: Konvertiert das PDF über den isolierten venv-Aufruf in Markdown."""
    print(f"--- Schritt 1: Starte Marker-OCR für {input_pdf_path} ---")
    os.makedirs(output_dir, exist_ok=True)

    abs_pdf_path = os.path.abspath(input_pdf_path)
    abs_output_dir = os.path.abspath(output_dir)

    if not os.path.exists(abs_pdf_path):
        raise FileNotFoundError(f"Die PDF-Datei wurde unter '{abs_pdf_path}' nicht gefunden.")

    venv_bin_dir = Path(sys.executable).parent
    local_marker_executable = venv_bin_dir / "marker_single"

    if local_marker_executable.exists():
        command = [str(local_marker_executable), str(abs_pdf_path), "--output_dir", str(abs_output_dir)]
        print(f"Nutze isolierten venv-Marker: {local_marker_executable}")
    else:
        command = ["marker_single", str(abs_pdf_path), "--output_dir", str(abs_output_dir)]
        print("Nutze Standard-Pfad für marker_single...")

    env = os.environ.copy()

    print("Führe OCR aus (das kann einen Moment dauern)...")
    try:
        subprocess.run(command, check=True, env=env, stdout=subprocess.DEVNULL, stderr=sys.stderr, shell=False)
        print("Marker erfolgreich ausgeführt.\n")
    except subprocess.CalledProcessError as e:
        print("\n[FEHLER] Marker-OCR fehlgeschlagen.")
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
    """Schritt 2b: Übersetzt den Text abschnittsweise nach den strengen Regeln."""
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


def split_into_level1_chapters(text: str) -> list:
    """
    Splittet den (bereits normalisierten) Markdown-Text an nummerierten ## Kapitelüberschriften.
    Jeder Eintrag: {'heading': str, 'full_text': str}.
    Erkennt Zeilen wie '## **1 EINLEITUNG**' oder '## 4 INDIVIDUELLE FAKTOREN'.
    """
    chapters = []
    current_heading = None
    current_lines = []

    for line in text.split('\n'):
        # Nummeriertes ## -Kapitel: "## **1 ..." oder "## 1 ..."
        m = re.match(r'^##\s+(?:\*\*)?(\d+)\s', line)
        if m:
            if current_heading is not None:
                chapters.append({'heading': current_heading, 'full_text': '\n'.join(current_lines)})
            # Heading-Text ohne Bold-Marker speichern
            current_heading = line[3:].strip().replace('**', '').strip()
            current_lines = [line]
        else:
            current_lines.append(line)

    if current_heading is not None:
        chapters.append({'heading': current_heading, 'full_text': '\n'.join(current_lines)})

    return chapters


def _summarize_single_chapter(heading: str, chapter_text: str) -> str:
    """Erstellt eine lernorientierte Zusammenfassung für ein einzelnes Kapitel."""
    prompt = (
        f"Erstelle eine lernorientierte Zusammenfassung für das folgende Kapitel: \"{heading}\"\n\n"
        "Pflichtanforderungen:\n"
        "1. ALLE Unterkapitel müssen vorhanden sein – kein einziges Unterkapitel darf fehlen!\n"
        "   Behalte die genauen Überschriften inkl. Nummerierung bei (z.B. '4.1.1 Hedonisches Wohlbefinden').\n"
        "2. Pro Unterkapitel: mindestens 3–5 Stichpunkte mit den wichtigsten Inhalten.\n"
        "3. Studienergebnisse IMMER erhalten: Metaanalysen, Effektstärken, Befundrichtung, Autoren & Jahr.\n"
        "4. Definitionen: wörtlich oder sehr nah am Original übernehmen.\n"
        "5. Keine neuen Informationen ergänzen.\n"
        "6. Abbildungen und Tabellen kurz erwähnen und ihren Inhalt beschreiben.\n"
        "7. Stichpunkte statt Fließtext (Ausnahme: Definitionen).\n"
        "8. Länge: maximal 40–50 % des Originals – ABER vollständige Unterkapitelabdeckung hat Vorrang vor Kürze.\n\n"
        "Selbstprüfung (am Ende anhängen):\n"
        "- Liste alle Unterkapitel des Originals auf\n"
        "- Markiere fehlende Unterkapitel oder fehlende Studienergebnisse\n\n"
        f"Kapiteltext:\n{chapter_text}"
    )
    try:
        response = call_gemini_with_retry(
            model_name='gemini-2.5-pro',
            contents=prompt,
            config=types.GenerateContentConfig(temperature=0.2)
        )
        return response.text
    except Exception as e:
        print(f"Fehler bei Zusammenfassung von '{heading}': {e}")
        raise


def generate_summary_by_chapter(text: str, out_dir: Path) -> str:
    """
    Schritt 4: Erstellt die Zusammenfassung kapitelweise (je Level-1-Kapitel ein API-Call).
    Nutzt Caching pro Kapitel: zusammenfassung_kap_XX.md in out_dir.
    Kombiniert am Ende zu zusammenfassung.md.
    """
    print("--- Schritt 4: Erstelle kapitelweise Zusammenfassung (via Gemini 2.5 Pro) ---")

    text = normalize_heading_levels(text)
    chapters = split_into_level1_chapters(text)

    if not chapters:
        print("   Keine Level-1-Kapitel gefunden, fasse Gesamttext zusammen...")
        fallback_path = out_dir / "zusammenfassung_kap_00.md"
        result = load_or_run(
            fallback_path,
            lambda: _summarize_single_chapter("Volltext", text),
            "Zusammenfassung (Volltext)"
        )
        (out_dir / "zusammenfassung.md").write_text(result, encoding="utf-8")
        return result

    summaries = []
    for i, chapter in enumerate(chapters, 1):
        cache_path = out_dir / f"zusammenfassung_kap_{i:02d}.md"
        label = f"Zusammenfassung Kap. {i}: {chapter['heading'][:60]}"
        chapter_summary = load_or_run(
            cache_path,
            lambda ch=chapter: _summarize_single_chapter(ch['heading'], ch['full_text']),
            label
        )
        summaries.append(chapter_summary)
        time.sleep(1)  # kurze Pause zwischen API-Calls

    combined = "\n\n".join(summaries)
    (out_dir / "zusammenfassung.md").write_text(combined, encoding="utf-8")
    print(f"   Zusammenfassung aus {len(chapters)} Kapiteln kombiniert.")
    return combined


def verify_with_questions(summary_text: str, questions_path: str) -> str:
    """Schritt 5: Qualitätssicherung der Zusammenfassung anhand von Leitfragen."""
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
        "- Zusätzlich 1–3 Schlüsselbegriffe aus die Textstelle nennen.\n"
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


# ---------------------------------------------------------------------------
# Markdown-Parsing & Word-Dokument-Aufbau
# ---------------------------------------------------------------------------

def parse_sections(text: str) -> list:
    """
    Zerlegt Markdown-Text in eine flache Liste von Abschnitten.
    Jeder Eintrag: {'level': int, 'heading': str, 'body': str}
    Abschnitte vor der ersten Überschrift haben level=0, heading='__preamble__'.
    """
    sections = []
    current = {'level': 0, 'heading': '__preamble__', 'lines': []}

    for line in text.split('\n'):
        m = re.match(r'^(#{1,6})\s+(.+)$', line)
        if m:
            body = '\n'.join(current['lines']).strip()
            if body or current['heading'] != '__preamble__':
                sections.append({
                    'level': current['level'],
                    'heading': current['heading'],
                    'body': body
                })
            current = {'level': len(m.group(1)), 'heading': m.group(2).strip(), 'lines': []}
        else:
            current['lines'].append(line)

    body = '\n'.join(current['lines']).strip()
    if body or current['heading'] != '__preamble__':
        sections.append({
            'level': current['level'],
            'heading': current['heading'],
            'body': body
        })

    return sections


def normalize_heading(h: str) -> str:
    """Für Matching: Bold-/Italic-Marker entfernen, lowercase. Nummern bleiben für eindeutige Keys."""
    h = re.sub(r'\*+', '', h)   # strip ** und *
    return h.lower().strip()


def normalize_heading_levels(text: str) -> str:
    """
    Normalisiert inkonsistente Markdown-Überschriftenebenen.
    Nummerierte Kapitelüberschriften erhalten konsistente Ebenen:
      "1 Titel"     → ## (H2)
      "4.1 Titel"   → ### (H3)
      "4.1.1 Titel" → #### (H4)
    Nicht-nummerierte Überschriften werden nicht verändert.
    """
    result = []
    for line in text.split('\n'):
        m = re.match(r'^(#{1,6})\s+(.+)$', line)
        if m:
            content = m.group(2).strip()
            clean = content.replace('**', '').strip()
            num_m = re.match(r'^(\d+)(\.(\d+)(\.(\d+))?)?(\s|$)', clean)
            if num_m:
                if num_m.group(3) is None:
                    new_level = '##'
                elif num_m.group(5) is None:
                    new_level = '###'
                else:
                    new_level = '####'
                result.append(f'{new_level} {content}')
            else:
                result.append(line)
        else:
            result.append(line)
    return '\n'.join(result)


def add_formatted_text(paragraph, text, default_color=None):
    """Parse Markdown-Fettungen (**text**) und füge sie als Word-Runs hinzu."""
    parts = re.split(r'(\*\*.*?\*\*)', text)
    for part in parts:
        if part.startswith('**') and part.endswith('**'):
            run = paragraph.add_run(part[2:-2])
            run.bold = True
        else:
            run = paragraph.add_run(part)

        if default_color:
            run.font.color.rgb = default_color


def _clean_unklar_cell(text: str) -> str:
    """
    Verbessert unleserliche OCR-Zellen: [UNKLAR: N<br>fo<br>k<br>t...]
    Fragmente werden zusammengefügt. Kurze Fragmente (≤2 Zeichen) werden
    direkt verbunden (Zeichen-OCR), längere mit Leerzeichen getrennt.
    Ergebnis: kompakter [OCR: ...]-Hinweis statt seitenlanger Garbage-Spalte.
    """
    def rebuild(m):
        inner = m.group(1).strip()
        parts = [p.strip() for p in re.split(r'<br\s*/?>', inner) if p.strip()]
        if not parts:
            return '[unleserlich]'
        avg_len = sum(len(p) for p in parts) / len(parts)
        joined = ''.join(parts) if avg_len <= 2.5 else ' '.join(parts)
        truncated = joined[:80] + ('…' if len(joined) > 80 else '')
        return f'[OCR: {truncated}]'
    return re.sub(r'\[UNKLAR:\s*(.*?)\]', rebuild, text, flags=re.DOTALL)


def _html_entities(text: str) -> str:
    """Wandelt einfache HTML-Entities in Plaintext um."""
    return (text
            .replace('&amp;', '&')
            .replace('&lt;', '<')
            .replace('&gt;', '>')
            .replace('&nbsp;', ' ')
            .replace('&#x27;', "'")
            .replace('&quot;', '"'))


def _set_cell_text(cell, raw_text: str, bold: bool = False, font_size: Pt = None):
    """Schreibt Text in eine Word-Tabellenzelle.
    Kopfzeilen (bold=True): <br> → Leerzeichen (kein Umbruch).
    Datenzellen: <br> → eigener Absatz je Teil.
    """
    font_size = font_size or Pt(9.5)
    raw_text = _clean_unklar_cell(raw_text)
    raw_text = _html_entities(raw_text)

    if bold:
        # Kopfzeile: alle Umbrüche als Leerzeichen, einzeiliger Text
        text = raw_text.replace('<br>', ' ').replace('\n', ' ').strip()
        para = cell.paragraphs[0]
        para.clear()
        run = para.add_run(text)
        run.bold = True
        run.font.size = font_size
    else:
        raw_text = raw_text.replace('<br>', '\n')
        parts = raw_text.split('\n')
        para = cell.paragraphs[0]
        para.clear()
        run = para.add_run(parts[0].strip())
        run.font.size = font_size
        for part in parts[1:]:
            para = cell.add_paragraph()
            run = para.add_run(part.strip())
            run.font.size = font_size


def _calc_col_widths(rows_data: list, num_cols: int) -> list:
    """
    Berechnet Spaltenbreiten in cm proportional zum längsten Zellinhalt je Spalte.
    Gesamtbreite: 16 cm (A4 minus Standardränder). Mindestbreite: 1,2 cm.
    """
    PAGE_CM = 16.0
    MIN_CM  = 1.2
    max_lens = []
    for c in range(num_cols):
        ml = max(
            (len(row[c].replace('<br>', ' ').replace('\n', ' ')) for row in rows_data if c < len(row)),
            default=1
        )
        max_lens.append(max(ml, 1))
    total = sum(max_lens)
    raw = [max(PAGE_CM * ml / total, MIN_CM) for ml in max_lens]
    scale = PAGE_CM / sum(raw)
    return [w * scale for w in raw]


def _shade_cell(cell, fill_hex: str):
    """Setzt Hintergrundfarbe einer Tabellenzelle (z.B. 'DDEEFF')."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)


def add_markdown_table_to_doc(doc, table_lines: list):
    """
    Konvertiert eine Liste von Markdown-Pipe-Zeilen in eine Word-Tabelle.
    Leere Zellen in derselben Spalte werden vertikal mit der Zelle darüber zusammengeführt.
    Tabellen werden nie ausgeblendet.
    """
    SEP_RE = re.compile(r'^\|[-|: ]+\|$')

    rows_data = []
    for raw_line in table_lines:
        line = raw_line.strip()
        if SEP_RE.match(line):
            continue
        cells = [c.strip() for c in line.split('|')]
        if cells and cells[0] == '':
            cells = cells[1:]
        if cells and cells[-1] == '':
            cells = cells[:-1]
        rows_data.append(cells)

    if not rows_data:
        return

    num_cols = max(len(r) for r in rows_data)
    num_rows = len(rows_data)

    # Zellinhalt normieren (alle Zeilen auf gleiche Spaltenanzahl)
    for r in rows_data:
        while len(r) < num_cols:
            r.append('')

    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.style = 'Table Grid'

    # Spaltenbreiten proportional zum Inhalt setzen
    col_widths_cm = _calc_col_widths(rows_data, num_cols)
    font_size = Pt(8) if num_cols > 5 else Pt(9.5)
    for ci, w_cm in enumerate(col_widths_cm):
        for cell in table.column_cells(ci):
            cell.width = Cm(w_cm)

    # Zellen befüllen
    for r_idx, row in enumerate(rows_data):
        for c_idx, cell_text in enumerate(row):
            cell = table.cell(r_idx, c_idx)
            is_header = (r_idx == 0)
            _set_cell_text(cell, cell_text, bold=is_header, font_size=font_size)
            if is_header:
                _shade_cell(cell, 'D9E2F3')

    # Leere Zellen vertikal zusammenführen (Spanning-Simulation)
    for c in range(num_cols):
        r = 1
        while r < num_rows:
            if rows_data[r][c] == '':
                try:
                    table.cell(r - 1, c).merge(table.cell(r, c))
                except Exception:
                    pass
            r += 1

    doc.add_paragraph()  # Abstand nach Tabelle


_CAPTION_RE = re.compile(
    r'^(\*\*)?(TABELLE|Tabelle|ABBILDUNG|Abbildung|TABLE|FIGURE|Figure|Abb\.|Tab\.)\b',
    re.IGNORECASE
)
_HRULE_RE = re.compile(r'^-{3,}$|^\*{3,}$|^_{3,}$')


def _hide_paragraph(p, indent_cm: float = 1.5):
    """Setzt alle Runs eines Absatzes auf hidden + fügt Einrückung hinzu."""
    p.paragraph_format.left_indent = Cm(indent_cm)
    for run in p.runs:
        run.font.hidden = True


def process_markdown_to_docx(doc, block_text, hide_text=False, base_path=None):
    """
    Interpretiert Markdown und fügt Inhalte dem Word-Dokument hinzu.
    - Tabellen und Bilder: IMMER sichtbar (nie ausgeblendet).
    - Abbildungs-/Tabellenbeschriftungen: sichtbar, auch wenn hide_text=True.
    - hide_text=True: Text wird hidden (nicht druckbar) + eingerückt; grau für Sichtbarkeit.
    - Horizontale Linien (---) werden übersprungen.
    """
    color_map = {
        'heading1': MM_HEADING_COLORS[1],
        'heading2': MM_HEADING_COLORS[2],
        'heading3': MM_HEADING_COLORS[3],
        'heading4': MM_HEADING_COLORS[4],
        'hidden_text':    RGBColor(0xBB, 0xBB, 0xBB),
        'hidden_heading': RGBColor(0x99, 0x99, 0x99),
    }

    lines = block_text.split('\n')
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # ── Horizontale Linie überspringen (--- / *** / ___) ──
        if _HRULE_RE.match(stripped):
            i += 1
            continue

        # ── Tabelle: immer sichtbar ──
        if stripped.startswith('|'):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                table_lines.append(lines[i].strip())
                i += 1
            add_markdown_table_to_doc(doc, table_lines)
            continue

        # ── Bild: immer sichtbar ──
        img_m = re.match(r'^!\[([^\]]*)\]\(([^)]+)\)$', stripped)
        if img_m:
            img_rel = img_m.group(2)
            img_path = os.path.join(base_path, img_rel) if base_path else img_rel
            if os.path.exists(img_path):
                try:
                    doc.add_picture(img_path, width=Inches(5.5))
                    doc.add_paragraph()
                except Exception:
                    doc.add_paragraph(f"[Bild nicht eingebettet: {img_rel}]")
            else:
                doc.add_paragraph(f"[Bild nicht gefunden: {img_rel}]")
            i += 1
            continue

        if not stripped:
            i += 1
            continue

        # Abbildungs-/Tabellenbeschriftungen nie ausblenden
        is_caption = bool(_CAPTION_RE.match(stripped))
        do_hide = hide_text and not is_caption

        # ── Überschriften ──
        if stripped.startswith('#### '):
            p = doc.add_heading(level=4)
            color = color_map['hidden_heading'] if do_hide else color_map['heading4']
            add_formatted_text(p, stripped[5:], default_color=color)
            if do_hide:
                _hide_paragraph(p)
        elif stripped.startswith('### '):
            p = doc.add_heading(level=3)
            color = color_map['hidden_heading'] if do_hide else color_map['heading3']
            add_formatted_text(p, stripped[4:], default_color=color)
            if do_hide:
                _hide_paragraph(p)
        elif stripped.startswith('## '):
            p = doc.add_heading(level=2)
            color = color_map['hidden_heading'] if do_hide else color_map['heading2']
            add_formatted_text(p, stripped[3:], default_color=color)
            if do_hide:
                _hide_paragraph(p)
        elif stripped.startswith('# '):
            p = doc.add_heading(level=1)
            color = color_map['hidden_heading'] if do_hide else color_map['heading1']
            add_formatted_text(p, stripped[2:], default_color=color)
            if do_hide:
                _hide_paragraph(p)
        # ── Aufzählung ──
        elif stripped.startswith('* ') or stripped.startswith('- '):
            p = doc.add_paragraph(style='List Bullet')
            color = color_map['hidden_text'] if do_hide else None
            add_formatted_text(p, stripped[2:], default_color=color)
            if do_hide:
                _hide_paragraph(p)
        # ── Normaler Text ──
        else:
            p = doc.add_paragraph()
            color = color_map['hidden_text'] if do_hide else None
            add_formatted_text(p, line, default_color=color)
            if do_hide:
                _hide_paragraph(p)
        i += 1


def _set_heading_color(heading_paragraph, color: RGBColor):
    """Setzt die Schriftfarbe aller Runs einer Überschrift."""
    for run in heading_paragraph.runs:
        run.font.color.rgb = color


# Überschriften, die standardmäßig übersprungen werden (Referenzen etc.)
_SKIP_HEADINGS = frozenset([
    'referenzen', 'literaturverzeichnis', 'references', 'bibliography',
    'literatur', 'quellen', 'quellenverzeichnis', 'endnoten', 'endnotes',
    'orcid', 'interessenkonflikt', 'conflict of interest',
    'erklärung zur datenverfügbarkeit', 'data availability statement',
    'danksagung', 'acknowledgments', 'kontrollliste',
])

def _is_skip_heading(heading: str) -> bool:
    """Gibt True zurück, wenn diese Überschrift standardmäßig übersprungen werden soll."""
    key = normalize_heading(heading)
    return key in _SKIP_HEADINGS


def _advance_counter(counters: list, level: int) -> str:
    """
    Erhöht den Zähler für `level` (1-basiert) und setzt tiefere Ebenen zurück.
    Gibt die kompakte Nummerierung zurück – Null-Ebenen werden übersprungen,
    damit Ebenensprünge (z.B. level 2 → level 4) keine '0.0'-Segmente erzeugen.
    Beispiel: counters=[4,0,0,1] → '4.1' statt '4.0.0.1'
    """
    counters[level - 1] += 1
    for i in range(level, len(counters)):
        counters[i] = 0
    return '.'.join(str(counters[i]) for i in range(level) if counters[i] > 0)


def _parent_level_from_chapter(parent_chapter: str) -> int:
    """
    Leitet den Heading-Level des Elternkapitels aus der Kapitelnummer ab.
    '4.2.1.2' → 4 Punkte-getrennte Teile → Heading 5
    Annahme: Top-Level-Kapitel (z.B. '1') sind Heading 2, da Heading 1 der Dokumenttitel ist.
    """
    return len(parent_chapter.split('.')) + 1


def _prefix_chapter_number(heading_text: str, prefix: str) -> str:
    """
    Setzt den Elternpräfix vor die Kapitelnummer einer Überschrift.
    '1 EINLEITUNG'  + '4.2.1.2' → '4.2.1.2.1 EINLEITUNG'
    '4.1 Abschnitt' + '4.2.1.2' → '4.2.1.2.4.1 Abschnitt'
    Überschriften ohne führende Zahl bleiben unverändert.
    """
    m = re.match(r'^(\d[\d\.]*)(.*)', heading_text.strip())
    if m:
        return f"{prefix}.{m.group(1)}{m.group(2)}"
    return heading_text


def build_translation_word_document(translated_text: str, output_path: str, base_path: str = None):
    """Zwischenschritt: Erstellt ein einfaches Word-Dokument aus dem übersetzten Text (kein Grau, keine Zusammenfassung)."""
    print(f"--- Erstelle Übersetzungs-Word-Dokument -> {output_path} ---")
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)
    process_markdown_to_docx(doc, normalize_heading_levels(translated_text), hide_text=False, base_path=base_path)
    doc.save(output_path)
    print("Übersetzungs-Word-Dokument erstellt.")


def build_interleaved_word_document(translated_text: str, summary_text: str, qa_text: str,
                                    output_path: str, base_path: str = None,
                                    parent_chapter: str = None, parent_level: int = None,
                                    skip_references: bool = True):
    """
    Erstellt ein Word-Dokument, bei dem Zusammenfassung und Originaltext
    kapitelweise verschränkt sind.

    parent_chapter: Elternkapitel im Zieldokument, z.B. '4.2.1.2'.
      Wenn angegeben, werden alle Kapitelnummern mit diesem Präfix versehen
      ('1 EINLEITUNG' → '4.2.1.2.1 EINLEITUNG') und Heading-Level entsprechend
      verschoben. 'Lernskript'-Titel und QA-Block werden dann weggelassen.
    parent_level:   Heading-Level des Elternkapitels (Standard: aus parent_chapter
      automatisch berechnet, z.B. '4.2.1.2' → 5).
    """
    print(f"--- Erstelle interleaved Word-Dokument -> {output_path} ---")

    # Elternkapitel-Logik
    if parent_chapter:
        lvl_shift = (parent_level if parent_level else _parent_level_from_chapter(parent_chapter))
        print(f"    Einfügemodus: Präfix '{parent_chapter}', Heading-Shift +{lvl_shift}")
    else:
        lvl_shift = 0

    counters = [0] * 9  # Zähler je Heading-Ebene für Auto-Nummerierung

    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)

    # --- QA-Block und Lernskript-Titel (nur im Standalone-Modus) ---
    if not parent_chapter:
        if qa_text and qa_text.strip() != "Keine Leitfragen zur Prüfung übergeben.":
            h = doc.add_heading("Qualitätsprüfung & Leitfragen-Abdeckung", level=1)
            _set_heading_color(h, MM_HEADING_COLORS[1])
            process_markdown_to_docx(doc, qa_text, hide_text=False, base_path=base_path)
            doc.add_page_break()
        h = doc.add_heading("Lernskript", level=1)
        _set_heading_color(h, MM_HEADING_COLORS[1])

    # --- Heading-Level normalisieren & Sections parsen ---
    translated_text = normalize_heading_levels(translated_text)
    summary_text    = normalize_heading_levels(summary_text)
    orig_sections   = parse_sections(translated_text)
    sum_sections    = parse_sections(summary_text)

    # Lookup: normalisierter Heading → Zusammenfassungstext
    sum_lookup = {}
    for s in sum_sections:
        if s['heading'] == '__preamble__':
            continue
        sum_lookup[normalize_heading(s['heading'])] = s['body']

    # --- Interleaved Aufbau ---
    for idx, section in enumerate(orig_sections):
        if section['heading'] == '__preamble__':
            if section['body']:
                process_markdown_to_docx(doc, section['body'], base_path=base_path)
            continue

        level     = section['level']
        heading   = section['heading']
        orig_body = section['body']

        # Überschriftentext: Markdown-Marker entfernen
        clean_heading = re.sub(r'\*+', '', heading).strip()

        # Referenzen und interne Sections ggf. überspringen
        if skip_references and _is_skip_heading(clean_heading):
            continue

        # Kapitelnummer vergeben (Einbettungsmodus)
        if parent_chapter:
            if re.match(r'^\d', clean_heading):
                clean_heading = _prefix_chapter_number(clean_heading, parent_chapter)
            else:
                local_num = _advance_counter(counters, level)
                clean_heading = f"{parent_chapter}.{local_num} {clean_heading}"

        # Heading-Level: aus der Kapitelnummer ableiten (korrekte Einrückung)
        if parent_chapter:
            num_m = re.match(r'^([\d.]+)\b', clean_heading)
            display_level = min(len(num_m.group(1).split('.')) + 1, 9) if num_m else min(level + lvl_shift, 9)
        else:
            display_level = level  # Standalone: originale Ebene

        # Zusammenfassung für dieses Kapitel
        sum_body = sum_lookup.get(normalize_heading(heading), '')

        # Leere Blatt-Kapitel überspringen (kein Inhalt, keine Unterkapitel)
        has_children = (
            idx + 1 < len(orig_sections) and
            orig_sections[idx + 1]['heading'] != '__preamble__' and
            orig_sections[idx + 1]['level'] > level
        )
        if not sum_body.strip() and not orig_body.strip() and not has_children:
            continue

        # Kapitelüberschrift
        h = doc.add_heading(clean_heading, level=display_level)
        _set_heading_color(h, MM_HEADING_COLORS.get(display_level, MM_HEADING_COLORS[9]))

        # Zusammenfassungstext (ohne Label)
        if sum_body.strip():
            process_markdown_to_docx(doc, sum_body, hide_text=False, base_path=base_path)

        # Originaltext: hidden (nicht druckbar) + eingerückt, kein Gliederungspunkt
        if orig_body.strip():
            process_markdown_to_docx(doc, orig_body, hide_text=True, base_path=base_path)

    doc.save(output_path)
    print("Word-Dokument erfolgreich erstellt.")


# ---------------------------------------------------------------------------
# Hauptprogramm
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="End-to-End PDF Translation & Learning Pipeline")
    parser.add_argument("pdf_path", type=str, help="Pfad zur Quell-PDF-Datei")
    parser.add_argument("--questions", type=str, default=None, help="Pfad zu den leseleitenden Fragen (optional)")
    parser.add_argument("--force", action="store_true", help="Alle Schritte neu berechnen (kein Resume)")
    parser.add_argument("--parent-chapter", type=str, default=None,
                        help="Elternkapitel im Zieldokument, z.B. '4.2.1.2'. "
                             "Präfixiert alle Kapitelnummern und verschiebt Heading-Level.")
    parser.add_argument("--parent-level", type=int, default=None,
                        help="Heading-Level des Elternkapitels (Standard: automatisch aus --parent-chapter).")
    parser.add_argument("--include-references", action="store_true",
                        help="Referenzen, Literaturverzeichnis etc. einbeziehen (Standard: werden übersprungen).")

    args = parser.parse_args()
    OUTPUT_BASE = "workspace/output"

    try:
        if not os.getenv("GEMINI_API_KEY"):
            raise ValueError("GEMINI_API_KEY fehlt in der .env-Datei!")

        if not os.path.exists(args.pdf_path):
            raise FileNotFoundError(f"Die Datei {args.pdf_path} wurde nicht gefunden.")

        pdf_stem = Path(args.pdf_path).stem
        out_dir  = Path(OUTPUT_BASE) / pdf_stem
        out_dir.mkdir(parents=True, exist_ok=True)

        # --- Schritt 1: OCR ---
        md_path = out_dir / f"{pdf_stem}.md"
        if args.force or not md_path.exists():
            raw_md_path = run_marker_ocr(args.pdf_path, OUTPUT_BASE)
            md_path = Path(raw_md_path)
        else:
            print(f"[SKIP] OCR – Markdown bereits vorhanden: {md_path}")

        raw_md = md_path.read_text(encoding="utf-8")

        # --- Schritt 2: Sprache prüfen & Übersetzen ---
        transl_path = out_dir / "de_uebersetzung.md"
        if args.force and transl_path.exists():
            transl_path.unlink()

        if not transl_path.exists():
            if check_if_english(raw_md):
                print("Text ist Englisch. Starte Übersetzung...")
                working_text = load_or_run(transl_path, lambda: translate_text(raw_md), "Übersetzung")
            else:
                print("Text ist bereits Deutsch. Keine Übersetzung notwendig.")
                working_text = raw_md
                transl_path.write_text(working_text, encoding="utf-8")
        else:
            print(f"[SKIP] Übersetzung – bereits vorhanden: {transl_path}")
            working_text = transl_path.read_text(encoding="utf-8")

        # --- Zwischenschritt: Übersetzung als eigenes Word-Dokument ---
        transl_docx_path = out_dir / f"{pdf_stem}_Uebersetzung.docx"
        if args.force or not transl_docx_path.exists():
            build_translation_word_document(working_text, str(transl_docx_path), base_path=str(out_dir))
        else:
            print(f"[SKIP] Übersetzungs-Docx – bereits vorhanden: {transl_docx_path}")

        # --- Schritt 4: Zusammenfassung (kapitelweise) ---
        sum_path = out_dir / "zusammenfassung.md"
        if args.force and sum_path.exists():
            sum_path.unlink()
            for f in out_dir.glob("zusammenfassung_kap_*.md"):
                f.unlink()

        if sum_path.exists():
            print(f"[SKIP] Zusammenfassung – bereits vorhanden: {sum_path}")
            summary_result = sum_path.read_text(encoding="utf-8")
        else:
            summary_result = generate_summary_by_chapter(working_text, out_dir)

        # --- Schritt 5: Qualitätssicherung (optional) ---
        qa_result = "Keine Leitfragen zur Prüfung übergeben."
        if args.questions:
            if os.path.exists(args.questions):
                qa_path = out_dir / "qa_ergebnis.md"
                if args.force and qa_path.exists():
                    qa_path.unlink()
                qa_result = load_or_run(
                    qa_path,
                    lambda: verify_with_questions(summary_result, args.questions),
                    "QA / Leitfragen"
                )
            else:
                print(f"Warnung: Fragen-Datei '{args.questions}' nicht gefunden. Überspringe QS.")

        # --- Word-Dokument zusammensetzen ---
        suffix = f"_Einbetten_{args.parent_chapter.replace('.', '-')}" if args.parent_chapter else "_Lernskript"
        final_docx_path = out_dir / f"{pdf_stem}{suffix}.docx"
        build_interleaved_word_document(
            working_text, summary_result, qa_result,
            str(final_docx_path), base_path=str(out_dir),
            parent_chapter=args.parent_chapter,
            parent_level=args.parent_level,
            skip_references=not args.include_references,
        )

        print(f"\n=== PIPELINE ERFOLGREICH BEENDET ===")
        print(f"Zwischenergebnisse:   {out_dir}")
        print(f"Übersetzung (Word):   {transl_docx_path}")
        print(f"Fertiges Dokument:    {final_docx_path}")
        if args.parent_chapter:
            print(f"  → Einfügemodus: Kapitelpräfix '{args.parent_chapter}', "
                  f"Heading-Shift +{args.parent_level or _parent_level_from_chapter(args.parent_chapter)}")

    except Exception as e:
        print(f"\nPipeline abgebrochen wegen: {e}")
