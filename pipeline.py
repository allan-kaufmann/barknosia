import os
import argparse
import subprocess
import sys
import time
import re
import json
from pathlib import Path
from dotenv import load_dotenv
from google import genai
from google.genai import types
from google.genai.errors import APIError
from lxml import etree
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.opc.part import Part
from docx.opc.packuri import PackURI

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

# Gemini API-Client: lazy initialisiert beim ersten API-Aufruf (nicht beim Import).
# Verhindert dass Tests ohne API-Key scheitern.
_gemini_client = None

def _get_gemini_client():
    global _gemini_client
    if _gemini_client is None:
        _gemini_client = genai.Client()
    return _gemini_client

def call_gemini_with_retry(model_name: str, contents, config, max_retries: int = 5, delay: int = 5):
    """Hilfsfunktion: Ruft Gemini auf und wiederholt den Versuch bei Serverüberlastung (503)."""
    client = _get_gemini_client()
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


def _clean_heading_text(line: str) -> str:
    """Extrahiert reinen Text einer Markdown-Überschrift (ohne HTML, Links, Bold-Marker, #)."""
    text = re.sub(r'<[^>]+>', '', line)                    # HTML-Tags (<span ...>) entfernen
    text = re.sub(r'\[([^\]]*)\]\([^)]*\)', r'\1', text)  # [Text](url) → Text
    text = re.sub(r'\*+', '', text)                        # **bold** → text
    text = re.sub(r'^#+\s*', '', text)                     # führende # entfernen
    return text.strip()


def extract_chapter(text: str, chapter_id: str) -> str:
    """
    Extrahiert ein bestimmtes Kapitel (inkl. aller Unterkapitel) aus einem Markdown-Text.
    chapter_id: z.B. '4.2' oder '1' — robuste Erkennung auch bei HTML-Tags in Überschriften.
    Unterstützt zwei Heading-Formate:
      - Numerisch:  "7 Titel…"  /  "7.2 Sub…"
      - Label:      "Kapitel 7" / "Chapter 7" / "Teil 7" / "Abschnitt 7"
    Extraktion endet bei der nächsten Überschrift gleicher/höherer Ebene die kein Unterkapitel ist.
    """
    escaped = re.escape(chapter_id)
    _label_pat = r'(?:Kapitel|Chapter|Teil|Abschnitt)'
    lines = text.split('\n')
    start_idx = None
    heading_level = None
    label_mode = False   # True wenn Kapitel via "Kapitel N"-Format gefunden

    for i, line in enumerate(lines):
        m = re.match(r'^(#{1,6})\s', line)
        if m:
            clean = _clean_heading_text(line)
            # Primär: "7 Titel" oder "7" allein — NICHT "7.1" oder "71"
            if re.match(rf'^{escaped}(\s|$)', clean):
                start_idx = i
                heading_level = len(m.group(1))
                break
            # Fallback: "Kapitel 7", "Chapter 7", "Teil 7", "Abschnitt 7"
            if re.match(rf'^{_label_pat}\s+{escaped}(\s|$)', clean, re.IGNORECASE):
                start_idx = i
                heading_level = len(m.group(1))
                label_mode = True
                break

    # Fallback: Kein Top-Level-Heading vorhanden (OCR hat es übersprungen),
    # aber Unterkapitel existieren (z.B. "8.1 ..." wenn "8 ..." fehlt).
    if start_idx is None:
        for i, line in enumerate(lines):
            m = re.match(r'^(#{1,6})\s', line)
            if m:
                clean = _clean_heading_text(line)
                if re.match(rf'^{escaped}\.', clean):
                    start_idx = i
                    heading_level = len(m.group(1))
                    break

    if start_idx is None:
        raise ValueError(f"Kapitel '{chapter_id}' nicht im Markdown gefunden. "
                         f"Tipp: --chapter mit exakter Nummer angeben (z.B. '1' oder '4.2').")

    result = []
    for i, line in enumerate(lines):
        if i < start_idx:
            continue
        if i > start_idx:
            m2 = re.match(r'^(#{1,6})\s', line)
            if m2:
                clean2 = _clean_heading_text(line)
                # Ende erst beim nächsten nummerierten Kapitel, das weder das Kapitel
                # selbst noch ein Unterkapitel (chapter_id.x) ist.
                if re.match(r'^\d', clean2) and not re.match(rf'^{escaped}(\.|\s|$)', clean2):
                    break
                # Im label_mode: auch bei "Kapitel M" (M ≠ chapter_id) stoppen.
                if label_mode and re.match(rf'^{_label_pat}\s+\d', clean2, re.IGNORECASE):
                    if not re.match(rf'^{_label_pat}\s+{escaped}(\s|$)', clean2, re.IGNORECASE):
                        break
        result.append(line)

    extracted = '\n'.join(result)
    print(f"[KAPITEL] '{chapter_id}' extrahiert: {len(result)} Zeilen, {len(extracted)} Zeichen")
    return extracted


def parse_qa_response(qa_text: str) -> list:
    """
    Parst den strukturierten QA-Output von verify_with_questions().
    Gibt eine Liste von Dicts zurück:
      [{'num': 1, 'antwort': '...', 'textgrundlage': '...', 'schluessel': '...', 'abdeckung': '...'}, ...]
    """
    items = []
    # Splitten an "Frage N" Zeilen
    blocks = re.split(r'(?m)^\*{0,2}(?:##\s*)?Frage\s+(\d+)\*{0,2}:?\s*$', qa_text)
    # blocks[0] = text vor Frage 1, dann abwechselnd: Fragenummer, Frageblock
    i = 1
    while i < len(blocks) - 1:
        num_str = blocks[i].strip()
        block = blocks[i + 1]
        block = re.sub(r'^\*\*([^*:]+:)\*\*', r'\1', block, flags=re.MULTILINE)
        i += 2
        try:
            num = int(num_str)
        except ValueError:
            continue

        def _extract(pattern, text, default='–'):
            m = re.search(pattern, text, re.MULTILINE | re.DOTALL)
            if m:
                return m.group(1).strip().split('\n')[0].strip()
            return default

        antwort     = _extract(r'^Antwort:\s*(.+?)(?=\n(?:Textgrundlage|Schlüsselbegriffe|Beleg|Abdeckung|$))', block)
        textgr      = _extract(r'^Textgrundlage:\s*(.+)', block)
        schluessel  = _extract(r'^Schlüsselbegriffe:\s*(.+)', block)
        beleg       = _extract(r'^Beleg:\s*(.+)', block, default='')
        abdeckung   = _extract(r'^Abdeckung:\s*(.+)', block)

        items.append({
            'num': num,
            'antwort': antwort,
            'textgrundlage': textgr,
            'schluessel': schluessel,
            'beleg': beleg,
            'abdeckung': abdeckung,
        })
    return items


def serialize_qa_items(items: list) -> str:
    """
    Serialisiert geparste QA-Items zurück ins Markdown-Format von verify_with_questions().
    Inverse zu parse_qa_response() – wird nach der Nachbearbeitung (Schritt 5b) genutzt,
    damit der Downstream-Pfad (Doc-Builder ruft parse_qa_response erneut auf) unverändert läuft.
    """
    blocks = []
    for it in items:
        lines = [
            f"**Frage {it['num']}**",
            f"**Antwort:** {it.get('antwort', '–')}",
            f"**Textgrundlage:** {it.get('textgrundlage', '–')}",
            f"**Schlüsselbegriffe:** {it.get('schluessel', '–')}",
        ]
        if it.get('beleg'):
            lines.append(f"**Beleg:** {it['beleg']}")
        lines.append(f"**Abdeckung:** {it.get('abdeckung', '–')}")
        blocks.append('\n'.join(lines))
    return '\n\n'.join(blocks)


def _add_comment_range_start(paragraph, comment_id: int):
    """Setzt w:commentRangeStart als erstes Kind eines Absatzes."""
    crs = OxmlElement('w:commentRangeStart')
    crs.set(qn('w:id'), str(comment_id))
    paragraph._p.insert(0, crs)


def _add_comment_range_end(paragraph, comment_id: int):
    """Setzt w:commentRangeEnd + w:commentReference ans Ende eines Absatzes."""
    p_elem = paragraph._p
    cre = OxmlElement('w:commentRangeEnd')
    cre.set(qn('w:id'), str(comment_id))
    p_elem.append(cre)
    run = OxmlElement('w:r')
    rpr = OxmlElement('w:rPr')
    rs = OxmlElement('w:rStyle')
    rs.set(qn('w:val'), 'CommentReference')
    rpr.append(rs)
    run.append(rpr)
    ref = OxmlElement('w:commentReference')
    ref.set(qn('w:id'), str(comment_id))
    run.append(ref)
    p_elem.append(run)


def _normalize_quote(text: str) -> str:
    """Normalisiert Text für Beleg-Matching: Markdown weg, Whitespace kollabiert, lowercase."""
    text = re.sub(r'[*_`"„“”»«\']', '', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip().lower()


def _find_para_by_quote(paras: list, quote: str):
    """
    Findet den Absatz in `paras`, dessen Text das (normalisierte) Beleg-Zitat enthält.
    Nutzt einen Präfix des Zitats (robust gegen leichte Abweichungen am Satzende).
    Gibt den Absatz zurück oder None.
    """
    nq = _normalize_quote(quote)
    if len(nq) < 8:
        return None
    needle = nq[:60]
    for p in paras:
        if not p.text.strip():
            continue
        if needle in _normalize_quote(p.text):
            return p
    return None


def _inject_comments_part(doc, comments_list: list):
    """
    Erstellt word/comments.xml und registriert es im OPC-Package.
    comments_list: [(id: int, text: str), ...]
    Muss VOR doc.save() aufgerufen werden.
    """
    WURI  = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    CT    = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'
    RT    = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
    WDATE = '2024-01-01T00:00:00Z'

    root = etree.Element(f'{{{WURI}}}comments',
                         nsmap={'w': WURI,
                                'wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
                                'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships'})
    for cid, text in comments_list:
        c = etree.SubElement(root, f'{{{WURI}}}comment')
        c.set(f'{{{WURI}}}id',       str(cid))
        c.set(f'{{{WURI}}}author',   'Lernfragen')
        c.set(f'{{{WURI}}}date',     WDATE)
        c.set(f'{{{WURI}}}initials', 'LF')
        p = etree.SubElement(c, f'{{{WURI}}}p')
        r = etree.SubElement(p, f'{{{WURI}}}r')
        t = etree.SubElement(r, f'{{{WURI}}}t')
        t.text = text

    xml_bytes = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    try:
        part = Part(PackURI('/word/comments.xml'), CT, xml_bytes, doc.part.package)
        doc.part.relate_to(part, RT)
        print(f"   Word-Kommentare: {len(comments_list)} eingebettet")
    except Exception as e:
        print(f"   Kommentare konnten nicht eingebettet werden: {e}")


# Menschlich lesbare Namen für die gängigsten Sprachcodes (für Übersetzungs-Prompts).
LANGUAGE_NAMES = {
    "de": "Deutsch",
    "en": "Englisch",
    "fr": "Französisch",
    "es": "Spanisch",
    "it": "Italienisch",
}


def _language_name(code: str) -> str:
    """Liefert den deutschen Klarnamen zu einem Sprachcode (Fallback: Code selbst)."""
    return LANGUAGE_NAMES.get((code or "").lower(), code)


def detect_language(text: str) -> str:
    """Schritt 2a: Erkennt die Hauptsprache des Dokuments und liefert einen ISO-Code
    (z.B. 'de', 'en').

    Nimmt drei Stichproben (Anfang, Mitte, Ende), um Dokumente mit z.B. deutschem Titel
    aber englischem Haupttext korrekt zu erkennen. Bei endgültigem Fehler wird 'unknown'
    zurückgegeben – die aufrufende Logik übersetzt dann NICHT (sichere Annahme:
    Quellsprache = Zielsprache), statt blind auf Englisch/Übersetzung auszuweichen.
    """
    print("--- Schritt 2a: Erkenne Sprache des Dokuments ---")
    n = len(text)
    samples = [
        text[:1500],
        text[max(0, n // 2 - 750): n // 2 + 750],
        text[max(0, n - 1500):],
    ]
    leseprobe = "\n\n---\n\n".join(s for s in samples if s.strip())
    prompt = (
        "Bestimme die Hauptsprache des folgenden Textes. "
        "Antworte mit exakt einem ISO-639-1-Sprachcode in Kleinbuchstaben "
        "(z.B. 'de' für Deutsch, 'en' für Englisch). "
        "Ignoriere vereinzelte fremdsprachige Wörter (Eigennamen, Zitate) – "
        "frage nur nach der Hauptsprache. Gib nur den Code aus, sonst nichts.\n\n"
        f"Text:\n{leseprobe}"
    )
    try:
        response = call_gemini_with_retry(
            model_name='gemini-2.5-flash',
            contents=prompt,
            config=types.GenerateContentConfig(temperature=0.0, max_output_tokens=5)
        )
        code = re.sub(r'[^a-z]', '', response.text.strip().lower())[:2]
        if code:
            print(f"   Erkannte Sprache: {code}")
            return code
        print("   Sprachcode nicht erkennbar – behandle als 'unknown'.")
        return "unknown"
    except Exception as e:
        print(f"Sprachprüfung fehlgeschlagen ({e}); behandle als 'unknown' (keine Übersetzung).")
        return "unknown"


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


def translate_text(text: str, source_lang: str = "en", target_lang: str = "de") -> str:
    """Schritt 2b: Übersetzt den Text abschnittsweise nach den strengen Regeln.

    source_lang/target_lang sind ISO-639-1-Codes; der Prompt wird daraus dynamisch gebaut,
    damit nie versehentlich in die falsche Richtung (z.B. Deutsch → Englisch) übersetzt wird.
    """
    src_name = _language_name(source_lang)
    tgt_name = _language_name(target_lang)
    print(f"--- Schritt 2b: Übersetze Text von {src_name} nach {tgt_name} (via Gemini 2.5 Pro) ---")

    system_prompt = (
        f"Übersetze den folgenden {src_name.lower()}en wissenschaftlichen Text "
        f"originalgetreu ins {tgt_name}.\n\n"
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
        "Nur die deutsche Übersetzung, kein Kommentar, keine Kontrollliste danach."
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
    Splittet den (bereits normalisierten) Markdown-Text an nummerierten Kapitelüberschriften.
    Adaptiv: Bei einem einzigen Top-Level-Kapitel (z.B. extrahiertes Einzelkapitel) wird eine
    Ebene tiefer gesplittet, sodass Unterkapitel als eigene Chunks erkannt werden.
    Jeder Eintrag: {'heading': str, 'full_text': str}.
    """
    lines = text.split('\n')
    _html_re = re.compile(r'<[^>]+>')

    def _numbered_level(line: str):
        """Gibt den #-Level zurück wenn die Zeile eine nummerierte Überschrift ist (HTML-ignorant)."""
        stripped = _html_re.sub('', line).replace('**', '')
        m = re.match(r'^(#{1,6})\s+(\d+(?:\.\d+)*)\s', stripped)
        return len(m.group(1)) if m else None

    numbered_levels = [nl for line in lines if (nl := _numbered_level(line)) is not None]

    if not numbered_levels:
        return []

    min_level = min(numbered_levels)
    top_count = sum(1 for l in numbered_levels if l == min_level)
    # Einziges Top-Kapitel → eine Ebene tiefer splitten (z.B. extrahiertes Kapitel 4)
    split_level = min_level + 1 if top_count == 1 else min_level

    chapters = []
    current_heading = None
    current_lines = []
    pre_split_lines: list[str] = []

    for line in lines:
        if _numbered_level(line) == split_level:
            if current_heading is not None:
                chapters.append({'heading': current_heading, 'full_text': '\n'.join(current_lines), 'level': split_level})
            stripped = _html_re.sub('', line).replace('**', '')
            current_heading = re.sub(r'^#{1,6}\s+', '', stripped).strip()
            current_lines = [line]
        else:
            if current_heading is None:
                pre_split_lines.append(line)
            else:
                current_lines.append(line)

    if current_heading is not None:
        chapters.append({'heading': current_heading, 'full_text': '\n'.join(current_lines), 'level': split_level})

    # Im adaptiven Modus (split_level = min_level+1) landet Text vor dem ersten Unterkapitel
    # (z.B. Einleitungstext nach "# 5 Kompetenzen..." vor "## 5.1 ...") in pre_split_lines.
    # Diesen ans erste Kapitel hängen, damit er mitverdichtet wird.
    if pre_split_lines and chapters:
        preamble_str = '\n'.join(pre_split_lines).strip()
        if preamble_str:
            # Erste Überschriftenzeile im Vorspann = Kapitel-Root (z.B. "3 Analyse …").
            _root_idx = next((k for k, ln in enumerate(pre_split_lines)
                              if re.match(r'^#{1,6}\s', _html_re.sub('', ln).replace('**', ''))), None)
            _root_body = '\n'.join(pre_split_lines[_root_idx + 1:]).strip() if _root_idx is not None else ''
            if _root_idx is not None and len(_root_body) > 300:
                # Kapitel-Root mit eigenem Einleitungstext → eigenes Kapitel, damit seine
                # Zusammenfassung dem Root-Heading zugeordnet werden kann (sonst unsichtbar in Kap. 1).
                _root_heading = re.sub(r'^#{1,6}\s+', '',
                                       _html_re.sub('', pre_split_lines[_root_idx]).replace('**', '')).strip()
                chapters.insert(0, {'heading': _root_heading,
                                    'full_text': preamble_str,
                                    'level': split_level - 1})
            else:
                chapters[0]['full_text'] = preamble_str + '\n\n' + chapters[0]['full_text']

    return chapters


def _split_at_level2(chapter_text: str) -> tuple[str, list[dict]]:
    """
    Teilt einen Kapiteltext an ## Überschriften (genau 2 Rauten, nicht ###) auf.
    Gibt zurück: (preamble_text, [{'heading': str, 'text': str}, ...])
    Der preamble ist der Text vor der ersten ## Überschrift.
    """
    lines = chapter_text.split('\n')
    preamble_lines: list[str] = []
    sections: list[dict] = []
    current_heading: str | None = None
    current_lines: list[str] = []

    for line in lines:
        m = re.match(r'^##\s+(.+)$', line)
        if m:
            if current_heading is not None:
                sections.append({'heading': current_heading, 'text': '\n'.join(current_lines)})
            elif current_lines:
                preamble_lines = current_lines[:]
            current_heading = m.group(1).strip()
            current_lines = [line]
        else:
            current_lines.append(line)

    if current_heading is not None:
        sections.append({'heading': current_heading, 'text': '\n'.join(current_lines)})
    elif current_lines and not preamble_lines:
        preamble_lines = current_lines

    return '\n'.join(preamble_lines), sections


def _split_at_level(chapter_text: str, level: int) -> tuple[str, list[dict]]:
    """Teilt text an Überschriften der angegebenen #-Tiefe. Gibt (preamble, sections) zurück."""
    prefix = '#' * level + ' '
    avoid_prefix = '#' * level + '##'
    lines = chapter_text.split('\n')
    preamble_lines: list[str] = []
    sections: list[dict] = []
    current_heading: str | None = None
    current_lines: list[str] = []

    for line in lines:
        if line.startswith(prefix) and not line.startswith(avoid_prefix):
            if current_heading is not None:
                sections.append({'heading': current_heading, 'text': '\n'.join(current_lines)})
            elif current_lines:
                preamble_lines = current_lines[:]
            current_heading = line[len(prefix):].strip()
            current_lines = [line]
        else:
            current_lines.append(line)

    if current_heading is not None:
        sections.append({'heading': current_heading, 'text': '\n'.join(current_lines)})
    elif current_lines and not preamble_lines:
        preamble_lines = current_lines

    return '\n'.join(preamble_lines), sections


def _detect_chapter_level(text: str) -> int:
    """Gibt den #-Level der ersten Heading-Zeile zurück (Fallback: 2)."""
    for line in text.splitlines():
        m = re.match(r'^(#{1,6})\s', line)
        if m:
            return len(m.group(1))
    return 2


def _find_sublevel(text: str, chapter_heading_level: int) -> int | None:
    """Findet die nächste Heading-Ebene unterhalb chapter_heading_level, oder None."""
    for level in range(chapter_heading_level + 1, 7):
        prefix = '#' * level + ' '
        avoid = '#' * level + '##'
        if any(line.startswith(prefix) and not line.startswith(avoid) for line in text.splitlines()):
            return level
    return None


def _group_into_chunks(preamble: str, sections: list[dict], max_chars: int = 15_000) -> list[dict]:
    """
    Fasst sections greedy zu Chunks von max max_chars Zeichen zusammen.
    Sections werden NIE in der Mitte geteilt – immer vollständig in einem Chunk.
    Preamble geht immer in den ersten Chunk.
    """
    chunks: list[dict] = []
    current_secs: list[dict] = []
    current_len = len(preamble)

    for sec in sections:
        sec_len = len(sec['text'])
        if current_secs and current_len + sec_len > max_chars:
            chunks.append({'preamble': preamble if not chunks else '', 'sections': current_secs})
            current_secs = []
            current_len = 0
        current_secs.append(sec)
        current_len += sec_len

    if current_secs:
        chunks.append({'preamble': preamble if not chunks else '', 'sections': current_secs})
    elif not chunks:
        chunks.append({'preamble': preamble, 'sections': []})

    return chunks


def _summarize_chapter_by_sections(ch: dict, out_dir: Path) -> str:
    """
    Für große Kapitel (> 3 Level-2-Sections): jede Section einzeln zusammenfassen.
    Caching: je Section eine eigene Datei zusammenfassung_kap_XX_YY.md.
    Kombinierter Output wird immer frisch aus Sub-Caches gebaut (kein äußeres load_or_run).
    """
    i = ch['index']
    preamble, sections = _split_at_level2(ch['full_text'])
    parts: list[str] = []

    if preamble.strip():
        preamble_file = out_dir / f"zusammenfassung_kap_{i:02d}_00.md"
        needs_call = not preamble_file.exists()
        preamble_summary = load_or_run(
            preamble_file,
            lambda: _summarize_single_chapter(ch['heading'], preamble),
            f"Zusammenfassung Kap. {i} Einleitung: {ch['heading'][:50]}"
        )
        parts.append(preamble_summary)
        if needs_call:
            time.sleep(1)

    for j, sec in enumerate(sections, start=1):
        sec_file = out_dir / f"zusammenfassung_kap_{i:02d}_{j:02d}.md"
        needs_call = not sec_file.exists()
        sec_summary = load_or_run(
            sec_file,
            lambda s=sec: _summarize_single_chapter(s['heading'], s['text']),
            f"Zusammenfassung Kap. {i}.{j:02d}: {sec['heading'][:50]}"
        )
        if not re.match(r'^#+\s', sec_summary.strip()):
            sec_summary = f"## {sec['heading']}\n\n{sec_summary}"
        parts.append(sec_summary)
        if needs_call:
            time.sleep(1)

    return '\n\n'.join(parts)


def _summarize_single_chapter(heading: str, chapter_text: str, output_lang: str = "de") -> str:
    """Erstellt eine lernorientierte Zusammenfassung für ein einzelnes Kapitel.

    output_lang (ISO-Code) erzwingt die Ausgabesprache: Auch fremdsprachige Passagen im
    Original (z.B. ein englisches Abstract) werden in dieser Sprache zusammengefasst, statt
    in der Originalsprache zu verbleiben.
    """
    lang_name = _language_name(output_lang)
    prompt = (
        f"Erstelle eine lernorientierte Zusammenfassung für das folgende Kapitel: \"{heading}\"\n\n"
        f"AUSGABESPRACHE: Der gesamte FLIESSTEXT (Stichpunkte, Definitionen, Beschreibungen) MUSS "
        f"auf {lang_name} verfasst sein. Liegt ein Teil des Originals in einer anderen Sprache vor "
        f"(z.B. ein englisches Abstract), übersetze dessen INHALT in der Zusammenfassung nach {lang_name}.\n"
        f"WICHTIG: Die ÜBERSCHRIFTEN dagegen EXAKT und unverändert aus dem Original übernehmen "
        f"(gleicher Wortlaut, gleiche Sprache, gleiche Nummerierung) – NICHT übersetzen und NICHT "
        f"umformulieren. Sie dienen als Zuordnungsschlüssel und müssen 1:1 mit dem Original übereinstimmen.\n\n"
        "Pflichtanforderungen:\n"
        f"1. ALLE Unterkapitel müssen vorhanden sein – kein einziges Unterkapitel darf fehlen!\n"
        f"   Übernimm jede Überschrift exakt wie im Original (Wortlaut inkl. evtl. Nummerierung, "
        f"z.B. '4.1.1 Hedonisches Wohlbefinden').\n"
        "2. Pro Unterkapitel: mindestens 3–5 Stichpunkte mit den wichtigsten Inhalten.\n"
        "3. Studienergebnisse IMMER erhalten: Metaanalysen, Effektstärken, Befundrichtung, Autoren & Jahr.\n"
        "4. Definitionen: wörtlich oder sehr nah am Original übernehmen.\n"
        "5. Keine neuen Informationen ergänzen.\n"
        "6. Abbildungen und Tabellen kurz erwähnen und ihren Inhalt beschreiben.\n"
        "7. Stichpunkte statt Fließtext (Ausnahme: Definitionen).\n"
        "8. Länge: maximal 40–50 % des Originals – ABER vollständige Unterkapitelabdeckung hat Vorrang vor Kürze.\n"
        "9. Kästen (Fokus/Studie/Definition/Beispiel) sind eigenständige Lernobjekte – "
        "fasse jeden als abgeschlossene Einheit unter seiner eigenen Überschrift zusammen.\n\n"
        "Selbstprüfung (am Ende anhängen):\n"
        "- Liste alle Unterkapitel des Originals auf\n"
        "- Markiere fehlende Unterkapitel oder fehlende Studienergebnisse\n\n"
        f"Kapiteltext:\n{chapter_text}"
    )
    # Pflichtliste aller Unterkapitel: alle Headings außer dem Kapitel-Heading selbst.
    _all_headings = re.findall(r'^#{2,6}\s+\*?\*?(.+?)\*?\*?\s*$', chapter_text, re.MULTILINE)
    _heading_norm = normalize_heading(heading)
    _required = [h for h in _all_headings if normalize_heading(h) != _heading_norm]
    if len(_required) >= 1:
        _rlist = '\n'.join(f'- {_h}' for _h in _required[:80])
        prompt += (
            f"\n\nPFLICHT-VOLLSTÄNDIGKEIT – folgende {len(_required)} Unterabschnitte MÜSSEN ALLE "
            f"als eigene Überschrift erscheinen:\n{_rlist}\n"
            "Jeder Abschnitt benötigt mindestens 1 Stichpunkt. Keiner darf fehlen!\n"
        )
    # Sehr große Kapitel (>50 Abschnitte, z.B. 263 Kompetenzen): strenge Kompression nötig
    # damit alle Abschnitte in den Output passen.
    # 263 Abschnitte × 3 Bullets × 15 Wörter ≈ 12.000 Wörter → passt in 65.536 Output-Token.
    if len(_required) > 50:
        prompt += (
            f"\n\nWICHTIG FÜR DIESES SEHR GROSSE KAPITEL ({len(_required)} Abschnitte):\n"
            "- Maximal 3 Stichpunkte pro Unterabschnitt (Ausnahme: Definition immer vollständig)\n"
            "- Maximal 20 Wörter pro Stichpunkt\n"
            "- Vollständigkeit (alle Abschnitte vorhanden) hat ABSOLUTEN Vorrang vor Ausführlichkeit\n"
            "- Keinen Abschnitt auslassen, lieber sehr kurz als gar nicht!\n"
        )
    # Sehr große Kapitel brauchen 65536 Output-Token; normale Kapitel 32768 für ausführliche Summaries.
    _out_tokens = 65536 if len(_required) > 50 else 32768
    try:
        response = call_gemini_with_retry(
            model_name='gemini-2.5-pro',
            contents=prompt,
            config=types.GenerateContentConfig(
                temperature=0.2,
                max_output_tokens=_out_tokens
            )
        )
        return response.text
    except Exception as e:
        print(f"Fehler bei Zusammenfassung von '{heading}': {e}")
        raise


def generate_summary_by_chapter(text: str, out_dir: Path, output_lang: str = "de") -> str:
    """
    Schritt 4: Erstellt die Zusammenfassung kapitelweise (je Level-1-Kapitel ein API-Call).
    Nutzt Caching pro Kapitel: zusammenfassung_kap_XX.md in out_dir.
    Kombiniert am Ende zu zusammenfassung.md.
    output_lang erzwingt die Ausgabesprache der Zusammenfassung.
    """
    print("--- Schritt 4: Erstelle kapitelweise Zusammenfassung (via Gemini 2.5 Pro) ---")

    text = normalize_heading_levels(text)

    # Preamble-Text (vor erstem nummerierten Heading) extrahieren.
    # split_into_level1_chapters() verwirft diesen Text sonst kommentarlos.
    _first_num = re.search(r'^#{1,6}\s+\d', text, re.MULTILINE)
    _preamble_text = text[:_first_num.start()].strip() if _first_num and _first_num.start() > 0 else ''

    chapters = split_into_level1_chapters(text)

    # Preamble zum ersten Kapitel hinzufügen damit Gemini es mitfasst.
    if _preamble_text and chapters:
        chapters[0]['full_text'] = _preamble_text + '\n\n' + chapters[0]['full_text']

    _CHUNK_LIMIT = 15_000  # ~30 Seiten; Kapitel über diesem Limit werden chunk-weise verarbeitet

    def _summarize_chapter_smart(ch: dict, ci: int) -> str:
        """Fasst ein Kapitel zusammen: direkt wenn klein, chunk-weise wenn groß."""
        # 'level' wird von split_into_level1_chapters gesetzt; Fallback für manuell erstellte Dicts.
        chap_level = ch.get('level', _detect_chapter_level(ch['full_text']))
        sublevel = _find_sublevel(ch['full_text'], chap_level)

        if sublevel is None or len(ch['full_text']) <= _CHUNK_LIMIT:
            return _summarize_single_chapter(ch['heading'], ch['full_text'], output_lang)

        preamble, sections = _split_at_level(ch['full_text'], sublevel)
        chunks = _group_into_chunks(preamble, sections, _CHUNK_LIMIT)

        if len(chunks) <= 1:
            return _summarize_single_chapter(ch['heading'], ch['full_text'], output_lang)

        print(f"   Kapitel {ci} ({ch['heading'][:40]}) → {len(chunks)} Chunks à max {_CHUNK_LIMIT} Zeichen")
        parts: list[str] = []
        for j, chunk in enumerate(chunks, start=1):
            chunk_file = out_dir / f"zusammenfassung_kap_{ci:02d}_{j:02d}.md"
            chunk_body = '\n\n'.join(s['text'] for s in chunk['sections'])
            chunk_text = (chunk['preamble'] + '\n\n' + chunk_body).strip()
            is_first = (j == 1)
            continuation_hint = (
                '' if is_first else
                f'\n\nHINWEIS: Dies ist Teil {j} von {len(chunks)} dieses Kapitels. '
                'Beginne direkt mit den Unterkapitelüberschriften ohne das übergeordnete Kapitelheading zu wiederholen.'
            )

            def _chunk_gen(ct=chunk_text, hd=ch['heading'], hint=continuation_hint):
                return _summarize_single_chapter(hd, ct + hint, output_lang)

            chunk_label = f"Zusammenfassung Kap. {ci} Teil {j}/{len(chunks)}: {ch['heading'][:40]}"
            chunk_summary = load_or_run(chunk_file, _chunk_gen, chunk_label)

            # Doppeltes Kapitel-Heading aus Chunks 2+ entfernen (falls KI es wiederholt)
            if not is_first and chunk_summary.strip():
                first_line = chunk_summary.strip().splitlines()[0]
                if re.match(r'^#{1,3}\s', first_line) and ch['heading'][:15] in first_line:
                    chunk_summary = chunk_summary.strip()[len(first_line):].lstrip('\n')

            parts.append(chunk_summary)
            time.sleep(1)

        return '\n\n'.join(parts)

    if not chapters:
        print("   Keine Level-1-Kapitel gefunden, fasse Gesamttext zusammen...")
        fallback_path = out_dir / "zusammenfassung_kap_00.md"
        # Wenn substanzieller Preamble-Inhalt existiert (Abstract etc.), einen fixen
        # "## Einleitung"-Abschnitt voranstellen, damit die KI ihn explizit zusammenfasst
        # und sum_lookup["einleitung"] später im Docx-Builder verwendet werden kann.
        _pre, _ = _split_at_level(text, 2)
        _pre_body = re.sub(r'^#{1,6}\s.*$', '', _pre, flags=re.MULTILINE).strip()
        full_text_for_summary = ('## Einleitung\n\n' + text) if len(_pre_body) > 300 else text
        fallback_ch = {'heading': 'Volltext', 'full_text': full_text_for_summary, 'level': 1}
        result = load_or_run(
            fallback_path,
            lambda: _summarize_chapter_smart(fallback_ch, 0),
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
            lambda ch=chapter, ci=i: _summarize_chapter_smart(ch, ci),
            label
        )
        summaries.append(chapter_summary)
        time.sleep(1)

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
        "- Keine groben Kapitelverweise, wenn Unterkapitel vorhanden sind.\n"
        "- Beleg: Ein wörtliches Zitat (5–15 Wörter), exakt aus der Wissensbasis kopiert "
        "– die Kernstelle, die die Frage belegt. Keine Paraphrase, keine Auslassungszeichen.\n\n"
        "Ausgabeformat pro Frage:\n"
        "Frage X\n"
        "Antwort: [max. 3 Sätze]\n"
        "Textgrundlage: [genaues Unterkapitel]\n"
        "Schlüsselbegriffe: [1–3 Begriffe]\n"
        "Beleg: [wörtliches Zitat aus der Zusammenfassung, 5–15 Wörter]\n"
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


_REWORK_LEVELS = {"teilweise", "nicht enthalten"}
_NO_SUPPLEMENT_MARKER = "KEINE ERGÄNZUNG MÖGLICH"


def rework_partial_answers(qa_text: str, working_text: str, questions_path: str = None):
    """
    Schritt 5b: Nacharbeit unvollständiger Antworten.

    Für jede Frage mit Abdeckung 'teilweise' oder 'nicht enthalten' wird der
    übersetzte Originaltext (working_text) der zugehörigen Textgrundlage geprüft
    und – falls dort vorhanden – die fehlende prüfungsrelevante Information ergänzt.
    Erfolgreiche Ergänzungen erhalten die Abdeckung 'vollständig durch Nachbearbeitung'.

    Rückgabe: (augmented_qa_text, supplement_map)
      augmented_qa_text: serialisierte QA (Markdown) mit aktualisierten Items
      supplement_map: {normalize_heading(textgrundlage): [ergaenzungstext, ...]}
    """
    items = parse_qa_response(qa_text)
    todo = [it for it in items if it.get('abdeckung', '').strip().lower() in _REWORK_LEVELS]
    if not todo:
        return qa_text, {}

    print(f"--- Schritt 5b: Nacharbeit {len(todo)} unvollständige(r) Antwort(en) ---")

    # Fragetexte laden (optional) – verbessert den Nacharbeitungs-Prompt
    questions_map: dict = {}
    if questions_path and os.path.exists(questions_path):
        with open(questions_path, encoding='utf-8') as qf:
            for line in qf:
                qm = re.match(r'^(\d+)\.\s+(.+)', line.strip())
                if qm:
                    questions_map[int(qm.group(1))] = qm.group(2)

    # Originaltext nach Sektionen indizieren (für gezieltes Lookup via Textgrundlage)
    sections = parse_sections(working_text)
    sec_lookup: dict = {}
    for s in sections:
        if s['heading'] == '__preamble__':
            continue
        sec_lookup[normalize_heading(s['heading'])] = s['body']

    def _section_context(textgrundlage: str) -> str:
        key = normalize_heading(textgrundlage)
        if key in sec_lookup:
            return sec_lookup[key]
        last = normalize_heading(textgrundlage.split('.')[-1])
        if last in sec_lookup:
            return sec_lookup[last]
        return working_text  # Fallback: gesamtes Kapitel

    supplement_map: dict = {}
    for it in todo:
        context = _section_context(it.get('textgrundlage', ''))
        q_text = questions_map.get(it['num'], '')
        prompt = (
            "Rolle:\nDu bist Lerncoach und Prüfer für Wirtschaftspsychologie.\n\n"
            "Situation:\nEine leseleitende Frage wurde anhand einer Zusammenfassung nur "
            f"'{it.get('abdeckung')}' beantwortet. Prüfe den nachstehenden Originaltext und "
            "ergänze ausschließlich die fehlende, prüfungsrelevante Kerninformation.\n\n"
            "Regeln:\n"
            "- Maximal 3 Sätze.\n"
            "- Nur Inhalte, die der Originaltext belegt – nicht raten, nichts erfinden.\n"
            "- Keine Wiederholung der bereits vorhandenen Antwort.\n"
            f"- Wenn der Originaltext die Lücke nicht schließt, antworte exakt: {_NO_SUPPLEMENT_MARKER}\n\n"
            f"Frage:\n{q_text or '(Fragetext nicht verfügbar – siehe bisherige Antwort)'}\n\n"
            f"Bisherige Antwort:\n{it.get('antwort', '')}\n\n"
            f"Originaltext (Textgrundlage '{it.get('textgrundlage', '')}'):\n{context}"
        )
        try:
            response = call_gemini_with_retry(
                model_name='gemini-2.5-pro',
                contents=prompt,
                config=types.GenerateContentConfig(temperature=0.1)
            )
            supplement = (response.text or '').strip()
        except Exception as e:
            print(f"  Frage {it['num']}: Nacharbeit fehlgeschlagen ({e}) – unverändert.")
            continue

        if not supplement or _NO_SUPPLEMENT_MARKER in supplement:
            print(f"  Frage {it['num']}: keine Ergänzung im Originaltext gefunden.")
            continue

        it['antwort'] = f"{it.get('antwort', '').rstrip()} {supplement}".strip()
        it['abdeckung'] = "vollständig durch Nachbearbeitung"
        tg = it.get('textgrundlage', '')
        for key in {normalize_heading(tg), normalize_heading(tg.split('.')[-1])}:
            supplement_map.setdefault(key, []).append(supplement)
        print(f"  Frage {it['num']}: ergänzt → vollständig durch Nachbearbeitung.")

    return serialize_qa_items(items), supplement_map


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


# Kasten-Labels für eingeschobene Lehrbuch-Boxen (Fokus/Studie/Definition/Beispiel/Exkurs).
_BOX_LABEL_RE = re.compile(r'^(fokus|studie|beispiel|definition|exkurs)\b', re.IGNORECASE)


def _is_box_heading(heading: str) -> bool:
    """True wenn die Überschrift ein eingeschobener Lehrbuch-Kasten ist (Fokus:/Studie: …)."""
    return bool(_BOX_LABEL_RE.match(_clean_heading_text(heading)))


def _split_paragraphs(body: str) -> list:
    """Body in Absätze splitten (Trenner: Leerzeile). Leere Absätze werden entfernt."""
    return [p.strip() for p in re.split(r'\n\s*\n', body) if p.strip()]


def _classify_box_boundaries(boxes: list) -> list:
    """
    LLM-Klassifikation: Für jeden Kasten die Anzahl der FÜHRENDEN Absätze, die wirklich
    zum Kasten gehören (Rest = irrtümlich von der OCR angehängter Fließtext des
    Elternabschnitts). boxes: [{'title': str, 'paragraphs': [str, ...]}, ...].
    Rückgabe: list[int] gleicher Länge, jeweils geclampt auf [1, len(paragraphs)].
    """
    lines = [
        "Du erhältst Textkästen aus einem Lehrbuch. Durch die OCR wurde an jeden Kasten "
        "fälschlich nachfolgender Fließtext angehängt, der zum umgebenden Kapitel gehört.\n",
        "Aufgabe: Bestimme für jeden Kasten, wie viele der ERSTEN Absätze tatsächlich zum "
        "Kasten gehören (thematisch zum Kastentitel passen). Die restlichen Absätze sind "
        "irrtümlich angehängter Fließtext.\n",
        "Regeln:\n- Ändere keinen Text.\n- Mindestens 1 Absatz gehört zum Kasten.\n"
        '- Gib NUR JSON zurück, Format {"0": n0, "1": n1, ...} (Kastenindex → Anzahl).\n',
        "Kästen:",
    ]
    for bi, box in enumerate(boxes):
        lines.append(f'\n=== Kasten {bi}: "{box["title"]}" ===')
        for pi, para in enumerate(box['paragraphs'], 1):
            lines.append(f'[Absatz {pi}] {para.replace(chr(10), " ")}')
    prompt = '\n'.join(lines)

    response = call_gemini_with_retry(
        model_name='gemini-2.5-pro',
        contents=prompt,
        config=types.GenerateContentConfig(temperature=0, response_mime_type='application/json'),
    )
    raw = response.text.strip()
    raw = re.sub(r'^```(?:json)?|```$', '', raw, flags=re.MULTILINE).strip()
    data = json.loads(raw)

    result = []
    for bi, box in enumerate(boxes):
        n = int(data.get(str(bi), 1))
        result.append(max(1, min(n, len(box['paragraphs']))))
    return result


def repair_box_structure(text: str) -> str:
    """
    Trennt eingeschobene Lehrbuch-Kästen (Fokus/Studie/Definition/Beispiel) von dem
    Fließtext, den die OCR fälschlich in den Kasten-Body gezogen hat, und führt diesen
    Fließtext zum Elternabschnitt zurück. Jeder Kasten wird dadurch ein eigenständiges,
    vollständiges Lernobjekt; der Elterntext bleibt zusammenhängend.

    Ohne Kästen → Originaltext (kein API-Call). Bei API-/Parse-Fehler → Originaltext
    unverändert (kein Crash, kein Datenverlust).
    """
    sections = parse_sections(text)
    box_indices = [
        i for i, s in enumerate(sections)
        if s['heading'] != '__preamble__' and _is_box_heading(s['heading'])
    ]
    if not box_indices:
        return text

    boxes = [
        {'title': _clean_heading_text(sections[i]['heading']),
         'paragraphs': _split_paragraphs(sections[i]['body'])}
        for i in box_indices
    ]
    try:
        counts = _classify_box_boundaries(boxes)
    except Exception as e:
        print(f"   [Box-Reparatur übersprungen] Klassifikation fehlgeschlagen: {e}")
        return text

    count_map = dict(zip(box_indices, counts))

    out_blocks: list[str] = []
    parent: dict | None = None  # {'head': str|None, 'paras': list[str], 'boxes': list[str]}

    def _flush(p):
        if p is None:
            return
        parts = []
        if p['head'] is not None:
            parts.append(p['head'])
        if p['paras']:
            parts.append('\n\n'.join(p['paras']))
        block = '\n\n'.join(parts).strip()
        if block:
            out_blocks.append(block)
        out_blocks.extend(p['boxes'])

    for i, s in enumerate(sections):
        head_line = None if s['heading'] == '__preamble__' else '#' * s['level'] + ' ' + s['heading']

        if i in count_map:
            paras = _split_paragraphs(s['body'])
            n = count_map[i]
            box_content, running = paras[:n], paras[n:]
            box_parts = [head_line] if head_line else []
            if box_content:
                box_parts.append('\n\n'.join(box_content))
            box_block = '\n\n'.join(box_parts).strip()
            if parent is not None:
                parent['paras'].extend(running)
                if box_block:
                    parent['boxes'].append(box_block)
            else:
                # Kasten ohne vorangehenden Elternabschnitt: eigenständig belassen.
                if box_block:
                    out_blocks.append(box_block)
                if running:
                    out_blocks.append('\n\n'.join(running))
        else:
            _flush(parent)
            parent = {'head': head_line, 'paras': _split_paragraphs(s['body']), 'boxes': []}

    _flush(parent)
    return '\n\n'.join(out_blocks).strip() + '\n'


def normalize_heading(h: str) -> str:
    """Für Matching: Bold-/Italic-Marker entfernen, lowercase. Nummern bleiben für eindeutige Keys."""
    h = re.sub(r'\*+', '', h)   # strip ** und *
    return h.lower().strip()


def normalize_heading_levels(text: str) -> str:
    """
    Normalisiert inkonsistente Markdown-Überschriftenebenen.
    Nummerierte Kapitelüberschriften erhalten konsistente Ebenen:
      "1 Titel"     → # (H1)
      "4.1 Titel"   → ## (H2)
      "4.1.1 Titel" → ### (H3)
    Nicht-nummerierte Überschriften werden nicht verändert.
    Tail-nummerierte Überschriften ("Titel 7.2.2") werden umgeordnet und
    ebenfalls normalisiert – mindestens 2 Punkte erforderlich (schließt
    einfache Abbildungsnummern wie "18.4" aus).
    """
    result = []
    for line in text.split('\n'):
        m = re.match(r'^(#{1,6})\s+(.+)$', line)
        if m:
            content = m.group(2).strip()
            clean = content.replace('**', '').strip()
            clean_nohtml = re.sub(r'<[^>]+>', '', clean)
            num_m = re.match(r'^(\d+(?:\.\d+)*)(\s|$)', clean_nohtml)
            if num_m:
                dots = num_m.group(1).count('.')
                new_level = '#' * min(dots + 1, 6)
                result.append(f'{new_level} {content}')
            else:
                # OCR-Artefakt: Kapitelnummer am Ende ("Titel 7.2.2" → "7.2.2 Titel").
                # Mindestens 2 Punkte (N.M.P) erforderlich – schließt Jahre und einfache
                # Abbildungsnummern aus ("Studie 2023", "Abbildung 18.4" bleiben unverändert).
                tail_m = re.match(r'^(.+?)\s+(\d+(?:\.\d+){2,})\s*$', clean_nohtml)
                if tail_m:
                    dots = tail_m.group(2).count('.')
                    new_level = '#' * min(dots + 1, 6)
                    reordered = f"{tail_m.group(2)} {tail_m.group(1).rstrip()}"
                    result.append(f'{new_level} {reordered}')
                else:
                    result.append(line)
        else:
            result.append(line)
    return '\n'.join(result)


def add_formatted_text(paragraph, text, default_color=None):
    """Parse Markdown: ***bold+italic***, **bold**, *italic*."""
    parts = re.split(r'(\*{3}[^*]+?\*{3}|\*{2}[^*]+?\*{2}|\*[^*]+?\*)', text)
    for part in parts:
        if part.startswith('***') and part.endswith('***') and len(part) > 6:
            run = paragraph.add_run(part[3:-3])
            run.bold = True
            run.italic = True
        elif part.startswith('**') and part.endswith('**') and len(part) > 4:
            run = paragraph.add_run(part[2:-2])
            run.bold = True
        elif part.startswith('*') and part.endswith('*') and len(part) > 2:
            run = paragraph.add_run(part[1:-1])
            run.italic = True
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
    Bullet-Concatenation (•A•B•C) → separate Absätze je Eintrag.
    """
    font_size = font_size or Pt(9.5)
    raw_text = _clean_unklar_cell(raw_text)
    raw_text = _html_entities(raw_text)

    if not bold and raw_text.count('•') > 1:
        bullet_parts = [p.strip() for p in raw_text.split('•') if p.strip()]
        raw_text = '\n'.join(bullet_parts)

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


_STATS_RE = re.compile(r'^(\d+\.\d+)\s+(\d+\.\d+)\s+(\.?\d+\.?\d*)$')

def _fix_concatenated_stats_cells(rows_data: list) -> list:
    """
    Korrigiert OCR-Artefakt: M, SD und ritc werden manchmal in einer Zelle
    zusammengefasst (z.B. '3.44 1.00 .49'). Wenn links und rechts der
    betreffenden Zelle leere Nachbarzellen existieren, werden die drei Werte
    auf die drei Zellen verteilt.
    """
    for row in rows_data:
        for j in range(len(row)):
            m = _STATS_RE.match(row[j].strip())
            if not m:
                continue
            # Prüfen ob Nachbarzellen leer sind (links+rechts oder nur rechts+rechts)
            if j >= 1 and j + 1 < len(row) and row[j - 1] == '' and row[j + 1] == '':
                row[j - 1], row[j], row[j + 1] = m.group(1), m.group(2), m.group(3)
            elif j + 2 < len(row) and row[j + 1] == '' and row[j + 2] == '':
                row[j], row[j + 1], row[j + 2] = m.group(1), m.group(2), m.group(3)
    return rows_data


def _drop_empty_columns(rows_data: list) -> list:
    """Entfernt Spalten, die in allen Zeilen leer sind (OCR-Artefakt)."""
    if not rows_data:
        return rows_data
    num_cols = max(len(r) for r in rows_data)
    keep = [
        c for c in range(num_cols)
        if any(c < len(r) and r[c].strip() for r in rows_data)
    ]
    return [[r[c] if c < len(r) else '' for c in keep] for r in rows_data]


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

    # Zellinhalt normieren (alle Zeilen auf gleiche Spaltenanzahl)
    for r in rows_data:
        while len(r) < num_cols:
            r.append('')

    rows_data = _fix_concatenated_stats_cells(rows_data)
    rows_data = _drop_empty_columns(rows_data)
    num_cols = max(len(r) for r in rows_data)
    num_rows = len(rows_data)

    # Titel-Zeile erkennen: Zeile 0 hat ≤1 Non-Empty-Cell bei mehreren Spalten
    non_empty_row0 = sum(1 for c in rows_data[0] if c.strip())
    has_title_row = (num_rows > 1 and num_cols > 1 and non_empty_row0 <= 1)
    header_row_idx = 1 if has_title_row else 0

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
            is_header = (r_idx == header_row_idx)
            is_title  = (has_title_row and r_idx == 0)
            _set_cell_text(cell, cell_text, bold=(is_header or is_title), font_size=font_size)
            if is_header:
                _shade_cell(cell, 'D9E2F3')

    # Titelzeile: alle Zellen in Zeile 0 zusammenführen
    if has_title_row:
        anchor = table.cell(0, 0)
        for c in range(1, num_cols):
            anchor.merge(table.cell(0, c))

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


def _clean_for_hidden(text: str) -> str:
    """Bereinigt Markdown-Text für den Hidden-Originaltext-Block.
    Entfernt Links, HTML-Tags, Seitenreferenzen und andere OCR-Artefakte.
    Bilder (![]()) bleiben erhalten."""
    text = re.sub(r'(?<!!)\[([^\]]*)\]\([^)]*\)', r'\1', text)  # [Text](url) → Text (nicht Bilder)
    text = re.sub(r'<[^>]+>', '', text)                          # HTML-Tags entfernen
    text = re.sub(r'^\[\d+\]\s*$', '', text, flags=re.MULTILINE)  # [1] allein → weg
    return text


def _strip_kontrollliste(text: str) -> str:
    """Entfernt Kontrolllisten-Blöcke aus der übersetzten Markdown-Datei.
    Muster: (optionales ---) gefolgt von **Kontrollliste** und Bullet-Zeilen."""
    text = re.sub(
        r'(?m)^---\s*\n\*\*Kontrollliste\*\*.*?(?=\n#{1,6}\s|\Z)',
        '',
        text,
        flags=re.DOTALL
    )
    # Kontrollliste ohne vorangehendes ---
    text = re.sub(
        r'(?m)^\*\*Kontrollliste\*\*.*?(?=\n#{1,6}\s|\Z)',
        '',
        text,
        flags=re.DOTALL
    )
    return text


def _strip_selbstpruefung(text: str) -> str:
    """Entfernt KI-interne Selbstprüfungs-Blöcke aus der Zusammenfassung.
    Diese Abschnitte sind KI-Validierungsartefakte und kein Lerninhalt.
    Unterstützt Varianten: '### **Selbstprüfung**', '**Selbstprüfung**', 'Selbstprüfung'."""
    # Vorangehendes --- wegstreifen (erscheint typischerweise direkt vor dem Block)
    text = re.sub(
        r'(?m)^---\s*\n(?=(?:#{1,6}\s+)?\*{0,3}\*?\*?[Ss]elbstpr)',
        '',
        text
    )
    text = re.sub(
        r'(?m)^(?:#{1,6}\s+)?\*{0,3}\*?\*?[Ss]elbstpr[uü]fung\*?\*?\*?.*?(?=\n#{1,6}\s|\Z)',
        '',
        text,
        flags=re.DOTALL
    )
    return text


def _compress_heading_levels(text: str) -> str:
    """
    Korrigiert inkonsistente OCR-Heading-Ebenen.
    Verhindert Sprünge > 1 Ebene; Geschwister-Überschriften (gleiche OCR-Ebene)
    erhalten denselben tatsächlichen Level – kein Kaskadeneffekt.
    """
    ocr_to_actual: dict = {}
    prev_actual = 0
    result = []
    for line in text.split('\n'):
        m = re.match(r'^(#{1,9})\s', line)
        if m:
            ocr_level = len(m.group(1))
            # Nummerierte Headings wurden bereits von normalize_heading_levels korrekt gesetzt
            if re.match(r'^#{1,9}\s+\*{0,2}\d', re.sub(r'<[^>]+>', '', line)):
                result.append(line)
                prev_actual = ocr_level
                continue
            if ocr_level in ocr_to_actual:
                actual = ocr_to_actual[ocr_level]
            elif ocr_level > prev_actual + 1:
                actual = prev_actual + 1
                ocr_to_actual[ocr_level] = actual
            else:
                actual = ocr_level
                ocr_to_actual[ocr_level] = actual
            prev_actual = actual
            result.append('#' * actual + line[ocr_level:])
        else:
            result.append(line)
    return '\n'.join(result)


def _strip_ocr_y_prefix(text: str) -> str:
    """Entfernt isoliertes 'y '-Präfix aus Headings und Bullet-Items.
    OCR-Artefakt: Sonderzeichen (Bullet-Pfeil) einer Sonderzeichenschrift wird als 'y' gelesen.
    Bereinigt außerdem trailing \\_ Sequenzen in Heading-Zeilen (OCR-Artefakt für Unterstriche)."""
    lines = []
    for line in text.split('\n'):
        if re.match(r'^#{1,6}\s', line):
            # Trailing \_ Sequenzen aus Heading-Text entfernen (z.B. "# Titel: \_ \_ \_")
            line = re.sub(r'(\s*\\_)+\s*$', '', line).rstrip()
        # Heading: ## y Titel  ODER  ## **y Titel** (y innerhalb Bold-Marker, wie KI es erzeugt)
        line = re.sub(r'^(#{1,6}\s+)(\**)y\s+', r'\1\2', line)
        # Bullet: beliebige Einrückung + Bullet-Marker + y (inkl. eingerückte Sublisten)
        line = re.sub(r'^(\s*[-*]\s+)y\s+', r'\1', line)
        lines.append(line)
    return '\n'.join(lines)


def _image_display_width(img_path: str, max_in: float = 5.5, assumed_dpi: int = 150):
    """Berechnet Anzeigebreite: min(natürliche Bildbreite, max_in).
    Nutzt PIL/Pillow falls verfügbar; andernfalls Dateigröße als Heuristik."""
    try:
        from PIL import Image as _PILImage
        with _PILImage.open(img_path) as im:
            natural_in = im.width / assumed_dpi
            return Inches(min(natural_in, max_in))
    except Exception:
        pass
    try:
        kb = os.path.getsize(img_path) / 1024
        if kb < 10:
            return Inches(1.5)
        if kb < 50:
            return Inches(3.0)
    except Exception:
        pass
    return Inches(max_in)


def _is_decorative_image(img_path: str, min_px: int = 100) -> bool:
    """True wenn das Bild wahrscheinlich dekorativ ist (Icon, Randsymbol).
    Kriterium: kleinste Seite < min_px (alle Kronen-Icons haben min ≤ 100 px;
    Inhaltsbilder wie 'Was ist das?'-Grafiken haben min > 100 px).
    Fallback ohne PIL: Dateigröße < 10 KB."""
    try:
        from PIL import Image as _PILImage
        with _PILImage.open(img_path) as im:
            w, h = im.width, im.height
            return min(w, h) < min_px
    except Exception:
        pass
    try:
        return os.path.getsize(img_path) < 10 * 1024
    except Exception:
        return False


_CAPTION_RE = re.compile(
    r'^(\*\*)?(TABELLE|Tabelle|ABBILDUNG|Abbildung|TABLE|FIGURE|Figure|Abb\.|Tab\.)\b',
    re.IGNORECASE
)
_HRULE_RE = re.compile(r'^-{3,}$|^\*{3,}$|^_{3,}$')


def _hide_paragraph(p, indent_cm: float = 1.5):
    """Setzt alle Runs eines Absatzes auf hidden + fügt Einrückung hinzu.
    Entfernt w:numPr (Bullet-Label) und setzt w:vanish in w:pPr/w:rPr."""
    p.paragraph_format.left_indent = Cm(indent_cm)
    for run in p.runs:
        run.font.hidden = True
    pPr = p._p.get_or_add_pPr()
    numPr = pPr.find(qn('w:numPr'))
    if numPr is not None:
        pPr.remove(numPr)
    rPr = pPr.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        pPr.append(rPr)
    vanish = OxmlElement('w:vanish')
    rPr.append(vanish)


def process_markdown_to_docx(doc, block_text, hide_text=False, base_path=None,
                             skip_images=False, headings_as_bold=False):
    """
    Interpretiert Markdown und fügt Inhalte dem Word-Dokument hinzu.
    - Tabellen und Bilder: IMMER sichtbar (nie ausgeblendet).
    - Abbildungs-/Tabellenbeschriftungen: sichtbar, auch wenn hide_text=True.
    - hide_text=True: Text wird hidden (nicht druckbar) + eingerückt; grau für Sichtbarkeit.
    - headings_as_bold=True: Markdown-Headings als fetter Normal-Text statt Word-Heading-Style
      (verhindert Einträge in der Navigationsleiste bei Summary-Body-Text).
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

    # Im Hidden-Block: Links, HTML und OCR-Artefakte bereinigen
    if hide_text:
        block_text = _clean_for_hidden(block_text)

    lines = block_text.split('\n')
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # ── HTML-Tags aus Fließtext entfernen (außer Bildzeilen) ──
        if not stripped.startswith('!'):
            stripped = re.sub(r'<[^>]+>', '', stripped)

        # ── Horizontale Linie überspringen (--- / *** / ___) ──
        if _HRULE_RE.match(stripped):
            i += 1
            continue

        # ── Tabelle ──
        if stripped.startswith('|'):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                table_lines.append(lines[i].strip())
                i += 1
            add_markdown_table_to_doc(doc, table_lines)
            continue

        # ── Bild: immer sichtbar (außer wenn skip_images=True oder dekorativ) ──
        img_m = re.match(r'^!\[([^\]]*)\]\(([^)]+)\)$', stripped)
        if img_m:
            if skip_images:
                i += 1
                continue
            img_rel = img_m.group(2)
            img_path = os.path.join(base_path, img_rel) if base_path else img_rel
            if os.path.exists(img_path):
                if _is_decorative_image(img_path):
                    # Icons, Logos, Randsymbole überspringen
                    i += 1
                    continue
                try:
                    doc.add_picture(img_path, width=_image_display_width(img_path))
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

        # ── Sub-Bullets (2+ führende Leerzeichen, Strich/Stern) ──
        sub_m = re.match(r'^(\s{2,})([-*])\s+(.+)', line)
        if sub_m:
            depth = len(sub_m.group(1)) // 2
            style = 'List Bullet 3' if depth >= 2 else 'List Bullet 2'
            p = doc.add_paragraph(style=style)
            color = color_map['hidden_text'] if do_hide else None
            add_formatted_text(p, sub_m.group(3), default_color=color)
            p.paragraph_format.space_after = Pt(0)
            if do_hide:
                _hide_paragraph(p)
            i += 1
            continue

        # ── Nummerierte Sub-Items (2+ führende Leerzeichen + Zahl) ──
        num_sub_m = re.match(r'^(\s{2,})(\d+)[.)]\s+(.+)', line)
        if num_sub_m:
            depth = len(num_sub_m.group(1)) // 2
            style = 'List Bullet 3' if depth >= 2 else 'List Bullet 2'
            p = doc.add_paragraph(style=style)
            color = color_map['hidden_text'] if do_hide else None
            add_formatted_text(p, num_sub_m.group(3), default_color=color)
            p.paragraph_format.space_after = Pt(0)
            if do_hide:
                _hide_paragraph(p)
            i += 1
            continue

        # ── Überschriften ──
        if stripped.startswith('#### '):
            hdg_text = stripped[5:]
            if headings_as_bold and not do_hide:
                p = doc.add_paragraph(style='Normal')
                p.add_run(hdg_text).bold = True
            else:
                p = doc.add_heading(level=4)
                color = color_map['hidden_heading'] if do_hide else color_map['heading4']
                add_formatted_text(p, hdg_text, default_color=color)
                if do_hide:
                    _hide_paragraph(p)
        elif stripped.startswith('### '):
            hdg_text = stripped[4:]
            if headings_as_bold and not do_hide:
                p = doc.add_paragraph(style='Normal')
                p.add_run(hdg_text).bold = True
            else:
                p = doc.add_heading(level=3)
                color = color_map['hidden_heading'] if do_hide else color_map['heading3']
                add_formatted_text(p, hdg_text, default_color=color)
                if do_hide:
                    _hide_paragraph(p)
        elif stripped.startswith('## '):
            hdg_text = stripped[3:]
            if headings_as_bold and not do_hide:
                p = doc.add_paragraph(style='Normal')
                p.add_run(hdg_text).bold = True
            else:
                p = doc.add_heading(level=2)
                color = color_map['hidden_heading'] if do_hide else color_map['heading2']
                add_formatted_text(p, hdg_text, default_color=color)
                if do_hide:
                    _hide_paragraph(p)
        elif stripped.startswith('# '):
            hdg_text = stripped[2:]
            if headings_as_bold and not do_hide:
                p = doc.add_paragraph(style='Normal')
                p.add_run(hdg_text).bold = True
            else:
                p = doc.add_heading(level=1)
                color = color_map['hidden_heading'] if do_hide else color_map['heading1']
                add_formatted_text(p, hdg_text, default_color=color)
                if do_hide:
                    _hide_paragraph(p)
        # ── Aufzählung ──
        elif stripped.startswith('* ') or stripped.startswith('- '):
            bullet_m = re.match(r'^[*-]\s+(.*)', stripped)
            bullet_content = bullet_m.group(1) if bullet_m else stripped[2:]
            if not bullet_content.strip():
                i += 1
                continue
            p = doc.add_paragraph(style='List Bullet')
            color = color_map['hidden_text'] if do_hide else None
            add_formatted_text(p, bullet_content, default_color=color)
            p.paragraph_format.space_after = Pt(0)
            if do_hide:
                _hide_paragraph(p)
        # ── Normaler Text ──
        else:
            p = doc.add_paragraph()
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            color = color_map['hidden_text'] if do_hide else None
            add_formatted_text(p, stripped, default_color=color)
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
    'history', 'historie',
])

_SKIP_HEADING_PATTERNS = ('m.sc', 'prof. dr', '@')

def _is_skip_heading(heading: str) -> bool:
    """Gibt True zurück, wenn diese Überschrift standardmäßig übersprungen werden soll."""
    key = normalize_heading(heading)
    if key in _SKIP_HEADINGS:
        return True
    return any(pat in key for pat in _SKIP_HEADING_PATTERNS)


# Front-Matter-Überschriften am Dokumentanfang (Impressum, Inhaltsverzeichnis),
# die standardmäßig übersprungen werden. Titel + Titelbild davor bleiben erhalten.
_FRONT_MATTER_HEADINGS = frozenset([
    'impressum', 'imprint',
    'inhalt', 'inhaltsverzeichnis', 'table of contents', 'contents',
])


def strip_front_matter(text: str) -> str:
    """Entfernt Front-Matter-Blöcke (Impressum, Inhaltsverzeichnis) am Dokumentanfang.

    Titel (erste Überschrift) und ggf. Titelbild davor bleiben erhalten. Ein Block läuft
    von seiner Überschrift bis zur nächsten gleich-/höherrangigen Überschrift, die selbst
    kein Front-Matter ist; tiefere Unterabschnitte des Blocks werden mit entfernt.
    Arbeitet zeilenbasiert, um HTML-Spans und Bild-Referenzen unverändert zu lassen.
    """
    out: list[str] = []
    skipping = False
    skip_level = 0
    for line in text.split('\n'):
        m = re.match(r'^(#{1,6})\s+(.*)$', line)
        if m:
            level = len(m.group(1))
            key = normalize_heading(_clean_heading_text(line))
            if skipping and level <= skip_level:
                # Block endet bei gleich-/höherrangiger Überschrift.
                skipping = False
            if key in _FRONT_MATTER_HEADINGS:
                skipping = True
                skip_level = level
                continue
            if skipping and level > skip_level:
                continue  # tieferes Front-Matter-Unterkapitel
        if skipping:
            continue
        out.append(line)
    return '\n'.join(out)


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


def _normalize_chapter_num(num: str) -> str:
    """Säubert eine (evtl. OCR-verschmutzte) Kapitelnummer.
    Kollabiert Mehrfachpunkte ('3...1' → '3.1') und entfernt führende/abschließende Punkte
    ('1.' → '1'). Liefert reine Punkt-getrennte Ziffernfolgen.
    """
    num = re.sub(r'\.{2,}', '.', num)   # 3...1 → 3.1
    return num.strip('.')               # 1.  → 1 ,  .1 → 1


def _join_number_and_rest(num: str, rest: str) -> str:
    """Fügt eine Kapitelnummer und den Resttext sauber zusammen (genau ein Trenn-Whitespace
    bei vorhandenem Titel, kein Trailing-Whitespace bei reiner Nummer)."""
    rest = rest or ""
    if rest and not rest[:1].isspace():
        rest = f" {rest}"
    return f"{num}{rest}"


def _prefix_chapter_number(heading_text: str, prefix: str) -> str:
    """
    Setzt den Elternpräfix vor die Kapitelnummer einer Überschrift.
    '1 EINLEITUNG'  + '4.2.1.2' → '4.2.1.2.1 EINLEITUNG'
    '4.1 Abschnitt' + '4.2.1.2' → '4.2.1.2.4.1 Abschnitt'
    '1. EINLEITUNG' + '6.2.4.1' → '6.2.4.1.1 EINLEITUNG'  (kein Trailing-Punkt)
    Überschriften ohne führende Zahl bleiben unverändert.
    """
    m = re.match(r'^(\d[\d\.]*)(.*)', heading_text.strip())
    if m:
        num = _normalize_chapter_num(m.group(1))
        return _join_number_and_rest(f"{prefix}.{num}", m.group(2))
    return heading_text


def _rebase_chapter_number(heading_text: str, chapter_root: str, parent_chapter: str) -> str:
    """
    Ersetzt die Nummer eines extrahierten Kapitels durch das Elternkapitel.
    chapter_root='1.3', parent='2.4.1.1':
      '1.3 Titel'   → '2.4.1.1 Titel'
      '1.3.2 Titel' → '2.4.1.1.2 Titel'
      '1.3'         → '2.4.1.1'        (auch reine Nummer ohne Titel)
    Nicht-Nachfahren werden regulär präfixiert (Fallback).
    """
    m = re.match(r'^(\d[\d.]*?)(\s.*|$)', heading_text.strip())
    if not m:
        return heading_text
    num, rest = _normalize_chapter_num(m.group(1)), m.group(2)
    if num == chapter_root:
        return _join_number_and_rest(parent_chapter, rest)
    if num.startswith(chapter_root + '.'):
        return _join_number_and_rest(f"{parent_chapter}.{num[len(chapter_root) + 1:]}", rest)
    return _join_number_and_rest(f"{parent_chapter}.{num}", rest)   # Fallback


def build_translation_word_document(translated_text: str, output_path: str, base_path: str = None):
    """Erstellt ein Word-Dokument aus dem übersetzten Text.
    Preamble (Abstract) eingerückt; Skip-Headings (Referenzen, Historie, Autor) gefiltert.
    """
    print(f"--- Erstelle Übersetzungs-Word-Dokument -> {output_path} ---")

    text = _strip_kontrollliste(translated_text)
    text = normalize_heading_levels(text)
    text = _strip_ocr_y_prefix(text)
    text = _compress_heading_levels(text)
    sections = parse_sections(text)

    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    for section in sections:
        heading = section['heading']
        body    = section['body']
        level   = section['level']

        if heading == '__preamble__':
            if body.strip():
                before = len(doc.paragraphs)
                process_markdown_to_docx(doc, body, hide_text=False, base_path=base_path, skip_images=True)
                for p in doc.paragraphs[before:]:
                    p.paragraph_format.left_indent = Cm(1.5)
            continue

        clean = _clean_heading_text(heading)
        if _is_skip_heading(clean):
            continue

        h = doc.add_heading(clean, level=level)
        _set_heading_color(h, MM_HEADING_COLORS.get(level, MM_HEADING_COLORS[9]))

        if body.strip():
            process_markdown_to_docx(doc, body, hide_text=False, base_path=base_path)

    doc.save(output_path)
    print("Übersetzungs-Word-Dokument erstellt.")


def build_interleaved_word_document(translated_text: str, summary_text: str, qa_text: str,
                                    output_path: str, base_path: str = None,
                                    parent_chapter: str = None, parent_level: int = None,
                                    skip_references: bool = True,
                                    questions_path: str = None,
                                    doc_title: str = "Lernskript",
                                    extracted_chapter: str = None,
                                    supplement_map: dict = None,
                                    title_as_parent: bool = False):
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

    # Strukturelle Neu-Nummerierung: Im reinen Einbettungs-Modus (parent_chapter, KEIN
    # extrahiertes Einzelkapitel) wird die gesamte Artikel-Gliederung allein aus der
    # Überschriften-Hierarchie neu nummeriert. Die OCR-Eigennummern (z.B. "1.", "3.") sind
    # nur teilweise vorhanden und inkonsistent → sie werden verworfen.
    structural_renumber = bool(parent_chapter and not extracted_chapter)

    counters = [0] * 9  # Zähler je Heading-Ebene für Auto-Nummerierung

    # --- QA vorbereiten: Textgrundlage-Map für Kommentare ---
    has_qa = qa_text and qa_text.strip() not in ("", "Keine Leitfragen zur Prüfung übergeben.")
    qa_items = parse_qa_response(qa_text) if has_qa else []
    qa_by_num: dict = {item['num']: item for item in qa_items}  # für Beleg-Verankerung
    textgrundlage_map: dict = {}  # normalize_heading(textgrundlage) → [fragenummern]
    for item in qa_items:
        for key in [normalize_heading(item['textgrundlage']),
                    normalize_heading(item['textgrundlage'].split('.')[-1])]:
            textgrundlage_map.setdefault(key, [])
            if item['num'] not in textgrundlage_map[key]:
                textgrundlage_map[key].append(item['num'])

    comment_list: list = []   # [(comment_id, comment_text)]
    comment_id: list  = [0]  # mutable int-wrapper

    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # --- Lernskript-Titel (Standalone-Modus) ---
    if not parent_chapter:
        h = doc.add_heading(doc_title, level=1)
        _set_heading_color(h, MM_HEADING_COLORS[1])

    # --- Heading-Level normalisieren & Sections parsen ---
    translated_text = _strip_kontrollliste(translated_text)
    translated_text = normalize_heading_levels(translated_text)
    translated_text = _strip_ocr_y_prefix(translated_text)
    summary_text    = _strip_kontrollliste(summary_text)
    summary_text    = _strip_selbstpruefung(summary_text)
    summary_text    = normalize_heading_levels(summary_text)
    summary_text    = _strip_ocr_y_prefix(summary_text)
    orig_sections   = parse_sections(translated_text)
    sum_sections    = parse_sections(summary_text)

    # Lookup: normalisierter Heading → Zusammenfassungstext
    sum_lookup = {}
    for s in sum_sections:
        if s['heading'] == '__preamble__':
            continue
        sum_lookup[normalize_heading(_clean_heading_text(s['heading']))] = s['body']

    # Scoped lookup: parent_key → {child_key → body}
    # Verhindert Conflation gleichnamiger Sub-Headings (z.B. "Off the job" unter jeder Kompetenz).
    # Ohne Scoping liefert sum_lookup["off the job"] immer den letzten Eintrag (letzte Kompetenz).
    sum_scoped: dict[str, dict[str, str]] = {}
    _scp_parent: str | None = None
    for _ss in sum_sections:
        if _ss['heading'] == '__preamble__':
            continue
        _ssk = normalize_heading(_clean_heading_text(_ss['heading']))
        if _ss['level'] <= 2:
            _scp_parent = _ssk
            sum_scoped[_ssk] = {'__self__': _ss['body']}
        elif _scp_parent is not None:
            sum_scoped[_scp_parent][_ssk] = _ss['body']

    # Dokument-Typ bestimmen: hat es nummerierte Kapitel?
    has_numbered_chapters = any(
        re.match(r'^\d', _clean_heading_text(s['heading']).strip())
        for s in orig_sections
        if s['heading'] != '__preamble__'
    )

    # Häufigkeit unnummerierter Überschriften zählen.
    # Einzigartige (freq=1) sind echte Kapitelthemen (z.B. "Agilität", "Analytisches Denken").
    # Wiederkehrende (freq>1) sind Frage-Muster unter jedem Unterkapitel ("Was ist das?").
    _unnumbered_freq: dict[str, int] = {}
    for _s in orig_sections:
        if _s['heading'] == '__preamble__':
            continue
        _k = normalize_heading(_clean_heading_text(_s['heading']))
        if not re.match(r'^\d', _k):
            _unnumbered_freq[_k] = _unnumbered_freq.get(_k, 0) + 1
    _auto_parent: str | None = None  # z. B. "5.3" – aktuelles nummeriertes Elternkapitel
    _auto_counter: int = 0

    # Vorberechnung: hat jede Section mindestens einen Nachfolger mit Summary-Inhalt?
    # Verhindert sichtbare Leer-Überschriften bei Elternkapiteln, deren Kinder alle
    # kein Summary haben (z.B. "5.3.27 Rückmeldung zu Konfliktverhalten" → nur Off/On the job).
    # Spiegelt dieselbe Kinder-Logik wie has_children:
    #   a) tiefere Ebene gilt immer als Nachfolger
    #   b) nummerierter Parent + unnummerierter Geschwister → konzeptuelles Kind
    #      (z.B. "5.3 Überblick" → "Agilität" auf gleicher Ebene)
    _any_visible_desc: list[bool] = [False] * len(orig_sections)
    for _pi in range(len(orig_sections)):
        _ps = orig_sections[_pi]
        if _ps['heading'] == '__preamble__':
            continue
        _plevel = _ps['level']
        _ps_numbered = bool(re.match(r'^\d', normalize_heading(_clean_heading_text(_ps['heading']))))
        for _ci in range(_pi + 1, len(orig_sections)):
            _cs = orig_sections[_ci]
            _cs_level = _cs['level']
            if _cs_level < _plevel:
                break  # Aufgestiegen = außerhalb des Scopes
            if _cs_level == _plevel:
                # Gleiche Ebene: nur als konzeptuelles Kind wenn parent nummeriert und
                # dieses Section unnummeriert ist (nummerierter → unnummerierter Folger).
                _cs_k = normalize_heading(_clean_heading_text(_cs['heading']))
                _cs_numbered = bool(re.match(r'^\d', _cs_k))
                if _cs_numbered or not _ps_numbered:
                    break  # Nummeriertes Geschwister oder keine konzeptuelle Eltern-Kind-Relation
                # Unnummeriert nach nummeriertem Parent → zählt als Kind, weiter suchen
            _clk = normalize_heading(_clean_heading_text(_cs['heading']))
            if sum_lookup.get(_clk, '').strip():
                _any_visible_desc[_pi] = True
                break

    # Pre-pass: welche Sections werden sichtbar? Nur diese erhalten sequentielle Auto-Nummern.
    # Ohne diesen Pre-pass entstehen Lücken wie 5.3.1, 5.3.5, 5.3.8 weil unsichtbare Sections
    # trotzdem nummeriert werden.
    _visible_section_indices: set[int] = set()
    for _vi, _vs in enumerate(orig_sections):
        if _vs['heading'] == '__preamble__':
            continue
        if re.search(r'^(?:\\_)+\s*$', _vs['body'], re.MULTILINE):
            continue  # Platzhalter-Section
        _vk = normalize_heading(_clean_heading_text(_vs['heading']))
        if sum_lookup.get(_vk, '').strip() or _any_visible_desc[_vi]:
            _visible_section_indices.add(_vi)

    # title_as_parent: Die Wurzel-Sektion (Artikeltitel, Tiefe 0) ist der Anker des
    # Dokuments und trägt direkt das Elternkapitel (z.B. 6.2.4.3). Sie immer als sichtbar
    # markieren, auch ohne eigene Zusammenfassung – sonst bekäme der Titel keine Nummer.
    if structural_renumber and title_as_parent:
        for _ti, _ts in enumerate(orig_sections):
            if _ts['heading'] != '__preamble__':
                _visible_section_indices.add(_ti)
                break

    # Tiefen-Prepass für die strukturelle Neu-Nummerierung: bildet die (inkonsistenten,
    # lückenhaften) OCR-Überschriftenebenen auf konsistente Tiefen 1,2,3… ab – analog zu
    # _compress_heading_levels, aber über die Sektions-Levels statt über Markdown-Zeilen und
    # ohne Sonderbehandlung nummerierter Headings (deren Eigennummern werden ohnehin verworfen).
    depth_map: dict[int, int] = {}
    if structural_renumber and title_as_parent:
        # Variante "Titel = Elternkapitel": Die erste Überschrift (Artikeltitel) ist die
        # permanente Wurzel (Tiefe 0 → trägt direkt das Elternkapitel, z.B. 6.2.4.1). Alle
        # weiteren Überschriften werden per Outline-Stack relativ verschachtelt, wobei die
        # Wurzel NIE vom Stack genommen wird → auch gleichrangige OCR-H1 (Studie 1/2/3) werden
        # zu Kindern des Titels (Tiefe 1), ihre Unterabschnitte zu Tiefe 2 usw.
        _stack: list[int] = []
        _first = True
        for _di, _ds in enumerate(orig_sections):
            if _ds['heading'] == '__preamble__':
                continue
            _lvl = _ds['level']
            if _first:
                depth_map[_di] = 0
                _first = False
                continue
            while _stack and _stack[-1] >= _lvl:
                _stack.pop()
            depth_map[_di] = len(_stack) + 1
            _stack.append(_lvl)
    elif structural_renumber:
        # Standard-Einbettung: bildet die (inkonsistenten, lückenhaften) OCR-Überschriftenebenen
        # auf konsistente Tiefen 1,2,3… ab – analog zu _compress_heading_levels, aber über die
        # Sektions-Levels statt über Markdown-Zeilen und ohne Sonderbehandlung nummerierter
        # Headings (deren Eigennummern werden ohnehin verworfen).
        _ocr_to_actual: dict[int, int] = {}
        _prev_actual = 0
        for _di, _ds in enumerate(orig_sections):
            if _ds['heading'] == '__preamble__':
                continue
            _lvl = _ds['level']
            if _lvl in _ocr_to_actual:
                _actual = _ocr_to_actual[_lvl]
            elif _lvl > _prev_actual + 1:
                _actual = _prev_actual + 1
                _ocr_to_actual[_lvl] = _actual
            else:
                _actual = _lvl
                _ocr_to_actual[_lvl] = _actual
            _prev_actual = _actual
            depth_map[_di] = _actual

    _current_competency_key: str | None = None  # aktuell verarbeitete Kompetenz für Scoped Lookup
    _struct_intro_attached = False  # Einleitungs-Summary im structural_renumber-Modus nur einmal anhängen

    # --- Interleaved Aufbau ---
    for idx, section in enumerate(orig_sections):
        if section['heading'] == '__preamble__':
            sum_preamble = next(
                (s['body'] for s in sum_sections if s['heading'] == '__preamble__'), ''
            )
            # KI-Metatext ("Absolut! Hier ist...") aus Summary-Preamble entfernen
            sum_preamble = re.sub(
                r'^[^\n]*(?:absolut|hier ist|lernorientierte)[^\n]*\n?',
                '', sum_preamble, flags=re.IGNORECASE
            ).strip()
            if sum_preamble:
                process_markdown_to_docx(doc, sum_preamble, hide_text=False, base_path=base_path)
            if section['body']:
                process_markdown_to_docx(doc, section['body'], hide_text=True, base_path=base_path)
            continue

        level     = section['level']
        heading   = section['heading']
        orig_body = section['body']

        clean_heading = _clean_heading_text(heading)
        lookup_key = normalize_heading(clean_heading)  # vor Präfix-Addition für Lookups

        if skip_references and _is_skip_heading(clean_heading):
            continue

        # Platzhalter-Sektionen (OCR-Notizlinien \_\_\_ im Body) komplett überspringen
        if re.search(r'^(?:\\_)+\s*$', orig_body, re.MULTILINE):
            continue

        # Normalisierung: "Titel 7.2.2" → "7.2.2 Titel" (OCR-Artefakt: Kapitelnummer am Zeilenende)
        if extracted_chapter:
            _ec_esc = re.escape(extracted_chapter)
            _m_tail = re.match(r'^(.+?)\s+(\d[\d.]*)\s*$', clean_heading)
            if _m_tail and re.match(rf'^{_ec_esc}\.', _m_tail.group(2)):
                clean_heading = f"{_m_tail.group(2)} {_m_tail.group(1).rstrip()}"

        # In eingebettetem Modus (parent_chapter gesetzt, keine nummerierten Kapitel):
        # Level-1-Headings sind OCR-Artefakte (Artikeltitel, Journal-Metadaten).
        # Originaltext als versteckten Text erhalten, aber KEINE Nav-Überschrift erzeugen,
        # damit der Counter für level-2+ sauber bei 1 beginnt (5.1.1, 5.1.2 ...).
        if parent_chapter and not structural_renumber and not has_numbered_chapters and level == 1:
            if len(orig_body.strip()) > 800:
                # Substanzieller Einleitungstext (Abstract/Intro): Summary anzeigen wenn vorhanden.
                _skipped_sum = sum_lookup.get(lookup_key, '') or sum_lookup.get('einleitung', '')
                if _skipped_sum.strip():
                    process_markdown_to_docx(doc, _skipped_sum, hide_text=False, base_path=base_path)
            if orig_body.strip():
                process_markdown_to_docx(doc, orig_body, hide_text=True, base_path=base_path)
            continue

        if structural_renumber:
            # OCR-Eigennummer verwerfen und allein aus der Hierarchie-Tiefe neu nummerieren.
            _title = re.sub(r'^\d[\d.]*\s+', '', clean_heading).strip()
            _depth = depth_map.get(idx, level)
            if idx in _visible_section_indices:
                if _depth <= 0:
                    # Wurzel/Artikeltitel (nur bei title_as_parent): trägt direkt das Elternkapitel.
                    clean_heading = _join_number_and_rest(parent_chapter, _title)
                else:
                    local_num = _advance_counter(counters, _depth)
                    number = f"{parent_chapter}.{local_num}" if local_num else parent_chapter
                    clean_heading = _join_number_and_rest(number, _title)
            else:
                # Unsichtbare Sektion (wird ausgeblendet): keine Nummer vergeben.
                clean_heading = _title
        elif parent_chapter:
            if re.match(r'^\d', clean_heading):
                if extracted_chapter:
                    clean_heading = _rebase_chapter_number(clean_heading, extracted_chapter, parent_chapter)
                else:
                    clean_heading = _prefix_chapter_number(clean_heading, parent_chapter)
            elif not has_numbered_chapters:
                # Unnummeriert ohne nummerierte Kapitel: hier auto-nummerieren.
                # Bei has_numbered_chapters übernimmt der Auto-Block unten
                # (verhindert Doppel-Nummerierung).
                local_num = _advance_counter(counters, level)
                number = f"{parent_chapter}.{local_num}" if local_num else parent_chapter
                clean_heading = _join_number_and_rest(number, clean_heading)

        if parent_chapter:
            num_m = re.match(r'^([\d.]+)\b', clean_heading)
            if num_m:
                display_level = min(len(num_m.group(1).split('.')) + 1, 9)
            elif _is_box_heading(clean_heading) or (extracted_chapter and level > 1 and not re.match(r'^\d', clean_heading)):
                # Kästen + nicht-nummerierte Level-2+-Headings im Einbette-Modus: 2 Ebenen unter Elternkapitel.
                display_level = min(lvl_shift + 2, 9)
            else:
                display_level = min(level + lvl_shift, 9)
        else:
            display_level = level

        sum_body = sum_lookup.get(lookup_key, '')
        # Structural-Renumber: Die Intro-Zusammenfassung (synthetische "Einleitung"-Sektion) matcht
        # keinen Original-Heading. Sie an die erste Tiefe-1-Sektion (Artikeltitel, z.B. 6.2.4.1.1)
        # anhängen, wenn diese selbst keine eigene Zusammenfassung hat.
        if (structural_renumber and not _struct_intro_attached
                and depth_map.get(idx) == (0 if title_as_parent else 1)
                and not sum_body.strip()):
            sum_body = sum_lookup.get('einleitung', '')
            _struct_intro_attached = True
        # clean_heading kann nach Rebase bereits eine Zahl vorne haben (tail-normalisierte Headings).
        originally_numbered = bool(re.match(r'^\d', lookup_key)) or bool(re.match(r'^\d', clean_heading))

        # Auto-Nummerierung: einzigartige unnummerierte Überschriften (freq=1) nach einem
        # nummerierten Kapitel erhalten automatisch eine Unterkapitelnummer (z.B. 5.3.1).
        # Wiederkehrende Überschriften ("Was ist das?" etc.) bleiben unnummeriert.
        # Nur sichtbare Sections (_visible_section_indices) erhalten eine Nummer → keine Lücken.
        # Im strukturellen Neu-Nummerierungs-Modus ist die Nummerierung oben abschließend erfolgt.
        if has_numbered_chapters and not structural_renumber:
            if originally_numbered:
                # clean_heading wurde oben bereits präfixiert/rebased (parent_chapter-Pfad),
                # daher die Elternnummer direkt daraus übernehmen – NICHT aus dem rohen
                # lookup_key. Sonst fehlt das Elternpräfix und ein Trailing-Punkt der
                # OCR-Nummer ("3.") erzeugt Doppelpunkte wie "3..1".
                _m = re.match(r'^(\d[\d.]*)', clean_heading.strip())
                _auto_parent = _normalize_chapter_num(_m.group(1)) if _m else None
                _auto_counter = 0
                _current_competency_key = None
            elif (_auto_parent is not None and _unnumbered_freq.get(lookup_key, 0) == 1
                  and not (extracted_chapter and _is_box_heading(clean_heading))
                  and not (extracted_chapter and level > 1)):
                # Unique unnummerierte Unterüberschrift: auto-nummerieren.
                # Kästen (Fokus/Studie/…) und Level-2+-Headings im Einbette-Modus ausgenommen
                # – sie erscheinen als eigenständige Lernobjekte ohne Kapitelnummer.
                originally_numbered = True
                _current_competency_key = lookup_key
                if idx in _visible_section_indices:
                    _auto_counter += 1
                    clean_heading = f"{_auto_parent}.{_auto_counter} {clean_heading}"
                _dots = clean_heading.split()[0].count('.')
                display_level = min(_dots + 2, 9)

        # Scoped lookup: Sub-Sections (freq>1, wiederkehrend) unter der aktuellen Kompetenz.
        # Verhindert dass sum_lookup["off the job"] immer den letzten Summary-Eintrag liefert.
        if not originally_numbered and _current_competency_key:
            _scoped = sum_scoped.get(_current_competency_key, {}).get(lookup_key, '')
            if _scoped.strip():
                sum_body = _scoped

        _next_sec = orig_sections[idx + 1] if idx + 1 < len(orig_sections) else None
        _next_is_unnumbered = (
            _next_sec is not None and
            _next_sec['heading'] != '__preamble__' and
            not re.match(r'^\d', _clean_heading_text(_next_sec['heading']).strip())
        )
        has_children = (
            _next_sec is not None and
            _next_sec['heading'] != '__preamble__' and
            # Echte Kinder: tiefere Ebene ODER nummerierte Section mit unnummeriertem Folger.
            # Geschwister (unnummeriert → unnummeriert gleicher Ebene) zählen NICHT als Kinder.
            (_next_sec['level'] > level or (originally_numbered and _next_is_unnumbered)) and
            # Nur als Elternknoten sichtbar wenn min. ein Nachfolger Summary-Inhalt hat —
            # verhindert goldene Leer-Überschriften ohne sichtbaren Folgeinhalt.
            _any_visible_desc[idx]
        )
        # Strukturwurzel im title_as_parent-Modus (Artikeltitel, Tiefe 0): trägt das
        # Elternkapitel und ist der Anker des Dokuments → immer als sichtbare Überschrift,
        # auch ohne eigene Zusammenfassung und ohne von der flachen OCR-Ebene erkannte Kinder.
        is_struct_root = structural_renumber and title_as_parent and depth_map.get(idx) == 0

        # Nur wirklich leere Sektionen überspringen (kein Original, kein Summary, keine Kinder).
        if not sum_body.strip() and not orig_body.strip() and not has_children and not is_struct_root:
            continue

        # Überschrift sichtbar: wenn Zusammenfassung vorhanden ODER Elternkapitel mit Kindern.
        # Sonst ausgeblendet — Originalinhalt bleibt im Dokument, ist im Summary-View aber
        # unsichtbar → kein leerer Gliederungspunkt, kein verlorener Inhalt.
        show_heading_visible = bool(sum_body.strip()) or has_children or is_struct_root
        if show_heading_visible:
            if has_numbered_chapters:
                nav_worthy = originally_numbered or (extracted_chapter and _is_box_heading(clean_heading))
            elif parent_chapter:
                nav_worthy = True
            else:
                nav_worthy = display_level <= 2
            if nav_worthy:
                h = doc.add_heading(clean_heading, level=display_level)
                _set_heading_color(h, MM_HEADING_COLORS.get(display_level, MM_HEADING_COLORS[9]))
            else:
                h = doc.add_paragraph(style='Normal')
                h.add_run(clean_heading).bold = True
        else:
            # Überschrift ausblenden wie den Originaltext darunter
            h = doc.add_paragraph(style='Normal')
            h.add_run(clean_heading).bold = True
            _hide_paragraph(h, indent_cm=1.5)

        # Zusammenfassung + Kommentar-Erkennung
        if sum_body.strip():
            before = len(doc.paragraphs)
            process_markdown_to_docx(doc, sum_body, hide_text=False, base_path=base_path,
                                     headings_as_bold=True)

            # Schlüssel zur Zuordnung Sektion → QA (vor-Präfix-Heading + Last-Part-Fallback)
            match_keys = [lookup_key]
            last_part = re.sub(r'^[\d.]+\s*', '', lookup_key).strip()
            if last_part:
                match_keys.append(last_part)

            # --- Ergänzungen aus der Nachbearbeitung (Schritt 5b) anfügen ---
            if supplement_map:
                seen_sup: set = set()
                for k in match_keys:
                    for sup in supplement_map.get(k, []):
                        if not sup or sup in seen_sup:
                            continue
                        seen_sup.add(sup)
                        p_sup = doc.add_paragraph()
                        r_tag = p_sup.add_run("[Ergänzt durch Nachbearbeitung] ")
                        r_tag.bold = True
                        r_tag.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
                        p_sup.add_run(sup)

            new_paras = doc.paragraphs[before:]
            first_para = next((p for p in new_paras if p.text.strip()), None)
            last_para  = next((p for p in reversed(new_paras) if p.text.strip()), first_para)

            # Word-Kommentar pro Frage – an der Beleg-Stelle verankert (Fallback: ganze Sektion)
            if first_para and textgrundlage_map:
                q_nums = []
                for k in match_keys:
                    q_nums.extend(textgrundlage_map.get(k, []))
                q_nums = sorted(set(q_nums))
                for n in q_nums:
                    beleg = (qa_by_num.get(n) or {}).get('beleg', '')
                    target = _find_para_by_quote(new_paras, beleg) if beleg else None
                    start_p = target or first_para
                    end_p   = target or last_para
                    cid = comment_id[0]
                    comment_id[0] += 1
                    _add_comment_range_start(start_p, cid)
                    _add_comment_range_end(end_p, cid)
                    comment_list.append((cid, f'Frage {n}'))

        if orig_body.strip():
            process_markdown_to_docx(doc, orig_body, hide_text=True, base_path=base_path)

    # --- Fragentext laden (optional) ---
    questions_map: dict = {}
    if questions_path and os.path.exists(questions_path):
        with open(questions_path, encoding='utf-8') as qf:
            for line in qf:
                qm = re.match(r'^(\d+)\.\s+(.+)', line.strip())
                if qm:
                    questions_map[int(qm.group(1))] = qm.group(2)

    # --- QA-Unterkapitel am Dokumentende ---
    if has_qa and qa_items:
        if parent_chapter:
            # Nach Fix 4 werden Level-1-Headings in nicht-nummerierten Dokumenten übersprungen
            # (OCR-Artefakte), sodass counters[0] = 0 bleibt. Die Top-Level-Inhalte liegen
            # bei Level 2 (counters[1]). Für nummerierte Dokumente zählt weiterhin counters[0].
            if structural_renumber:
                # Strukturelle Neu-Nummerierung: Top-Level-Sektionen zählen counters[0];
                # Lernfragen ist das nächste Geschwister auf Tiefe 1.
                qa_top_num = counters[0] + 1
            elif extracted_chapter and has_numbered_chapters:
                # Rebase-Modus: Unterkapitel werden über _auto_counter gezählt;
                # Lernfragen ist das nächste Geschwister.
                qa_top_num = _auto_counter + 1
            elif not has_numbered_chapters:
                qa_top_num = counters[1] + 1
            else:
                qa_top_num = counters[0] + 1
            qa_hdg_text = f"{parent_chapter}.{qa_top_num} Lernfragen"
            qa_level    = min(lvl_shift + 1, 9)
            q_sub_level = min(lvl_shift + 2, 9)
        else:
            qa_hdg_text = "Lernfragen"
            qa_level    = 1
            q_sub_level = 2

        h = doc.add_heading(qa_hdg_text, level=qa_level)
        _set_heading_color(h, MM_HEADING_COLORS.get(qa_level, MM_HEADING_COLORS[9]))

        for q_idx, item in enumerate(qa_items, start=1):
            if parent_chapter:
                fq_hdg = f"{parent_chapter}.{qa_top_num}.{q_idx} Frage {item['num']}"
            else:
                fq_hdg = f"Frage {item['num']}"
            h_f = doc.add_heading(fq_hdg, level=q_sub_level)
            _set_heading_color(h_f, MM_HEADING_COLORS.get(q_sub_level, MM_HEADING_COLORS[9]))

            # Fragetext anzeigen wenn verfügbar
            q_text = questions_map.get(item['num'])
            if q_text:
                p_q = doc.add_paragraph()
                r_q = p_q.add_run(q_text)
                r_q.bold = True
                r_q.italic = True

            doc.add_paragraph(item['antwort'])

            meta = doc.add_paragraph()
            r = meta.add_run(
                f"Quelle: {item['textgrundlage']}  |  "
                f"Schlüsselbegriffe: {item['schluessel']}  |  "
                f"Abdeckung: {item['abdeckung']}"
            )
            r.font.size = Pt(9)
            r.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    # --- Kommentare einbetten und speichern ---
    if comment_list:
        _inject_comments_part(doc, comment_list)
    doc.save(output_path)
    print(f"Word-Dokument erstellt ({len(comment_list)} Kommentare).")


# ---------------------------------------------------------------------------
# Hauptprogramm
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="End-to-End PDF Translation & Learning Pipeline",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Beispiele:\n"
            "  python pipeline.py dok.pdf\n"
            "  python pipeline.py dok.pdf --parent-chapter 6.2.4.3 --title-as-parent --questions fragen.txt\n"
            "  python pipeline.py dok.pdf --chapter 4.2 --parent-chapter 2.4.1.2\n"
            "  python pipeline.py dok.pdf --no-translate --no-summary\n"
            "  python pipeline.py dok.pdf --force   # alle Zwischenergebnisse neu berechnen\n"
        ),
    )
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
    parser.add_argument("--include-front-matter", action="store_true",
                        help="Front-Matter (Impressum, Inhaltsverzeichnis) NICHT überspringen "
                             "(Standard: wird übersprungen; Titel + Titelbild bleiben immer erhalten).")
    parser.add_argument("--chapter", type=str, default=None,
                        help="Nur dieses Kapitel extrahieren, z.B. '4.2' oder '3'. "
                             "Sucht im OCR-Markdown nach der Überschrift und extrahiert das Kapitel "
                             "inkl. aller Unterkapitel bis zur nächsten gleichrangigen Überschrift.")
    parser.add_argument("--no-summary", action="store_true",
                        help="Nur Übersetzung ausgeben – keine Zusammenfassung, kein ausgeblendeter Text, "
                             "kein interleaved-Dokument. Das Übersetzungs-Docx ist das finale Ergebnis.")
    parser.add_argument("--target-language", type=str, default=None,
                        help="Zielsprache als ISO-Code, z.B. 'de' oder 'en'. "
                             "Standard: 'de' (englische Quellen werden automatisch nach Deutsch übersetzt). "
                             "Ist Quell- = Zielsprache, wird NICHT übersetzt.")
    parser.add_argument("--source-language", type=str, default=None,
                        help="Quellsprache als ISO-Code erzwingen (überspringt die automatische Erkennung).")
    parser.add_argument("--no-translate", action="store_true",
                        help="Keine Übersetzung – Originalsprache beibehalten, unabhängig von der Erkennung.")
    parser.add_argument("--title-as-parent", action="store_true",
                        help="Nur mit --parent-chapter (ohne --chapter): Der Artikeltitel trägt direkt "
                             "die Elternkapitelnummer (z.B. 6.2.4.1) statt 6.2.4.1.1; alle Abschnitte "
                             "inkl. Studien werden zu dessen Unterpunkten (6.2.4.1.1, 6.2.4.1.2 …).")

    args = parser.parse_args()
    OUTPUT_BASE = "workspace/output"

    try:
        if not os.getenv("GEMINI_API_KEY"):
            raise ValueError("GEMINI_API_KEY fehlt in der .env-Datei!")

        if not os.path.exists(args.pdf_path):
            raise FileNotFoundError(f"Die Datei {args.pdf_path} wurde nicht gefunden.")

        pdf_stem  = Path(args.pdf_path).stem
        doc_title = pdf_stem.replace('_', ' ')
        out_dir   = Path(OUTPUT_BASE) / pdf_stem      # Dokument-Ordner: enthält die finalen .docx
        out_dir.mkdir(parents=True, exist_ok=True)

        # Arbeits-/Zwischendateien (OCR-Markdown + Bilder, Übersetzung, Zusammenfassung, QA)
        # liegen gebündelt in out_dir/work, getrennt von den finalen Ergebnissen.
        work_dir = out_dir / "work"
        work_dir.mkdir(parents=True, exist_ok=True)

        # Kapitel-spezifisches Cache-Verzeichnis innerhalb des Arbeitsordners
        chapter_safe = args.chapter.replace('.', '_') if args.chapter else None
        cache_dir = work_dir / f"kap{chapter_safe}" if chapter_safe else work_dir
        cache_dir.mkdir(parents=True, exist_ok=True)

        # --- Schritt 1: OCR ---
        # Marker schreibt nach work_dir/<pdf_stem>/<pdf_stem>.md (inkl. Bilder).
        md_path = work_dir / pdf_stem / f"{pdf_stem}.md"
        if args.force or not md_path.exists():
            raw_md_path = run_marker_ocr(args.pdf_path, str(work_dir))
            md_path = Path(raw_md_path)
        else:
            print(f"[SKIP] OCR – Markdown bereits vorhanden: {md_path}")

        raw_md = md_path.read_text(encoding="utf-8")
        # Bilder liegen im selben Ordner wie das OCR-Markdown → base_path für Bildauflösung.
        image_base = str(md_path.parent)

        # --- Schritt 1a: Front-Matter überspringen (Impressum, Inhaltsverzeichnis) ---
        # Standardmäßig aktiv; Titel + Titelbild bleiben erhalten. Spart außerdem
        # Übersetzungs-/Zusammenfassungs-Kosten für irrelevante Vorseiten.
        if not args.include_front_matter:
            before_len = len(raw_md)
            raw_md = strip_front_matter(raw_md)
            if len(raw_md) < before_len:
                print(f"[FRONT-MATTER] Impressum/Inhaltsverzeichnis übersprungen "
                      f"({before_len - len(raw_md)} Zeichen entfernt).")

        # --- Schritt 1b: Kapitel-Filter (optional) ---
        if args.chapter:
            raw_md = extract_chapter(raw_md, args.chapter)

        # --- Schritt 2: Sprache prüfen & ggf. übersetzen ---
        # Übersetzt wird NUR, wenn Quell- und Zielsprache verschieden sind. Default-Ziel ist
        # Deutsch: englische Quellen werden automatisch übersetzt, deutsche bleiben deutsch.
        # Eine explizite --target-language oder --source-language überschreibt die Automatik.
        transl_path = cache_dir / "de_uebersetzung.md"
        if args.force and transl_path.exists():
            transl_path.unlink()

        target_lang = (args.target_language or "de").lower()
        is_translated = False
        # content_lang = Sprache des working_text → steuert die Ausgabesprache der Zusammenfassung.
        content_lang = "de"

        if args.no_translate:
            print("--no-translate: Übersetzung übersprungen, Originalsprache bleibt erhalten.")
            working_text = raw_md
            content_lang = (args.source_language or "de").lower()
        else:
            source_lang = (args.source_language or detect_language(raw_md)).lower()
            if source_lang == "unknown":
                print("Sprache unklar – sicherheitshalber keine Übersetzung (Original bleibt erhalten).")
                working_text = raw_md
                content_lang = "de"
            elif source_lang == target_lang:
                print(f"Quelle ist bereits {_language_name(target_lang)}. Keine Übersetzung notwendig.")
                working_text = raw_md
                content_lang = target_lang
            elif transl_path.exists():
                # Gecachte Übersetzung aus einem vorherigen Lauf wiederverwenden – aber nur,
                # nachdem feststeht, dass tatsächlich übersetzt werden soll.
                print(f"[SKIP] Übersetzung – bereits vorhanden: {transl_path}")
                working_text = transl_path.read_text(encoding="utf-8")
                is_translated = True
                content_lang = target_lang
            else:
                print(f"Quelle ist {_language_name(source_lang)}. Starte Übersetzung nach "
                      f"{_language_name(target_lang)}...")
                working_text = translate_text(raw_md, source_lang=source_lang, target_lang=target_lang)
                transl_path.write_text(working_text, encoding="utf-8")
                print(f"       Übersetzung gespeichert: {transl_path}")
                is_translated = True
                content_lang = target_lang

        # --- Zwischenschritt: Übersetzungs-Docx (nur bei englischer Quelle) ---
        kap_infix = f"_kap{chapter_safe}" if chapter_safe else ""
        if is_translated:
            transl_docx_path = out_dir / f"{pdf_stem}{kap_infix}_Uebersetzung.docx"
            if args.force or not transl_docx_path.exists():
                build_translation_word_document(working_text, str(transl_docx_path), base_path=image_base)
            else:
                print(f"[SKIP] Übersetzungs-Docx – bereits vorhanden: {transl_docx_path}")

            if args.no_summary:
                print(f"\n=== PIPELINE ERFOLGREICH BEENDET (nur Übersetzung) ===")
                print(f"Zwischenergebnisse: {cache_dir}")
                print(f"Fertiges Dokument:  {transl_docx_path}")
                sys.exit(0)
        elif args.no_summary:
            print(f"\n=== PIPELINE ERFOLGREICH BEENDET (Text bereits Deutsch – keine Übersetzung erstellt) ===")
            print(f"Zwischenergebnisse: {cache_dir}")
            sys.exit(0)

        # --- Schritt 3: Box-Strukturreparatur (nur wenn Lehrbuch-Kästen vorhanden) ---
        # Trennt eingeschobene Kästen (Fokus/Studie/Definition/Beispiel) vom Fließtext,
        # den die OCR fälschlich in den Kasten gezogen hat. Speist Summary + Interleaved.
        struct_path = cache_dir / "de_strukturiert.md"
        if args.force and struct_path.exists():
            struct_path.unlink()
        if struct_path.exists():
            print(f"[SKIP] Box-Strukturreparatur – bereits vorhanden: {struct_path}")
            working_text = struct_path.read_text(encoding="utf-8")
        else:
            print("--- Schritt 3: Box-Strukturreparatur (Kästen vom Fließtext trennen) ---")
            working_text = repair_box_structure(working_text)
            struct_path.write_text(working_text, encoding="utf-8")

        # --- Schritt 4: Zusammenfassung (kapitelweise) ---
        sum_path = cache_dir / "zusammenfassung.md"
        if args.force and sum_path.exists():
            sum_path.unlink()
            for f in cache_dir.glob("zusammenfassung_kap_*.md"):
                f.unlink()

        if sum_path.exists():
            print(f"[SKIP] Zusammenfassung – bereits vorhanden: {sum_path}")
            summary_result = sum_path.read_text(encoding="utf-8")
        else:
            summary_result = generate_summary_by_chapter(working_text, cache_dir, output_lang=content_lang)

        # --- Schritt 5: Qualitätssicherung (optional) ---
        qa_result = "Keine Leitfragen zur Prüfung übergeben."
        supplement_map: dict = {}
        if args.questions:
            if os.path.exists(args.questions):
                qa_path = cache_dir / "qa_ergebnis.md"
                if args.force and qa_path.exists():
                    qa_path.unlink()
                qa_result = load_or_run(
                    qa_path,
                    lambda: verify_with_questions(summary_result, args.questions),
                    "QA / Leitfragen"
                )

                # --- Schritt 5b: Nacharbeit unvollständiger Antworten ---
                rework_path = cache_dir / "qa_nachbearbeitung.md"
                sup_path    = cache_dir / "qa_ergaenzungen.json"
                if args.force:
                    for p in (rework_path, sup_path):
                        if p.exists():
                            p.unlink()
                if rework_path.exists() and sup_path.exists():
                    print(f"[SKIP] Nacharbeit – bereits vorhanden: {rework_path}")
                    qa_result = rework_path.read_text(encoding="utf-8")
                    supplement_map = json.loads(sup_path.read_text(encoding="utf-8"))
                else:
                    qa_result, supplement_map = rework_partial_answers(
                        qa_result, working_text, args.questions
                    )
                    rework_path.write_text(qa_result, encoding="utf-8")
                    sup_path.write_text(json.dumps(supplement_map, ensure_ascii=False, indent=2),
                                        encoding="utf-8")
            else:
                print(f"Warnung: Fragen-Datei '{args.questions}' nicht gefunden. Überspringe QS.")

        # --- Word-Dokument zusammensetzen ---
        suffix = f"_Einbetten_{args.parent_chapter.replace('.', '-')}" if args.parent_chapter else "_Lernskript"
        final_docx_path = out_dir / f"{pdf_stem}{kap_infix}{suffix}.docx"
        build_interleaved_word_document(
            working_text, summary_result, qa_result,
            str(final_docx_path), base_path=image_base,
            parent_chapter=args.parent_chapter,
            parent_level=args.parent_level,
            skip_references=not args.include_references,
            questions_path=args.questions,
            doc_title=doc_title,
            extracted_chapter=args.chapter,
            supplement_map=supplement_map,
            title_as_parent=args.title_as_parent,
        )

        print(f"\n=== PIPELINE ERFOLGREICH BEENDET ===")
        print(f"Zwischenergebnisse:   {cache_dir}")
        if is_translated:
            print(f"Übersetzung (Word):   {transl_docx_path}")
        print(f"Fertiges Dokument:    {final_docx_path}")
        if args.parent_chapter:
            print(f"  → Einfügemodus: Kapitelpräfix '{args.parent_chapter}', "
                  f"Heading-Shift +{args.parent_level or _parent_level_from_chapter(args.parent_chapter)}")

    except Exception as e:
        print(f"\nPipeline abgebrochen wegen: {e}")
