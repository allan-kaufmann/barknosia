"""
Unit-Tests für pipeline.py

Ausführung:
    python -m pytest test_pipeline.py -v

Jede Funktion, die mit test_ beginnt, wird von pytest automatisch gefunden und ausgeführt.
"""
import pytest
from docx import Document

from pipeline import (
    _clean_for_hidden,
    _clean_heading_text,
    normalize_heading,
    normalize_heading_levels,
    parse_sections,
    parse_qa_response,
    _advance_counter,
    _prefix_chapter_number,
    _is_skip_heading,
    split_into_level1_chapters,
    add_formatted_text,
    process_markdown_to_docx,
    _strip_kontrollliste,
    _compress_heading_levels,
    _fix_concatenated_stats_cells,
    _drop_empty_columns,
    _hide_paragraph,
    add_markdown_table_to_doc,
    _set_cell_text,
    _strip_ocr_y_prefix,
    build_interleaved_word_document,
)


# ---------------------------------------------------------------------------
# Gruppe 1: _clean_for_hidden  (Fix: Bilder werden nicht entfernt)
# ---------------------------------------------------------------------------

def test_clean_for_hidden_preserves_images():
    result = _clean_for_hidden("Bild: ![Abbildung 1](figures/img.png) im Text.")
    assert "![Abbildung 1](figures/img.png)" in result


def test_clean_for_hidden_removes_text_links():
    result = _clean_for_hidden("[Quelle](https://example.com)")
    assert result == "Quelle"


def test_clean_for_hidden_text_link_does_not_affect_image():
    md = "[Link](https://url.com) und ![Bild](bild.jpeg)"
    result = _clean_for_hidden(md)
    assert "Link" in result           # Link-Text bleibt
    assert "https://url.com" not in result  # URL weg
    assert "![Bild](bild.jpeg)" in result   # Bild unverändert


def test_clean_for_hidden_removes_html_tags():
    result = _clean_for_hidden('<span id="page-12">Text</span>')
    assert "<span" not in result
    assert "Text" in result


def test_clean_for_hidden_removes_standalone_page_refs():
    result = _clean_for_hidden("[1]\n\nText dahinter")
    assert "[1]" not in result
    assert "Text dahinter" in result


# ---------------------------------------------------------------------------
# Gruppe 2: _clean_heading_text + normalize_heading  (Fix: Lookup-Key)
# ---------------------------------------------------------------------------

def test_clean_heading_removes_html():
    raw = '<span id="page-12-0"></span>**1.2 Definition**'
    assert _clean_heading_text(raw) == "1.2 Definition"


def test_clean_heading_removes_markdown_link():
    raw = '**1.2 [Definition](https://dorsch.hogrefe.com/stichwort/definition)**'
    assert _clean_heading_text(raw) == "1.2 Definition"


def test_clean_heading_combined_html_and_link():
    raw = '<span id="page-12-0"></span><span id="page-12-1"></span>**1.2 [Definition](https://dorsch.de)**'
    assert _clean_heading_text(raw) == "1.2 Definition"


def test_normalize_heading_strips_bold_and_lowercases():
    assert normalize_heading("**1.2 Definition**") == "1.2 definition"


def test_normalize_heading_strips_italic():
    assert normalize_heading("*Abschnitt*") == "abschnitt"


def test_lookup_key_ocr_heading_matches_summary_heading():
    """
    Kerntest für den Empty-Chapter-Fix:
    Ein OCR-Heading mit HTML und Links muss denselben lookup_key ergeben
    wie der saubere Summary-Heading.
    """
    ocr_heading     = '<span id="p">**1.2 [Definition](https://dorsch.de)**'
    summary_heading = "**1.2 Definition**"
    assert normalize_heading(_clean_heading_text(ocr_heading)) == normalize_heading(summary_heading)


# ---------------------------------------------------------------------------
# Gruppe 3: add_formatted_text  (Fixes: italic, triple asterisk)
# ---------------------------------------------------------------------------

def test_add_formatted_text_bold():
    para = Document().add_paragraph()
    add_formatted_text(para, "**fett**")
    bold_runs = [r for r in para.runs if r.bold]
    assert len(bold_runs) == 1
    assert bold_runs[0].text == "fett"


def test_add_formatted_text_italic():
    para = Document().add_paragraph()
    add_formatted_text(para, "*kursiv*")
    italic_runs = [r for r in para.runs if r.italic]
    assert len(italic_runs) == 1
    assert italic_runs[0].text == "kursiv"


def test_add_formatted_text_bold_italic_triple_asterisk():
    """Fix: ***Ausbildung:*** wurde früher als Literal-Text mit Sternchen ausgegeben."""
    para = Document().add_paragraph()
    add_formatted_text(para, "***Ausbildung:***")
    bi_runs = [r for r in para.runs if r.bold and r.italic]
    assert len(bi_runs) == 1
    assert bi_runs[0].text == "Ausbildung:"
    assert "*" not in bi_runs[0].text


def test_add_formatted_text_plain():
    para = Document().add_paragraph()
    add_formatted_text(para, "normaler Text")
    assert para.runs[0].text == "normaler Text"
    assert not para.runs[0].bold
    assert not para.runs[0].italic


def test_add_formatted_text_mixed():
    para = Document().add_paragraph()
    add_formatted_text(para, "Vor ***bold+italic*** Nach")
    texts = [r.text for r in para.runs]
    assert "Vor " in texts
    assert "bold+italic" in texts
    assert " Nach" in texts
    bi_runs = [r for r in para.runs if r.bold and r.italic]
    assert len(bi_runs) == 1
    assert bi_runs[0].text == "bold+italic"


def test_add_formatted_text_no_asterisk_in_output():
    """Kein Sternchen darf als sichtbarer Text im Dokument landen."""
    para = Document().add_paragraph()
    add_formatted_text(para, "***fett-kursiv*** und **nur-fett** und *nur-kursiv*")
    for run in para.runs:
        assert "*" not in run.text, f"Sternchen in Run: '{run.text}'"


# ---------------------------------------------------------------------------
# Gruppe 4: process_markdown_to_docx – Bullet-Handling (mehrere Fixes)
# ---------------------------------------------------------------------------

def test_bullet_no_leading_spaces():
    """Fix: '*   Text' (Stern + 3 Leerzeichen) erzeugte '  Text' mit führenden Spaces."""
    doc = Document()
    process_markdown_to_docx(doc, "*   Lernen in Unternehmen")
    bullets = [p for p in doc.paragraphs if p.style.name == 'List Bullet']
    assert len(bullets) == 1
    assert bullets[0].text == "Lernen in Unternehmen"


def test_bullet_single_space_after_star():
    doc = Document()
    process_markdown_to_docx(doc, "* Einfacher Bullet")
    bullets = [p for p in doc.paragraphs if p.style.name == 'List Bullet']
    assert bullets[0].text == "Einfacher Bullet"


def test_empty_bullet_skipped():
    """Fix: '*   ' ohne Text erzeugte leeren Bullet-Absatz."""
    doc = Document()
    process_markdown_to_docx(doc, "*   \n* Inhalt")
    bullets = [p for p in doc.paragraphs if p.style.name == 'List Bullet']
    assert len(bullets) == 1
    assert bullets[0].text == "Inhalt"


def test_numbered_sub_item_as_sub_bullet():
    """Fix: '    1.  Text' landete als Plaintext statt eingerücktem Sub-Bullet."""
    doc = Document()
    process_markdown_to_docx(doc, "*   Liste:\n    1.  Erster Punkt\n    2.  Zweiter Punkt")
    sub_bullets = [p for p in doc.paragraphs if p.style.name in ('List Bullet 2', 'List Bullet 3')]
    assert len(sub_bullets) == 2
    assert sub_bullets[0].text == "Erster Punkt"
    assert sub_bullets[1].text == "Zweiter Punkt"


def test_sub_bullet_with_asterisk():
    doc = Document()
    process_markdown_to_docx(doc, "*   Eltern\n    *   Kind-Bullet")
    sub = [p for p in doc.paragraphs if p.style.name in ('List Bullet 2', 'List Bullet 3')]
    assert len(sub) == 1
    assert sub[0].text == "Kind-Bullet"


def test_triple_asterisk_in_bullet():
    """Fix: '***Ausbildung:*** Text' in Bullet durfte keine Sternchen zeigen."""
    doc = Document()
    process_markdown_to_docx(doc, "*   ***Ausbildung:*** Curricular organisierte Aktivitäten")
    bullets = [p for p in doc.paragraphs if p.style.name == 'List Bullet']
    assert len(bullets) == 1
    for run in bullets[0].runs:
        assert "*" not in run.text


# ---------------------------------------------------------------------------
# Gruppe 5: normalize_heading_levels + parse_sections
# ---------------------------------------------------------------------------

def test_normalize_level1_to_h1():
    result = normalize_heading_levels("# 1 Einleitung\nText")
    assert result.startswith("# 1 Einleitung")


def test_normalize_level2_to_h2():
    result = normalize_heading_levels("# 1.1 Abschnitt\nText")
    assert result.startswith("## 1.1 Abschnitt")


def test_normalize_level3_to_h3():
    result = normalize_heading_levels("# 1.1.1 Unterabschnitt\nText")
    assert result.startswith("### 1.1.1 Unterabschnitt")


def test_normalize_non_numbered_unchanged():
    result = normalize_heading_levels("# Einleitung ohne Nummer\nText")
    assert result.startswith("# Einleitung ohne Nummer")


def test_parse_sections_finds_headings():
    md = "## 2 Kapitel\nErster Absatz.\n\n## 3 Folgekapitel\nZweiter Absatz."
    sections = parse_sections(md)
    heads = [s['heading'] for s in sections if s['heading'] != '__preamble__']
    assert "2 Kapitel" in heads
    assert "3 Folgekapitel" in heads


def test_parse_sections_preamble():
    md = "Einleitungstext\n## 1 Kapitel\nInhalt"
    sections = parse_sections(md)
    assert sections[0]['heading'] == '__preamble__'
    assert "Einleitungstext" in sections[0]['body']


def test_parse_sections_body_content():
    md = "## 1 Kapitel\nZeile 1\nZeile 2"
    sections = parse_sections(md)
    kap = next(s for s in sections if s['heading'] != '__preamble__')
    assert "Zeile 1" in kap['body']
    assert "Zeile 2" in kap['body']


# ---------------------------------------------------------------------------
# Bonus: Weitere pure Helfer
# ---------------------------------------------------------------------------

def test_advance_counter_increments_level():
    counters = [0] * 9
    result = _advance_counter(counters, 1)
    assert result == "1"
    assert counters[0] == 1


def test_advance_counter_resets_deeper_levels():
    counters = [3, 2, 1, 0, 0, 0, 0, 0, 0]
    _advance_counter(counters, 2)
    assert counters[1] == 3
    assert counters[2] == 0


def test_advance_counter_skips_zero_segments():
    counters = [4, 0, 0, 0, 0, 0, 0, 0, 0]
    _advance_counter(counters, 2)
    result = _advance_counter(counters, 2)
    assert '.' in result
    assert '0' not in result.split('.')[0]


def test_prefix_chapter_number_numbered():
    assert _prefix_chapter_number("1 EINLEITUNG", "4.2") == "4.2.1 EINLEITUNG"


def test_prefix_chapter_number_numbered_sub():
    assert _prefix_chapter_number("1.1 Abschnitt", "4.2") == "4.2.1.1 Abschnitt"


def test_prefix_chapter_number_non_numbered_unchanged():
    assert _prefix_chapter_number("Einleitung ohne Zahl", "4.2") == "Einleitung ohne Zahl"


def test_is_skip_heading_referenzen():
    assert _is_skip_heading("Referenzen") is True
    assert _is_skip_heading("referenzen") is True


def test_is_skip_heading_literaturverzeichnis():
    assert _is_skip_heading("Literaturverzeichnis") is True


def test_is_skip_heading_normal_chapter():
    assert _is_skip_heading("1 Einleitung") is False
    assert _is_skip_heading("Methoden") is False


def test_parse_qa_response_basic():
    qa = (
        "Frage 1\n"
        "Antwort: Die Antwort auf Frage 1.\n"
        "Textgrundlage: 1.1 Abschnitt\n"
        "Schlüsselbegriffe: Begriff A, Begriff B\n"
        "Abdeckung: vollständig\n\n"
        "Frage 2\n"
        "Antwort: Die Antwort auf Frage 2.\n"
        "Textgrundlage: 1.2 Unterabschnitt\n"
        "Schlüsselbegriffe: Begriff C\n"
        "Abdeckung: teilweise\n"
    )
    items = parse_qa_response(qa)
    assert len(items) == 2
    assert items[0]['num'] == 1
    assert "Frage 1" in items[0]['antwort'] or "Antwort" in items[0]['antwort']
    assert items[0]['textgrundlage'] == "1.1 Abschnitt"
    assert items[1]['num'] == 2
    assert items[1]['abdeckung'] == "teilweise"


def test_split_into_level1_chapters():
    md = (
        "## 1 Einleitung\nText Kap 1\n"
        "### 1.1 Unterkapitel\nText 1.1\n"
        "## 2 Methoden\nText Kap 2\n"
    )
    chapters = split_into_level1_chapters(md)
    assert len(chapters) == 2
    headings = [c['heading'] for c in chapters]
    assert any("1 Einleitung" in h for h in headings)
    assert any("2 Methoden" in h for h in headings)


# ---------------------------------------------------------------------------
# Gruppe 7: _strip_kontrollliste
# ---------------------------------------------------------------------------

def test_strip_kontrollliste_removes_block_with_hr():
    text = "Übersetzungstext.\n\n---\n**Kontrollliste**\n- Absätze: 5\n- Hinweise: Keine.\n\n## Kapitel 2\nWeiterer Text."
    result = _strip_kontrollliste(text)
    assert "Kontrollliste" not in result
    assert "Kapitel 2" in result
    assert "Übersetzungstext" in result


def test_strip_kontrollliste_no_hr():
    text = "Text.\n\n**Kontrollliste**\n- Absätze: 3\n- Hinweise: Keine.\n\n## Nächstes Kapitel\nText."
    result = _strip_kontrollliste(text)
    assert "Kontrollliste" not in result
    assert "Nächstes Kapitel" in result


def test_strip_kontrollliste_no_kontrollliste():
    text = "# Kapitel\n\nNormaler Text ohne Kontrollliste."
    result = _strip_kontrollliste(text)
    assert result == text


# ---------------------------------------------------------------------------
# Gruppe 8: _compress_heading_levels
# ---------------------------------------------------------------------------

def test_compress_heading_levels_siblings():
    """H1 gefolgt von drei H4-Geschwistern → H1, H2, H2, H2."""
    md = "# Titel\n\n#### Sub A\n\n#### Sub B\n\n#### Sub C\n"
    result = _compress_heading_levels(md)
    lines = [l for l in result.split('\n') if l.startswith('#')]
    assert lines[0] == '# Titel'
    assert lines[1] == '## Sub A'
    assert lines[2] == '## Sub B'
    assert lines[3] == '## Sub C'


def test_compress_heading_levels_no_change_needed():
    """Korrekte Hierarchie bleibt unverändert."""
    md = "# H1\n\n## H2\n\n### H3\n"
    result = _compress_heading_levels(md)
    assert result == md


def test_compress_heading_levels_resets_on_new_h1():
    """Nach einem neuen H1 wird die Mapping-Logik neu berechnet."""
    md = "# Abschnitt A\n\n#### Sub A1\n\n# Abschnitt B\n\n#### Sub B1\n"
    result = _compress_heading_levels(md)
    lines = [l for l in result.split('\n') if l.startswith('#')]
    assert lines[0] == '# Abschnitt A'
    assert lines[1] == '## Sub A1'
    assert lines[2] == '# Abschnitt B'
    assert lines[3] == '## Sub B1'


# ---------------------------------------------------------------------------
# Gruppe 9: _is_skip_heading Erweiterungen
# ---------------------------------------------------------------------------

def test_is_skip_heading_history():
    assert _is_skip_heading("History") is True
    assert _is_skip_heading("Historie") is True


def test_is_skip_heading_author_email():
    assert _is_skip_heading("Timo Kortsch t.kortsch@tu-bs.de") is True


def test_is_skip_heading_msc():
    assert _is_skip_heading("Timo Kortsch, M.Sc. Prof. Dr. Simone Kauffeld") is True


def test_is_skip_heading_normal_not_skipped():
    assert _is_skip_heading("Theoretischer Hintergrund") is False


# ---------------------------------------------------------------------------
# Gruppe 10: process_markdown_to_docx – HTML-Stripping
# ---------------------------------------------------------------------------

def test_process_strips_sup_tags():
    doc = Document()
    process_markdown_to_docx(doc, "Text mit Fußnote<sup>1</sup> hier.")
    texts = [p.text for p in doc.paragraphs]
    combined = ' '.join(texts)
    assert '<sup>' not in combined
    assert 'Text mit Fußnote' in combined
    assert '1' in combined


def test_process_strips_span_tags():
    doc = Document()
    process_markdown_to_docx(doc, "Ein <span id='x'>wichtiger</span> Begriff.")
    texts = ' '.join(p.text for p in doc.paragraphs)
    assert '<span' not in texts
    assert 'wichtiger' in texts


# ---------------------------------------------------------------------------
# Gruppe 11: Bildfilterung in process_markdown_to_docx
# ---------------------------------------------------------------------------

def test_image_skip_images_flag(tmp_path):
    """skip_images=True: Bildzeile wird komplett ignoriert, kein Platzhalter."""
    doc = Document()
    process_markdown_to_docx(doc, "![Logo](logo.png)", base_path=str(tmp_path), skip_images=True)
    texts = ' '.join(p.text for p in doc.paragraphs)
    assert 'logo.png' not in texts
    assert 'Bild nicht' not in texts


def test_image_skip_images_false_adds_placeholder(tmp_path):
    """skip_images=False (Standard): fehlendes Bild → Platzhalter-Text."""
    doc = Document()
    process_markdown_to_docx(doc, "![Abb](missing.png)", base_path=str(tmp_path), skip_images=False)
    texts = ' '.join(p.text for p in doc.paragraphs)
    assert 'missing.png' in texts


# ---------------------------------------------------------------------------
# Gruppe 12: _fix_concatenated_stats_cells
# ---------------------------------------------------------------------------

def test_fix_concatenated_stats_cells_left_right_empty():
    """Werte links und rechts leer → auf drei Spalten verteilen."""
    rows = [['Item', 'Übersetzung', '', '3.44 1.00 .49', '']]
    result = _fix_concatenated_stats_cells(rows)
    assert result[0][2] == '3.44'
    assert result[0][3] == '1.00'
    assert result[0][4] == '.49'


def test_fix_concatenated_stats_cells_right_empty():
    """Werte nur rechts leer → auf rechte drei Spalten verteilen."""
    rows = [['Item', '3.44 1.00 .49', '', '']]
    result = _fix_concatenated_stats_cells(rows)
    assert result[0][1] == '3.44'
    assert result[0][2] == '1.00'
    assert result[0][3] == '.49'


def test_fix_concatenated_stats_cells_no_change_needed():
    """Zeile ohne OCR-Artefakt bleibt unverändert."""
    rows = [['Item', 'Übersetzung', '3.44', '1.00', '.49']]
    result = _fix_concatenated_stats_cells(rows)
    assert result[0] == ['Item', 'Übersetzung', '3.44', '1.00', '.49']


# ---------------------------------------------------------------------------
# Gruppe 13: normalize_heading_levels – beliebige Tiefe
# ---------------------------------------------------------------------------

def test_normalize_4level_depth():
    """4.1.1.1 (3 Punkte) → H4 (####)."""
    result = normalize_heading_levels("# 4.1.1.1 Tief\nText")
    assert result.startswith("#### 4.1.1.1 Tief")


def test_normalize_5level_depth():
    """4 Punkte → H5 (#####)."""
    result = normalize_heading_levels("# 1.2.3.4.5 Sehr tief\nText")
    assert result.startswith("##### 1.2.3.4.5 Sehr tief")


# ---------------------------------------------------------------------------
# Gruppe 14: _compress_heading_levels – nummerierte Headings unverändert
# ---------------------------------------------------------------------------

def test_compress_skips_numbered():
    """Nummerierte Headings (nach normalize) werden von _compress nicht verändert."""
    md = "## 4 Kapitel\n### 4.1 Sub\n#### 4.1.1 SubSub\n### 4.2 Sub2\n"
    result = _compress_heading_levels(md)
    assert "## 4 Kapitel" in result
    assert "### 4.1 Sub" in result
    assert "#### 4.1.1 SubSub" in result
    assert "### 4.2 Sub2" in result


# ---------------------------------------------------------------------------
# Gruppe 15: split_into_level1_chapters – adaptiver Split
# ---------------------------------------------------------------------------

def test_split_chapters_adaptive_h3():
    """Extrahiertes Kapitel: einziges ## → Unterkapitel auf ### werden gesplittet."""
    md = (
        "## 4 Vorgehen\nText\n"
        "### 4.1 Bedarfserhebung\nText 4.1\n"
        "### 4.2 Konzeptentwicklung\nText 4.2\n"
    )
    chapters = split_into_level1_chapters(md)
    headings = [c['heading'] for c in chapters]
    assert len(chapters) == 2
    assert any("4.1" in h for h in headings)
    assert any("4.2" in h for h in headings)


def test_split_chapters_multi_top_stays_at_h2():
    """Mehrere ## Kapitel → weiterhin an ## splitten, nicht tiefer."""
    md = (
        "## 1 Einleitung\nText Kap 1\n"
        "### 1.1 Unterkapitel\nText 1.1\n"
        "## 2 Methoden\nText Kap 2\n"
    )
    chapters = split_into_level1_chapters(md)
    assert len(chapters) == 2
    headings = [c['heading'] for c in chapters]
    assert any("1 Einleitung" in h for h in headings)
    assert any("2 Methoden" in h for h in headings)


# ---------------------------------------------------------------------------
# Gruppe 16: _drop_empty_columns
# ---------------------------------------------------------------------------

def test_drop_empty_columns_removes_blank():
    """Spalte, die in allen Zeilen leer ist, wird entfernt."""
    rows = [['A', 'B', ''], ['1', '2', ''], ['3', '4', '']]
    result = _drop_empty_columns(rows)
    assert all(len(r) == 2 for r in result)
    assert result[0] == ['A', 'B']


def test_drop_empty_columns_keeps_nonempty():
    """Spalte mit mindestens einem nicht-leeren Wert bleibt erhalten."""
    rows = [['A', 'B', 'C'], ['1', '', '3']]
    result = _drop_empty_columns(rows)
    assert all(len(r) == 3 for r in result)


def test_drop_empty_columns_empty_input():
    assert _drop_empty_columns([]) == []


# ---------------------------------------------------------------------------
# Gruppe 17: _hide_paragraph – Aufzählungszeichen ausblenden
# ---------------------------------------------------------------------------

def test_hide_paragraph_hides_bullet():
    """_hide_paragraph setzt w:vanish in pPr/rPr, um das Bullet-Zeichen auszublenden."""
    from docx.oxml.ns import qn as _qn
    doc = Document()
    p = doc.add_paragraph(style='List Bullet')
    p.add_run("Bullet Text")
    _hide_paragraph(p)
    pPr = p._p.find(_qn('w:pPr'))
    assert pPr is not None, "pPr fehlt"
    rPr = pPr.find(_qn('w:rPr'))
    assert rPr is not None, "rPr in pPr fehlt"
    vanish = rPr.find(_qn('w:vanish'))
    assert vanish is not None, "w:vanish fehlt – Bullet-Zeichen wird nicht ausgeblendet"


# ---------------------------------------------------------------------------
# Gruppe 18: _hide_paragraph – w:numPr wird entfernt
# ---------------------------------------------------------------------------

def test_hide_paragraph_removes_numpr():
    """_hide_paragraph entfernt w:numPr, damit das List-Label (• + Tab) nicht sichtbar bleibt."""
    from docx.oxml.ns import qn as _qn
    doc = Document()
    p = doc.add_paragraph(style='List Bullet')
    p.add_run("Bullet Text")
    # List Bullet-Style setzt w:numPr → muss nach _hide_paragraph weg sein
    _hide_paragraph(p)
    pPr = p._p.find(_qn('w:pPr'))
    assert pPr is not None, "pPr fehlt"
    numPr = pPr.find(_qn('w:numPr'))
    assert numPr is None, "w:numPr ist noch vorhanden – Bullet-Label bleibt sichtbar"


# ---------------------------------------------------------------------------
# Gruppe 19: add_markdown_table_to_doc – Titel-Zeile erkennen
# ---------------------------------------------------------------------------

def test_table_title_row_merged():
    """Erste Zeile mit nur 1 Non-Empty-Cell → wird zur gemergten Titelzeile; Zeile 1 bekommt Shading."""
    from docx.oxml.ns import qn as _qn
    doc = Document()
    table_lines = [
        "| Lange Fragestellung |  |  |",
        "|---------------------|--|--|",
        "| Spalte A | Spalte B | Spalte C |",
        "| Wert 1   | Wert 2   | Wert 3   |",
    ]
    add_markdown_table_to_doc(doc, table_lines)
    # Zuletzt eingefügte Tabelle
    tbl = doc.tables[-1]
    # Zeile 0 soll nach merge nur noch 1 Zelle haben (alle zusammengeführt)
    row0_cells = tbl.rows[0].cells
    assert row0_cells[0] is row0_cells[-1], "Titelzeile wurde nicht zusammengeführt (Zellen nicht identisch)"
    # Zeile 1 (echter Header) soll Shading haben
    tc1 = tbl.rows[1].cells[0]._tc
    tcPr = tc1.find(_qn('w:tcPr'))
    assert tcPr is not None, "tcPr in Header-Zeile fehlt"
    shd = tcPr.find(_qn('w:shd'))
    assert shd is not None, "w:shd fehlt – Header-Zeile hat kein Shading"


def test_table_normal_header_unchanged():
    """Tabelle mit mehreren Non-Empty-Cells in Zeile 0 → normale Header-Logik (Zeile 0 = Header)."""
    from docx.oxml.ns import qn as _qn
    doc = Document()
    table_lines = [
        "| Name | Wert |",
        "|------|------|",
        "| A    | 1    |",
    ]
    add_markdown_table_to_doc(doc, table_lines)
    tbl = doc.tables[-1]
    # Zeile 0 soll Shading haben (normaler Header)
    tc0 = tbl.rows[0].cells[0]._tc
    tcPr = tc0.find(_qn('w:tcPr'))
    assert tcPr is not None
    shd = tcPr.find(_qn('w:shd'))
    assert shd is not None, "w:shd fehlt in Zeile 0 – normaler Header nicht erkannt"


# ---------------------------------------------------------------------------
# Gruppe 20: _set_cell_text – Bullet-Concatenation
# ---------------------------------------------------------------------------

def test_set_cell_text_splits_bullets():
    """•A•B•C in einer Zelle → 3 separate Paragraphen."""
    doc = Document()
    tbl = doc.add_table(rows=1, cols=1)
    cell = tbl.cell(0, 0)
    _set_cell_text(cell, "•Wiedererkennen•Benennen•Abrufen")
    texts = [p.text.strip() for p in cell.paragraphs if p.text.strip()]
    assert len(texts) == 3, f"Erwartet 3 Paragraphen, erhalten: {texts}"
    assert "Wiedererkennen" in texts
    assert "Benennen" in texts
    assert "Abrufen" in texts


def test_set_cell_text_single_bullet_unchanged():
    """Einzelnes • am Anfang (kein Concat) → kein Split."""
    doc = Document()
    tbl = doc.add_table(rows=1, cols=1)
    cell = tbl.cell(0, 0)
    _set_cell_text(cell, "•Eintrag ohne Split")
    texts = [p.text.strip() for p in cell.paragraphs if p.text.strip()]
    assert len(texts) == 1, f"Kein Split erwartet, erhalten: {texts}"


# ---------------------------------------------------------------------------
# Gruppe 21: process_markdown_to_docx – headings_as_bold
# ---------------------------------------------------------------------------

def test_process_markdown_headings_as_bold():
    """headings_as_bold=True → ### Titel wird als Normal-Paragraph mit bold=True eingefügt."""
    doc = Document()
    process_markdown_to_docx(doc, "### Durchführung\nText darunter.", headings_as_bold=True)
    # Erster Nicht-Leer-Paragraph soll Normal-Style (kein Heading) sein
    paras = [p for p in doc.paragraphs if p.text.strip()]
    assert len(paras) >= 1
    first = paras[0]
    assert first.style.name == 'Normal', f"Erwartet Normal, erhalten: {first.style.name}"
    assert first.runs and first.runs[0].bold, "Heading-Text soll bold=True sein"


def test_process_markdown_headings_default_uses_heading_style():
    """headings_as_bold=False (Standard) → ### Titel wird als Heading 3 eingefügt."""
    doc = Document()
    process_markdown_to_docx(doc, "### Titel", headings_as_bold=False)
    paras = [p for p in doc.paragraphs if p.text.strip()]
    assert len(paras) >= 1
    assert 'Heading' in paras[0].style.name or paras[0].style.name.startswith('berschrift'), \
        f"Erwartet Heading-Style, erhalten: {paras[0].style.name}"


# ---------------------------------------------------------------------------
# Gruppe 22: _strip_ocr_y_prefix – Headings
# ---------------------------------------------------------------------------

def test_strip_y_prefix_heading():
    """'## y **Titel**' → '## **Titel**' (y-Präfix entfernt)."""
    text = "## y **Woran erkenne ich das?**"
    result = _strip_ocr_y_prefix(text)
    assert result == "## **Woran erkenne ich das?**"


def test_strip_y_prefix_all_heading_levels():
    """y-Präfix wird bei allen Heading-Levels (# bis ######) entfernt."""
    for level in range(1, 7):
        hashes = '#' * level
        text = f"{hashes} y Titel"
        result = _strip_ocr_y_prefix(text)
        assert result == f"{hashes} Titel", f"Fehler bei Level {level}: {result!r}"


# ---------------------------------------------------------------------------
# Gruppe 23: _strip_ocr_y_prefix – Bullets und Sonderfälle
# ---------------------------------------------------------------------------

def test_strip_y_prefix_bullet():
    """'- y Inhalt' → '- Inhalt' (y-Präfix bei Bullet-Items entfernt)."""
    text = "- y erfasst neue Situationen"
    result = _strip_ocr_y_prefix(text)
    assert result == "- erfasst neue Situationen"


def test_strip_y_prefix_no_false_positive_word():
    """'- young person' bleibt unverändert (y ist Teil eines Wortes, kein isoliertes Präfix)."""
    text = "- young person"
    result = _strip_ocr_y_prefix(text)
    assert result == "- young person"


def test_strip_y_prefix_regular_text_unchanged():
    """Fließtext mit 'y' bleibt unverändert (kein Heading/Bullet-Präfix)."""
    text = "Das Wort yesterday enthält y, aber das ist kein Präfix."
    result = _strip_ocr_y_prefix(text)
    assert result == text


def test_strip_y_prefix_multiline():
    """Mehrere Zeilen: nur Zeilen mit y-Präfix werden bereinigt."""
    text = "## y Kapitel\nNormaler Text mit y drin.\n- y Bullet\n- Normales Bullet"
    result = _strip_ocr_y_prefix(text)
    lines = result.split('\n')
    assert lines[0] == "## Kapitel"
    assert lines[1] == "Normaler Text mit y drin."
    assert lines[2] == "- Bullet"
    assert lines[3] == "- Normales Bullet"


# ---------------------------------------------------------------------------
# Gruppe 24: build_interleaved_word_document – unnummerierte Headings in Nav
# ---------------------------------------------------------------------------

def _run_interleaved(orig_md: str, summary_md: str) -> list:
    """Hilfsfunktion: build_interleaved_word_document in Temp-Datei ausführen, Paragraphs zurückgeben."""
    import tempfile, os
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        tmp_path = tmp.name
    try:
        build_interleaved_word_document(
            translated_text=orig_md,
            summary_text=summary_md,
            qa_text="",
            output_path=tmp_path,
        )
        from docx import Document as _Doc
        doc = _Doc(tmp_path)
        return [p for p in doc.paragraphs if p.text.strip()]
    finally:
        os.unlink(tmp_path)


def test_interleaved_unnumbered_heading_not_in_nav():
    """Wiederkehrende unnummerierte Überschrift (freq>1) → Normal+Bold, nicht in Nav."""
    # "Was ist das?" erscheint unter 5.1 UND 5.2 → freq=2 → kein Auto-Nummerierung
    orig_md = (
        "## 5.1 Kapitel\n\n### Was ist das?\n\nDefinition A.\n\n"
        "## 5.2 Anderes\n\n### Was ist das?\n\nDefinition B."
    )
    sum_md = "## 5.1 Kapitel\n\nZusammenfassung.\n\n### Was ist das?\n\nKurze Zusammenfassung."
    paras = _run_interleaved(orig_md, sum_md)
    heading_para = next((p for p in paras if "Was ist das" in p.text), None)
    assert heading_para is not None, "Überschrift 'Was ist das?' nicht im Dokument gefunden"
    assert heading_para.style.name == 'Normal', (
        f"Wiederkehrende Überschrift soll Normal-Style haben, erhalten: {heading_para.style.name!r}"
    )
    assert heading_para.runs and heading_para.runs[0].bold, \
        "Wiederkehrende Überschrift soll bold=True sein"


def test_interleaved_numbered_heading_in_nav():
    """Nummerierte Sektion mit Summary → Heading-Style (erscheint in Nav)."""
    orig_md = "## 5.1 Einführung\n\nHier steht der Originaltext."
    sum_md = "## 5.1 Einführung\n\nKurze Zusammenfassung."
    paras = _run_interleaved(orig_md, sum_md)
    heading_para = next((p for p in paras if "5.1 Einführung" in p.text), None)
    assert heading_para is not None, "Überschrift '5.1 Einführung' nicht im Dokument gefunden"
    assert 'Heading' in heading_para.style.name or heading_para.style.name.startswith('berschrift'), (
        f"Nummerierte Überschrift soll Heading-Style haben, erhalten: {heading_para.style.name!r}"
    )


def test_interleaved_unnumbered_h2_in_nav_when_no_numbered():
    """Unnummerierte H2 in Dokument OHNE nummerierte Kapitel → Heading-Style (smart nav)."""
    orig_md = "## Einführung\n\nHier steht der Originaltext."
    sum_md = "## Einführung\n\nKurze Zusammenfassung."
    paras = _run_interleaved(orig_md, sum_md)
    heading_para = next((p for p in paras if "Einführung" in p.text), None)
    assert heading_para is not None, "Überschrift 'Einführung' nicht im Dokument gefunden"
    assert 'Heading' in heading_para.style.name or heading_para.style.name.startswith('berschrift'), (
        f"Unnummerierte H2 (kein nummeriertes Dokument) soll Heading-Style haben, erhalten: {heading_para.style.name!r}"
    )


def test_interleaved_unnumbered_h2_not_in_nav_when_numbered_doc():
    """Wiederkehrende unnummerierte H2 mit Summary in numm. Dokument → Normal+Bold (nicht in Nav)."""
    # "Woran erkenne ich das?" erscheint zweimal → freq=2 → keine Auto-Nummerierung.
    # Summary enthält "Woran erkenne ich das?" → wird angezeigt, aber als Normal+Bold (nicht in Nav).
    orig_md = (
        "## 5.1 Einführung\n\nKapiteltext.\n\n## Woran erkenne ich das?\n\nText A.\n\n"
        "## 5.2 Weiteres\n\nKapiteltext.\n\n## Woran erkenne ich das?\n\nText B."
    )
    sum_md = (
        "## 5.1 Einführung\n\nZusammenfassung.\n\n"
        "## 5.2 Weiteres\n\nZusammenfassung 2.\n\n"
        "## Woran erkenne ich das?\n\nKompetenz-Merkmale."
    )
    paras = _run_interleaved(orig_md, sum_md)
    heading_para = next((p for p in paras if "Woran erkenne ich" in p.text), None)
    assert heading_para is not None, "Überschrift 'Woran erkenne ich das?' nicht gefunden"
    assert heading_para.style.name == 'Normal', (
        f"Wiederkehrende H2 in numm. Dokument soll Normal-Style haben, erhalten: {heading_para.style.name!r}"
    )
    assert heading_para.runs and heading_para.runs[0].bold, "Soll bold=True sein"


# ---------------------------------------------------------------------------
# Gruppe 25: has_children erkennt unnummerierte Folge-Sektion als konzept. Kind
# ---------------------------------------------------------------------------

def test_numbered_section_not_skipped_when_followed_by_unnumbered():
    """Nummerierte Sektion ohne Body/Summary erscheint, wenn direkt danach unnummerierte Sektion folgt."""
    # 5.3 hat keinen Body und kein Summary, aber "Agilität" (unnummeriert) folgt direkt
    orig_md = "## 5.3 Überblick\n\n## Agilität\n\nKompetenz-Inhalt hier."
    sum_md = "## Agilität\n\nAgilität-Zusammenfassung."
    paras = _run_interleaved(orig_md, sum_md)
    # 5.3 muss im Dokument erscheinen (nicht übersprungen)
    heading_para = next((p for p in paras if "5.3 Überblick" in p.text), None)
    assert heading_para is not None, "5.3 Überblick soll nicht übersprungen werden (has_children via unnummerierte Folgesektion)"
    assert 'Heading' in heading_para.style.name or heading_para.style.name.startswith('berschrift'), \
        f"5.3 soll Heading-Style haben, erhalten: {heading_para.style.name!r}"


# ---------------------------------------------------------------------------
# Gruppe 26: Auto-Nummerierung einzigartiger Unterkapitel + Skip-Fix für Orig-Body
# ---------------------------------------------------------------------------

def test_auto_numbering_unique_heading_gets_number_and_nav():
    """Einzigartige unnummerierte Überschrift nach numm. Kapitel → auto 5.3.1, in Nav."""
    orig_md = "## 5.3 Überblick\n\n## Agilität\n\nAgilität ist die Fähigkeit..."
    sum_md = "## Agilität\n\nZusammenfassung Agilität."
    paras = _run_interleaved(orig_md, sum_md)
    heading_para = next((p for p in paras if "5.3.1" in p.text and "Agilität" in p.text), None)
    assert heading_para is not None, "Einzigartige Überschrift soll als '5.3.1 Agilität' erscheinen"
    assert 'Heading' in heading_para.style.name or heading_para.style.name.startswith('berschrift'), \
        f"Auto-nummerierte Überschrift soll Heading-Style haben, erhalten: {heading_para.style.name!r}"


def test_auto_numbering_recurring_heading_stays_bold_not_nav():
    """Wiederkehrende unnummerierte Überschrift mit Summary → Normal+Bold, keine Auto-Nummerierung."""
    # "Was ist das?" erscheint zweimal → freq=2 → NICHT auto-nummeriert → Normal+Bold (nicht Heading-Style).
    # Summary enthält "Was ist das?" damit es nicht übersprungen wird.
    orig_md = (
        "## 5.3 Überblick\n\n"
        "## Agilität\n\n"
        "## Was ist das?\n\nDefinition Agilität.\n\n"
        "## Analytisches Denken\n\n"
        "## Was ist das?\n\nDefinition Analytik."
    )
    sum_md = (
        "## Agilität\n\nAgilität-Zusammenfassung.\n\n"
        "## Analytisches Denken\n\nAnalytik-Zusammenfassung.\n\n"
        "## Was ist das?\n\nDefinition."
    )
    paras = _run_interleaved(orig_md, sum_md)
    was_paras = [p for p in paras if "Was ist das?" in p.text]
    assert len(was_paras) >= 1, "Mindestens eine 'Was ist das?'-Überschrift soll im Dokument erscheinen"
    for wp in was_paras:
        assert wp.style.name == 'Normal', \
            f"'Was ist das?' soll Normal-Style haben (wiederkehrend), erhalten: {wp.style.name!r}"


def test_unnumbered_sibling_heading_hidden_when_no_summary():
    """Unnummerierte Geschwister-Sektion ohne Summary → Überschrift ausgeblendet.

    'Off the job' und 'On the job' erscheinen mehrfach (freq>1) → keine Auto-Nummerierung.
    Ohne Summary → show_heading_visible=False → Überschrift ausgeblendet (w:vanish).
    Die vorherige falsche Logik setzte has_children=True für jede unnummerierte Folge-Sektion
    (auch Geschwister), was diese sichtbar machte.
    """
    orig_md = (
        "## 5.3.1 Durchsetzen\n\n"
        "## Off the job\n\nMaßnahme A.\n\n"
        "## On the job\n\nMaßnahme B.\n\n"
        "## 5.3.2 Empathie\n\n"
        "## Off the job\n\nMaßnahme C."
    )
    sum_md = (
        "## 5.3.1 Durchsetzen\n\nZusammenfassung Durchsetzen.\n\n"
        "## 5.3.2 Empathie\n\nZusammenfassung Empathie."
    )
    paras = _run_interleaved(orig_md, sum_md)

    from lxml import etree
    from docx.oxml.ns import qn

    off_paras = [p for p in paras if "Off the job" in p.text]
    assert len(off_paras) >= 1, "'Off the job' muss im Dokument vorhanden sein (Originaltext erhalten)"
    for p in off_paras:
        pPr = p._p.find(qn('w:pPr'))
        has_vanish = False
        if pPr is not None:
            rPr = pPr.find(qn('w:rPr'))
            if rPr is not None and rPr.find(qn('w:vanish')) is not None:
                has_vanish = True
        assert has_vanish, \
            f"'Off the job' ohne Summary soll ausgeblendet (w:vanish) sein, Style: {p.style.name!r}"


# ---------------------------------------------------------------------------
# Gruppe 27: _any_visible_desc – kein sichtbarer Elternknoten ohne sichtbare Kinder
# ---------------------------------------------------------------------------

def test_numbered_parent_hidden_when_all_children_have_no_summary():
    """Nummerierter Parent ohne eigenes Summary und ohne Summary in Kindern → nicht sichtbar.

    Szenario analog zu '5.3.27 Rückmeldung zu Konfliktverhalten' im echten Dokument:
    Das OCR hat 'Rückmeldung zu Konfliktverhalten' als eigene Kompetenz mit Originaltext,
    die KI-Summary enthält sie NICHT als eigene Überschrift. Auch Kinder ('Off the job',
    'On the job') haben kein Summary.
    Früher: has_children=True → sichtbare goldene Leer-Überschrift.
    Nach Fix: _any_visible_desc=False → has_children=False → Überschrift ausgeblendet.

    Außerdem: '5.3 Überblick' hat 'Agilität' als konzeptuelles Kind (gleiche Ebene,
    nummerierter Parent + unnummerierter Folger). 'Agilität' hat Summary →
    _any_visible_desc[5.3]=True → '5.3 Überblick' bleibt sichtbar (kein Rückschritt).
    """
    orig_md = (
        "## 5.3 Überblick\n\n"
        "## Rückmeldung zu Konfliktverhalten\n\nWas ist das? Originaltext hier.\n\n"
        "## Off the job\n\nMaßnahme A.\n\n"
        "## On the job\n\nMaßnahme B.\n\n"
        "## Agilität\n\nAgilität ist die Fähigkeit...\n\n"
        "## Off the job\n\nMaßnahme C.\n\n"
        "## On the job\n\nMaßnahme D."
    )
    # Summary hat KEINE Einträge für Rückmeldung, Off/On the job.
    # Agilität hat ein Summary → Elternknoten "5.3 Überblick" bleibt sichtbar (hat sichtbares Kind).
    sum_md = (
        "## 5.3 Überblick\n\nÜberblick-Text.\n\n"
        "## Agilität\n\nAgilität-Zusammenfassung."
    )
    paras = _run_interleaved(orig_md, sum_md)

    from docx.oxml.ns import qn

    # "Rückmeldung zu Konfliktverhalten" muss im Dokument vorhanden sein (Originaltext erhalten)
    rueckmeldung_paras = [p for p in paras if "Rückmeldung zu Konfliktverhalten" in p.text]
    assert len(rueckmeldung_paras) >= 1, "'Rückmeldung zu Konfliktverhalten' fehlt im Dokument"

    # Muss als hidden (w:vanish) markiert sein – keine sichtbare Leer-Überschrift
    for p in rueckmeldung_paras:
        pPr = p._p.find(qn('w:pPr'))
        has_vanish = False
        if pPr is not None:
            rPr = pPr.find(qn('w:rPr'))
            if rPr is not None and rPr.find(qn('w:vanish')) is not None:
                has_vanish = True
        assert has_vanish, (
            "'Rückmeldung zu Konfliktverhalten' ohne Summary und ohne sichtbare Kinder "
            f"soll ausgeblendet sein, Style: {p.style.name!r}"
        )

    # Zur Sicherheit: "Agilität" MUSS sichtbar sein (hat Summary)
    agilitat_paras = [p for p in paras if "Agilität" in p.text and "Zusammenfassung" not in p.text]
    assert any(
        p._p.find(qn('w:pPr')) is None or
        (p._p.find(qn('w:pPr')).find(qn('w:rPr')) is None) or
        (p._p.find(qn('w:pPr')).find(qn('w:rPr')).find(qn('w:vanish')) is None)
        for p in agilitat_paras
    ), "Agilität mit Summary muss sichtbar sein"


# ---------------------------------------------------------------------------
# Gruppe 28: Sequentielle Nummerierung (keine Lücken)
# ---------------------------------------------------------------------------

def test_sequential_numbering_no_gaps():
    """3 Kompetenzen, nur 2 in Summary → Navigation zeigt 5.3.1, 5.3.2 statt 5.3.1, 5.3.3."""
    orig_md = (
        "## 5.3 Überblick\n\nEinleitungstext.\n\n"
        "## Agilität\n\nAgilität Originaltext.\n\n"
        "## Belastbarkeit\n\nBelastbarkeit Originaltext.\n\n"
        "## Durchsetzungsvermögen\n\nDurchsetzung Originaltext."
    )
    sum_md = (
        "## 5.3 Überblick\n\nÜberblick-Zusammenfassung.\n\n"
        "## Agilität\n\nAgilität-Zusammenfassung.\n\n"
        # Belastbarkeit fehlt im Summary
        "## Durchsetzungsvermögen\n\nDurchsetzung-Zusammenfassung."
    )
    paras = _run_interleaved(orig_md, sum_md)
    visible_texts = [p.text for p in paras]

    # 5.3.1 und 5.3.2 müssen vorhanden sein (keine Lücke 5.3.1, 5.3.3)
    assert any("5.3.1" in t for t in visible_texts), "5.3.1 fehlt"
    assert any("5.3.2" in t for t in visible_texts), "5.3.2 fehlt (Lücke statt sequentiell)"
    # 5.3.3 darf nicht existieren (es gibt nur 2 sichtbare Kompetenzen)
    assert not any("5.3.3" in t for t in visible_texts), "5.3.3 sollte nicht existieren"
    # Belastbarkeit bekommt keine sichtbare Nummer – darf nicht in Nav erscheinen
    belastbarkeit_nav = [t for t in visible_texts if "5.3.2 Belastbarkeit" in t or "5.3.3 Belastbarkeit" in t]
    assert len(belastbarkeit_nav) == 0, "Belastbarkeit ohne Summary darf keine Nummer haben"


# ---------------------------------------------------------------------------
# Gruppe 29: Scoped Lookup (Sub-Sections kompetenzbezogen)
# ---------------------------------------------------------------------------

def test_scoped_lookup_off_the_job_scoped_per_competency():
    """'Off the job' unter Kompetenz A soll Inhalt von A zeigen, nicht von B."""
    orig_md = (
        "## 5.3 Überblick\n\nEinleitungstext.\n\n"
        "## Agilität\n\nAgilität Originaltext.\n\n"
        "## Off the job\n\nMaßnahme Agilität original.\n\n"
        "## Belastbarkeit\n\nBelastbarkeit Originaltext.\n\n"
        "## Off the job\n\nMaßnahme Belastbarkeit original."
    )
    sum_md = (
        "## 5.3 Überblick\n\nÜberblick-Zusammenfassung.\n\n"
        "## Agilität\n\nAgilität-Zusammenfassung.\n\n"
        "#### Off the job\n\nAgilität-Maßnahme-A.\n\n"
        "## Belastbarkeit\n\nBelastbarkeit-Zusammenfassung.\n\n"
        "#### Off the job\n\nBelastbarkeit-Maßnahme-B."
    )
    paras = _run_interleaved(orig_md, sum_md)
    all_texts = " | ".join(p.text for p in paras)

    # Beide Kompetenzen müssen vorhanden sein
    assert "Agilität" in all_texts, "Agilität fehlt"
    assert "Belastbarkeit" in all_texts, "Belastbarkeit fehlt"
    # Beide Off-the-job-Inhalte müssen vorhanden sein (nicht nur der letzte)
    assert "Agilität-Maßnahme-A" in all_texts, "Agilität Off-the-job-Inhalt fehlt"
    assert "Belastbarkeit-Maßnahme-B" in all_texts, "Belastbarkeit Off-the-job-Inhalt fehlt"


# ---------------------------------------------------------------------------
# Gruppe 30: Platzhalter-Sektionen überspringen
# ---------------------------------------------------------------------------

def test_placeholder_section_skipped():
    """Section mit \\_\\_\\_ im Body (OCR-Notizlinien) wird komplett übersprungen."""
    orig_md = (
        "## 5.3 Überblick\n\nEinleitungstext.\n\n"
        "## Agilität\n\nAgilität Originaltext.\n\n"
        "## Meine persönlichen Anmerkungen\n\n\\_\\_\\_\\_\\_\\_\\_\\_\\_\\_\n\n"
        "## Wo finde ich noch mehr darüber\n\nLiteraturliste."
    )
    sum_md = (
        "## 5.3 Überblick\n\nÜberblick-Zusammenfassung.\n\n"
        "## Agilität\n\nAgilität-Zusammenfassung.\n\n"
        "## Wo finde ich noch mehr darüber\n\nBuch A, Buch B."
    )
    paras = _run_interleaved(orig_md, sum_md)
    all_texts = [p.text for p in paras]

    # "Meine persönlichen Anmerkungen" darf weder sichtbar noch als Heading erscheinen
    assert not any("Meine persönlichen Anmerkungen" in t for t in all_texts), (
        "'Meine persönlichen Anmerkungen' (Platzhalter) darf nicht im Dokument erscheinen"
    )
    # "Wo finde ich" darf erscheinen (hat Summary)
    assert any("Wo finde ich" in t for t in all_texts), "'Wo finde ich' fehlt im Dokument"

