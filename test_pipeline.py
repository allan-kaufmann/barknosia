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

def test_normalize_level1_to_h2():
    result = normalize_heading_levels("# 1 Einleitung\nText")
    assert result.startswith("## 1 Einleitung")


def test_normalize_level2_to_h3():
    result = normalize_heading_levels("# 1.1 Abschnitt\nText")
    assert result.startswith("### 1.1 Abschnitt")


def test_normalize_level3_to_h4():
    result = normalize_heading_levels("# 1.1.1 Unterabschnitt\nText")
    assert result.startswith("#### 1.1.1 Unterabschnitt")


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
    """4.1.1.1 (3 Punkte) → H5 (#####)."""
    result = normalize_heading_levels("# 4.1.1.1 Tief\nText")
    assert result.startswith("##### 4.1.1.1 Tief")


def test_normalize_5level_depth():
    """5 Punkte → max H6 (######)."""
    result = normalize_heading_levels("# 1.2.3.4.5 Sehr tief\nText")
    assert result.startswith("###### 1.2.3.4.5 Sehr tief")


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

def test_interleaved_unnumbered_heading_not_in_nav():
    """Unnummerierte Sektion mit Summary → Normal-Style (kein Heading), erster Run bold."""
    orig_md = "## Was ist das?\n\nHier steht der Originaltext."
    sum_lookup = {"Was ist das?": "Kurze Zusammenfassung des Inhalts."}
    doc = build_interleaved_word_document(
        translated_text=orig_md,
        summary_text="## Was ist das?\n\nKurze Zusammenfassung des Inhalts.",
        qa_result="",
        output_path=None,
    )
    non_empty = [p for p in doc.paragraphs if p.text.strip()]
    heading_para = non_empty[0]
    assert heading_para.style.name == 'Normal', (
        f"Unnummerierte Überschrift soll Normal-Style haben, erhalten: {heading_para.style.name!r}"
    )
    assert heading_para.runs and heading_para.runs[0].bold, \
        "Unnummerierte Überschrift soll bold=True sein"


def test_interleaved_numbered_heading_in_nav():
    """Nummerierte Sektion mit Summary → Heading-Style (erscheint in Nav)."""
    orig_md = "## 5.1 Einführung\n\nHier steht der Originaltext."
    doc = build_interleaved_word_document(
        translated_text=orig_md,
        summary_text="## 5.1 Einführung\n\nKurze Zusammenfassung.",
        qa_result="",
        output_path=None,
    )
    non_empty = [p for p in doc.paragraphs if p.text.strip()]
    heading_para = non_empty[0]
    assert 'Heading' in heading_para.style.name or heading_para.style.name.startswith('berschrift'), (
        f"Nummerierte Überschrift soll Heading-Style haben, erhalten: {heading_para.style.name!r}"
    )
