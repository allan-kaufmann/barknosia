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

def test_image_filter_skips_picture(tmp_path):
    """Bild mit 'Picture' im Namen wird übersprungen (kein Paragraph mit Pfad)."""
    fake_img = tmp_path / "_page_0_Picture_2.jpeg"
    fake_img.write_bytes(b"")
    doc = Document()
    process_markdown_to_docx(doc, f"![]({fake_img.name})", base_path=str(tmp_path))
    texts = ' '.join(p.text for p in doc.paragraphs)
    assert fake_img.name not in texts


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
