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
