# Barknosia — Anforderungen (User Stories)

## Systemübersicht

**Barknosia** ist eine Automatisierungs-Pipeline für akademische und Lehr-PDFs. Sie wandelt PDFs (typischerweise englischsprachige Fachartikel oder Buchkapitel) in aufbereitete, deutsche Lernunterlagen im Word-Format um.

**Kernidee:** Das fertige Word-Dokument enthält zwei Ebenen gleichzeitig:
1. **Sichtbare Ebene** — eine KI-generierte, stichpunktartige Zusammenfassung je Abschnitt
2. **Referenz-Ebene** — der vollständige Originaltext, ausgeblendet (nur sichtbar bei "ausgeblendeten Text anzeigen")

**Technische Bausteine:**
| Baustein | Funktion |
|---|---|
| Marker OCR | PDF → Markdown-Extraktion |
| Gemini 2.5 Pro | Spracherkennung, Übersetzung, Zusammenfassung, QA |
| python-docx | Word-Dokument-Erzeugung |

---

## Epic 1 — PDF-Einlesen (OCR)

### US-1.1 — PDF einlesen
**Als** Nutzer  
**möchte ich** ein beliebiges PDF als Eingabe angeben,  
**damit** der Text automatisch extrahiert und weiterverarbeitet werden kann.

**Akzeptanzkriterien:**
- Aufruf: `python pipeline.py mein_dokument.pdf`
- Marker OCR erzeugt eine Markdown-Datei im Ausgabeordner (`workspace/output/{name}/{name}.md`)
- Fehlerhafte oder fehlende OCR-Datei bricht die Pipeline mit einer klaren Fehlermeldung ab

### US-1.2 — Einzelnes Kapitel extrahieren
**Als** Nutzer  
**möchte ich** optional nur ein bestimmtes Kapitel verarbeiten,  
**damit** ich nicht immer das gesamte Dokument neu berechnen muss.

**Akzeptanzkriterien:**
- Aufruf: `python pipeline.py dok.pdf --chapter "4.2"`
- Nur das angegebene Kapitel und alle seine Unterkapitel werden extrahiert und verarbeitet
- Ausgabedateien erhalten einen Kapitel-Suffix

---

## Epic 2 — Übersetzung

### US-2.1 — Automatische Spracherkennung
**Als** Nutzer  
**möchte ich** dass die Pipeline automatisch erkennt, ob der Text Englisch oder Deutsch ist,  
**damit** ich nicht manuell angeben muss, ob eine Übersetzung nötig ist.

**Akzeptanzkriterien:**
- Englischer Text wird automatisch übersetzt
- Bereits deutscher Text überspringt den Übersetzungsschritt

### US-2.2 — Übersetzung Englisch → Deutsch
**Als** Nutzer  
**möchte ich** eine qualitativ hochwertige deutsche Übersetzung des Originaltexts,  
**damit** ich englische Fachliteratur auf Deutsch bearbeiten kann.

**Akzeptanzkriterien:**
- Übersetzung via Gemini 2.5 Pro, kapitelweise
- Fachbegriffe werden einheitlich übersetzt
- Ausgabe: `de_uebersetzung.md` im Ausgabeordner

### US-2.3 — Kapitelweises Caching
**Als** Nutzer  
**möchte ich** dass abgebrochene Läufe fortsetzbar sind,  
**damit** ich nicht bei einem API-Fehler von vorne beginnen muss.

**Akzeptanzkriterien:**
- Jedes übersetzte Kapitel wird separat gecacht (`zusammenfassung_kap_XX.md`)
- Bereits gecachte Kapitel werden beim nächsten Lauf übersprungen
- `--force` erzwingt vollständige Neuberechnung (ignoriert Cache)

### US-2.4 — Übersetzungs-Dokument ausgeben
**Als** Nutzer  
**möchte ich** die Übersetzung auch als formatiertes Word-Dokument erhalten,  
**damit** ich sie direkt lesen oder weiterverwenden kann.

**Akzeptanzkriterien:**
- Ausgabe: `{name}_Uebersetzung.docx`
- Farbige Überschriften im MM-Skript-Farbschema (siehe Epic 6)
- Mit `--no-summary` wird nur dieses Dokument erzeugt (kein Lernmittel)

---

## Epic 3 — Zusammenfassung

### US-3.1 — Lernzusammenfassung generieren
**Als** Nutzer  
**möchte ich** für jeden Abschnitt des Dokuments eine prägnante Zusammenfassung,  
**damit** ich die wesentlichen Inhalte effizient erfassen kann.

**Akzeptanzkriterien:**
- KI (Gemini 2.5 Pro) fasst jeden Abschnitt als Stichpunkte zusammen
- Zusammenfassung hält die Gliederungsstruktur des Originaldokuments ein
- Ausgabe: `zusammenfassung.md`
- **Qualitätsanforderung:** Kein Abschnitt darf eine leere Zusammenfassung haben — wenn ein Abschnitt im Original Inhalt hat, muss die Zusammenfassung etwas enthalten

### US-3.2 — Zusammenfassung ist stichpunktartig
**Als** Nutzer  
**möchte ich** die Zusammenfassung als prägnante Stichpunkte,  
**damit** ich schnell die Kernaussagen überblicken kann (kein Fließtext).

---

## Epic 4 — Qualitätsprüfung (optional)

### US-4.1 — Lernfragen beantworten
**Als** Nutzer  
**möchte ich** eigene Lernfragen gegen die Zusammenfassung prüfen lassen,  
**damit** ich sehen kann, ob die wichtigen Konzepte abgedeckt sind.

**Akzeptanzkriterien:**
- Aufruf: `python pipeline.py dok.pdf --questions meine_fragen.txt`
- Gemini beantwortet jede Frage unter Angabe der Textgrundlage
- Ausgabe: strukturierte `qa_ergebnis.md` mit Frage, Antwort, Textgrundlage, Schlüsselbegriffen, Abdeckungsgrad
- QA-Ergebnisse erscheinen am Ende des Lernmittel-Dokuments

---

## Epic 5 — Word-Ausgabe: Übersetzungsdokument

### US-5.1 — Überschriften überspringen
**Als** Nutzer  
**möchte ich** dass irrelevante Abschnitte (Referenzen, Autorenblöcke) automatisch übersprungen werden,  
**damit** das Dokument nur Lerninhalt enthält.

**Akzeptanzkriterien:**
- Übersprungen werden: "Referenzen", "Literaturverzeichnis", "History"/"Historie", Abschnitte mit E-Mail-Adressen oder akademischen Titeln (Prof. Dr., M.Sc.)
- Mit `--include-references` werden Referenzen dennoch einbezogen
- Normale Kapitelüberschriften werden nie übersprungen

---

## Epic 6 — Word-Ausgabe: Lernmittel *(Kern-Epic)*

Dies ist der wichtigste und komplexeste Teil der Pipeline. Alle Regeln hier sind verbindlich.

### US-6.1 — Originaltext vollständig erhalten
**Als** Nutzer  
**möchte ich** dass der vollständige Originaltext immer im Dokument enthalten ist,  
**damit** ich bei Bedarf den genauen Wortlaut nachschlagen kann.

**Akzeptanzkriterien:**
- **Kein einziges Wort des Originaltexts wird gelöscht oder weggelassen**
- Der Originaltext wird als ausgeblendeter Text dargestellt (Word: "Hidden")
- Formatierung: 1,5 cm Einzug, Schriftfarbe Grau (RGB 187,187,187)
- Sichtbar wenn der Nutzer in Word "Alle Formatierungszeichen anzeigen" / "Ausgeblendeten Text drucken" aktiviert
- Bilder und Bildunterschriften werden NIE ausgeblendet

### US-6.2 — Zusammenfassung sichtbar, Original ausgeblendet
**Als** Nutzer  
**möchte ich** im normalen Lesemodus nur die Zusammenfassung sehen,  
**damit** ich das Dokument als kompaktes Lernmittel nutzen kann.

**Akzeptanzkriterien:**
- Zusammenfassungstext ist sichtbar (keine versteckte Formatierung)
- Originaltext ist ausgeblendet (direkt nach der zugehörigen Zusammenfassung)
- Der Originaltext wird NIEMALS als sichtbarer Fallback angezeigt — auch nicht wenn die Zusammenfassung fehlt

### US-6.3 — Abschnitte ohne Zusammenfassung: Überschrift und Inhalt ausblenden
**Als** Nutzer  
**möchte ich** dass Abschnitte ohne Zusammenfassung im normalen Lesemodus nicht sichtbar sind,  
**damit** keine leeren Gliederungspunkte im Dokument erscheinen.

**Akzeptanzkriterien:**
- Hat ein Abschnitt keinen Zusammenfassungstext → Überschrift UND Originaltext werden ausgeblendet
- Im normalen Lesemodus ist dieser Abschnitt komplett unsichtbar (keine leere Überschrift)
- Im Referenz-Modus (ausgeblendeten Text anzeigen) ist alles sichtbar
- **Ausnahme:** Elternkapitel ohne eigenen Text, aber mit Unterkapiteln → Überschrift bleibt sichtbar (sie strukturiert die Unterkapitel)

### US-6.4 — Wirklich leere Abschnitte überspringen
**Als** Nutzer  
**möchte ich** dass Abschnitte ohne jeglichen Inhalt (kein Original, keine Zusammenfassung, keine Unterkapitel) übersprungen werden,  
**damit** das Dokument keine unnötigen Leerzeilen enthält.

**Akzeptanzkriterien:**
- Kein Original UND kein Summary UND keine Unterkapitel → komplett übersprungen (nicht im Dokument)
- Diese Regel greift nur wenn wirklich gar nichts vorhanden ist

### US-6.5 — Gliederungsstruktur erhalten
**Als** Nutzer  
**möchte ich** dass die Gliederung des Originaldokuments im Word-Navigationspanel erkennbar ist,  
**damit** ich schnell zwischen Kapiteln navigieren kann.

**Akzeptanzkriterien:**

| Überschrift-Typ | Navigationspanel (Word Outline) | Formatierung |
|---|---|---|
| Nummerierte Überschrift (z.B. `5.3.1`) mit Zusammenfassung | **Ja** — als Heading-Style | Farbig (MM-Schema) |
| Auto-nummerierte Überschrift (einmalig, nach nummerierten Kapitel) | **Ja** — als Heading-Style | Farbig (MM-Schema) |
| Wiederkehrende unnummerierte Überschrift (z.B. "Was ist das?") | **Nein** | Normal + Fett |
| Überschrift ohne Zusammenfassung (ausgeblendet) | **Nein** | Normal + Fett (ausgeblendet) |
| In Dokumenten ohne Nummerierung: H1/H2 | **Ja** | Heading-Style |

**MM-Skript-Farbschema (Heading-Ebenen):**
| Ebene | Farbe | RGB |
|---|---|---|
| H1 | Dunkelgold | 196, 154, 0 |
| H2–H3 | Hellgold | 255, 202, 8 |
| H4–H7 | Dunkelamber | 130, 102, 0 |
| H8–H9 | Fast-Schwarz | 39, 39, 39 |

### US-6.6 — Auto-Nummerierung einmaliger Überschriften
**Als** Nutzer  
**möchte ich** dass thematisch eigenständige Unterabschnitte ohne eigene Nummerierung automatisch nummeriert werden,  
**damit** sie in der Gliederung auffindbar sind.

**Akzeptanzkriterien:**
- Kontext: Dokument hat nummerierte Kapitel (z.B. 5.3)
- Eine unnummerierte Überschrift, die im gesamten Dokument **nur einmal** vorkommt, nach einem nummerierten Kapitel → erhält automatisch die nächste Nummer (z.B. `5.3.1 Agilität`, `5.3.2 Analytisches Denken`)
- Eine Überschrift, die **mehrfach** vorkommt (z.B. "Was ist das?", "Off the job", "On the job") → wird NICHT nummeriert, erscheint als Normal+Fett, NICHT in Gliederung
- Die Häufigkeit wird vor dem Durchlauf über das gesamte Dokument gezählt

### US-6.7 — Ebenenberechnung aus Nummerierung
**Als** Nutzer  
**möchte ich** dass Überschriften-Ebenen automatisch aus der Nummerierung berechnet werden,  
**damit** die Hierarchie korrekt im Word-Dokument dargestellt wird.

**Akzeptanzkriterien:**
- Formel: Anzahl Punkte + 1 = Heading-Ebene  
  (`5` → H1, `5.1` → H2, `5.1.1` → H3, `5.1.1.1` → H4)
- Maximale Ebene: H9
- Auto-nummerierte Überschriften erhalten die Ebene ihrer berechneten Nummer

### US-6.8 — Verdichteter Modus für nicht-prüfungsrelevante Zusatzliteratur (`--condensed`)
**Als** Nutzer  
**möchte ich** umfangreiche, nicht prüfungsrelevante Zusatzliteratur (z.B. Sammelbände/Festschriften mit
mehreren eigenständigen Artikeln) nur verdichtet in meine Lernunterlage übernehmen,  
**damit** ich nicht den vollständigen Originaltext von Inhalten mitschleppen muss, die für die Prüfung
irrelevant sind.

**Bewusste Ausnahme von US-6.1:** Für diesen Modus gilt die Regel "kein Wort des Originaltexts geht
verloren" **nicht**. Statt des vollständigen Originaltexts wird pro Artikel nur eine kurze
Kernaussagen-Übersicht sowie 2–3 wörtliche Zitate im ausgeblendeten Text abgelegt.

**Akzeptanzkriterien:**
- Artikel-Erkennung erfolgt präfix-unabhängig über die flachste vorkommende Heading-Ebene
  (`split_into_articles`) — funktioniert auch bei römisch nummerierten oder unnummerierten
  Artikelüberschriften, im Gegensatz zum ziffernbasierten `split_into_level1_chapters`
- Pro Artikel: sichtbare Kurzübersicht (3–5 Stichpunkte: Thema, zentrale Erkenntnis/Schlussfolgerung)
- Pro Artikel: 2–3 wörtliche Zitate aus dem Original, ausschließlich als ausgeblendeter Text
  (nicht der vollständige Artikeltext)
- Einbettbar als Unterkapitel in eine bestehende Lernunterlage über den bestehenden
  `--parent-chapter`/`--title-as-parent`-Mechanismus (z.B. neues Unterkapitel 6.9 mit 6.9.1, 6.9.2, …)
- Nicht kombinierbar mit `--questions`, `--chapter`, `--from`/`--to`
- Der bestehende Hauptmodus (`build_interleaved_word_document`, `generate_summary_by_chapter`,
  `split_into_level1_chapters`, `_summarize_single_chapter`) bleibt davon vollständig unberührt

---

## Epic 7 — Tabellenverarbeitung

### US-7.1 — Tabellen korrekt darstellen
**Als** Nutzer  
**möchte ich** dass Tabellen aus dem Originaltext korrekt im Word-Dokument erscheinen,  
**damit** ich tabellarische Daten (z.B. Studienresultate) lesen kann.

**Akzeptanzkriterien:**
- Titelzeilen (erste Zeile hat nur eine nicht-leere Zelle) werden als gemergte Überschrift dargestellt
- Leere Spalten werden entfernt
- Statistische Mehrfachwerte in einer Zelle (`3.44 1.00 .49`) werden auf einzelne Zellen aufgeteilt
- Aufeinanderfolgende gleiche Zellwerte in einer Spalte werden vertikal verbunden

---

## Epic 8 — OCR-Artefakt-Bereinigung

### US-8.1 — OCR-Fehler korrigieren
**Als** Nutzer  
**möchte ich** dass typische OCR-Artefakte automatisch bereinigt werden,  
**damit** der Text lesbar ist ohne manuelle Nachbearbeitung.

**Akzeptanzkriterien:**
- Y-Präfix-Fehler: `## y **Titel**` → `## **Titel**` (für alle Überschrifts-Ebenen und Aufzählungszeichen)
- HTML-Tags in Überschriften werden entfernt (z.B. `<span id="page-12-0">`)
- Markdown-Links in Überschriften werden zu Klartext (`[Definition](url)` → `Definition`)
- Seitenreferenz-Marker (`[1]` allein in einer Zeile) werden aus ausgeblendetem Text entfernt

### US-8.2 — Dekorative Bilder entfernen
**Als** Nutzer  
**möchte ich** dass kleine Symbole und Logos automatisch herausgefiltert werden,  
**damit** das Dokument nicht mit irrelevanten Grafiken überladen wird.

**Akzeptanzkriterien:**
- Bilder mit einer minimalen Ausdehnung unter 100 Pixel (Breite oder Höhe) werden übersprungen
- Bilder ohne Pixelinfo und Dateigröße < 8 KB werden übersprungen
- Echte inhaltliche Abbildungen (≥ 100 Pixel) bleiben immer erhalten

---

## Epic 9 — CLI & Konfiguration

### US-9.1 — Standardmodus
**Als** Nutzer  
**möchte ich** mit einem einzigen Befehl das vollständige Lernmittel erzeugen,  
**damit** die Nutzung so einfach wie möglich ist.

**Akzeptanzkriterien:**
- `python pipeline.py mein_dokument.pdf` erzeugt alle Ausgaben (Übersetzung + Zusammenfassung + Lernmittel-DOCX)
- Ausgabeordner wird automatisch angelegt: `workspace/output/{pdf-name}/`

### US-9.2 — Einbetten in übergeordnetes Dokument
**Als** Nutzer  
**möchte ich** ein verarbeitetes Kapitel in ein bestehendes übergeordnetes Dokument einbetten,  
**damit** ich Teilkapitel aus verschiedenen PDFs in einer Gesamtunterlage zusammenführen kann.

**Akzeptanzkriterien:**
- Aufruf: `python pipeline.py dok.pdf --parent-chapter "4.2.1.2"`
- Alle Kapitelnummern werden mit dem Präfix versehen (`1 Einleitung` → `4.2.1.2.1 Einleitung`)
- Alle Heading-Ebenen werden entsprechend verschoben
- Ausgabedatei: `{name}_Einbetten_{parent_chapter}.docx`

### US-9.3 — Konfigurationsoptionen im Überblick
| Option | Beschreibung |
|---|---|
| `--no-summary` | Nur Übersetzungs-DOCX erzeugen; Zusammenfassung und Lernmittel überspringen |
| `--chapter STR` | Nur dieses Kapitel extrahieren und verarbeiten (inkl. Unterkapitel) |
| `--parent-chapter STR` | In übergeordnetes Kapitel einbetten (Nummern-Präfix, Ebenen verschieben) |
| `--parent-level INT` | Heading-Ebene des Elternkapitels (wird sonst automatisch berechnet) |
| `--questions FILE` | QA-Lernfragen einbinden |
| `--force` | Cache ignorieren, alle Schritte neu berechnen |
| `--include-references` | Referenzen/Literaturverzeichnis nicht überspringen |
| `--condensed` | Verdichteter Modus für nicht-prüfungsrelevante Zusatzliteratur (siehe US-6.8): pro Artikel nur Kurzübersicht + 2–3 Zitate statt vollständigem Original |

---

## Nicht-funktionale Anforderungen

| Anforderung | Beschreibung |
|---|---|
| **Vollständigkeit** | Kein Inhalt des Originaltexts darf verloren gehen — weder durch Löschen noch durch Ersetzen |
| **Wiederaufnahme** | Abgebrochene Läufe können fortgesetzt werden; bereits gecachte Kapitel werden nicht neu berechnet |
| **API-Resilienz** | Automatische Wiederholung bei API-Fehlern (503, 429) mit exponentiellem Backoff |
| **Nachvollziehbarkeit** | Alle Zwischenergebnisse (OCR-MD, Übersetzung, Zusammenfassung, QA) werden als Dateien gespeichert |

---

## Glossar

| Begriff | Bedeutung |
|---|---|
| Lernmittel | Das finale Word-Dokument mit interleaved Zusammenfassung + ausgeblendetem Originaltext |
| Interleaved | Abwechselnd: Zusammenfassung sichtbar, Originaltext ausgeblendet, Abschnitt für Abschnitt |
| Auto-Nummerierung | Automatische Vergabe von Kapitelnummern für einmalig vorkommende, unnummerierte Überschriften |
| Elternkapitel | Nummeriertes Kapitel ohne eigenen Textinhalt, dessen Inhalt ausschließlich in Unterkapiteln steht |
| Referenz-Modus | Word-Ansicht mit eingeblendetem "ausgeblendetem Text" — zeigt den vollständigen Originaltext |
