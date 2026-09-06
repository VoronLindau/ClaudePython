#!/usr/bin/env python3
"""
docx_chapter_compare.py
Version 2.6 / 2026-09-05 / Grund: Bugfix - Bilder wurden VOR der
    Kapitel-Klassifizierung des jeweiligen Absatzes angehaengt. Enthielt ein
    Absatz gleichzeitig eine neue (formatvorlagen-numerierte) Ueberschrift
    UND ein Bild, landete das Bild faelschlich in einem soeben dafuer
    erzeugten Fallback-Kapitel ("_N") statt im tatsaechlichen Kapitel - und
    Fallback-Kapitel werden ans Ende sortiert, wodurch das Bild optisch an
    der falschen Stelle (ganz am Schluss) auftauchte. Bilder werden jetzt
    NACH der Klassifizierung angehaengt, landen also im jeweils gerade
    bestimmten/erzeugten Kapitel.

Vergleicht zwei Word-Dokumente (.docx) auf Basis von Kapitelnummern als
Fixpunkten und erzeugt einen eigenstaendigen HTML-Report:
- linke Spalte = Dokument A, rechte Spalte = Dokument B
- jede Zeile = ein Kapitel (gematcht ueber Kapitelnummer)
- farbige Verbindungslinie in der Mittelspalte zeigt die Beziehung:
    gruen   = unveraendert
    orange  = geaendert (inkl. Wort-Diff-Highlighting im Text)
    blau    = neu (nur in Dokument B)
    rot     = geloescht (nur in Dokument A)
- Statistik-Leiste oben mit Anzahl je Kategorie
- "Nur Unterschiede anzeigen"-Filter (Client-seitig, kein Server noetig)

Nutzung (CLI):
    python docx_chapter_compare.py alt.docx neu.docx -o report.html
    (Report danach einfach per Doppelklick im Browser oeffnen)

Nutzung (GUI):
    python docx_chapter_compare_gui.py
"""

import argparse
import base64
import difflib
import glob
import hashlib
import html
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Twips
from lxml import etree

SCRIPT_VERSION = "2.6"
REVIEW_SCHEMA_VERSION = "1.0"

REVIEW_STATUS_OPTIONS = [
    ("", "— nicht bewertet —"),
    ("accepted", "Accepted"),
    ("not_accepted", "Not accepted"),
    ("refinement_customer", "Refinement needed from customer"),
    ("internal_clarification", "Internal clarification needed"),
]

# ---------------------------------------------------------------------------
# Kapitel-Erkennung
# ---------------------------------------------------------------------------

# DOORS/Jazz-Word-Exporte kodieren die Kapitel-/Anforderungsnummer meist NICHT
# im Heading-Text, sondern als "NUMMER<TAB>Text" direkt im Fliesstext, z.B.:
#   "2.1.1\tein (Fiedel) gemaess Kapitel 3.1, ..."
# Word-Heading-Formatvorlagen sind dabei oft gar nicht gesetzt oder tragen
# selbst keine Nummer. Manche Exporte sind zusaetzlich stark fragmentiert:
# eine Nummer wie "2.1.4" kann ueber mehrere Absaetze verteilt sein
# ("2", "2", ".1.4", ...). Die Erkennung unten deckt beide Faelle ab:
#  1) sauberer Fall: NUMMER<TAB>Text in einem Absatz
#  2) fragmentierter Fall: einzelne Zahl-Fragmente werden in einer Warteschlange
#     gesammelt und durch nachfolgende ".N"-Fortsetzungen oder durch normalen
#     Text vervollstaendigt

ANCHOR_TAB_RE = re.compile(r"^(\d+(?:\.\d+)*[a-zA-Z]?)\s*\t\s*(.*)$")
BARE_NUMBER_RE = re.compile(r"^(\d+[a-zA-Z]?)$")
DOT_CONTINUATION_RE = re.compile(r"^\.(\d+(?:\.\d+)*[a-zA-Z]?)\s*\t?\s*(.*)$")

# Bildformate, die ein Browser direkt per <img>/data-URI darstellen kann.
# Word bettet eingefuegte Grafiken (v.a. per Copy&Paste aus anderen Office-Apps)
# haeufig als EMF/WMF (Vektor-Metafile) ein - das kann kein Browser rendern.
# Solche Bilder werden trotzdem auf Aenderungen geprueft (Hash-Vergleich),
# im Report aber nur als Hinweis statt als Vorschau angezeigt.
WEB_SAFE_IMAGE_TYPES = {"image/png", "image/jpeg", "image/jpg", "image/gif", "image/bmp", "image/webp"}

# ---------------------------------------------------------------------------
# Formatvorlagen-verknuepfte automatische Nummerierung ("Ueberschrift 1/2/...")
# ---------------------------------------------------------------------------
#
# Viele Word-Vorlagen nummerieren Kapitelueberschriften NICHT als Text,
# sondern ueber eine automatische, mit der Formatvorlage verknuepfte Liste
# (word/numbering.xml: <w:abstractNum><w:lvl><w:pStyle val="Heading1"/>...).
# Die Zahl ("2", "2.1", ...) existiert dann NUR als Rendering-Ergebnis und
# steht nirgends im gespeicherten Absatztext - sie muss durch Nachbilden der
# Word-Zaehllogik rekonstruiert werden (pro Ebene hochzaehlen, tiefere Ebenen
# zuruecksetzen, wenn eine flachere Ebene wieder auftaucht).

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _load_style_numbering(docx_path):
    """Liest word/numbering.xml (falls vorhanden) und liefert
    {style_id: {"ilvl": int, "start": int}} fuer alle Ueberschriften-Ebenen,
    deren automatische Nummerierung ueber 'Formatvorlage verknuepfen'
    (pStyle innerhalb <w:lvl>) funktioniert. Nur numFmt='decimal' wird
    unterstuetzt (roemisch/alphabetisch etc. werden ausgelassen). Liefert
    {} wenn keine Numerierung vorhanden ist oder die Datei nicht lesbar ist."""
    try:
        with zipfile.ZipFile(docx_path) as z:
            if "word/numbering.xml" not in z.namelist():
                return {}
            xml_bytes = z.read("word/numbering.xml")
    except Exception:
        return {}

    try:
        root = etree.fromstring(xml_bytes)
    except Exception:
        return {}

    ns = {"w": W_NS}
    style_map = {}
    for lvl in root.findall(".//w:abstractNum/w:lvl", ns):
        pstyle_el = lvl.find("w:pStyle", ns)
        if pstyle_el is None:
            continue
        style_id = pstyle_el.get(f"{{{W_NS}}}val")
        if not style_id:
            continue
        numfmt_el = lvl.find("w:numFmt", ns)
        numfmt = numfmt_el.get(f"{{{W_NS}}}val") if numfmt_el is not None else "decimal"
        if numfmt != "decimal":
            continue
        ilvl = int(lvl.get(f"{{{W_NS}}}ilvl", "0"))
        start_el = lvl.find("w:start", ns)
        start_val = int(start_el.get(f"{{{W_NS}}}val", "1")) if start_el is not None else 1
        style_map[style_id] = {"ilvl": ilvl, "start": start_val}
    return style_map


class _HeadingNumberer:
    """Bildet Words Zaehllogik fuer formatvorlagen-verknuepfte Nummerierung
    nach: pro Ebene hochzaehlen, tiefere Ebenen zuruecksetzen sobald eine
    flachere Ebene erneut auftritt. Liefert None fuer Absaetze, deren Style
    nicht Teil der Nummerierung ist."""

    def __init__(self, style_numbering):
        self.style_numbering = style_numbering
        self.counters = {}  # ilvl -> aktueller Zaehlerstand

    def number_for_style(self, style_id):
        info = self.style_numbering.get(style_id)
        if info is None:
            return None
        ilvl = info["ilvl"]
        if ilvl in self.counters:
            self.counters[ilvl] += 1
        else:
            self.counters[ilvl] = info["start"]
        for deeper in [l for l in self.counters if l > ilvl]:
            del self.counters[deeper]
        parts = [str(self.counters.get(l, 1)) for l in range(ilvl + 1)]
        return ".".join(parts)


def _style_id(paragraph):
    try:
        return paragraph.style.style_id
    except Exception:
        return None


def _extract_images(paragraph):
    """Liefert eingebettete Bilder eines Absatzes als Liste von
    {hash, blob, content_type}. Erkennt sowohl moderne DrawingML-Bilder
    (<w:drawing>//<a:blip>) als auch aeltere VML-Bilder (<w:pict>//<v:imagedata>)."""
    images = []
    part = paragraph.part
    p_xml = paragraph._p

    blips = p_xml.findall(".//" + qn("a:blip"))
    for blip in blips:
        rId = blip.get(qn("r:embed")) or blip.get(qn("r:link"))
        if not rId or rId not in part.rels:
            continue
        try:
            image_part = part.rels[rId].target_part
            blob = image_part.blob
            content_type = image_part.content_type
        except Exception:
            continue
        images.append({"hash": hashlib.sha1(blob).hexdigest(), "blob": blob, "content_type": content_type})

    # Aeltere VML-Bilder (v:imagedata) - kommt in aelteren/kompatiblen Exporten vor.
    # Der VML-Namespace ist in python-docx' Standard-nsmap nicht enthalten,
    # daher hier direkt als Clark-Notation-URI.
    VML_NS = "urn:schemas-microsoft-com:vml"
    imagedatas = p_xml.findall(f".//{{{VML_NS}}}imagedata")
    for imgdata in imagedatas:
        rId = imgdata.get(qn("r:id"))
        if not rId or rId not in part.rels:
            continue
        try:
            image_part = part.rels[rId].target_part
            blob = image_part.blob
            content_type = image_part.content_type
        except Exception:
            continue
        h = hashlib.sha1(blob).hexdigest()
        if not any(im["hash"] == h for im in images):
            images.append({"hash": h, "blob": blob, "content_type": content_type})

    return images


def extract_chapters(docx_path):
    """Liest ein docx und liefert eine Liste von Kapiteln/Anforderungen in
    Dokumentreihenfolge: {number, title, paragraphs: [...], text: str, images: [...]}.

    Kapitel ohne erkennbare Nummer bekommen einen Fallback-Key ("_1", "_2", ...);
    diese werden spaeter ueber Text-Aehnlichkeit nachtraeglich gematcht (siehe
    build_comparison), da echte Nummern zwischen Exportversionen fehlen koennen.
    """
    doc = Document(docx_path)
    style_numbering = _load_style_numbering(docx_path)
    numberer = _HeadingNumberer(style_numbering)

    chapters = []
    current = None
    pending_numbers = []  # FIFO-Warteschlange fuer fragmentierte Zahl-Reste
    fallback_counter = 0
    current_para_idx = -1

    def new_chapter(number):
        nonlocal current, fallback_counter
        if number is None:
            fallback_counter += 1
            number = f"_{fallback_counter}"
        # first_para_index merkt sich, an welchem Word-Absatz (0-basiert)
        # dieses Kapitel beginnt - wird fuer die exakte Seiten-Abfrage per
        # MS-Word-COM-Automation gebraucht (siehe attach_pages).
        current = {"number": number, "title": "", "paragraphs": [], "images": [],
                   "first_para_index": current_para_idx}
        chapters.append(current)

    def append_text(text):
        nonlocal current
        if current is None:
            new_chapter(None)
        if text:
            current["paragraphs"].append(text)

    def append_images(imgs):
        nonlocal current
        if not imgs:
            return
        if current is None:
            new_chapter(None)
        current["images"].extend(imgs)

    for para_idx, para in enumerate(doc.paragraphs):
        current_para_idx = para_idx
        text = para.text.strip()
        imgs = _extract_images(para)

        # WICHTIG: Bilder werden ERST NACH der Kapitel-Klassifizierung
        # angehaengt (ganz am Ende dieser Iteration), nicht vorher. Ein
        # Absatz kann gleichzeitig eine neue Kapitel-Ueberschrift UND ein
        # Bild enthalten (z.B. eine formatvorlagen-numerierte Ueberschrift
        # mit eingebettetem Logo/Bild) - wuerde das Bild VOR der Erkennung
        # des neuen Kapitels angehaengt, landet es faelschlich in einem
        # gerade erst erzeugten Fallback-Kapitel (das dann als "_N" ans Ende
        # sortiert wird) statt im tatsaechlichen, gerade begonnenen Kapitel.

        # Formatvorlagen-verknuepfte automatische Nummerierung hat Vorrang:
        # wenn dieser Absatz-Style Teil einer numerierten Ueberschriften-Kette
        # ist (z.B. "Heading1"/"Heading2"), ist die Zahl im Text selbst NICHT
        # vorhanden und muss hier rekonstruiert werden.
        auto_number = numberer.number_for_style(_style_id(para))
        if auto_number is not None:
            pending_numbers.clear()
            new_chapter(auto_number)
            if text:
                append_text(text)
        elif not text:
            pass  # nur Bild, kein Text -> haengt sich unten ans aktuelle Kapitel
        else:
            m_tab = ANCHOR_TAB_RE.match(text)
            m_dot = DOT_CONTINUATION_RE.match(text)
            m_bare = BARE_NUMBER_RE.match(text)

            if m_tab:
                pending_numbers.clear()
                new_chapter(m_tab.group(1))
                append_text(m_tab.group(2))
            elif m_dot and pending_numbers:
                base = pending_numbers.pop(0)
                new_chapter(f"{base}.{m_dot.group(1)}")
                append_text(m_dot.group(2))
            elif m_bare:
                pending_numbers.append(m_bare.group(1))
            elif pending_numbers:
                number = pending_numbers.pop(0)
                pending_numbers.clear()  # uebrige Duplikate/Reste verwerfen
                new_chapter(number)
                append_text(text)
            else:
                append_text(text)

        if imgs:
            append_images(imgs)

    for ch in chapters:
        ch["text"] = "\n".join(ch["paragraphs"])
        ch["level"] = 1

    return chapters


def index_by_key(chapters):
    """Baut ein Dict number->chapter, dedupliziert Mehrfachnummern robust."""
    idx = {}
    counts = {}
    for ch in chapters:
        num = ch["number"]
        counts[num] = counts.get(num, 0) + 1
        key = num if counts[num] == 1 else f"{num}__{counts[num]}"
        ch["key"] = key
        idx[key] = ch
    return idx


# ---------------------------------------------------------------------------
# Seiten-Erkennung ("Buchform") - OPTIONAL, benoetigt LibreOffice
# ---------------------------------------------------------------------------
#
# Word speichert die tatsaechliche Seitenzahl eines Absatzes NICHT in der
# Datei - sie ergibt sich erst beim Layout/Rendern (abhaengig von Schriftart,
# Raendern, Zoom etc.) und ist damit aus der reinen XML-Struktur nicht
# ableitbar. Als Naeherung wird das Dokument per LibreOffice (soffice
# --headless) zu PDF gerendert und jedes Kapitel per Textabgleich der
# tatsaechlich gerenderten PDF-Seite zugeordnet. Das ist eine Annaeherung:
# Word kann geringfuegig anders umbrechen als LibreOffice (andere
# Schriftmetriken), i.d.R. stimmen die Seitenzahlen aber gut genug ueberein,
# um Kapitel visuell nach "Seite" zu gruppieren.

# Windows-Installationen koennen auf beliebigen Laufwerken liegen (nicht nur
# C:) - z.B. "D:\Program Files\LibreOffice\program\soffice.exe". Kandidaten
# und Glob-Muster werden daher ueber die gaengigen Laufwerksbuchstaben
# generiert, statt nur C: fest anzunehmen.
_WIN_DRIVES = ["C", "D", "E", "F"]

SOFFICE_CANDIDATES = [
    "soffice",
    "libreoffice",
    "/usr/bin/soffice",
    "/opt/libreoffice/program/soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
] + [
    fr"{drive}:\Program Files\LibreOffice\program\soffice.exe" for drive in _WIN_DRIVES
] + [
    fr"{drive}:\Program Files (x86)\LibreOffice\program\soffice.exe" for drive in _WIN_DRIVES
]

# Glob-Muster fuer versionierte/portable Installationen, z.B.
# "D:\Program Files\LibreOffice 7.6\program\soffice.exe" oder Installationen
# unter %LOCALAPPDATA%\Programs.
SOFFICE_GLOB_PATTERNS = [
    fr"{drive}:\Program Files\LibreOffice*\program\soffice.exe" for drive in _WIN_DRIVES
] + [
    fr"{drive}:\Program Files (x86)\LibreOffice*\program\soffice.exe" for drive in _WIN_DRIVES
] + [
    str(Path.home() / "AppData/Local/Programs/LibreOffice*/program/soffice.exe"),
]


def find_soffice(explicit_path=None):
    """Sucht ein lauffaehiges LibreOffice/soffice-Binary. Gibt None zurueck,
    wenn keins gefunden wird (Seiten-Gruppierung wird dann einfach ausgelassen).
    Reihenfolge: expliziter Pfad > Umgebungsvariable SOFFICE_PATH >
    PATH/Standardpfade > Glob-Suche nach versionierten/portablen
    Windows-Installationen."""
    if explicit_path:
        return explicit_path if Path(explicit_path).exists() else None
    env_path = os.environ.get("SOFFICE_PATH")
    if env_path and Path(env_path).exists():
        return env_path
    for candidate in SOFFICE_CANDIDATES:
        if Path(candidate).is_absolute():
            if Path(candidate).exists():
                return candidate
        else:
            found = shutil.which(candidate)
            if found:
                return found
    for pattern in SOFFICE_GLOB_PATTERNS:
        matches = sorted(glob.glob(pattern), reverse=True)  # neueste Version zuerst
        if matches:
            return matches[0]
    return None


def render_page_texts(docx_path, soffice_path, timeout=90):
    """Rendert ein docx per LibreOffice zu PDF und liefert eine Liste mit dem
    (whitespace-normalisierten) Text je Seite. Liefert None bei jedem Fehler
    (kein soffice, Timeout, Konvertierungsfehler, PDF nicht lesbar) - der
    Aufrufer behandelt das als "Seiten-Gruppierung nicht verfuegbar"."""
    try:
        import pdfplumber
    except ImportError:
        return None
    try:
        with tempfile.TemporaryDirectory() as tmp:
            subprocess.run(
                [soffice_path, "--headless", "--convert-to", "pdf", "--outdir", tmp, str(docx_path)],
                check=True, timeout=timeout, capture_output=True,
            )
            pdf_path = Path(tmp) / (Path(docx_path).stem + ".pdf")
            if not pdf_path.exists():
                return None
            with pdfplumber.open(pdf_path) as pdf:
                return [normalize_whitespace(p.extract_text() or "") for p in pdf.pages]
    except Exception:
        return None


def assign_pages_to_chapters(chapters, page_texts, key_len=40):
    """Ordnet jedem Kapitel per Substring-Suche (auf normalisiertem Text) die
    PDF-Seite zu, auf der sein Text-Anfang vorkommt. Sucht positions-bewusst
    vorwaerts (Cursor innerhalb der aktuellen Seite + Seiten-Index), damit
    WIEDERHOLTER/DUPLIZIERTER Text (z.B. zwei inhaltlich identische
    Abschnitte) nicht faelschlich immer der ersten Fundstelle zugeordnet
    wird - jede bereits gefundene Textstelle wird "verbraucht" und kann
    nicht nochmal treffen. Setzt ch['page'] direkt auf jedem Kapitel-Dict."""
    if not page_texts:
        for ch in chapters:
            ch["page"] = None
        return

    page_idx = 0
    cursor = 0  # Zeichen-Position innerhalb der aktuellen Seite, ab der weitergesucht wird
    for ch in chapters:
        key = normalize_whitespace(ch["text"])[:key_len]
        if key:
            while True:
                pos = page_texts[page_idx].find(key, cursor)
                if pos != -1:
                    cursor = pos + len(key)
                    break
                if page_idx + 1 < len(page_texts):
                    page_idx += 1
                    cursor = 0
                else:
                    break  # letzte Seite erreicht, nicht gefunden - Kapitel bleibt hier
        ch["page"] = page_idx + 1


def find_word_com():
    """Prueft, ob MS Word per COM-Automation ansprechbar sein KOENNTE (nur
    Windows, pywin32 installiert). Das ist nur ein Verfuegbarkeits-Check auf
    das Python-Paket - ob tatsaechlich Word installiert ist, zeigt sich erst
    beim eigentlichen Verbindungsversuch in assign_pages_via_word_com."""
    if sys.platform != "win32":
        return False
    try:
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        return False


def assign_pages_via_word_com(chapters, docx_path, timeout=120):
    """Fragt MS Word SELBST (COM-Automation, nur Windows, benoetigt
    'pip install pywin32' und eine lokale Word-Installation) fuer jedes
    Kapitel die tatsaechlich von Word berechnete Seitenzahl ab - keine
    Annaeherung wie beim LibreOffice/PDF-Weg, sondern das Original.

    Setzt ch['page'] direkt auf jedem Kapitel-Dict. Gibt True bei Erfolg
    zurueck, False wenn Word/pywin32 nicht verfuegbar ist oder irgendein
    Fehler auftrat (dann bleibt ch['page'] unveraendert, damit der Aufrufer
    z.B. noch auf LibreOffice ausweichen kann).

    HINWEIS: Dieser Pfad ist speziell fuer Rechner ohne LibreOffice, aber mit
    installiertem MS Office (z.B. Firmenrechner) gedacht. Nutzt die COM-
    Konstante wdActiveEndPageNumber (=3) von Word.Range.Information(...).
    """
    if sys.platform != "win32":
        return False
    try:
        import pywintypes
        import win32com.client
    except ImportError:
        return False

    wanted_indices = sorted({ch["first_para_index"] for ch in chapters if ch.get("first_para_index", -1) >= 0})
    if not wanted_indices:
        return False

    word = None
    doc = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(
            str(Path(docx_path).resolve()), ReadOnly=True, AddToRecentFiles=False, Visible=False,
        )
        WD_ACTIVE_END_PAGE_NUMBER = 3
        page_by_index = {}
        wanted_set = set(wanted_indices)
        max_wanted = wanted_indices[-1]
        for i, para in enumerate(doc.Paragraphs):
            if i > max_wanted:
                break
            if i in wanted_set:
                try:
                    page_by_index[i] = para.Range.Information(WD_ACTIVE_END_PAGE_NUMBER)
                except Exception:
                    pass
        for ch in chapters:
            ch["page"] = page_by_index.get(ch.get("first_para_index", -1))
        return True
    except Exception:
        return False
    finally:
        try:
            if doc is not None:
                doc.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass


def attach_pages(chapters_a, chapters_b, path_a, path_b, soffice_path=None):
    """Versucht, beiden Kapitel-Listen echte Seitenzahlen zuzuordnen.
    Reihenfolge: (1) MS Word per COM (exakt, nur Windows mit installiertem
    Word), (2) LibreOffice-Rendering + Textabgleich (Naeherung, aber
    plattformunabhaengig), (3) keine Seiteninfo. Gibt einen Status-String
    zurueck: 'word_com', 'libreoffice' oder None (nicht verfuegbar)."""
    if find_word_com():
        ok_a = assign_pages_via_word_com(chapters_a, path_a)
        ok_b = assign_pages_via_word_com(chapters_b, path_b) if ok_a else False
        if ok_a and ok_b:
            return "word_com"
        # Bei Teilerfolg lieber sauber zuruecksetzen und den anderen Weg probieren
        for ch in chapters_a:
            ch.pop("page", None)
        for ch in chapters_b:
            ch.pop("page", None)

    soffice_path = find_soffice(soffice_path)
    if soffice_path:
        pages_a = render_page_texts(path_a, soffice_path)
        pages_b = render_page_texts(path_b, soffice_path)
        assign_pages_to_chapters(chapters_a, pages_a)
        assign_pages_to_chapters(chapters_b, pages_b)
        if pages_a and pages_b:
            return "libreoffice"

    for ch in chapters_a:
        ch["page"] = None
    for ch in chapters_b:
        ch["page"] = None
    return None


# ---------------------------------------------------------------------------
# Diff-Logik
# ---------------------------------------------------------------------------

_TOKEN_RE = re.compile(r"\s+|\S+")


def _tokenize(text):
    return _TOKEN_RE.findall(text)


def render_diff_pair(text_a, text_b):
    """Liefert (html_a, html_b) mit <span class="del"> / <span class="ins">
    Markierungen fuer die jeweils andere Seite (Wort-Diff)."""
    tokens_a = _tokenize(text_a)
    tokens_b = _tokenize(text_b)
    sm = difflib.SequenceMatcher(None, tokens_a, tokens_b, autojunk=False)

    left_parts, right_parts = [], []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        seg_a = html.escape("".join(tokens_a[i1:i2]))
        seg_b = html.escape("".join(tokens_b[j1:j2]))
        if tag == "equal":
            left_parts.append(seg_a)
            right_parts.append(seg_b)
        elif tag == "delete":
            left_parts.append(f'<span class="del">{seg_a}</span>')
        elif tag == "insert":
            right_parts.append(f'<span class="ins">{seg_b}</span>')
        elif tag == "replace":
            left_parts.append(f'<span class="del">{seg_a}</span>')
            right_parts.append(f'<span class="ins">{seg_b}</span>')

    return (
        "".join(left_parts).replace("\n", "<br>"),
        "".join(right_parts).replace("\n", "<br>"),
    )


def plain_html(text):
    return html.escape(text).replace("\n", "<br>")


def image_hash_set(chapter):
    return sorted(img["hash"] for img in chapter.get("images", []))


def classify(ch_a, ch_b):
    """Liefert (status, ratio, images_changed). images_changed ist True, wenn
    sich die Menge der eingebetteten Grafiken zwischen beiden Seiten
    unterscheidet (Hash-Vergleich, ordnungsunabhaengig)."""
    images_changed = image_hash_set(ch_a) != image_hash_set(ch_b)
    if ch_a["text"] == ch_b["text"] and ch_a["title"] == ch_b["title"] and not images_changed:
        return "unchanged", 1.0, False
    ratio = difflib.SequenceMatcher(None, ch_a["text"], ch_b["text"], autojunk=False).ratio()
    return "changed", ratio, images_changed


def natural_sort_key(number, order_index):
    """Numerische Kapitelnummern (1, 1.2, 2.10 ...) korrekt sortieren;
    nicht-numerische Fallback-Keys (_preamble, _3 ...) ans Ende, stabil
    nach urspruenglicher Dokumentreihenfolge."""
    if number.startswith("_"):
        return (1, order_index, ())
    parts = number.split(".")
    try:
        parsed = tuple(int(p) for p in parts)
    except ValueError:
        return (1, order_index, ())
    return (0, 0, parsed)


def _similarity(a, b):
    if not a or not b:
        return 0.0
    return difflib.SequenceMatcher(None, a, b, autojunk=False).ratio()


def fallback_match_unnumbered(deleted_only, new_only, threshold=0.6):
    """Versucht, Kapitel ohne (uebereinstimmende) Nummer per Text-/Titel-
    Aehnlichkeit einander zuzuordnen. Das faengt Faelle ab, in denen eine
    Kapitelnummer in einer der beiden Exportversionen komplett fehlt (echter
    Fallback-Key "_..."), der Kapiteltext/-titel aber erkennbar aehnlich ist.

    WICHTIG: Kandidaten mit einer (ggf. falsch rekonstruierten, aber nicht
    leeren) Nummer wie "2" oder "2.1.2a" werden hier NICHT beruecksichtigt -
    nur echte Fallback-Keys ("_1", "_2", ...). Sonst werden bei stark
    fragmentierten Dokumenten faelschlich inhaltlich unpassende Kapitel
    zusammengeklebt, nur weil beide Seiten "uebrig" waren.
    """
    candidates_a = [c for c in deleted_only if c["number"].startswith("_")]
    candidates_b = [c for c in new_only if c["number"].startswith("_")]

    pairs = []
    used_b_keys = set()
    for ca in candidates_a:
        best, best_score = None, 0.0
        for cb in candidates_b:
            if cb["key"] in used_b_keys:
                continue
            score = max(_similarity(ca["title"], cb["title"]), _similarity(ca["text"], cb["text"]))
            if score > best_score:
                best_score, best = score, cb
        if best is not None and best_score >= threshold:
            used_b_keys.add(best["key"])
            pairs.append((ca, best, best_score))
    return pairs


def normalize_whitespace(text):
    return re.sub(r"\s+", " ", text).strip()


def _normalize_chapters_for_comparison(chapters):
    """Erstellt Kopien der Kapitel mit auf Einzelzeilen normalisiertem
    Text/Titel (Zeilenumbrueche/Mehrfach-Leerzeichen -> ein Leerzeichen),
    fuer den optionalen 'Zeilenumbrueche ignorieren'-Modus. Bilder bleiben
    unangetastet (Referenz wird uebernommen)."""
    out = []
    for ch in chapters:
        new_ch = dict(ch)
        new_ch["text"] = normalize_whitespace(ch["text"])
        new_ch["title"] = normalize_whitespace(ch["title"]) if ch["title"] else ch["title"]
        out.append(new_ch)
    return out


def build_comparison(chapters_a, chapters_b, ignore_linebreaks=True):
    """ignore_linebreaks (Standard: True): Zeilen-/Absatzumbrueche werden vor
    dem Vergleich zu einzelnen Leerzeichen normalisiert, sodass reines
    Neu-Umbrechen von Text (z.B. durch Nachbearbeitung oder unterschiedliche
    Absatzstruktur) nicht als inhaltliche Aenderung gewertet wird. Mit
    ignore_linebreaks=False wird strikt inklusive Absatzgrenzen verglichen."""
    if ignore_linebreaks:
        chapters_a = _normalize_chapters_for_comparison(chapters_a)
        chapters_b = _normalize_chapters_for_comparison(chapters_b)

    idx_a = index_by_key(chapters_a)
    idx_b = index_by_key(chapters_b)
    order_a = {ch["key"]: i for i, ch in enumerate(chapters_a)}
    order_b = {ch["key"]: i for i, ch in enumerate(chapters_b)}

    all_keys = list(dict.fromkeys(list(idx_a.keys()) + list(idx_b.keys())))
    rows = []
    deleted_only, new_only = [], []

    for key in all_keys:
        ca = idx_a.get(key)
        cb = idx_b.get(key)

        if ca and cb:
            status, ratio, images_changed = classify(ca, cb)
            if status == "changed":
                html_a, html_b = render_diff_pair(ca["text"], cb["text"])
            else:
                html_a, html_b = plain_html(ca["text"]), plain_html(cb["text"])
            rows.append({
                "key": key, "number": ca["number"], "status": status, "ratio": ratio,
                "title_a": ca["title"], "title_b": cb["title"],
                "html_a": html_a, "html_b": html_b,
                "images_a": ca.get("images", []), "images_b": cb.get("images", []),
                "images_changed": images_changed,
                "page_a": ca.get("page"), "page_b": cb.get("page"),
            })
        elif ca and not cb:
            deleted_only.append(ca)
        else:
            new_only.append(cb)

    # Zweiter Durchgang: verbleibende unmatched Kapitel per Text-/Titel-
    # Aehnlichkeit zusammenfuehren (faengt fehlende/verschobene Nummern ab).
    fallback_pairs = fallback_match_unnumbered(deleted_only, new_only)
    matched_a_keys = {ca["key"] for ca, _, _ in fallback_pairs}
    matched_b_keys = {cb["key"] for _, cb, _ in fallback_pairs}

    for ca, cb, score in fallback_pairs:
        status, ratio, images_changed = classify(ca, cb)
        if status == "changed":
            html_a, html_b = render_diff_pair(ca["text"], cb["text"])
        else:
            html_a, html_b = plain_html(ca["text"]), plain_html(cb["text"])
        number = ca["number"] if not ca["number"].startswith("_") else cb["number"]
        rows.append({
            "key": ca["key"], "number": number, "status": status, "ratio": ratio,
            "title_a": ca["title"], "title_b": cb["title"],
            "html_a": html_a, "html_b": html_b,
            "images_a": ca.get("images", []), "images_b": cb.get("images", []),
            "images_changed": images_changed,
            "page_a": ca.get("page"), "page_b": cb.get("page"),
        })
        order_a[ca["key"]] = order_a.get(ca["key"], order_a.get(cb["key"], 0))

    for ca in deleted_only:
        if ca["key"] in matched_a_keys:
            continue
        rows.append({
            "key": ca["key"], "number": ca["number"], "status": "deleted", "ratio": 0.0,
            "title_a": ca["title"], "title_b": None,
            "html_a": plain_html(ca["text"]), "html_b": None,
            "images_a": ca.get("images", []), "images_b": [],
            "images_changed": bool(ca.get("images")),
            "page_a": ca.get("page"), "page_b": None,
        })

    for cb in new_only:
        if cb["key"] in matched_b_keys:
            continue
        rows.append({
            "key": cb["key"], "number": cb["number"], "status": "new", "ratio": 0.0,
            "title_a": None, "title_b": cb["title"],
            "html_a": None, "html_b": plain_html(cb["text"]),
            "images_a": [], "images_b": cb.get("images", []),
            "images_changed": bool(cb.get("images")),
            "page_a": None, "page_b": cb.get("page"),
        })

    def sort_key(row):
        oi = order_a.get(row["key"])
        if oi is None:
            oi = 10_000_000 + order_b.get(row["key"], 0)
        return natural_sort_key(row["number"], oi)

    rows.sort(key=sort_key)
    return rows


def compute_stats(chapters_a, chapters_b, rows):
    return {
        "total_a": len(chapters_a),
        "total_b": len(chapters_b),
        "unchanged": sum(1 for r in rows if r["status"] == "unchanged"),
        "changed": sum(1 for r in rows if r["status"] == "changed"),
        "new": sum(1 for r in rows if r["status"] == "new"),
        "deleted": sum(1 for r in rows if r["status"] == "deleted"),
    }


# ---------------------------------------------------------------------------
# HTML-Report
# ---------------------------------------------------------------------------

STATUS_LABEL = {
    "unchanged": "Unverändert",
    "changed": "Geändert",
    "new": "Neu",
    "deleted": "Gelöscht",
}

STATUS_ICON = {
    "unchanged": "\u2713",  # check
    "changed": "\u270E",    # pencil
    "new": "\u271A",        # plus
    "deleted": "\u2716",    # cross
}

STATUS_LABEL_SHORT = {
    "unchanged": "Unverändert",
    "changed": "Geändert",
    "new": "Neu",
    "deleted": "Gelöscht",
}

# Farben je Status - abgestimmt auf die Farben im HTML-Report (siehe CSS
# --c-unchanged/--c-changed/--c-new/--c-deleted), damit HTML und DOCX
# optisch konsistent wirken.
STATUS_COLORS_DOCX = {
    "unchanged": {"rgb": "16A34A", "fill": "DCFCE7"},
    "changed": {"rgb": "C2410C", "fill": "FFEDD5"},
    "new": {"rgb": "1D4ED8", "fill": "DBEAFE"},
    "deleted": {"rgb": "B91C1C", "fill": "FEE2E2"},
}


def _plain_preview(html_text, max_len=220):
    """Wandelt den HTML-Diff-Text eines Kapitels in reinen Klartext fuer den
    DOCX-Report um (Tags entfernen, kuerzen)."""
    if not html_text:
        return ""
    plain = re.sub(r"<[^>]+>", "", html_text)
    plain = html.unescape(plain).strip()
    plain = re.sub(r"\s+", " ", plain)
    if len(plain) > max_len:
        plain = plain[:max_len].rsplit(" ", 1)[0] + " …"
    return plain


def _set_cell_shading(cell, fill_hex):
    """python-docx hat keine eingebaute API fuer Zellenschattierung - wird
    daher direkt als <w:shd>-Element in die Zellen-XML eingehaengt."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    tcPr.append(shd)


def _set_col_widths(table, widths_dxa):
    """Setzt Spaltenbreiten sowohl auf der Tabelle als auch auf jeder Zelle -
    beides ist noetig, sonst werden die Breiten in manchen Word-Versionen
    (und Google Docs) ignoriert."""
    table.autofit = False
    for row in table.rows:
        for cell, width in zip(row.cells, widths_dxa):
            cell.width = Twips(width)
    for col, width in zip(table.columns, widths_dxa):
        col.width = Twips(width)


def generate_docx_report(rows, stats, name_a, name_b, output_path, meta_a=None, meta_b=None, reviews=None):
    """Erzeugt einen kompakten Word-Report fuer die schnelle Orientierung
    (z.B. Weitergabe im Unternehmen an Leute ohne Zugriff auf den
    HTML-Report): Zusammenfassung oben, danach eine Tabelle mit allen
    Kapiteln, die sich geaendert haben/neu/geloescht sind (mit kurzem
    Text-Auszug beider Seiten), am Ende eine kompakte Gesamtliste aller
    Kapitel zur Nachvollziehbarkeit.

    reviews: optionales Dict {key: {"status":..., "comment":...}} aus einer
    zuvor exportierten Review-JSON-Datei - wird dann als zusaetzliche Spalte
    mit aufgenommen."""
    doc = Document()

    doc.add_heading("Kapitelvergleich", level=0)
    meta_a = meta_a or {"name": name_a, "modified": ""}
    meta_b = meta_b or {"name": name_b, "modified": ""}
    p = doc.add_paragraph()
    p.add_run(f"Dokument A: {meta_a['name']}").bold = True
    if meta_a.get("modified"):
        p.add_run(f"  (Stand: {meta_a['modified']})")
    p2 = doc.add_paragraph()
    p2.add_run(f"Dokument B: {meta_b['name']}").bold = True
    if meta_b.get("modified"):
        p2.add_run(f"  (Stand: {meta_b['modified']})")
    doc.add_paragraph(
        f"Erzeugt am {datetime.now().strftime('%d.%m.%Y %H:%M')} · "
        f"docx_chapter_compare.py Version {SCRIPT_VERSION}"
    ).runs[0].font.size = Pt(9)

    # --- Zusammenfassung -----------------------------------------------
    doc.add_heading("Zusammenfassung", level=1)
    summary = doc.add_table(rows=2, cols=6)
    summary.style = "Light Grid Accent 1"
    labels = ["Kapitel A", "Kapitel B", "Unverändert", "Geändert", "Neu", "Gelöscht"]
    values = [stats["total_a"], stats["total_b"], stats["unchanged"],
              stats["changed"], stats["new"], stats["deleted"]]
    for col, label in enumerate(labels):
        summary.cell(0, col).text = label
        summary.cell(0, col).paragraphs[0].runs[0].bold = True
    for col, value in enumerate(values):
        cell = summary.cell(1, col)
        cell.text = str(value)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        status_key = ["total_a", "total_b", "unchanged", "changed", "new", "deleted"][col]
        if status_key in STATUS_COLORS_DOCX:
            colors = STATUS_COLORS_DOCX[status_key]
            _set_cell_shading(cell, colors["fill"])
            run = cell.paragraphs[0].runs[0]
            run.font.color.rgb = RGBColor.from_string(colors["rgb"])
            run.bold = True

    changed_rows = [r for r in rows if r["status"] != "unchanged"]

    # --- Aenderungen im Detail ------------------------------------------
    doc.add_heading(f"Änderungen im Detail ({len(changed_rows)})", level=1)
    if not changed_rows:
        doc.add_paragraph("Keine Änderungen gefunden - beide Dokumente sind inhaltlich identisch.")
    else:
        has_reviews = bool(reviews)
        col_count = 5 if has_reviews else 4
        table = doc.add_table(rows=1, cols=col_count)
        table.style = "Light Grid Accent 1"
        headers = ["Nr.", "Status", f"Dokument A", f"Dokument B"]
        if has_reviews:
            headers.append("Review")
        for col, h in enumerate(headers):
            cell = table.cell(0, col)
            cell.text = h
            cell.paragraphs[0].runs[0].bold = True
        widths = [900, 1300, 3400, 3400] + ([2300] if has_reviews else [])
        _set_col_widths(table, widths)

        for r in changed_rows:
            row_cells = table.add_row().cells
            row_cells[0].text = r["number"]
            status_cell = row_cells[1]
            status_cell.text = STATUS_LABEL_SHORT[r["status"]]
            colors = STATUS_COLORS_DOCX[r["status"]]
            _set_cell_shading(status_cell, colors["fill"])
            status_run = status_cell.paragraphs[0].runs[0]
            status_run.font.color.rgb = RGBColor.from_string(colors["rgb"])
            status_run.bold = True
            if r["status"] == "changed":
                status_cell.add_paragraph(f"{int(r['ratio'] * 100)}% gleich").runs[0].font.size = Pt(8)

            row_cells[2].text = _plain_preview(r.get("html_a")) or "—"
            row_cells[3].text = _plain_preview(r.get("html_b")) or "—"

            if has_reviews:
                entry = reviews.get(r["key"], {})
                status_val = entry.get("status", "")
                comment_val = entry.get("comment", "")
                label = dict(REVIEW_STATUS_OPTIONS).get(status_val, "— nicht bewertet —")
                review_cell = row_cells[4]
                review_cell.text = label
                if comment_val:
                    review_cell.add_paragraph(comment_val).runs[0].font.size = Pt(8)

            for cell in row_cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        if run.font.size is None:
                            run.font.size = Pt(9)

    # --- Vollstaendige Kapitelliste (kompakt, zur Nachvollziehbarkeit) --
    doc.add_heading(f"Vollständige Kapitelliste ({len(rows)})", level=1)
    overview = doc.add_table(rows=1, cols=3)
    overview.style = "Light List Accent 1"
    for col, h in enumerate(["Nr.", "Status", "Auszug"]):
        cell = overview.cell(0, col)
        cell.text = h
        cell.paragraphs[0].runs[0].bold = True
    _set_col_widths(overview, [900, 1600, 6500])
    for r in rows:
        row_cells = overview.add_row().cells
        row_cells[0].text = r["number"]
        status_cell = row_cells[1]
        status_cell.text = STATUS_LABEL_SHORT[r["status"]]
        run = status_cell.paragraphs[0].runs[0]
        run.font.color.rgb = RGBColor.from_string(STATUS_COLORS_DOCX[r["status"]]["rgb"])
        preview = _plain_preview(r.get("html_b") or r.get("html_a"), max_len=110)
        row_cells[2].text = preview
        for cell in row_cells:
            for para in cell.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(9)

    doc.save(output_path)


def _safe_id(key):
    """Wandelt einen Kapitel-Key in eine gueltige HTML-id/Fragment-Zeichenkette um."""
    return re.sub(r"[^a-zA-Z0-9_-]", "_", key)


def doc_metadata(path):
    """Name + letzter-Aenderungszeitpunkt einer Datei, fuers Review-JSON."""
    path = Path(path)
    try:
        modified = datetime.fromtimestamp(path.stat().st_mtime).isoformat(timespec="seconds")
    except OSError:
        modified = ""
    return {"name": path.name, "modified": modified}


def _preview_title(title, body_html, max_len=70):
    """Wenn kein separater Titel vorhanden ist (typisch bei DOORS-Requirement-
    Zeilen ohne Ueberschrift), aus dem Fliesstext eine kurze Vorschau ableiten."""
    if title:
        return title
    if not body_html:
        return ""
    plain = re.sub(r"<[^>]+>", "", body_html).replace("&amp;", "&")
    plain = html.unescape(plain)
    plain = plain.strip()
    if len(plain) > max_len:
        plain = plain[:max_len].rsplit(" ", 1)[0] + " …"
    return plain


def _images_html(images):
    """Rendert eine Thumbnail-Galerie fuer eine Liste von Bildern. Web-taugliche
    Formate (PNG/JPEG/GIF/BMP/WebP) werden als data-URI eingebettet; andere
    (v.a. EMF/WMF-Vektorgrafiken, die Browser nicht darstellen koennen)
    bekommen stattdessen einen Hinweistext."""
    if not images:
        return ""
    parts = ['<div class="img-gallery">']
    for img in images:
        ct = img.get("content_type", "")
        if ct in WEB_SAFE_IMAGE_TYPES:
            b64 = base64.b64encode(img["blob"]).decode("ascii")
            parts.append(f'<img class="img-thumb" src="data:{html.escape(ct)};base64,{b64}" alt="Grafik">')
        else:
            size_kb = max(1, len(img["blob"]) // 1024)
            label = ct or "unbekanntes Format"
            parts.append(
                f'<div class="img-placeholder" title="Keine Browser-Vorschau moeglich">'
                f'🖼 {html.escape(label)} ({size_kb} KB)</div>'
            )
    parts.append("</div>")
    return "".join(parts)


def _page_classes(page, first, last):
    """CSS-Klassen fuer die Seiten-Box: durchgaengiger Rahmen ueber mehrere
    Zeilen hinweg, wenn sie zur selben (gerenderten) Seite gehoeren. first/last
    bestimmen, an welcher Kante die Box eine sichtbare obere/untere Umrandung
    + Eckenrundung bekommt (mittendrin bleibt die Kante offen, damit die Box
    ueber mehrere Zeilen hinweg nahtlos wirkt)."""
    if page is None:
        return ""
    band = "pageband-even" if page % 2 == 0 else "pageband-odd"
    if first and last:
        grp = "pg-both"
    elif first:
        grp = "pg-first"
    elif last:
        grp = "pg-last"
    else:
        grp = "pg-mid"
    return f" {band} pg-grouped {grp}"


def _cell(number, title, body_html, images, side, status, page=None, first_in_group=False, last_in_group=False):
    page_class = _page_classes(page, first_in_group, last_in_group)
    page_tag = f'<div class="page-tag">📄 Seite {page}</div>' if (page is not None and first_in_group) else ""

    if body_html is None:
        inner = (
            f'<div class="cell cell-empty cell-{side}{page_class}" data-status="{status}">'
            f'<span class="empty-hint">— kein Kapitel {html.escape(number)} in diesem Dokument —</span>'
            f"</div>"
        )
    else:
        display_title = _preview_title(title, body_html)
        header = f'<span class="chnum">{html.escape(number)}</span> <span class="chtitle">{html.escape(display_title)}</span>'
        inner = (
            f'<details class="cell cell-{side}{page_class}" data-status="{status}" {"open" if status != "unchanged" else ""}>'
            f"<summary>{header}</summary>"
            f'<div class="chbody">{body_html}</div>'
            f"{_images_html(images)}"
            f"</details>"
        )
    # WICHTIG: pro Grid-Spalte muss genau EIN direktes Kind-Element ans
    # .grid-row (display:grid) uebergeben werden - die Seiten-Marke wird
    # daher zusammen mit der Zelle in einen gemeinsamen Slot-Wrapper gepackt,
    # statt als zusaetzliches Geschwister-Element, sonst verschiebt sich die
    # 3-Spalten-Zuordnung (links/Connector/rechts) bei jeder Seiten-Marke.
    return f'<div class="cell-slot">{page_tag}{inner}</div>'


def _compute_group_flags(page_seq):
    """Liefert (first_flags, last_flags) - je True, wenn die Zeile an dieser
    Stelle die erste/letzte einer zusammenhaengenden Seiten-Gruppe ist.
    Seiten=None (keine Seiteninfo bzw. "kein Kapitel X in diesem Dokument")
    werden fuer die Gruppierung uebersprungen/ignoriert, statt eine neue
    Gruppe zu erzwingen - sonst reisst eine Luecke (Kapitel nur auf einer
    Seite vorhanden) die Box unnoetig auseinander."""
    n = len(page_seq)
    first_flags, last_flags = [False] * n, [False] * n

    next_real = [None] * n
    upcoming = None
    for i in range(n - 1, -1, -1):
        next_real[i] = upcoming
        if page_seq[i] is not None:
            upcoming = page_seq[i]

    last_real = None
    for i, p in enumerate(page_seq):
        if p is None:
            continue
        first_flags[i] = (p != last_real)
        last_flags[i] = (p != next_real[i])
        last_real = p
    return first_flags, last_flags


def render_html(rows, stats, name_a, name_b, ignore_linebreaks=True, meta_a=None, meta_b=None,
                 pages_method=None, diagnostics=None):
    meta_a = meta_a or {"name": name_a, "modified": ""}
    meta_b = meta_b or {"name": name_b, "modified": ""}

    review_options_html = "".join(
        f'<option value="{v}">{html.escape(label)}</option>' for v, label in REVIEW_STATUS_OPTIONS
    )

    pages_a_seq = [r.get("page_a") for r in rows]
    pages_b_seq = [r.get("page_b") for r in rows]
    first_a, last_a = _compute_group_flags(pages_a_seq)
    first_b, last_b = _compute_group_flags(pages_b_seq)

    row_html = []
    for idx, r in enumerate(rows):
        status = r["status"]
        page_a, page_b = r.get("page_a"), r.get("page_b")
        left = _cell(r["number"], r["title_a"], r["html_a"], r.get("images_a", []), "left", status,
                     page=page_a, first_in_group=first_a[idx], last_in_group=last_a[idx])
        right = _cell(r["number"], r["title_b"], r["html_b"], r.get("images_b", []), "right", status,
                       page=page_b, first_in_group=first_b[idx], last_in_group=last_b[idx])
        pct = f'{int(r["ratio"] * 100)}%' if status == "changed" else ""
        img_badge = '<span class="conn-img-badge" title="Grafik geändert">🖼</span>' if r.get("images_changed") else ""
        spine_class = " has-page-spine" if (page_a is not None or page_b is not None) else ""
        connector = (
            f'<div class="connector conn-{status}{spine_class}">'
            f'<span class="conn-icon">{STATUS_ICON[status]}</span>'
            f"{img_badge}"
            f'<span class="conn-pct">{pct}</span>'
            f"</div>"
        )
        safe_id = _safe_id(r["key"])
        review_box = f"""
        <div class="review-box" data-review-key="{html.escape(r['key'], quote=True)}">
          <label for="rs-{safe_id}">Review:</label>
          <select id="rs-{safe_id}" class="review-select" onchange="onReviewChange('{safe_id}')">{review_options_html}</select>
          <input type="text" id="rc-{safe_id}" class="review-comment" placeholder="Kommentar…" oninput="onReviewChange('{safe_id}')">
        </div>"""
        # Abstand zur vorherigen Zeile nur einfuegen, wenn mindestens eine
        # Seite hier tatsaechlich eine neue Seiten-Gruppe beginnt - so
        # verschmelzen die Boxen ueber mehrere Zeilen optisch nahtlos.
        continues_group = idx > 0 and not first_a[idx] and not first_b[idx]
        wrapper_class = "row-wrapper continues-group" if continues_group else "row-wrapper"
        row_html.append(
            f'<div class="{wrapper_class}" data-status="{status}">'
            f'<div class="grid-row row-{status}" data-status="{status}">{left}{connector}{right}</div>'
            f"{review_box}"
            f"</div>"
        )

    stats_html = f"""
      <div class="stat stat-total"><div class="stat-num">{stats['total_a']}</div><div class="stat-label">Kapitel A</div></div>
      <div class="stat stat-total"><div class="stat-num">{stats['total_b']}</div><div class="stat-label">Kapitel B</div></div>
      <div class="stat stat-unchanged"><div class="stat-num">{stats['unchanged']}</div><div class="stat-label">Unverändert</div></div>
      <div class="stat stat-changed"><div class="stat-num">{stats['changed']}</div><div class="stat-label">Geändert</div></div>
      <div class="stat stat-new"><div class="stat-num">{stats['new']}</div><div class="stat-label">Neu</div></div>
      <div class="stat stat-deleted"><div class="stat-num">{stats['deleted']}</div><div class="stat-label">Gelöscht</div></div>
    """

    linebreak_note = (
        "Zeilenumbrüche werden beim Vergleich ignoriert (nur Wortinhalt zählt)."
        if ignore_linebreaks else
        "Zeilenumbrüche werden strikt mitverglichen (inkl. Absatzgrenzen)."
    )
    if pages_method == "word_com":
        page_note = "📄 Kapitel sind nach der von MS Word berechneten Seite gruppiert (exakt, per COM-Automation)."
    elif pages_method == "libreoffice":
        page_note = "📄 Kapitel sind nach gerenderter Seite gruppiert (via LibreOffice, Näherung – Word kann geringfügig anders umbrechen)."
    elif pages_method == "unavailable":
        page_note = "📄 Seiten-Gruppierung nicht verfügbar (weder MS Word/COM noch LibreOffice/soffice gefunden, oder Rendern fehlgeschlagen)."
    else:
        page_note = ""

    # Preflight-Diagnose-Box: immer im Report gespeichert (nicht nur bei
    # Fehlern), damit man nicht extra --diagnose-pages auf der Kommandozeile
    # laufen lassen muss, um zu sehen, woran eine nicht verfuegbare
    # Seiten-Gruppierung liegt. Standardmaessig eingeklappt, wenn alles
    # funktioniert hat, automatisch aufgeklappt, wenn etwas fehlgeschlagen ist.
    diagnostics_html = ""
    if diagnostics:
        diag_items = "".join(f"<li>{html.escape(line)}</li>" for line in diagnostics)
        diag_open = "open" if pages_method == "unavailable" else ""
        diagnostics_html = f"""
  <details class="diagnostics-box" {diag_open}>
    <summary>🔧 Preflight-Diagnose (Seiten-Gruppierung)</summary>
    <ul>{diag_items}</ul>
  </details>"""

    # Sicher als JS-Objekt-Literale einbetten (json.dumps escaped Anfuehrungszeichen,
    # Backslashes etc. korrekt - kein manuelles String-Basteln noetig).
    doc_meta_json = json.dumps({"a": meta_a, "b": meta_b}, ensure_ascii=False)
    script_version_json = json.dumps(SCRIPT_VERSION)
    review_schema_json = json.dumps(REVIEW_SCHEMA_VERSION)

    return f"""<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="utf-8">
<title>Dokumentvergleich: {html.escape(name_a)} vs {html.escape(name_b)}</title>
<style>
  :root {{
    --c-unchanged: #16a34a;
    --c-changed: #ea580c;
    --c-new: #2563eb;
    --c-deleted: #dc2626;
    --bg: #f7f7f8;
    --border: #e2e2e6;
  }}
  * {{ box-sizing: border-box; }}
  body {{
    font-family: -apple-system, Segoe UI, Roboto, Arial, sans-serif;
    margin: 0; background: var(--bg); color: #1a1a1a;
  }}
  header {{
    position: sticky; top: 0; z-index: 10;
    background: #fff; border-bottom: 1px solid var(--border);
    padding: 14px 20px;
  }}
  h1 {{ font-size: 16px; margin: 0 0 10px 0; font-weight: 600; }}
  .doc-names {{ font-size: 13px; color: #555; margin-bottom: 10px; }}
  .doc-names b {{ color: #111; }}
  .stats-bar {{ display: flex; gap: 10px; flex-wrap: wrap; align-items: stretch; }}
  .stat {{
    background: #fafafa; border: 1px solid var(--border); border-radius: 8px;
    padding: 6px 14px; min-width: 74px; text-align: center;
  }}
  .stat-num {{ font-size: 20px; font-weight: 700; }}
  .stat-label {{ font-size: 11px; color: #666; margin-top: 2px; }}
  .stat-unchanged .stat-num {{ color: var(--c-unchanged); }}
  .stat-changed .stat-num {{ color: var(--c-changed); }}
  .stat-new .stat-num {{ color: var(--c-new); }}
  .stat-deleted .stat-num {{ color: var(--c-deleted); }}
  .toolbar {{ margin-top: 10px; display: flex; gap: 14px; align-items: center; font-size: 13px; }}
  .legend {{ display: flex; gap: 12px; font-size: 12px; color: #444; flex-wrap: wrap; }}
  .legend span {{ display: inline-flex; align-items: center; gap: 4px; }}
  .legend i {{ width: 10px; height: 10px; border-radius: 50%; display: inline-block; }}
  .li-unchanged {{ background: var(--c-unchanged); }}
  .li-changed {{ background: var(--c-changed); }}
  .li-new {{ background: var(--c-new); }}
  .li-deleted {{ background: var(--c-deleted); }}

  .compare-grid {{ padding: 16px 20px 60px; }}
  .row-wrapper {{ margin-bottom: 14px; }}
  .row-wrapper.continues-group {{ margin-bottom: 0; }}
  .row-wrapper.hidden-by-filter {{ display: none; }}
  @keyframes deltaFlash {{
    0%   {{ box-shadow: 0 0 0 4px rgba(234, 88, 12, 0.65); }}
    100% {{ box-shadow: 0 0 0 4px rgba(234, 88, 12, 0); }}
  }}
  .delta-flash {{ animation: deltaFlash 1.1s ease-out; border-radius: 8px; }}
  .grid-row {{
    display: grid;
    grid-template-columns: 1fr 70px 1fr;
    gap: 0;
    align-items: stretch;
  }}
  .cell {{
    background: #fff; border: 1px solid var(--border); border-radius: 6px;
    padding: 8px 12px; font-size: 13px; line-height: 1.5; overflow-wrap: break-word;
  }}
  .cell-slot {{ min-width: 0; }}
  .cell-left {{ border-right: none; border-radius: 6px 0 0 6px; }}
  .cell-right {{ border-left: none; border-radius: 0 6px 6px 0; }}
  .cell-empty {{ display: flex; align-items: center; justify-content: center; color: #999; font-size: 12px; font-style: italic; background: #fbfbfb; }}
  .page-tag {{
    font-size: 18px; color: #1e293b; font-weight: 800; letter-spacing: 0.04em;
    margin: 16px 0 6px 6px; text-transform: uppercase;
  }}

  /* Seiten-Bandierung (abwechselnder Hintergrund je Seite) bleibt hier -
     der eigentliche Rahmen (pg-*) steht WEITER UNTEN, NACH den status-
     spezifischen border-color-Regeln (.row-unchanged/.row-changed/...),
     damit er nicht von der (viel helleren) Statusfarbe ueberschrieben wird
     (gleiche CSS-Spezifitaet, es gewinnt die spaeter stehende Regel). */
  .pageband-odd {{ background: #f4f6fa; }}
  summary {{ cursor: pointer; font-weight: 600; }}
  .chnum {{ color: #555; font-variant-numeric: tabular-nums; }}
  .chtitle {{ color: #111; }}
  .chbody {{ margin-top: 6px; color: #333; white-space: normal; }}
  .del {{ background: #fee2e2; color: #991b1b; text-decoration: line-through; }}
  .ins {{ background: #dcfce7; color: #14532d; }}

  .connector {{
    display: flex; flex-direction: column; align-items: center; justify-content: center;
    position: relative;
  }}
  .connector::before {{
    content: ""; position: absolute; left: 0; right: 0; top: 50%; height: 3px;
    transform: translateY(-50%);
  }}
  .conn-unchanged::before {{ background: var(--c-unchanged); }}
  .conn-changed::before {{ background: var(--c-changed); }}
  .conn-new::before {{ background: linear-gradient(to right, transparent 50%, var(--c-new) 50%); }}
  .conn-deleted::before {{ background: linear-gradient(to right, var(--c-deleted) 50%, transparent 50%); }}
  /* Seiten-Spange: durchgehende vertikale Linie in der Connector-Spalte,
     solange fuer diese Zeile auf mindestens einer Seite eine Seitenzahl
     bekannt ist - macht sichtbar, dass die linke und rechte Seiten-Box
     zusammengehoeren (ein Paar bilden), statt zwei unabhaengige Kaesten zu
     wirken. Ragt bewusst leicht ueber die Zeilenhoehe hinaus (top/bottom
     negativ), damit sie ueber den Zeilen-Abstand hinweg optisch nahtlos
     mit der Spange der Nachbarzeile verschmilzt.
  */
  .connector.has-page-spine::after {{
    content: ""; position: absolute; left: 50%; top: -8px; bottom: -8px;
    width: 3px; background: #1e3a8a; opacity: 0.4; transform: translateX(-50%); z-index: 0;
  }}
  .conn-icon {{ z-index: 1; background: #fff; border-radius: 50%; width: 22px; height: 22px; display: flex; align-items: center; justify-content: center; font-size: 12px; border: 2px solid; }}
  .conn-unchanged .conn-icon {{ border-color: var(--c-unchanged); color: var(--c-unchanged); }}
  .conn-changed .conn-icon {{ border-color: var(--c-changed); color: var(--c-changed); }}
  .conn-new .conn-icon {{ border-color: var(--c-new); color: var(--c-new); }}
  .conn-deleted .conn-icon {{ border-color: var(--c-deleted); color: var(--c-deleted); }}
  .conn-pct {{ font-size: 10px; color: #666; margin-top: 2px; z-index: 1; }}
  .conn-img-badge {{ font-size: 12px; z-index: 1; margin-top: 2px; }}

  .img-gallery {{ display: flex; flex-wrap: wrap; gap: 6px; margin-top: 8px; }}
  .img-thumb {{ max-width: 160px; max-height: 120px; border: 1px solid var(--border); border-radius: 4px; object-fit: contain; background: #fff; }}
  .img-placeholder {{
    font-size: 11px; color: #92400e; background: #fffbeb; border: 1px dashed #fcd34d;
    border-radius: 4px; padding: 6px 10px; max-width: 220px;
  }}

  .row-unchanged .cell-left, .row-unchanged .cell-right {{ border-color: #bbf7d0; }}
  .row-changed .cell-left, .row-changed .cell-right {{ border-color: #fed7aa; }}
  .row-new .cell-right {{ border-color: #bfdbfe; }}
  .row-deleted .cell-left {{ border-color: #fecaca; }}

  /* Seiten-Box: kraeftiger, unuebersehbarer Rahmen ueber mehrere Zeilen
     hinweg, wenn sie zur selben (gerenderten) Seite gehoeren. Steht
     BEWUSST nach den Status-Farbregeln oben, damit er nicht von der
     (viel helleren) Statusfarbe ueberschrieben wird (gleiche
     CSS-Spezifitaet - hier gewinnt die zuletzt stehende Regel).
     pg-first/last oeffnen die Box oben/unten mit Eckenrundung, pg-mid
     laesst die Kante nahtlos offen, pg-both rundet eine alleinstehende
     Ein-Zeilen-Seite komplett. */
  .cell.pg-grouped {{
    border-left-width: 6px !important; border-left-color: #1e3a8a !important;
    background: #eef2ff;
  }}
  .cell-left.pg-grouped {{ margin-left: 6px; }}
  .cell-right.pg-grouped {{ margin-right: 6px; }}
  .cell.pg-mid {{
    border-top: none !important; border-bottom: none !important; border-radius: 0 !important;
  }}
  .cell.pg-first {{
    border-top: 5px solid #1e3a8a !important; border-bottom: none !important;
    box-shadow: 0 -3px 8px -2px rgba(30, 58, 138, 0.35);
  }}
  .cell.pg-last {{
    border-bottom: 5px solid #1e3a8a !important; border-top: none !important;
    box-shadow: 0 3px 8px -2px rgba(30, 58, 138, 0.35);
  }}
  .cell.pg-both {{
    border-top: 5px solid #1e3a8a !important; border-bottom: 5px solid #1e3a8a !important;
    box-shadow: 0 0 8px -1px rgba(30, 58, 138, 0.35);
  }}
  .cell-left.pg-first, .cell-left.pg-both {{ border-top-left-radius: 10px; }}
  .cell-left.pg-last, .cell-left.pg-mid {{ border-top-left-radius: 0; }}
  .cell-left.pg-last, .cell-left.pg-both {{ border-bottom-left-radius: 10px; }}
  .cell-left.pg-first, .cell-left.pg-mid {{ border-bottom-left-radius: 0; }}
  .cell-right.pg-first, .cell-right.pg-both {{ border-top-right-radius: 10px; }}
  .cell-right.pg-last, .cell-right.pg-mid {{ border-top-right-radius: 0; }}
  .cell-right.pg-last, .cell-right.pg-both {{ border-bottom-right-radius: 10px; }}
  .cell-right.pg-first, .cell-right.pg-mid {{ border-bottom-right-radius: 0; }}

  button.filter-btn {{
    border: 1px solid var(--border); background: #fff; border-radius: 6px;
    padding: 5px 12px; font-size: 12px; cursor: pointer;
  }}
  button.filter-btn.active {{ background: #111; color: #fff; border-color: #111; }}

  .review-box {{
    display: flex; align-items: center; gap: 8px;
    background: #fafafa; border: 1px solid var(--border); border-top: none;
    border-radius: 0 0 6px 6px; padding: 6px 12px; font-size: 12px;
  }}
  .review-box label {{ color: #666; white-space: nowrap; }}
  .review-select {{ font-size: 12px; padding: 3px 4px; border-radius: 4px; border: 1px solid var(--border); background: #fff; }}
  .review-comment {{ flex: 1; font-size: 12px; padding: 4px 8px; border-radius: 4px; border: 1px solid var(--border); }}
  .review-box.rv-accepted {{ background: #f0fdf4; }}
  .review-box.rv-not_accepted {{ background: #fef2f2; }}
  .review-box.rv-refinement_customer {{ background: #eff6ff; }}
  .review-box.rv-internal_clarification {{ background: #fff7ed; }}
  .version-line {{ font-size: 11px; color: #888; }}
  .diagnostics-box {{
    margin-top: 8px; font-size: 12px; background: #f8fafc; border: 1px solid var(--border);
    border-radius: 6px; padding: 6px 12px;
  }}
  .diagnostics-box summary {{ cursor: pointer; color: #475569; font-weight: 600; }}
  .diagnostics-box ul {{ margin: 8px 0 4px 0; padding-left: 20px; color: #334155; }}
  .diagnostics-box li {{ margin-bottom: 4px; }}
  #review-import-warning {{
    display: none; background: #fffbeb; border: 1px solid #fcd34d; color: #92400e;
    padding: 8px 14px; border-radius: 6px; font-size: 12px; margin-top: 8px; white-space: pre-line;
  }}
</style>
</head>
<body>
<header>
  <h1>Kapitelvergleich (Fixpunkt: Kapitelnummer)</h1>
  <div class="doc-names">Dokument A: <b>{html.escape(name_a)}</b> &nbsp;|&nbsp; Dokument B: <b>{html.escape(name_b)}</b></div>
  <div class="doc-names">ℹ️ {html.escape(linebreak_note)}</div>
  {f'<div class="doc-names">{html.escape(page_note)}</div>' if page_note else ''}
  {diagnostics_html}
  <div class="version-line">Tool-Version {SCRIPT_VERSION} · Review-Schema {REVIEW_SCHEMA_VERSION}</div>
  <div class="stats-bar">{stats_html}</div>
  <div class="toolbar">
    <button class="filter-btn active" id="btn-all" onclick="setFilter('all')">Alle</button>
    <button class="filter-btn" id="btn-diff" onclick="setFilter('diff')">Nur Unterschiede</button>
    <button class="filter-btn" id="btn-unreviewed" onclick="setFilter('unreviewed')">Nur unbewertet</button>
    <button class="filter-btn" id="btn-expand" onclick="toggleExpandAll()">Alle auf-/zuklappen</button>
    <button class="filter-btn" onclick="jumpToDelta(-1)" title="Zum vorherigen Unterschied springen">⏮ Vorheriges Delta</button>
    <button class="filter-btn" onclick="jumpToDelta(1)" title="Zum naechsten Unterschied springen">⏭ Nächstes Delta</button>
    <button class="filter-btn" onclick="exportReviews()">⬇ Review exportieren (JSON)</button>
    <button class="filter-btn" onclick="document.getElementById('review-file-input').click()">⬆ Review importieren (JSON)</button>
    <input type="file" id="review-file-input" accept=".json,application/json" style="display:none" onchange="importReviewsFile(event)">
    <span class="legend">
      <span><i class="li-unchanged"></i> unverändert</span>
      <span><i class="li-changed"></i> geändert</span>
      <span><i class="li-new"></i> neu</span>
      <span><i class="li-deleted"></i> gelöscht</span>
      <span>🖼 Grafik geändert</span>
    </span>
  </div>
  <div id="review-import-warning"></div>
</header>
<div class="compare-grid" id="grid">
  {''.join(row_html)}
</div>
<script>
  const DOC_META = {doc_meta_json};
  const SCRIPT_VERSION = {script_version_json};
  const REVIEW_SCHEMA_VERSION = {review_schema_json};
  const REVIEW_STATUS_VALUES = ['accepted', 'not_accepted', 'refinement_customer', 'internal_clarification'];

  const REVIEW_BOXES = {{}};
  document.querySelectorAll('.review-box').forEach(function(box) {{
    REVIEW_BOXES[box.getAttribute('data-review-key')] = box;
  }});

  let filterMode = 'all';
  let expanded = false;
  let deltaIndex = -1;

  function setFilter(mode) {{
    filterMode = mode;
    deltaIndex = -1;  // Filter geaendert - Delta-Navigation faengt neu an
    ['all', 'diff', 'unreviewed'].forEach(function(m) {{
      document.getElementById('btn-' + m).classList.toggle('active', m === mode);
    }});
    document.querySelectorAll('.row-wrapper').forEach(function(row) {{
      const status = row.getAttribute('data-status');
      let show = true;
      if (mode === 'diff') {{
        show = status !== 'unchanged';
      }} else if (mode === 'unreviewed') {{
        const select = row.querySelector('.review-select');
        show = !select || !select.value;
      }}
      row.classList.toggle('hidden-by-filter', !show);
    }});
  }}

  function toggleExpandAll() {{
    expanded = !expanded;
    document.querySelectorAll('details.cell').forEach(function(d) {{ d.open = expanded; }});
  }}

  function getDeltaRows() {{
    return Array.from(document.querySelectorAll('.row-wrapper')).filter(function(row) {{
      return row.getAttribute('data-status') !== 'unchanged' && !row.classList.contains('hidden-by-filter');
    }});
  }}

  function jumpToDelta(direction) {{
    const rows = getDeltaRows();
    if (!rows.length) {{ return; }}
    deltaIndex += direction;
    if (deltaIndex >= rows.length) {{ deltaIndex = 0; }}
    if (deltaIndex < 0) {{ deltaIndex = rows.length - 1; }}
    const row = rows[deltaIndex];
    row.scrollIntoView({{behavior: 'smooth', block: 'center'}});
    row.classList.remove('delta-flash');
    void row.offsetWidth;  // Reflow erzwingen, damit die Animation bei erneutem Treffer neu startet
    row.classList.add('delta-flash');
  }}

  function onReviewChange(id) {{
    const select = document.getElementById('rs-' + id);
    const box = select.closest('.review-box');
    REVIEW_STATUS_VALUES.forEach(function(v) {{ box.classList.remove('rv-' + v); }});
    if (select.value) {{ box.classList.add('rv-' + select.value); }}
    if (filterMode === 'unreviewed') {{ setFilter('unreviewed'); }}
  }}

  function collectReviews() {{
    const result = {{}};
    Object.keys(REVIEW_BOXES).forEach(function(key) {{
      const box = REVIEW_BOXES[key];
      const status = box.querySelector('.review-select').value;
      const comment = box.querySelector('.review-comment').value.trim();
      if (status || comment) {{
        result[key] = {{status: status, comment: comment}};
      }}
    }});
    return result;
  }}

  function sanitizeFilename(name) {{
    return (name || 'doc').replace(/[^a-zA-Z0-9_.-]+/g, '_').slice(0, 60);
  }}

  function exportReviews() {{
    const data = {{
      schema_version: REVIEW_SCHEMA_VERSION,
      generated_by: 'docx_chapter_compare.py v' + SCRIPT_VERSION,
      doc_a: DOC_META.a,
      doc_b: DOC_META.b,
      reviews: collectReviews()
    }};
    const blob = new Blob([JSON.stringify(data, null, 2)], {{type: 'application/json'}});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'review_' + sanitizeFilename(DOC_META.a.name) + '_vs_' + sanitizeFilename(DOC_META.b.name) + '.json';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }}

  function importReviewsFile(evt) {{
    const file = evt.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function(e) {{
      let data;
      try {{ data = JSON.parse(e.target.result); }}
      catch (err) {{ alert('Ungültige JSON-Datei: ' + err.message); return; }}
      applyReviews(data);
    }};
    reader.readAsText(file);
    evt.target.value = '';
  }}

  function applyReviews(data) {{
    data = data || {{}};
    const reviews = data.reviews || {{}};
    let matched = 0, unmatched = 0;
    Object.keys(reviews).forEach(function(key) {{
      const entry = reviews[key] || {{}};
      const box = REVIEW_BOXES[key];
      if (!box) {{ unmatched++; return; }}
      const select = box.querySelector('.review-select');
      const comment = box.querySelector('.review-comment');
      select.value = entry.status || '';
      comment.value = entry.comment || '';
      onReviewChange(select.id.replace('rs-', ''));
      matched++;
    }});

    const warnBox = document.getElementById('review-import-warning');
    const da = data.doc_a || {{}}, db = data.doc_b || {{}};
    let warnings = [];
    if (da.name && DOC_META.a.name && da.name !== DOC_META.a.name) {{
      warnings.push('Dokument A: JSON nennt "' + da.name + '", aktueller Report vergleicht "' + DOC_META.a.name + '".');
    }}
    if (da.modified && DOC_META.a.modified && da.modified !== DOC_META.a.modified) {{
      warnings.push('Dokument A wurde seit dem Review-Export geändert (Zeitstempel weicht ab).');
    }}
    if (db.name && DOC_META.b.name && db.name !== DOC_META.b.name) {{
      warnings.push('Dokument B: JSON nennt "' + db.name + '", aktueller Report vergleicht "' + DOC_META.b.name + '".');
    }}
    if (db.modified && DOC_META.b.modified && db.modified !== DOC_META.b.modified) {{
      warnings.push('Dokument B wurde seit dem Review-Export geändert (Zeitstempel weicht ab).');
    }}
    if (unmatched > 0) {{
      warnings.push(unmatched + ' Kapitel aus der JSON-Datei wurden im aktuellen Report nicht gefunden (evtl. andere Kapitelstruktur) und übersprungen.');
    }}
    if (warnings.length) {{
      warnBox.textContent = '⚠ ' + warnings.join('\\n⚠ ');
      warnBox.style.display = 'block';
    }} else {{
      warnBox.style.display = 'none';
    }}
    alert('Review importiert: ' + matched + ' Kapitel übernommen.' + (warnings.length ? ' Siehe Hinweis oben.' : ''));
  }}
</script>
</body>
</html>
"""


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def diagnose_page_detection(soffice_explicit=None):
    """Liefert eine Liste menschenlesbarer Diagnosezeilen, WARUM die Seiten-
    Gruppierung nicht verfuegbar ist - fuer die Fehlersuche auf einem
    konkreten Rechner (z.B. 'pywin32 fehlt' vs. 'soffice nicht gefunden')."""
    lines = []

    if sys.platform != "win32":
        lines.append(f"MS Word/COM: nicht Windows (Plattform: {sys.platform}), daher nicht verfuegbar.")
    else:
        try:
            import win32com.client
        except ImportError:
            lines.append("MS Word/COM: pywin32 ist NICHT installiert. Installieren mit: pip install pywin32")
        else:
            # Echten Verbindungsversuch machen statt nur zu raten - liefert
            # die tatsaechliche COM-Fehlermeldung (z.B. "Word nicht
            # installiert" vs. einen anderen, spezifischeren Fehler).
            word = None
            try:
                word = win32com.client.DispatchEx("Word.Application")
                lines.append("MS Word/COM: Verbindung zu Word erfolgreich aufgebaut - sollte also "
                              "funktionieren. Falls trotzdem 'unavailable', liegt es an einem Fehler "
                              "beim Oeffnen/Auslesen der konkreten Datei (evtl. geschuetzt oder korrupt).")
            except Exception as exc:
                lines.append(f"MS Word/COM: pywin32 ist installiert, aber die Verbindung zu Word "
                              f"schlug fehl: {exc!r}. Das deutet meist darauf hin, dass MS Word auf "
                              f"diesem Rechner gar nicht installiert ist (z.B. nur LibreOffice) - "
                              f"dann ist dieser Weg hier erwartungsgemaess nicht nutzbar, LibreOffice "
                              f"sollte aber greifen (siehe unten).")
            finally:
                try:
                    if word is not None:
                        word.Quit()
                except Exception:
                    pass

    soffice_path = find_soffice(soffice_explicit)
    if soffice_path:
        lines.append(f"LibreOffice: soffice gefunden unter '{soffice_path}'.")
        try:
            import pdfplumber  # noqa: F401
            pdfplumber_ok = True
            lines.append("pdfplumber: ist installiert - PDF-Seiten koennen gelesen werden.")
        except ImportError:
            pdfplumber_ok = False
            lines.append("pdfplumber: ist NICHT installiert - das ist vermutlich der Grund, "
                          "warum die Seiten-Gruppierung trotz gefundenem soffice nicht "
                          "funktioniert! Installieren mit: pip install pdfplumber")

        if pdfplumber_ok:
            # Echten End-to-End-Test: kleines Test-Dokument erzeugen, konvertieren,
            # PDF lesen - deckt Konvertierungsfehler/Timeouts auf, die reine
            # Existenzpruefung des Pfads nicht zeigen wuerde.
            try:
                with tempfile.TemporaryDirectory() as tmp:
                    test_docx = Path(tmp) / "diagnose_test.docx"
                    Document_ = Document()
                    Document_.add_paragraph("1\tTestabsatz fuer Diagnose")
                    Document_.save(test_docx)
                    result = render_page_texts(test_docx, soffice_path, timeout=60)
                if result:
                    lines.append("End-zu-End-Test: Testkonvertierung erfolgreich - die Seiten-"
                                  "Gruppierung sollte funktionieren. Falls es bei den echten "
                                  "Dokumenten trotzdem 'unavailable' zeigt, koennte es an genau "
                                  "diesen Dateien liegen (z.B. Kennwortschutz, sehr grosse Datei, "
                                  "Sonderzeichen im Dateinamen).")
                else:
                    lines.append("End-zu-End-Test: Testkonvertierung ist FEHLGESCHLAGEN "
                                  "(render_page_texts lieferte kein Ergebnis). Moegliche "
                                  "Ursachen: soffice startet nicht richtig (evtl. laeuft schon "
                                  "eine andere LibreOffice-Instanz und blockiert), oder das PDF "
                                  "konnte nicht gelesen werden.")
            except Exception as exc:
                lines.append(f"End-zu-End-Test: Fehler beim Testen: {exc!r}")
    else:
        lines.append("LibreOffice: soffice wurde NICHT gefunden (weder im PATH noch an den "
                      "ueblichen Installationspfaden). Pfad manuell angeben mit --soffice-path "
                      "\"C:\\Pfad\\zu\\soffice.exe\" (den Pfad findest du z.B. per Rechtsklick auf "
                      "die LibreOffice-Verknuepfung -> Dateipfad oeffnen).")
    return lines


def main():
    print(f"docx_chapter_compare.py Version {SCRIPT_VERSION} ({Path(__file__).resolve()})")
    parser = argparse.ArgumentParser(
        description="Vergleicht zwei Word-Dokumente anhand von Kapitelnummern und erzeugt einen HTML-Report."
    )
    parser.add_argument("doc_a", nargs="?", help="Pfad zum ersten (alten) .docx")
    parser.add_argument("doc_b", nargs="?", help="Pfad zum zweiten (neuen) .docx")
    parser.add_argument("-o", "--output", default="vergleich.html", help="Pfad der HTML-Ausgabedatei")
    parser.add_argument(
        "--keep-linebreaks", action="store_true",
        help="Zeilen-/Absatzumbrueche NICHT ignorieren (strikter Vergleich). "
             "Standard: Umbrueche werden ignoriert, nur der Wortinhalt zaehlt.",
    )
    parser.add_argument(
        "--no-pages", action="store_true",
        help="Seiten-Gruppierung ('Buchform') abschalten. Standard: an, sofern "
             "LibreOffice (soffice) gefunden wird.",
    )
    parser.add_argument(
        "--soffice-path", default=None,
        help="Expliziter Pfad zu soffice/soffice.exe, falls automatische Suche fehlschlaegt.",
    )
    parser.add_argument(
        "--diagnose-pages", action="store_true",
        help="Nur pruefen, warum die Seiten-Gruppierung (nicht) verfuegbar ist, und die "
             "Diagnosezeilen ausgeben - ohne doc_a/doc_b, ohne Report zu erzeugen.",
    )
    parser.add_argument(
        "--docx", nargs="?", const="", default=None, metavar="PFAD",
        help="Zusaetzlich einen kompakten Word-Report erzeugen (fuer schnelle Orientierung, "
             "z.B. Weitergabe im Unternehmen ohne Browser). Ohne PFAD wird der Name der "
             "HTML-Ausgabe mit .docx-Endung verwendet.",
    )
    parser.add_argument(
        "--review-json", default=None, metavar="PFAD",
        help="Zuvor per 'Review exportieren' gespeicherte JSON-Datei einlesen und die "
             "Bewertungen (Status/Kommentar) in den Word-Report mit aufnehmen.",
    )
    args = parser.parse_args()

    if args.diagnose_pages:
        print("Diagnose Seiten-Gruppierung:")
        for line in diagnose_page_detection(args.soffice_path):
            print(f"  - {line}")
        sys.exit(0)

    if not args.doc_a or not args.doc_b:
        parser.error("doc_a und doc_b sind erforderlich (ausser bei --diagnose-pages).")

    ignore_linebreaks = not args.keep_linebreaks

    path_a, path_b = Path(args.doc_a), Path(args.doc_b)
    for p in (path_a, path_b):
        if not p.exists():
            print(f"Datei nicht gefunden: {p}", file=sys.stderr)
            sys.exit(1)

    chapters_a = extract_chapters(path_a)
    chapters_b = extract_chapters(path_b)

    if not chapters_a or not chapters_b:
        print("Warnung: In mindestens einem Dokument wurden keine Kapitel erkannt "
              "(keine Heading-Formatvorlagen und kein Kapitelnummern-Muster gefunden).",
              file=sys.stderr)

    pages_method = None
    diagnostics = None
    if not args.no_pages:
        print("Preflight-Check Seiten-Gruppierung ...")
        diagnostics = diagnose_page_detection(args.soffice_path)
        for line in diagnostics:
            print(f"  - {line}")

        print("Ermittle Seiten (MS Word COM, sonst LibreOffice) - kann einige Sekunden dauern ...")
        method = attach_pages(chapters_a, chapters_b, path_a, path_b, soffice_path=args.soffice_path)
        pages_method = method if method is not None else "unavailable"
        if method == "word_com":
            print("Seiten via MS Word ermittelt (exakt).")
        elif method == "libreoffice":
            print("Seiten via LibreOffice ermittelt (Näherung).")
        else:
            print("Hinweis: Weder MS Word (COM/pywin32) noch LibreOffice (soffice) verfügbar - "
                  "Seiten-Gruppierung wird ausgelassen. Details siehe Preflight-Check oben "
                  "(auch im HTML-Report unter 'Preflight-Diagnose' zu finden).", file=sys.stderr)

    rows = build_comparison(chapters_a, chapters_b, ignore_linebreaks=ignore_linebreaks)
    stats = compute_stats(chapters_a, chapters_b, rows)

    out_html = render_html(
        rows, stats, path_a.name, path_b.name, ignore_linebreaks=ignore_linebreaks,
        meta_a=doc_metadata(path_a), meta_b=doc_metadata(path_b),
        pages_method=pages_method, diagnostics=diagnostics,
    )
    out_path = Path(args.output)
    out_path.write_text(out_html, encoding="utf-8")

    print(f"Report geschrieben: {out_path.resolve()}")
    print(f"Kapitel A: {stats['total_a']} | Kapitel B: {stats['total_b']} | "
          f"unveraendert: {stats['unchanged']} | geaendert: {stats['changed']} | "
          f"neu: {stats['new']} | geloescht: {stats['deleted']}")

    if args.docx is not None:
        docx_path = Path(args.docx) if args.docx else out_path.with_suffix(".docx")
        reviews = None
        if args.review_json:
            try:
                review_data = json.loads(Path(args.review_json).read_text(encoding="utf-8"))
                reviews = review_data.get("reviews", {})
            except Exception as exc:
                print(f"Warnung: Review-JSON konnte nicht gelesen werden ({exc}) - "
                      f"Word-Report wird ohne Review-Spalte erzeugt.", file=sys.stderr)
        generate_docx_report(
            rows, stats, path_a.name, path_b.name, docx_path,
            meta_a=doc_metadata(path_a), meta_b=doc_metadata(path_b), reviews=reviews,
        )
        print(f"Word-Report geschrieben: {docx_path.resolve()}")


if __name__ == "__main__":
    main()
