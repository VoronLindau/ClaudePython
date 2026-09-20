#!/usr/bin/env python3
"""
docx_chapter_compare_gui.py
Version 2.3 / 2026-09-20 / Grund: Wichtiger Bugfix (Nutzer bestaetigte: auch
    eine unveraenderte alte Version zeigt inzwischen dasselbe Problem - also
    eine Umgebungsaenderung auf dem Rechner, kein Code-Bug) - manche
    (v.a. streng konfigurierte Firmen-)Rechner behandeln lokale file://-
    Dateien aus Sicherheitsgruenden anders als normale Webseiten und zeigen
    HTML nur als Rohtext statt gerendert an. open_in_browser() startet jetzt
    zuerst einen minimalen lokalen HTTP-Server (NUR 127.0.0.1, zufaelliger
    Port) und oeffnet den Bericht ueber http://127.0.0.1:PORT/... statt
    file://, was diese Einschraenkung umgeht (echter Content-Type-Header
    text/html statt Datei-Endungs-Raten). Aus Sicherheitsgruenden (ggf.
    vertrauliche Dokumente im selben Ordner wie der Bericht) wird NICHT der
    komplette Ausgabeordner ausgeliefert, sondern die Report-Datei zuvor in
    ein frisches, isoliertes Temp-Verzeichnis kopiert - der Server liefert
    ausschliesslich diese eine Datei aus (mit echtem Request getestet:
    Verzeichnis-Traversal auf das Elternverzeichnis schlaegt fehl/liefert
    nur die eigene Datei). Faellt bei Fehlschlag weiterhin auf file:// +
    Registry-/Pfad-Suche + Windows-Standard zurueck wie in v2.2.

Desktop-GUI (Tkinter, keine Zusatz-Installation noetig) fuer den
Kapitelvergleich zweier Word-Dokumente. Nutzt dieselbe Vergleichslogik und
denselben HTML-Report wie docx_chapter_compare.py (Browser-Diff-Darstellung
mit Statistik-Leiste, farbigen Verbindungslinien und Wort-Diff).

Start:
    python docx_chapter_compare_gui.py

Voraussetzung: python-docx muss installiert sein (siehe README/Chatverlauf):
    pip install python-docx

Diese Datei muss im selben Ordner liegen wie docx_chapter_compare.py.
"""

import json
import os
import subprocess
import sys
import threading
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

GUI_VERSION = "2.3"

try:
    import docx_chapter_compare as _core
    from docx_chapter_compare import (
        attach_pages,
        build_comparison,
        compute_stats,
        detect_possible_moves,
        doc_metadata,
        extract_chapters,
        generate_docx_report,
        render_html,
    )
except ImportError as exc:
    # Haeufigste Ursache: python-docx fehlt, oder docx_chapter_compare.py
    # liegt nicht im selben Verzeichnis wie dieses Skript.
    print(f"Fehler beim Import von docx_chapter_compare.py: {exc}", file=sys.stderr)
    print("Pruefe, ob 'python-docx' installiert ist (pip install python-docx) "
          "und ob docx_chapter_compare.py im selben Ordner liegt.", file=sys.stderr)
    raise

CORE_VERSION = getattr(_core, "SCRIPT_VERSION", "?")
CORE_PATH = getattr(_core, "__file__", "?")
GUI_PATH = str(Path(__file__).resolve())

# Beim Start IMMER in der Konsole ausgeben - unabhaengig davon, ob die GUI
# ueberhaupt geoeffnet wird (z.B. falls ein Fehler vor dem Fenster auftritt).
print(f"docx_chapter_compare_gui.py Version {GUI_VERSION} ({GUI_PATH})")
print(f"docx_chapter_compare.py     Version {CORE_VERSION} ({CORE_PATH})")

CONFIG_PATH = Path.home() / ".docx_chapter_compare_gui.json"


def default_output_dir():
    """Documents-Ordner, falls vorhanden - sonst Fallback aufs Home-Verzeichnis.
    Vermeidet, Report-Dateien direkt ins Profil-Root (C:\\Users\\<Name>\\) zu
    kippen, was auf Windows unueblich/unaufgeraeumt wirkt."""
    docs = Path.home() / "Documents"
    return docs if docs.is_dir() else Path.home()


# Bekannte Installationspfade gaengiger Browser (Windows) - werden VOR der
# Windows-Dateizuordnung probiert. Grund: manche (v.a. restriktiv
# konfigurierte Firmen-)Rechner haben keine .html-Dateizuordnung gesetzt;
# webbrowser.open() loest dann ueber die Windows-Shell auf und zeigt den
# "Wie soll diese Datei geoeffnet werden?"-Dialog mit Apps wie Editor/Paint/
# Gimp (oder faellt auf einen veralteten Internet Explorer als Standard-App
# zurueck, der lokale HTML-Dateien teils nur als Rohtext statt gerendert
# anzeigt) statt direkt einen modernen Browser zu starten. Ein direkt
# gestarteter Browser umgeht diese Zuordnung komplett.
_BROWSER_EXE_NAMES = ["msedge.exe", "chrome.exe", "firefox.exe"]
_BROWSER_CANDIDATES = [
    r"%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe",
    r"%ProgramFiles%\Microsoft\Edge\Application\msedge.exe",
    r"%ProgramFiles%\Google\Chrome\Application\chrome.exe",
    r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe",
    r"%LocalAppData%\Google\Chrome\Application\chrome.exe",
    r"%ProgramFiles%\Mozilla Firefox\firefox.exe",
    r"%ProgramFiles(x86)%\Mozilla Firefox\firefox.exe",
]


def _find_browser_via_registry():
    """Fragt Windows selbst (App Paths-Registrierung) nach dem tatsaechlichen
    Installationsort von Edge/Chrome/Firefox - das ist der Mechanismus, den
    Windows intern auch benutzt, und findet Browser zuverlaessiger als
    geratene Standardpfade (z.B. bei individuell konfigurierten
    Firmenrechnern, benutzerspezifischen statt systemweiten Installationen,
    o.ae.). Gibt den ersten gefundenen Pfad zurueck, oder None."""
    if sys.platform != "win32":
        return None
    try:
        import winreg
    except ImportError:
        return None
    for exe_name in _BROWSER_EXE_NAMES:
        key_path = rf"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{exe_name}"
        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(hive, key_path) as key:
                    path, _ = winreg.QueryValueEx(key, "")
                    if path and Path(path).exists():
                        return path
            except (FileNotFoundError, OSError):
                continue
    return None


def _start_local_server(source_file):
    """Startet einen minimalen lokalen HTTP-Server (NUR auf 127.0.0.1,
    zufaelliger freier Port) im Hintergrund. Manche (v.a. streng
    konfigurierte Firmen-)Rechner behandeln lokale file://-Dateien aus
    Sicherheitsgruenden anders als normale Webseiten - z.B. wird HTML dann
    nur als Rohtext angezeigt statt gerendert (beobachtet: aendert sich auf
    einem Rechner im Zeitverlauf durch ein Windows-/Edge-Update oder eine
    IT-Richtlinie, unabhaengig von diesem Skript). Ueber einen echten - wenn
    auch rein lokalen, nach aussen nicht erreichbaren - HTTP-Server umgeht
    man file://-spezifische Einschraenkungen komplett.

    WICHTIG (Sicherheit): Liefert NICHT den kompletten Ordner der Report-
    Datei aus (der koennte z.B. der Documents-Ordner mit anderen, ggf.
    vertraulichen Dateien sein) - stattdessen wird die Report-Datei in ein
    frisches, temporaeres Verzeichnis kopiert und NUR das ausgeliefert.
    Selbst wenn ein anderer lokaler Prozess den Server abfragen wuerde,
    kaeme er nur an genau die eine Datei, die ohnehin gleich im Browser
    angezeigt wird - nicht an den Rest des Ordners.

    Gibt (port, temp_dir) zurueck, oder (None, None) wenn nicht gestartet
    werden konnte."""
    try:
        import functools
        import http.server
        import shutil
        import tempfile
    except ImportError:
        return None, None
    try:
        temp_dir = tempfile.mkdtemp(prefix="docx_compare_view_")
        shutil.copy2(source_file, Path(temp_dir) / source_file.name)
        handler = functools.partial(http.server.SimpleHTTPRequestHandler, directory=temp_dir)
        httpd = http.server.HTTPServer(("127.0.0.1", 0), handler)
        port = httpd.server_address[1]
        thread = threading.Thread(target=httpd.serve_forever, daemon=True)
        thread.start()
        return port, temp_dir
    except Exception:
        return None, None


def open_in_browser(path):
    """Oeffnet eine Datei in einem echten Browser. Reihenfolge:
    (1) ueber einen lokalen Mini-HTTP-Server (http://127.0.0.1:PORT/...) -
        das ist der ROBUSTESTE Weg, da manche Firmenrechner file://-Inhalte
        speziell einschraenken/nur als Rohtext anzeigen, http://-Inhalte
        aber normal behandeln;
    (2) falls der Server nicht gestartet werden konnte: direkt per file://,
        ueber Windows' eigene App-Paths-Registrierung (am zuverlaessigsten)
        oder bekannte Standard-Installationspfade;
    (3) als letzter Ausweg: webbrowser.open() (Windows-Dateizuordnung).
    Gibt (erfolg: bool, methode: str) zurueck, damit der Aufrufer sichtbar
    machen kann, WELCHER Weg gegriffen hat."""
    server_port, _server_temp_dir = _start_local_server(path)
    if server_port is not None:
        url = f"http://127.0.0.1:{server_port}/{path.name}"
    else:
        url = path.resolve().as_uri()

    if sys.platform == "win32":
        registry_hit = _find_browser_via_registry()
        if registry_hit:
            try:
                subprocess.Popen([registry_hit, url])
                return True, ("server" if server_port else "registry")
            except Exception:
                pass
        for template in _BROWSER_CANDIDATES:
            exe = os.path.expandvars(template)
            if Path(exe).exists():
                try:
                    subprocess.Popen([exe, url])
                    return True, ("server" if server_port else "candidate_path")
                except Exception:
                    continue
    try:
        webbrowser.open(url)
        return True, ("server" if server_port else "os_default")
    except Exception:
        return False, "failed"


def load_config():
    try:
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_config(cfg):
    try:
        CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass  # Konfiguration ist rein komfortbezogen - Fehler hier sind unkritisch


class CompareApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"DOCX Kapitelvergleich  —  GUI v{GUI_VERSION} · Core v{CORE_VERSION}")
        self.geometry("640x500")
        self.minsize(560, 420)
        self.resizable(True, False)

        self.cfg = load_config()

        self.var_a = tk.StringVar(value=self.cfg.get("doc_a", ""))
        self.var_b = tk.StringVar(value=self.cfg.get("doc_b", ""))
        self.var_out = tk.StringVar(value=self.cfg.get("out", str(default_output_dir() / "vergleich.html")))
        self.var_status = tk.StringVar(value="Bereit.")
        self.var_ignore_linebreaks = tk.BooleanVar(value=self.cfg.get("ignore_linebreaks", True))
        self.var_detect_pages = tk.BooleanVar(value=self.cfg.get("detect_pages", True))
        self.var_detect_moves = tk.BooleanVar(value=self.cfg.get("detect_moves", True))
        self.var_export_docx = tk.BooleanVar(value=self.cfg.get("export_docx", False))

        self._build_ui()

    # ------------------------------------------------------------------
    def _build_ui(self):
        frame = ttk.Frame(self)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Kapitelvergleich zweier Word-Dokumente", font=("", 12, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=10, pady=(12, 4)
        )
        ttk.Label(
            frame,
            text="Kapitel-/Anforderungsnummer wird als Fixpunkt genutzt. Ergebnis öffnet sich im Browser.",
            foreground="#555",
        ).grid(row=1, column=0, columnspan=3, sticky="w", padx=10, pady=(0, 10))

        version_frame = ttk.Frame(frame, relief="groove", borderwidth=1)
        version_frame.grid(row=2, column=0, columnspan=3, sticky="ew", padx=10, pady=(0, 10))
        ttk.Label(
            version_frame,
            text=f"GUI v{GUI_VERSION}  →  {GUI_PATH}",
            font=("Consolas", 8), foreground="#333",
        ).pack(anchor="w", padx=6, pady=(4, 0))
        ttk.Label(
            version_frame,
            text=f"Core v{CORE_VERSION}  →  {CORE_PATH}",
            font=("Consolas", 8), foreground="#333",
        ).pack(anchor="w", padx=6, pady=(0, 4))

        self._file_row(frame, row=3, label="Dokument A (alt):", var=self.var_a, command=self.pick_a)
        self._file_row(frame, row=4, label="Dokument B (neu):", var=self.var_b, command=self.pick_b)
        self._file_row(frame, row=5, label="Report-Ausgabe:", var=self.var_out, command=self.pick_out, save=True)

        ttk.Checkbutton(
            frame, text="Zeilenumbrüche ignorieren (nur Wortinhalt vergleichen)",
            variable=self.var_ignore_linebreaks,
        ).grid(row=6, column=0, columnspan=3, sticky="w", padx=10, pady=(2, 0))

        ttk.Checkbutton(
            frame, text="Seiten-Gruppierung / Buchform (MS Word oder LibreOffice, etwas langsamer)",
            variable=self.var_detect_pages,
        ).grid(row=7, column=0, columnspan=3, sticky="w", padx=10, pady=(2, 0))

        ttk.Checkbutton(
            frame, text="Mögliche Kapitel-Verschiebungen erkennen",
            variable=self.var_detect_moves,
        ).grid(row=8, column=0, columnspan=3, sticky="w", padx=10, pady=(2, 0))

        ttk.Checkbutton(
            frame, text="Zusätzlich Word-Report erzeugen (kompakt, für schnelle Weitergabe im Unternehmen)",
            variable=self.var_export_docx,
        ).grid(row=9, column=0, columnspan=3, sticky="w", padx=10, pady=(2, 0))

        frame.columnconfigure(1, weight=1)

        self.btn_compare = ttk.Button(frame, text="Vergleichen ▶", command=self.run_compare)
        self.btn_compare.grid(row=10, column=0, padx=10, pady=(16, 4), sticky="w")

        self.progress = ttk.Progressbar(frame, mode="determinate", maximum=1, value=0)
        self.progress.grid(row=10, column=1, columnspan=2, padx=10, pady=(16, 4), sticky="ew")

        self.progress_label = ttk.Label(frame, text="", foreground="#666", font=("", 9))
        self.progress_label.grid(row=11, column=0, columnspan=3, sticky="w", padx=10)

        self.stats_label = ttk.Label(frame, text="", justify="left")
        self.stats_label.grid(row=12, column=0, columnspan=3, sticky="w", padx=10, pady=(6, 0))

        status_bar = ttk.Label(self, textvariable=self.var_status, relief="sunken", anchor="w")
        status_bar.pack(fill="x", side="bottom")

    def _file_row(self, parent, row, label, var, command, save=False):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=10, pady=6)
        entry = ttk.Entry(parent, textvariable=var)
        entry.grid(row=row, column=1, sticky="ew", padx=(0, 6), pady=6)
        ttk.Button(parent, text="Speichern unter…" if save else "Durchsuchen…", command=command).grid(
            row=row, column=2, sticky="e", padx=(0, 10), pady=6
        )

    # ------------------------------------------------------------------
    def pick_a(self):
        path = filedialog.askopenfilename(
            title="Dokument A (alt) auswählen",
            filetypes=[("Word-Dokumente", "*.docx"), ("Alle Dateien", "*.*")],
        )
        if path:
            self.var_a.set(path)

    def pick_b(self):
        path = filedialog.askopenfilename(
            title="Dokument B (neu) auswählen",
            filetypes=[("Word-Dokumente", "*.docx"), ("Alle Dateien", "*.*")],
        )
        if path:
            self.var_b.set(path)

    def pick_out(self):
        path = filedialog.asksaveasfilename(
            title="Report speichern unter",
            defaultextension=".html",
            filetypes=[("HTML-Datei", "*.html")],
            initialfile="vergleich.html",
        )
        if path:
            self.var_out.set(path)

    # ------------------------------------------------------------------
    def run_compare(self):
        path_a = Path(self.var_a.get().strip())
        path_b = Path(self.var_b.get().strip())
        out_path = Path(self.var_out.get().strip() or (default_output_dir() / "vergleich.html"))

        if not self.var_a.get().strip() or not self.var_b.get().strip():
            messagebox.showwarning("Fehlende Angabe", "Bitte beide Dokumente auswählen.")
            return
        if not path_a.exists():
            messagebox.showerror("Datei nicht gefunden", f"Dokument A nicht gefunden:\n{path_a}")
            return
        if not path_b.exists():
            messagebox.showerror("Datei nicht gefunden", f"Dokument B nicht gefunden:\n{path_b}")
            return

        ignore_linebreaks = self.var_ignore_linebreaks.get()
        detect_pages = self.var_detect_pages.get()
        detect_moves = self.var_detect_moves.get()
        export_docx = self.var_export_docx.get()

        # Schritt-Liste VORHER exakt so aufbauen, wie sie im Worker durchlaufen
        # wird - daraus ergibt sich die Gesamtzahl fuer die Fortschrittsanzeige.
        steps = ["Dokument A einlesen", "Dokument B einlesen"]
        if detect_pages:
            steps += ["Preflight-Check Seiten-Gruppierung", "Seiten ermitteln"]
        steps += ["Vergleich berechnen"]
        if detect_moves:
            steps += ["Verschiebungen erkennen"]
        steps += ["Report schreiben"]
        if export_docx:
            steps += ["Word-Report erzeugen"]

        self.btn_compare.config(state="disabled")
        self.progress.config(mode="determinate", maximum=len(steps), value=0)
        self.progress_label.config(text=f"[0/{len(steps)}] Bereit …")
        self.var_status.set("Vergleiche Dokumente …")
        self.stats_label.config(text="")

        thread = threading.Thread(
            target=self._worker,
            args=(path_a, path_b, out_path, ignore_linebreaks, detect_pages, detect_moves, export_docx, len(steps)),
            daemon=True,
        )
        thread.start()

    def _set_progress(self, i, total, label):
        self.progress["value"] = i
        self.progress_label.config(text=f"[{i}/{total}] {label} …")

    def _worker(self, path_a, path_b, out_path, ignore_linebreaks, detect_pages, detect_moves, export_docx, total_steps):
        step_counter = [0]

        def advance(label):
            step_counter[0] += 1
            i = step_counter[0]
            self.after(0, lambda: self._set_progress(i, total_steps, label))

        try:
            chapters_a = extract_chapters(path_a)
            advance("Dokument A einlesen")
            chapters_b = extract_chapters(path_b)
            advance("Dokument B einlesen")

            if not chapters_a or not chapters_b:
                self.after(0, lambda: messagebox.showwarning(
                    "Keine Kapitel erkannt",
                    "In mindestens einem Dokument wurden keine Kapitel/Anforderungen erkannt.\n"
                    "Ergebnis kann unvollständig sein.",
                ))

            pages_method = None
            diagnostics = None
            if detect_pages:
                diagnostics = _core.diagnose_page_detection()
                advance("Preflight-Check Seiten-Gruppierung")
                method = attach_pages(chapters_a, chapters_b, path_a, path_b)
                pages_method = method if method is not None else "unavailable"
                advance("Seiten ermitteln")

            rows = build_comparison(chapters_a, chapters_b, ignore_linebreaks=ignore_linebreaks)
            advance("Vergleich berechnen")

            moves = None
            moves_complete = True
            if detect_moves:
                moves, moves_complete = detect_possible_moves(rows)
                advance("Verschiebungen erkennen")

            stats = compute_stats(chapters_a, chapters_b, rows)
            out_html = render_html(
                rows, stats, path_a.name, path_b.name, ignore_linebreaks=ignore_linebreaks,
                meta_a=doc_metadata(path_a), meta_b=doc_metadata(path_b),
                pages_method=pages_method, diagnostics=diagnostics, moves=moves, moves_complete=moves_complete,
            )
            out_path.parent.mkdir(parents=True, exist_ok=True)
            out_path.write_text(out_html, encoding="utf-8")
            advance("Report schreiben")

            docx_path = None
            if export_docx:
                docx_path = out_path.with_suffix(".docx")
                generate_docx_report(
                    rows, stats, path_a.name, path_b.name, docx_path,
                    meta_a=doc_metadata(path_a), meta_b=doc_metadata(path_b),
                )
                advance("Word-Report erzeugen")

            # Konfiguration fuer naechsten Start merken
            save_config({
                "doc_a": str(path_a), "doc_b": str(path_b), "out": str(out_path),
                "ignore_linebreaks": ignore_linebreaks, "detect_pages": detect_pages,
                "detect_moves": detect_moves, "export_docx": export_docx,
            })

            self.after(0, lambda: self._on_success(out_path, stats, docx_path,
                                                     len(moves) if moves else 0, moves_complete))
        except Exception as exc:  # noqa: BLE001 - Fehler dem Nutzer anzeigen statt zu verschlucken
            self.after(0, lambda: self._on_error(exc))

    def _on_success(self, out_path, stats, docx_path=None, move_count=0, moves_complete=True):
        self.btn_compare.config(state="normal")
        status_text = f"Report erstellt: {out_path}"
        if docx_path:
            status_text += f"  |  Word-Report: {docx_path}"
        self.var_status.set(status_text)
        move_note = ""
        if move_count:
            move_note = f"  🔀 {move_count} mögliche Verschiebung(en) erkannt."
        if not moves_complete:
            move_note += "  ⚠ Verschiebungs-Erkennung wegen Zeitbudget unvollständig."
        self.progress_label.config(text="Fertig." + move_note)
        self.stats_label.config(
            text=(
                f"Kapitel A: {stats['total_a']}   Kapitel B: {stats['total_b']}   "
                f"Unverändert: {stats['unchanged']}   Geändert: {stats['changed']}   "
                f"Neu: {stats['new']}   Gelöscht: {stats['deleted']}"
            )
        )
        opened, method = open_in_browser(out_path)
        if not opened:
            messagebox.showinfo(
                "Bericht bereit",
                f"Der Bericht wurde erstellt, konnte aber nicht automatisch geöffnet werden:\n\n"
                f"{out_path}\n\nBitte die Datei manuell doppelklicken oder in einen Browser ziehen.",
            )
        elif method == "os_default":
            messagebox.showwarning(
                "Bericht geöffnet über Windows-Standard",
                "Weder über die Windows-Registrierung noch über bekannte Installationspfade "
                "konnte Edge/Chrome/Firefox gefunden werden. Der Bericht wurde stattdessen über "
                "die Windows-Standardzuordnung geöffnet - falls sich dabei ein alter/falscher "
                "Browser (z.B. Internet Explorer) geöffnet hat und die Seite nur als Rohtext "
                "zeigt, bitte die Datei manuell mit Edge/Chrome öffnen:\n\n"
                f"{out_path}",
            )

    def _on_error(self, exc):
        self.btn_compare.config(state="normal")
        self.var_status.set("Fehler beim Vergleich.")
        self.progress_label.config(text="Abgebrochen.")
        messagebox.showerror("Fehler", f"Beim Vergleich ist ein Fehler aufgetreten:\n\n{exc}")


def main():
    app = CompareApp()
    app.mainloop()


if __name__ == "__main__":
    main()
