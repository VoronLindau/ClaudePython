#!/usr/bin/env python3
"""
docx_chapter_compare_gui.py
Version 1.6 / 2026-09-05 / Grund: Checkbox "Zusätzlich Word-Report erzeugen"
    ergänzt (nutzt generate_docx_report aus docx_chapter_compare.py v2.4) -
    Ausgabe landet neben dem HTML-Report mit gleichem Namen, .docx-Endung.

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
import sys
import threading
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

GUI_VERSION = "1.6"

try:
    import docx_chapter_compare as _core
    from docx_chapter_compare import (
        attach_pages,
        build_comparison,
        compute_stats,
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
        self.geometry("640x460")
        self.minsize(560, 390)
        self.resizable(True, False)

        self.cfg = load_config()

        self.var_a = tk.StringVar(value=self.cfg.get("doc_a", ""))
        self.var_b = tk.StringVar(value=self.cfg.get("doc_b", ""))
        self.var_out = tk.StringVar(value=self.cfg.get("out", str(Path.home() / "vergleich.html")))
        self.var_status = tk.StringVar(value="Bereit.")
        self.var_ignore_linebreaks = tk.BooleanVar(value=self.cfg.get("ignore_linebreaks", True))
        self.var_detect_pages = tk.BooleanVar(value=self.cfg.get("detect_pages", True))
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
            frame, text="Zusätzlich Word-Report erzeugen (kompakt, für schnelle Weitergabe im Unternehmen)",
            variable=self.var_export_docx,
        ).grid(row=8, column=0, columnspan=3, sticky="w", padx=10, pady=(2, 0))

        frame.columnconfigure(1, weight=1)

        self.btn_compare = ttk.Button(frame, text="Vergleichen ▶", command=self.run_compare)
        self.btn_compare.grid(row=9, column=0, padx=10, pady=16, sticky="w")

        self.progress = ttk.Progressbar(frame, mode="indeterminate")
        self.progress.grid(row=9, column=1, columnspan=2, padx=10, pady=16, sticky="ew")

        self.stats_label = ttk.Label(frame, text="", justify="left")
        self.stats_label.grid(row=10, column=0, columnspan=3, sticky="w", padx=10)

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
        out_path = Path(self.var_out.get().strip() or (Path.home() / "vergleich.html"))

        if not self.var_a.get().strip() or not self.var_b.get().strip():
            messagebox.showwarning("Fehlende Angabe", "Bitte beide Dokumente auswählen.")
            return
        if not path_a.exists():
            messagebox.showerror("Datei nicht gefunden", f"Dokument A nicht gefunden:\n{path_a}")
            return
        if not path_b.exists():
            messagebox.showerror("Datei nicht gefunden", f"Dokument B nicht gefunden:\n{path_b}")
            return

        self.btn_compare.config(state="disabled")
        self.progress.start(12)
        self.var_status.set("Vergleiche Dokumente …")
        self.stats_label.config(text="")

        ignore_linebreaks = self.var_ignore_linebreaks.get()
        detect_pages = self.var_detect_pages.get()
        export_docx = self.var_export_docx.get()

        thread = threading.Thread(
            target=self._worker,
            args=(path_a, path_b, out_path, ignore_linebreaks, detect_pages, export_docx),
            daemon=True,
        )
        thread.start()

    def _worker(self, path_a, path_b, out_path, ignore_linebreaks, detect_pages, export_docx):
        try:
            chapters_a = extract_chapters(path_a)
            chapters_b = extract_chapters(path_b)

            if not chapters_a or not chapters_b:
                self.after(0, lambda: messagebox.showwarning(
                    "Keine Kapitel erkannt",
                    "In mindestens einem Dokument wurden keine Kapitel/Anforderungen erkannt.\n"
                    "Ergebnis kann unvollständig sein.",
                ))

            pages_method = None
            diagnostics = None
            if detect_pages:
                self.after(0, lambda: self.var_status.set("Preflight-Check Seiten-Gruppierung …"))
                diagnostics = _core.diagnose_page_detection()
                self.after(0, lambda: self.var_status.set("Ermittle Seiten (MS Word, sonst LibreOffice) …"))
                method = attach_pages(chapters_a, chapters_b, path_a, path_b)
                pages_method = method if method is not None else "unavailable"
                self.after(0, lambda: self.var_status.set("Vergleiche Dokumente …"))

            rows = build_comparison(chapters_a, chapters_b, ignore_linebreaks=ignore_linebreaks)
            stats = compute_stats(chapters_a, chapters_b, rows)
            out_html = render_html(
                rows, stats, path_a.name, path_b.name, ignore_linebreaks=ignore_linebreaks,
                meta_a=doc_metadata(path_a), meta_b=doc_metadata(path_b),
                pages_method=pages_method, diagnostics=diagnostics,
            )
            out_path.parent.mkdir(parents=True, exist_ok=True)
            out_path.write_text(out_html, encoding="utf-8")

            docx_path = None
            if export_docx:
                self.after(0, lambda: self.var_status.set("Erzeuge Word-Report …"))
                docx_path = out_path.with_suffix(".docx")
                generate_docx_report(
                    rows, stats, path_a.name, path_b.name, docx_path,
                    meta_a=doc_metadata(path_a), meta_b=doc_metadata(path_b),
                )

            # Konfiguration fuer naechsten Start merken
            save_config({
                "doc_a": str(path_a), "doc_b": str(path_b), "out": str(out_path),
                "ignore_linebreaks": ignore_linebreaks, "detect_pages": detect_pages,
                "export_docx": export_docx,
            })

            self.after(0, lambda: self._on_success(out_path, stats, docx_path))
        except Exception as exc:  # noqa: BLE001 - Fehler dem Nutzer anzeigen statt zu verschlucken
            self.after(0, lambda: self._on_error(exc))

    def _on_success(self, out_path, stats, docx_path=None):
        self.progress.stop()
        self.btn_compare.config(state="normal")
        status_text = f"Report erstellt: {out_path}"
        if docx_path:
            status_text += f"  |  Word-Report: {docx_path}"
        self.var_status.set(status_text)
        self.stats_label.config(
            text=(
                f"Kapitel A: {stats['total_a']}   Kapitel B: {stats['total_b']}   "
                f"Unverändert: {stats['unchanged']}   Geändert: {stats['changed']}   "
                f"Neu: {stats['new']}   Gelöscht: {stats['deleted']}"
            )
        )
        webbrowser.open(out_path.resolve().as_uri())

    def _on_error(self, exc):
        self.progress.stop()
        self.btn_compare.config(state="normal")
        self.var_status.set("Fehler beim Vergleich.")
        messagebox.showerror("Fehler", f"Beim Vergleich ist ein Fehler aufgetreten:\n\n{exc}")


def main():
    app = CompareApp()
    app.mainloop()


if __name__ == "__main__":
    main()
