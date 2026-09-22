#!/usr/bin/env python3
"""
docx_chapter_compare_gui.py
Version 2.8 / 2026-09-22 / Grund: Update passend zu docx_chapter_compare.py Version 3.11
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

GUI_VERSION = "2.8"

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
    print(f"Fehler beim Import von docx_chapter_compare.py: {exc}", file=sys.stderr)
    print("Pruefe, ob 'python-docx' installiert ist (pip install python-docx) "
          "und ob docx_chapter_compare.py im selben Ordner liegt.", file=sys.stderr)
    raise

CORE_VERSION = getattr(_core, "SCRIPT_VERSION", "?")
CORE_PATH = getattr(_core, "__file__", "?")
GUI_PATH = str(Path(__file__).resolve())

print(f"docx_chapter_compare_gui.py Version {GUI_VERSION} ({GUI_PATH})")
print(f"docx_chapter_compare.py     Version {CORE_VERSION} ({CORE_PATH})")

CONFIG_PATH = Path.home() / ".docx_chapter_compare_gui.json"

def default_output_dir():
    docs = Path.home() / "Documents"
    return docs if docs.is_dir() else Path.home()

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
        pass


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
        self._show_splash_and_start()


    def _show_splash_and_start(self):
        self.withdraw()  
        
        splash = tk.Toplevel(self)
        splash.overrideredirect(True)
        splash.attributes('-topmost', True)
        
        width, height = 450, 300
        x = (splash.winfo_screenwidth() - width) // 2
        y = (splash.winfo_screenheight() - height) // 2
        splash.geometry(f"{width}x{height}+{x}+{y}")
        splash.configure(background="white")
        
        frame = tk.Frame(splash, bg="white")
        frame.pack(fill="both", expand=True)
        
        script_dir = Path(__file__).resolve().parent
        logo_path = script_dir / "FirmenLogo.JPG"
        logo_loaded = False
        
        self._splash_label = tk.Label(frame, bg="white")
        self._splash_label.place(x=0, y=0, width=width, height=height)
        
        version_label = tk.Label(
            frame, 
            text=f"Lade Version {GUI_VERSION} (Core {CORE_VERSION})...", 
            font=("Arial", 10, "bold"), 
            bg="#555555",
            fg="white",
            padx=10,
            pady=4
        )
        version_label.place(relx=0.5, rely=0.9, anchor="center")
        
        if logo_path.exists():
            try:
                from PIL import Image, ImageTk, ImageDraw
                import math  
                
                img = Image.open(logo_path).convert("RGBA")
                resample_filter = getattr(Image, 'Resampling', Image).LANCZOS 
                img = img.resize((width, height), resample_filter)
                w, h = img.size
                
                spot_size = int(max(w, h) * 0.8)  
                spot = Image.new("RGBA", (spot_size, spot_size), (255, 255, 255, 0))
                draw = ImageDraw.Draw(spot)
                
                cx, cy = spot_size // 2, spot_size // 2
                radius = spot_size // 2
                
                for r_step in range(radius, 0, -2):
                    dist = r_step / radius
                    alpha = int(170 * (1 - dist**2))  
                    draw.ellipse(
                        [(cx - r_step, cy - r_step), (cx + r_step, cy + r_step)], 
                        fill=(255, 255, 255, alpha)
                    )
                
                self._anim_frame = 0
                self._anim_max_frames = 100  
                
                def update_animation():
                    if not splash.winfo_exists():
                        return
                        
                    progress = (self._anim_frame % self._anim_max_frames) / self._anim_max_frames
                    
                    if progress < 0.65:  
                        eff_progress = progress / 0.65
                        current_x = int(-spot_size + eff_progress * (w + spot_size))
                        base_y = (h - spot_size) / 2
                        current_y = int(base_y + math.sin(eff_progress * math.pi) * (h * 0.15))
                        
                        overlay = Image.new("RGBA", (w, h), (255, 255, 255, 0))
                        overlay.paste(spot, (current_x, current_y), spot)
                        
                        composite = Image.alpha_composite(img, overlay)
                        photo = ImageTk.PhotoImage(composite)
                    else:
                        photo = ImageTk.PhotoImage(img)
                        
                    self._splash_label.config(image=photo)
                    self._splash_label.image = photo 
                    
                    self._anim_frame += 1
                    splash.after(33, update_animation) 
                
                update_animation()
                logo_loaded = True
                
            except ImportError:
                self._splash_label.config(text="FirmenLogo.JPG gefunden, aber 'Pillow' fehlt!\n'pip install Pillow'", fg="red")
            except Exception as e:
                self._splash_label.config(text=f"Fehler beim Laden:\n{e}", fg="red")
        else:
            self._splash_label.config(text="Kein FirmenLogo.JPG im Ordner gefunden.", fg="gray")
            
        if not logo_loaded:
            fallback_label = tk.Label(frame, text="Kapitelvergleich Tool", font=("Arial", 16, "bold"), bg="white")
            fallback_label.place(relx=0.5, rely=0.4, anchor="center")
            
        self.after(5000, lambda: self._close_splash(splash))
 
 
    def _close_splash(self, splash):
        splash.destroy()
        self.deiconify() 

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
            frame, text="Zusätzlich Word-Report erzeugen (kompakt, für schnelle Weitergabe)",
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

    def run_compare(self):
        path_a = Path(self.var_a.get().strip())
        path_b = Path(self.var_b.get().strip())
        out_path = Path(self.var_out.get().strip() or (default_output_dir() / "vergleich.html"))
        
        if out_path.suffix.lower() != ".html":
            out_path = out_path.with_name(out_path.name + ".html") if out_path.suffix else out_path.with_suffix(".html")
            self.var_out.set(str(out_path))

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
            toc_info = None
            if detect_pages:
                diagnostics = _core.diagnose_page_detection()
                advance("Preflight-Check Seiten-Gruppierung")
                method, toc_info = attach_pages(chapters_a, chapters_b, path_a, path_b)
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
                toc_info=toc_info,
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

            save_config({
                "doc_a": str(path_a), "doc_b": str(path_b), "out": str(out_path),
                "ignore_linebreaks": ignore_linebreaks, "detect_pages": detect_pages,
                "detect_moves": detect_moves, "export_docx": export_docx,
            })

            self.after(0, lambda: self._on_success(out_path, stats, docx_path,
                                                     len(moves) if moves else 0, moves_complete))
        except Exception as exc: 
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
                "Browser nicht erkannt. Falls sich der Bericht nicht korrekt öffnet, ziehe ihn manuell in Microsoft Edge oder Chrome:\n\n"
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