"""
aas_pcf_query.py
================
Liest WAGO-Artikelnummern aus einer Excel-Datei, fragt fuer jede Nummer den
PCF-Wert (Product Carbon Footprint, kg CO2e) aus dem WAGO AAS Portal ab
und schreibt die Ergebnisse in eine neue Excel-Datei.

Ausfuehren:
    python aas_pcf_query.py

Technische Grundlage:
    Das WAGO AAS Portal stellt fuer jede Artikelnummer ein CarbonFootprint-
    Submodel bereit. Die URL laesst sich direkt aus der Artikelnummer
    konstruieren:

        AAS-Submodel-ID : https://wago.com/ids/submodel/{article}/PCF
        Endpoint-URL    : https://c1.api.wago.com/smartdata-aas-env/submodels/{base64(ID)}

    Der PCF-Wert befindet sich im Feld PcfCO2eq innerhalb der
    SubmodelElementList ProductCarbonFootprints.

Performance-Hinweise (2026-08):
    Benchmark mit 60 Artikeln ergab zwei Klassen langsamer Requests:
    1. ConnectionError-Artikel (~20s): Server antwortet nicht → Max-Retries erschoepft.
       Ursache: Bestimmte Artikelnummern loesen serverseitig keine Antwort aus.
    2. ReadTimeout-dann-Retry-Artikel (~10s): Erster Request laueft auf Timeout,
       zweiter Retry ist sofort erfolgreich (Server-seitiges Cache-Warmup).

    Optimierungen:
    - REQUEST_TIMEOUT (5, 10) statt (10, 30): Schlaegt bei Timeout 3x schneller fehl.
    - Retry(total=2, read=1, backoff_factor=0.2): Weniger Wartezeit zwischen Versuchen.
    - Thread-lokale Sessions: Keine geteilten Verbindungspools zwischen Worker-Threads.
    - ThreadPoolExecutor (WORKERS=5): Parallele Verarbeitung mehrerer Artikel.
    - SLOW_REQUEST_THRESHOLD: Logging langsamer Requests zur Diagnose.
"""

import base64
import os
import subprocess
import threading
import time
import tkinter as tk
from concurrent.futures import ThreadPoolExecutor, as_completed
from tkinter import filedialog, messagebox, ttk

import pandas as pd
import requests
import urllib3

# ── Konfiguration ────────────────────────────────────────────────────────────

ARTICLE_COLUMN = "Article Number"   # Spaltenname mit Artikelnummern in der Excel-Datei

# WAGO AAS Environment Basis-URL
AAS_ENV_BASE = "https://c1.api.wago.com/smartdata-aas-env"

# TLS-Pruefung deaktivieren (WAGO Portal nutzt selbstsignierte Zertifikate)
CERT_VERIFY = False

# Timeout in Sekunden (connect, read)
# Begruendung: Benchmark zeigt bimodale Verteilung: schnell (<0.3s) oder Timeout.
# 10s read-timeout triggert Retry schnell genug ohne echte Antworten zu verpassen.
REQUEST_TIMEOUT = (5, 10)

# Parallele Worker-Threads fuer AAS-Abfragen.
# 5 ist ein guter Kompromiss zwischen Geschwindigkeit und API-Belastung.
WORKERS = 5

# Schwellwert fuer Slow-Query-Logging (Sekunden)
SLOW_REQUEST_THRESHOLD = 3.0

# Status-Konstanten
STATUS_PCF_AVAILABLE = "PCF available"
STATUS_AAS_NO_PCF    = "AAS available - PCF not available"
STATUS_NOT_FOUND     = "Article not found"
STATUS_QUERY_ERROR   = "Query error"

# Sentinel-Werte fuer _jget
_NOT_FOUND = object()
_ERROR     = object()

# ── HTTP-Setup ───────────────────────────────────────────────────────────────

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Thread-lokaler Speicher: jeder Worker-Thread bekommt seine eigene Session.
# Begruendung: requests.Session ist nicht vollstaendig thread-safe bei
# gleichzeitiger Verwendung des Verbindungspools aus mehreren Threads.
_thread_local = threading.local()


def _get_session():
    """
    Gibt eine thread-lokale requests.Session mit Retry-Logik zurueck.

    Jeder Thread bekommt seine eigene Session, um Verbindungspool-Konflikte
    bei paralleler Verarbeitung zu vermeiden.
    """
    if not hasattr(_thread_local, "session") or _thread_local.session is None:
        from requests.adapters import HTTPAdapter, Retry
        s = requests.Session()
        retry = Retry(
            total=2,               # Max. 2 Versuche gesamt (war 3)
            connect=2,
            read=1,                # 1 Read-Retry (war 3) - reduziert Timeout-Wartezeit
            backoff_factor=0.2,    # Wartezeit zwischen Retries (war 0.5)
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=["GET"],
        )
        s.mount("https://", HTTPAdapter(max_retries=retry))
        s.headers.update({
            "User-Agent": "WAGO-AAS-PCF-Query/1.0",
            "Accept": "application/json",
        })
        s.verify = CERT_VERIFY
        _thread_local.session = s
    return _thread_local.session


def _jget(url, timeout=REQUEST_TIMEOUT):
    """
    HTTP GET -> JSON-Dict, _NOT_FOUND (404), oder _ERROR (techn. Fehler).

    Unterscheidung:
        404              -> _NOT_FOUND  (Artikel im AAS nicht vorhanden)
        anderer Fehler   -> _ERROR      (technischer Fehler, keine sichere Aussage)
        Erfolg           -> dict        (JSON-Antwort)
    """
    try:
        r = _get_session().get(url, timeout=timeout)
        if r.status_code == 404:
            return _NOT_FOUND
        r.raise_for_status()
        return r.json()
    except requests.exceptions.HTTPError as e:
        print(f"  Warning: HTTP {e.response.status_code} fuer {url}")
        return _ERROR
    except requests.RequestException as e:
        print(f"  Warning: {type(e).__name__}: {str(e)[:100]}")
        return _ERROR


# ── AAS-Hilfsfunktionen ──────────────────────────────────────────────────────

def _submodel_url(article_number):
    """
    Konstruiert die direkte AAS-Submodel-URL fuer das CarbonFootprint-Submodel
    der gegebenen Artikelnummer.

    URL-Muster:
        ID      : https://wago.com/ids/submodel/{article}/PCF
        Endpoint: {AAS_ENV_BASE}/submodels/{base64url(ID)}
    """
    submodel_id = f"https://wago.com/ids/submodel/{article_number}/PCF"
    encoded = base64.b64encode(submodel_id.encode()).decode()
    return f"{AAS_ENV_BASE}/submodels/{encoded}"


def _extract_all_values(elements, path=""):
    """
    Traversiert rekursiv alle submodelElements (inkl. SubmodelElementList und
    SubmodelElementCollection) und gibt ein Dict {Pfad: Wert} zurueck.

    Kompatibel mit AAS v3 (SubmodelElementList ohne idShort bei Listenelementen).
    """
    result = {}
    for el in (elements or []):
        idshort  = el.get("idShort") or ""
        new_path = f"{path}/{idshort}".lstrip("/") if path else idshort

        # Kinder: submodelElements (Submodel-Ebene) oder value (Collection/List)
        children = el.get("submodelElements")
        if children is None and isinstance(el.get("value"), list):
            children = el["value"]

        if children:
            result.update(_extract_all_values(children, new_path))
        else:
            result[new_path] = el.get("value")

    return result


def _parse_float(raw):
    """Parst einen Zahlenwert; akzeptiert Komma als Dezimaltrenner."""
    if raw is None:
        return None
    s = str(raw).strip().replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


# ── Kernfunktion ─────────────────────────────────────────────────────────────

def _evaluate_pcf_result(result):
    if result is _NOT_FOUND:
        return None, STATUS_NOT_FOUND
    if result is _ERROR or not isinstance(result, dict):
        return None, STATUS_QUERY_ERROR
        
    all_values = _extract_all_values(result.get("submodelElements", []))
    for path, value in all_values.items():
        if "pcfco2eq" in path.lower():
            parsed = _parse_float(value)
            if parsed is not None:
                return parsed, STATUS_PCF_AVAILABLE
    return None, STATUS_AAS_NO_PCF

def get_pcf_result(article_number):
    """
    Gibt ein Dict mit PCF-Wert und Status fuer die gegebene WAGO-Artikelnummer zurueck.
    Nutzt 'Fast Path' + 'Safe Fallback'.
    """
    article_number = str(article_number).strip()
    
    # 1. Fast Path
    url = _submodel_url(article_number)
    fast_result = _jget(url)
    
    pcf_val, pcf_status = _evaluate_pcf_result(fast_result)
    if pcf_status == STATUS_PCF_AVAILABLE:
        return {"pcf": pcf_val, "status": pcf_status}
        
    # 2. Safe Fallback
    import json
    import urllib.parse
    asset_id = f"https://wago.com/ids/asset/{article_number}"
    asset_id_encoded = base64.b64encode(json.dumps([{"name": "globalAssetId", "value": asset_id}]).encode()).decode()
    reg_url = f"https://c1.api.wago.com/smartdata-aas-registry/shell-descriptors?assetIds={urllib.parse.quote(asset_id_encoded)}"
    
    # Registry is slow, use 30s timeout
    reg_res = _jget(reg_url, timeout=(10, 30))
    if isinstance(reg_res, dict) and "result" in reg_res:
        shells = reg_res.get("result", [])
        for shell in shells:
            submodels = shell.get("submodelDescriptors", [])
            for sm in submodels:
                id_short = (sm.get("idShort") or "").lower()
                if "pcf" in id_short or "carbonfootprint" in id_short:
                    sm_id = sm.get("id")
                    if sm_id:
                        encoded_sm = base64.b64encode(sm_id.encode()).decode()
                        sm_url = f"{AAS_ENV_BASE}/submodels/{encoded_sm}"
                        sm_res = _jget(sm_url, timeout=(10, 30))
                        fb_val, fb_status = _evaluate_pcf_result(sm_res)
                        if fb_status == STATUS_PCF_AVAILABLE:
                            return {"pcf": fb_val, "status": fb_status}
                            
    # Fallback hat auch nichts gefunden, urspruengliches Fast-Path-Ergebnis zurueckgeben
    return {"pcf": pcf_val, "status": pcf_status}


def get_pcf(article_number):
    """
    Rueckwaerts-kompatibler Wrapper: gibt nur den PCF-Wert (float | None) zurueck.
    """
    return get_pcf_result(article_number)["pcf"]


def _query_with_timing(article_number):
    """
    Fuehrt get_pcf_result() aus und loggt langsame Requests.
    Geeignet fuer Worker-Threads.
    """
    t0 = time.perf_counter()
    try:
        res = get_pcf_result(article_number)
    except Exception as e:
        print(f"  Unerwarteter Fehler fuer {article_number}: {e}")
        res = {"pcf": None, "status": STATUS_QUERY_ERROR}
    dur = time.perf_counter() - t0
    if dur >= SLOW_REQUEST_THRESHOLD:
        print(f"  Slow query: {article_number} - {dur:.1f}s [{res['status']}]")
    return res


# ── Hauptprogramm (Batch, ohne GUI) ──────────────────────────────────────────

def main(input_file, output_file, progress_callback=None):
    """
    Verarbeitet eine Excel-Datei mit WAGO-Artikelnummern.

    Parameter:
        input_file        : Pfad zur Eingabedatei (str)
        output_file       : Pfad zur Ausgabedatei (str)
        progress_callback : Optionale Funktion(current, total, article, result_dict)
                            Wird nach jeder verarbeiteten Artikelnummer aufgerufen.

    Rueckgabe:
        dict mit Zusammenfassung oder None bei Fehler
    """
    print(f"Lese {input_file} ...")
    try:
        df = pd.read_excel(input_file, dtype=str)
    except FileNotFoundError:
        print(f"Fehler: Datei nicht gefunden: {input_file}")
        return None
    except Exception as e:
        print(f"Fehler beim Lesen von {input_file}: {e}")
        return None

    if ARTICLE_COLUMN not in df.columns:
        print(f"Fehler: Spalte '{ARTICLE_COLUMN}' nicht in {input_file} gefunden.")
        return None

    df[ARTICLE_COLUMN] = df[ARTICLE_COLUMN].fillna("").astype(str).str.strip()

    articles_all    = df[ARTICLE_COLUMN].tolist()
    unique_articles = list(dict.fromkeys(a for a in articles_all if a))
    unique_total    = len(unique_articles)
    total           = len(articles_all)

    print(f"Gefunden: {total} Zeilen, {unique_total} eindeutige Artikelnummern")
    print(f"Starte Abfrage mit {WORKERS} parallelen Threads ...\n")

    cache      = {}
    cache_lock = threading.Lock()
    done_count = [0]

    def _worker(article):
        res = _query_with_timing(article)
        with cache_lock:
            cache[article] = res
            done_count[0] += 1
            done = done_count[0]
        if progress_callback:
            progress_callback(done, unique_total, article, res)
        return article, res

    with ThreadPoolExecutor(max_workers=WORKERS) as executor:
        futures = {executor.submit(_worker, a): a for a in unique_articles}
        for future in as_completed(futures):
            article, res = future.result()
            pcf_str = f"{res['pcf']} kg CO2e" if res["pcf"] is not None else "not available"
            print(f"  {article:<25} {pcf_str:<20} [{res['status']}]")

    # Ergebnislisten in Originalreihenfolge aufbauen
    pcf_values   = []
    pcf_statuses = []
    for article in articles_all:
        if not article:
            pcf_values.append("not available")
            pcf_statuses.append("")
        else:
            res = cache[article]
            pcf_values.append(res["pcf"] if res["pcf"] is not None else "not available")
            pcf_statuses.append(res["status"])

    df["PCF"]        = pcf_values
    df["PCF Status"] = pcf_statuses

    try:
        df.to_excel(output_file, index=False)
    except Exception as e:
        print(f"\nFehler beim Schreiben von {output_file}: {e}")
        return None

    counts = {s: 0 for s in [STATUS_PCF_AVAILABLE, STATUS_AAS_NO_PCF,
                               STATUS_NOT_FOUND, STATUS_QUERY_ERROR]}
    for s in pcf_statuses:
        if s in counts:
            counts[s] += 1

    summary = {
        "total_rows":     total,
        "unique_queried": unique_total,
        "pcf_available":  counts[STATUS_PCF_AVAILABLE],
        "aas_no_pcf":     counts[STATUS_AAS_NO_PCF],
        "not_found":      counts[STATUS_NOT_FOUND],
        "query_errors":   counts[STATUS_QUERY_ERROR],
        "output_file":    output_file,
    }

    print(f"""
Finished.

Rows total       : {total}
Unique queried   : {unique_total}
PCF available    : {summary['pcf_available']}
AAS - no PCF     : {summary['aas_no_pcf']}
Not found        : {summary['not_found']}
Query errors     : {summary['query_errors']}

Output written to:
{output_file}
""")

    return summary


# ── GUI ───────────────────────────────────────────────────────────────────────

class PCFQueryApp(tk.Tk):
    """
    Vollstaendige Tkinter-GUI fuer das WAGO AAS PCF Query Tool.
    Verwendet ThreadPoolExecutor fuer parallele AAS-Abfragen.
    """

    C_BG      = "#1e2228"
    C_PANEL   = "#262c35"
    C_BORDER  = "#353d47"
    C_ACCENT  = "#8cc63f"
    C_ACCENT_HOVER = "#9dd94f"
    C_SUCCESS = "#4caf73"
    C_WARN    = "#f0a500"
    C_ERROR   = "#e05252"
    C_TEXT    = "#e8eaed"
    C_MUTED   = "#8a9bb0"
    C_BTN_FG  = "#ffffff"

    FONT_HEAD  = ("Segoe UI", 13, "bold")
    FONT_LABEL = ("Segoe UI", 9)
    FONT_BOLD  = ("Segoe UI", 9, "bold")
    FONT_MONO  = ("Consolas", 9)
    FONT_SMALL = ("Segoe UI", 8)

    def __init__(self):
        super().__init__()
        self.title("WAGO AAS PCF Query")
        self.configure(bg=self.C_BG)
        self.resizable(False, False)

        self._input_path   = None
        self._output_path  = None
        self._df           = None
        self._articles     = []
        self._query_thread = None
        self._stop_event   = threading.Event()

        self._build_ui()
        self._center_window()

    # ── UI-Aufbau ─────────────────────────────────────────────────────────────

    def _build_ui(self):
        tk.Frame(self, bg=self.C_ACCENT, height=4).pack(fill="x")

        title_frame = tk.Frame(self, bg=self.C_BG)
        title_frame.pack(fill="x", padx=16, pady=(12, 4))
        
        text_frame = tk.Frame(title_frame, bg=self.C_BG)
        text_frame.pack(side="left")
        
        tk.Label(
            text_frame, text="WAGO AAS PCF Query",
            font=("Segoe UI", 16, "bold"),
            fg=self.C_ACCENT, bg=self.C_BG,
        ).pack(anchor="w")
        tk.Label(
            text_frame, text="Product Carbon Footprint – AAS Portal Query Tool",
            font=self.FONT_SMALL, fg=self.C_MUTED, bg=self.C_BG,
        ).pack(anchor="w")
        
        import sys
        import os
        from PIL import Image, ImageTk
        
        def get_resource_path(relative_path):
            if hasattr(sys, '_MEIPASS'):
                return os.path.join(sys._MEIPASS, relative_path)
            return os.path.join(os.path.abspath(os.path.dirname(__file__)), relative_path)
            
        try:
            logo_path = get_resource_path("WAGO_LOGO.webp")
            if os.path.exists(logo_path):
                img = Image.open(logo_path)
                h_new = 36
                w_new = int(img.width * (h_new / img.height))
                img = img.resize((w_new, h_new), Image.Resampling.LANCZOS)
                self._logo_img = ImageTk.PhotoImage(img)
                tk.Label(title_frame, image=self._logo_img, bg=self.C_BG).pack(side="right", anchor="e")
        except Exception as e:
            print(f"Logo not loaded: {e}")

        self._divider()

        # ── Dateiauswahl ──────────────────────────────────────────────────────
        self._section("Excel File")
        btn_row = tk.Frame(self, bg=self.C_BG)
        btn_row.pack(fill="x", padx=16, pady=(2, 4))
        self._btn_select = self._make_button(
            btn_row, "  Select Excel File", self._on_select_file
        )
        self._btn_select.pack(side="left")

        self._lbl_path = tk.Label(
            self, text="No file selected.",
            font=self.FONT_MONO, fg=self.C_MUTED, bg=self.C_BG,
            wraplength=460, justify="left",
        )
        self._lbl_path.pack(anchor="w", padx=16, pady=(0, 4))

        # ── File-Zusammenfassung ──────────────────────────────────────────────
        self._divider()
        self._section("File Summary")
        frm_summary = self._make_panel()
        frm_summary.pack(fill="x", padx=16, pady=(2, 4))
        self._lbl_summary = tk.Label(
            frm_summary, text="—",
            font=self.FONT_MONO, fg=self.C_MUTED, bg=self.C_PANEL,
            justify="left", anchor="w",
        )
        self._lbl_summary.pack(anchor="w", padx=10, pady=6)

        # ── Vorschau ──────────────────────────────────────────────────────────
        self._divider()
        self._section("Preview  (first 10 article numbers)")
        frm_prev = self._make_panel()
        frm_prev.pack(fill="x", padx=16, pady=(2, 4))
        self._lbl_preview = tk.Label(
            frm_prev, text="—",
            font=self.FONT_MONO, fg=self.C_MUTED, bg=self.C_PANEL,
            justify="left", anchor="w",
        )
        self._lbl_preview.pack(anchor="w", padx=10, pady=6)

        # ── Start-Button ──────────────────────────────────────────────────────
        self._divider()
        btn_start_frame = tk.Frame(self, bg=self.C_BG)
        btn_start_frame.pack(fill="x", padx=16, pady=8)
        self._btn_start = self._make_button(
            btn_start_frame, "  Start PCF Query", self._on_start_query, accent=True
        )
        self._btn_start.pack(side="left")
        self._btn_start.config(state="disabled")

        # ── Fortschritt ───────────────────────────────────────────────────────
        self._divider()
        self._section("Progress")
        frm_prog = tk.Frame(self, bg=self.C_BG)
        frm_prog.pack(fill="x", padx=16, pady=(2, 4))

        self._progress_var = tk.DoubleVar(value=0)
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(
            "WAGO.Horizontal.TProgressbar",
            troughcolor=self.C_BORDER,
            background=self.C_ACCENT,
            bordercolor=self.C_PANEL,
            lightcolor=self.C_ACCENT,
            darkcolor=self.C_ACCENT,
        )
        ttk.Progressbar(
            frm_prog,
            variable=self._progress_var,
            maximum=100,
            style="WAGO.Horizontal.TProgressbar",
            length=460,
        ).pack(pady=(0, 4))

        self._lbl_prog_count = tk.Label(
            frm_prog, text="—",
            font=self.FONT_LABEL, fg=self.C_MUTED, bg=self.C_BG,
        )
        self._lbl_prog_count.pack(anchor="w")

        self._lbl_prog_article = tk.Label(
            frm_prog, text="",
            font=self.FONT_MONO, fg=self.C_TEXT, bg=self.C_BG,
        )
        self._lbl_prog_article.pack(anchor="w")

        self._lbl_eta = tk.Label(
            frm_prog, text="",
            font=self.FONT_LABEL, fg=self.C_MUTED, bg=self.C_BG,
        )
        self._lbl_eta.pack(anchor="w", pady=(0, 4))

        # ── Ergebnisse ────────────────────────────────────────────────────────
        self._divider()
        self._section("Results")
        frm_res = self._make_panel()
        frm_res.pack(fill="x", padx=16, pady=(2, 4))

        self._result_rows = {}
        result_defs = [
            ("pcf_available", "PCF available:",          self.C_SUCCESS),
            ("aas_no_pcf",    "AAS available - no PCF:", self.C_WARN),
            ("not_found",     "Article not found:",      self.C_MUTED),
            ("query_errors",  "Query errors:",           self.C_ERROR),
        ]
        for key, label, color in result_defs:
            row = tk.Frame(frm_res, bg=self.C_PANEL)
            row.pack(fill="x", padx=10, pady=2)
            tk.Label(
                row, text=label, font=self.FONT_LABEL,
                fg=self.C_MUTED, bg=self.C_PANEL, width=28, anchor="w",
            ).pack(side="left")
            val_lbl = tk.Label(
                row, text="—", font=self.FONT_BOLD,
                fg=color, bg=self.C_PANEL, width=6, anchor="e",
            )
            val_lbl.pack(side="left")
            self._result_rows[key] = val_lbl
        tk.Frame(frm_res, bg=self.C_PANEL, height=4).pack()

        # ── Output-Buttons ────────────────────────────────────────────────────
        self._divider()
        frm_out = tk.Frame(self, bg=self.C_BG)
        frm_out.pack(fill="x", padx=16, pady=8)

        self._btn_open_folder = self._make_button(
            frm_out, "  Open output folder", self._on_open_folder
        )
        self._btn_open_folder.pack(side="left", padx=(0, 8))
        self._btn_open_folder.config(state="disabled")

        self._btn_open_excel = self._make_button(
            frm_out, "  Open Excel file", self._on_open_excel
        )
        self._btn_open_excel.pack(side="left")
        self._btn_open_excel.config(state="disabled")

        self._divider()
        frm_footer = tk.Frame(self, bg=self.C_BG)
        frm_footer.pack(fill="x", padx=16, pady=(0, 8))
        
        lbl_footer = tk.Label(
            frm_footer, text="Developed by Alexander Flekler",
            font=self.FONT_SMALL, fg=self.C_MUTED, bg=self.C_BG
        )
        lbl_footer.pack(side="left")
        
        btn_about = tk.Button(
            frm_footer, text="About", command=self._on_about,
            font=self.FONT_SMALL, fg=self.C_MUTED, bg=self.C_BG,
            relief="flat", cursor="hand2", bd=0, activebackground=self.C_BG, activeforeground=self.C_TEXT
        )
        btn_about.pack(side="right")
        btn_about.bind("<Enter>", lambda e: btn_about.config(fg=self.C_TEXT))
        btn_about.bind("<Leave>", lambda e: btn_about.config(fg=self.C_MUTED))

    # ── Widget-Helfer ─────────────────────────────────────────────────────────

    def _divider(self):
        tk.Frame(self, bg=self.C_BORDER, height=1).pack(fill="x", padx=8, pady=4)

    def _section(self, text):
        tk.Label(
            self, text=text.upper(),
            font=("Segoe UI", 7, "bold"), fg=self.C_MUTED, bg=self.C_BG,
        ).pack(anchor="w", padx=16, pady=(4, 0))

    def _make_panel(self):
        return tk.Frame(
            self, bg=self.C_PANEL, relief="flat",
            highlightbackground=self.C_BORDER, highlightthickness=1,
        )

    def _make_button(self, parent, text, command, accent=False):
        bg = self.C_ACCENT if accent else self.C_PANEL
        btn = tk.Button(
            parent, text=text, command=command,
            font=self.FONT_BOLD,
            bg=bg, fg=self.C_BTN_FG,
            activebackground=self.C_ACCENT_HOVER if accent else self.C_BORDER,
            activeforeground=self.C_BTN_FG,
            relief="flat", cursor="hand2",
            padx=12, pady=6,
            bd=0,
        )
        btn.bind("<Enter>", lambda e: btn.config(
            bg=self.C_ACCENT_HOVER if accent else self.C_BORDER))
        btn.bind("<Leave>", lambda e: btn.config(bg=bg))
        return btn

    def _center_window(self):
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")

    # ── Event-Handler ─────────────────────────────────────────────────────────

    def _on_about(self):
        from tkinter import messagebox
        messagebox.showinfo(
            "About",
            "WAGO AAS PCF Query\nDeveloped by Alexander Flekler\nCorporate Sustainability\nVersion 1.0"
        )

    def _on_select_file(self):
        path = filedialog.askopenfilename(
            title="WAGO PCF Query – Excel-Datei auswaehlen",
            filetypes=[("Excel-Dateien", "*.xlsx"), ("Alle Dateien", "*.*")],
        )
        if path:
            self._load_file(path)

    def _load_file(self, path):
        self._input_path = path
        self._btn_start.config(state="disabled")
        self._lbl_path.config(text=path, fg=self.C_TEXT)
        self._df = None

        try:
            df = pd.read_excel(path, dtype=str)
        except Exception as e:
            self._lbl_summary.config(
                text=f"Could not read file:\n{e}", fg=self.C_ERROR,
            )
            self._lbl_preview.config(text="—", fg=self.C_MUTED)
            return

        if ARTICLE_COLUMN not in df.columns:
            cols = ", ".join(df.columns.tolist())
            self._lbl_summary.config(
                text=(
                    f'Column "{ARTICLE_COLUMN}" not found in this file.\n'
                    f"Available columns: {cols}"
                ),
                fg=self.C_ERROR,
            )
            self._lbl_preview.config(text="—", fg=self.C_MUTED)
            return

        df[ARTICLE_COLUMN] = df[ARTICLE_COLUMN].fillna("").astype(str).str.strip()
        articles_all = df[ARTICLE_COLUMN].tolist()
        valid        = [a for a in articles_all if a]
        unique       = list(dict.fromkeys(valid))
        duplicates   = len(valid) - len(unique)
        empty_count  = len(articles_all) - len(valid)

        self._df       = df
        self._articles = articles_all

        stem = os.path.splitext(os.path.basename(path))[0]
        self._output_path = os.path.join(os.path.dirname(path), f"{stem}_PCF.xlsx")

        summary_text = (
            f"  Article column :  {ARTICLE_COLUMN}\n"
            f"  Rows           :  {len(articles_all)}\n"
            f"  Valid articles :  {len(valid)}\n"
            f"  Unique articles:  {len(unique)}\n"
            f"  Duplicates     :  {duplicates}\n"
            f"  Empty values   :  {empty_count}"
        )
        self._lbl_summary.config(text=summary_text, fg=self.C_TEXT)

        preview_items = unique[:10]
        if len(unique) > 10:
            preview_text = "\n".join(f"  {a}" for a in preview_items) + \
                           f"\n  ... ({len(unique) - 10} more)"
        else:
            preview_text = "\n".join(f"  {a}" for a in preview_items)
        self._lbl_preview.config(
            text=preview_text if preview_items else "  (no articles found)",
            fg=self.C_TEXT if preview_items else self.C_MUTED,
        )

        if valid:
            self._btn_start.config(state="normal")
        else:
            self._lbl_summary.config(
                text=summary_text + "\n\n  No valid article numbers found.",
                fg=self.C_WARN,
            )

    def _on_start_query(self):
        if self._df is None:
            return

        self._set_ui_running(True)
        self._progress_var.set(0)
        self._lbl_prog_count.config(text="Starting...", fg=self.C_MUTED)
        self._lbl_prog_article.config(text="")
        self._lbl_eta.config(text="Estimated time remaining: calculating...")
        for lbl in self._result_rows.values():
            lbl.config(text="—")
        self._btn_open_folder.config(state="disabled")
        self._btn_open_excel.config(state="disabled")

        self._stop_event.clear()
        self._query_thread = threading.Thread(target=self._run_query, daemon=True)
        self._query_thread.start()

    def _run_query(self):
        """
        Orchestrierungs-Thread: verteilt Artikel auf Worker via ThreadPoolExecutor,
        sammelt Ergebnisse und sendet GUI-Updates thread-safe via after().
        """
        start_time = time.time()

        counts = {
            STATUS_PCF_AVAILABLE: 0,
            STATUS_AAS_NO_PCF:    0,
            STATUS_NOT_FOUND:     0,
            STATUS_QUERY_ERROR:   0,
        }

        articles_all    = self._articles
        valid_articles  = [a for a in articles_all if a]
        unique_articles = list(dict.fromkeys(valid_articles))
        unique_total    = len(unique_articles)

        cache      = {}
        cache_lock = threading.Lock()
        done_count = [0]

        def _worker(article):
            """Laeuft in einem Worker-Thread des ThreadPoolExecutor."""
            if self._stop_event.is_set():
                return article, {"pcf": None, "status": STATUS_QUERY_ERROR}

            res = _query_with_timing(article)

            with cache_lock:
                cache[article] = res
                done_count[0] += 1
                done = done_count[0]
                if res["status"] in counts:
                    counts[res["status"]] += 1

            # ETA berechnen
            elapsed = time.time() - start_time
            if done >= 3 and done < unique_total:
                avg    = elapsed / done
                remain = avg * (unique_total - done)
                mins   = int(remain // 60)
                secs   = int(remain % 60)
                eta_str = f"Estimated time remaining: {mins} min {secs:02d} sec"
            elif done >= unique_total:
                eta_str = "Finalizing..."
            else:
                eta_str = "Estimated time remaining: calculating..."

            pct = (done / unique_total * 100) if unique_total > 0 else 100

            # Lokale Kopien fuer Lambda-Closure (thread-safe)
            _p = pct
            _u = done
            _t = unique_total
            _a = article
            _e = eta_str
            _c = dict(counts)
            self.after(0, lambda p=_p, u=_u, t=_t, a=_a, e=_e, c=_c:
                       self._update_progress(p, u, t, a, e, c))

            return article, res

        with ThreadPoolExecutor(max_workers=WORKERS) as executor:
            futures = {executor.submit(_worker, a): a for a in unique_articles}
            for _ in as_completed(futures):
                pass   # Ergebnisse werden direkt im Worker verarbeitet

        if self._stop_event.is_set():
            self.after(0, lambda: self._on_query_done(None))
            return

        # Ergebnislisten in Originalreihenfolge aufbauen
        pcf_values   = []
        pcf_statuses = []
        df = self._df.copy()

        for article in articles_all:
            if not article:
                pcf_values.append("not available")
                pcf_statuses.append("")
            else:
                res = cache.get(article, {"pcf": None, "status": STATUS_QUERY_ERROR})
                pcf_values.append(res["pcf"] if res["pcf"] is not None else "not available")
                pcf_statuses.append(res["status"])

        df["PCF"]        = pcf_values
        df["PCF Status"] = pcf_statuses

        summary = None
        try:
            df.to_excel(self._output_path, index=False)
            summary = {
                "unique_queried": unique_total,
                "pcf_available":  counts[STATUS_PCF_AVAILABLE],
                "aas_no_pcf":     counts[STATUS_AAS_NO_PCF],
                "not_found":      counts[STATUS_NOT_FOUND],
                "query_errors":   counts[STATUS_QUERY_ERROR],
                "output_file":    self._output_path,
            }
        except Exception as e:
            print(f"Fehler beim Schreiben der Ausgabedatei: {e}")
            err = str(e)
            self.after(0, lambda msg=err: messagebox.showerror(
                "Write Error", f"Could not write output file:\n{msg}",
            ))

        self.after(0, lambda s=summary: self._on_query_done(s))

    # ── GUI-Update-Callbacks (Main-Thread via after()) ─────────────────────────

    def _update_progress(self, pct, unique_done, unique_total, article, eta_str, counts):
        self._progress_var.set(pct)
        self._lbl_prog_count.config(
            text=f"Processing {unique_done} / {unique_total} unique articles",
            fg=self.C_TEXT,
        )
        self._lbl_prog_article.config(
            text=f"Current article: {article}", fg=self.C_ACCENT,
        )
        self._lbl_eta.config(text=eta_str, fg=self.C_MUTED)

        self._result_rows["pcf_available"].config(
            text=str(counts.get(STATUS_PCF_AVAILABLE, 0)))
        self._result_rows["aas_no_pcf"].config(
            text=str(counts.get(STATUS_AAS_NO_PCF, 0)))
        self._result_rows["not_found"].config(
            text=str(counts.get(STATUS_NOT_FOUND, 0)))
        self._result_rows["query_errors"].config(
            text=str(counts.get(STATUS_QUERY_ERROR, 0)))

    def _on_query_done(self, summary):
        self._set_ui_running(False)
        self._progress_var.set(100)
        self._lbl_prog_article.config(text="")
        self._lbl_eta.config(text="")

        if summary is None:
            self._lbl_prog_count.config(
                text="Query finished with errors. Check console output.",
                fg=self.C_ERROR,
            )
            return

        self._lbl_prog_count.config(
            text=(
                f"Query completed — "
                f"{summary['unique_queried']} unique articles processed"
            ),
            fg=self.C_SUCCESS,
        )

        self._result_rows["pcf_available"].config(text=str(summary["pcf_available"]))
        self._result_rows["aas_no_pcf"].config(text=str(summary["aas_no_pcf"]))
        self._result_rows["not_found"].config(text=str(summary["not_found"]))
        self._result_rows["query_errors"].config(text=str(summary["query_errors"]))

        self._btn_open_folder.config(state="normal")
        self._btn_open_excel.config(state="normal")

    def _set_ui_running(self, running: bool):
        state = "disabled" if running else "normal"
        self._btn_select.config(state=state)
        self._btn_start.config(state=state)

    # ── Output-Aktionen ───────────────────────────────────────────────────────

    def _on_open_folder(self):
        if self._output_path:
            import platform
            from pathlib import Path
            folder = Path(self._output_path).parent.resolve()
            
            try:
                if platform.system() == "Windows":
                    os.startfile(str(folder))
                else:
                    subprocess.Popen(['xdg-open', str(folder)])
            except Exception as e:
                messagebox.showerror("Error", f"Could not open folder:\n{e}")

    def _on_open_excel(self):
        if self._output_path and os.path.isfile(self._output_path):
            os.startfile(self._output_path)
        else:
            messagebox.showwarning(
                "File not found",
                f"Output file not found:\n{self._output_path}",
            )


# ── Einstiegspunkt ────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = PCFQueryApp()
    app.mainloop()
