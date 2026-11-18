"""
aas_pcf_query.py
================
WAGO Product Carbon Footprint (PCF) Query & Portfolio Analysis Tool.

Unterstuetzt drei zentrale Workflows:
1. Single Query          : Direkte Abfrage einzelner WAGO-Artikelnummern (Kurz- oder Langformat).
2. Batch Query           : Struktur-agnostischer Excel-Import (.xlsx & .xls) inkl. Drag-and-Drop
                           und standardisiertem Export in das Worksheet 'PCF Results' bei
                           vollstaendigem Erhalt aller Original-Worksheets.
3. PCF Portfolio Coverage: Interne Portfolio-Analyse ueber alle erreichbaren AAS-Shells
                           inkl. Abgleich mit dem eingebetteten Q4-2026-Referenzdatenbestand.

Ausfuehren:
    python aas_pcf_query.py
"""

from __future__ import annotations

import base64
import ctypes
from ctypes import wintypes
import datetime
import email
from email import policy as email_policy
import html
from html.parser import HTMLParser
import io
import json
import os
import queue
import quopri
import re
import shutil
import subprocess
import sys
import threading
import time
import tkinter as tk
import urllib.parse
import webbrowser
import xml.etree.ElementTree as ET
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable, Dict, List, Optional, Set, Tuple, Union

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import requests
import urllib3

try:
    import xlrd
except ImportError:
    xlrd = None  # type: ignore

# ── Konfiguration (Zentralisiert) ────────────────────────────────────────────

ARTICLE_COLUMN = "Article Number"   # Standard-Spaltenname (Fallback/Default)

PCF_RESULTS_SHEET_NAME = "PCF Results"

COL_ARTICLE_SHORT = "Article Number (Short)"
COL_ARTICLE_LONG  = "Article Number (Long)"
COL_PCF           = "PCF"
COL_PCF_UNIT      = "PCF Unit"
COL_PCF_STATUS    = "PCF Status"    # Nur noch fuer interne/Legacy-Kompatibilitaet

DEFAULT_PCF_UNIT  = "kg CO2e"
PORTAL_URL        = "https://aas.wago.com/"

# WAGO SmartService AAS API Basis-URL (AAS Part 2 v3.2)
AAS_BASE_URL      = os.environ.get("WAGO_AAS_BASE_URL", "https://c1.api.wago.com/smartservice-aas").rstrip("/")
AAS_ENV_BASE      = AAS_BASE_URL   # Abwaertskompatibler Alias
AAS_REGISTRY_BASE = os.environ.get("WAGO_AAS_REGISTRY_URL", AAS_BASE_URL).rstrip("/")

# AAS Identifier Templates
SUBMODEL_ID_TEMPLATE        = "https://www.wago.com/submodel/{article}/CarbonFootprint"
LEGACY_SUBMODEL_ID_TEMPLATE = "https://wago.com/ids/submodel/{article}/PCF"
SHELL_ID_TEMPLATE           = "https://wago.com/ids/aas/{article}"
ASSET_ID_TEMPLATE           = "https://wago.com/ids/asset/{article}"

# Optionale Authentifizierung (falls API-Gateway Token oder API-Key erfordert)
API_AUTH_TOKEN = os.environ.get("WAGO_AAS_AUTH_TOKEN", "")
API_KEY        = os.environ.get("WAGO_AAS_API_KEY", "")

# TLS-Pruefung deaktivieren (WAGO Portal nutzt selbstsignierte Zertifikate)
CERT_VERIFY = False

# Timeout in Sekunden (connect, read)
REQUEST_TIMEOUT = (5, 10)

# Parallele Worker-Threads fuer AAS-Abfragen
WORKERS = 5

# Schwellwert fuer Slow-Query-Logging (Sekunden)
SLOW_REQUEST_THRESHOLD = 3.0

# ── WAGO Corporate Design Farben (Section 8.2) ───────────────────────────────

WAGO_GREEN         = "#6EC800"
WAGO_GREEN_HOVER   = "#5EB000"
WAGO_WHITE         = "#FFFFFF"
WAGO_LIGHT_GREY    = "#EFF0F1"
WAGO_ANTHRACITE_15 = "#DEDFE1"
WAGO_ANTHRACITE_40 = "#A5A8AF"
WAGO_ANTHRACITE_70 = "#616A73"
WAGO_ANTHRACITE    = "#1F2837"
WAGO_BLACK         = "#000000"

# ── Status-Konstanten ────────────────────────────────────────────────────────

# Kundenfreundliche Statuswerte
STATUS_PCF_AVAILABLE     = "PCF available"
STATUS_NOT_YET_AVAILABLE = "Not yet available"
STATUS_NOT_AVAILABLE     = STATUS_NOT_YET_AVAILABLE  # Alias
STATUS_QUERY_ERROR       = "Query error"

# Interne Diagnose-Statuswerte (fuer Logging, Troubleshooting & Portfolio Coverage)
INTERNAL_STATUS_PCF_AVAILABLE = "PCF available"
INTERNAL_STATUS_AAS_NO_PCF    = "AAS available - PCF not available"
INTERNAL_STATUS_NOT_FOUND     = "Article not found"
INTERNAL_STATUS_INVALID_ART   = "Invalid article number"
INTERNAL_STATUS_AUTH_ERROR    = "Authentication error"
INTERNAL_STATUS_TIMEOUT       = "Request timeout"
INTERNAL_STATUS_SERVER_ERROR  = "Server error"
INTERNAL_STATUS_INVALID_RESP  = "Unexpected API response"
INTERNAL_STATUS_QUERY_ERROR   = "Query error"

# Abwaertskompatible Konstanten-Namen
STATUS_AAS_NO_PCF = INTERNAL_STATUS_AAS_NO_PCF
STATUS_NOT_FOUND  = INTERNAL_STATUS_NOT_FOUND

# Sentinel-Werte fuer _jget
_NOT_FOUND = object()
_ERROR     = object()

# Betriebsmodi (Online = WAGO AAS API, Offline = Eingebetteter Q4 2026 Snapshot)
MODE_ONLINE  = "online"
MODE_OFFLINE = "offline"
MODE_LABELS  = {
    MODE_ONLINE:  "Online (WAGO AAS)",
    MODE_OFFLINE: "Offline (Q4 2026 Snapshot)",
}

# Workflow-Tab-Kennungen
WORKFLOW_SINGLE    = "single"
WORKFLOW_BATCH     = "batch"
WORKFLOW_PORTFOLIO = "portfolio"

SUPPORTED_EXCEL_EXTENSIONS = (".xlsx", ".xls")

# ── WAGO Artikelnummern-Normalisierung & Konvertierung ───────────────────────

# Primaerer Block: <Serie>-<Artikel>[Suffix], z. B. 221-412, 0221-0412, 2000-115, 770-103-JPE
_RE_PRIMARY_NUMERIC = re.compile(
    r"^(\d{2,4})-(\d{3,4})((?:_[A-Za-z0-9]+|-[A-Za-z][A-Za-z0-9]*))?$"
)
# Primaerer Block mit alphanumerischer Artikelkennung, z. B. 835-0FAR, 0835-0RSC
_RE_PRIMARY_ALPHANUM = re.compile(
    r"^(\d{3,4})-([0-9][A-Za-z0-9]{3})$"
)
# Varianten-Block (nach '/'): numerisch, z. B. 046-000, 0046-0000, 000-100, 281-413_HI
_RE_VARIANT_NUMERIC = re.compile(
    r"^(\d{1,4})-(\d{1,4})((?:_[A-Za-z0-9]+|-[A-Za-z][A-Za-z0-9]*))?$"
)
# Kunden-/Sonder-Varianten-Block (nach '/'): z. B. K010-2937, K000-0001, K000-002, Z000-1809, Z000-046
_RE_VARIANT_KZ = re.compile(
    r"^[A-Za-z]\d{3}-\d{3,4}(?:_[A-Za-z0-9]+)?$"
)
# Komplexer alphanumerischer Varianten-Block nach '/': z. B. 000-00200PW25-000 oder 0000-00200PW25-0000
_RE_VARIANT_COMPLEX = re.compile(
    r"^(\d{3,4})-(\d*[A-Za-z][A-Za-z0-9]*)-(\d{3,4})$"
)
# Kurzer alphabetischer Varianten-Suffix nach '/': z. B. KN
_RE_VARIANT_ALPHA = re.compile(
    r"^[A-Za-z]{1,6}$"
)


def _format_part_pair(p0: str, p1: str, suffix: str = "") -> Tuple[str, str, str]:
    """
    Konvertiert zwei numerische Teilbloecke (p0, p1) in (short_block, long_block, block_fmt).

    WAGO-Konvention:
    - Long-Format : 4-stellig mit fuehrender Null (z. B. 0221-0412, 0221-2411, 2000-0115)
    - Short-Format: 3-stellig (wenn 4-stellig mit '0' beginnt, wird genau eine fuehrende '0'
                    entfernt: 0221 -> 221, 0002 -> 002, 0000 -> 000) bzw. 4-stellig (wenn >= 1000:
                    2411 -> 2411, 2000 -> 2000).
    """
    p0_long = p0.zfill(4)
    p1_long = p1.zfill(4)
    p0_short = p0_long[1:] if p0_long.startswith("0") else p0_long
    p1_short = p1_long[1:] if p1_long.startswith("0") else p1_long

    short_block = f"{p0_short}-{p1_short}{suffix}"
    long_block  = f"{p0_long}-{p1_long}{suffix}"
    raw_block   = f"{p0}-{p1}{suffix}"

    if raw_block == short_block and raw_block == long_block:
        fmt = "both"
    elif raw_block == short_block:
        fmt = "short"
    elif raw_block == long_block:
        fmt = "long"
    else:
        fmt = "mixed"

    return short_block, long_block, fmt


def _is_plausible_unhyphenated_wago_8digit(d8: str, was_7digit: bool = False) -> bool:
    """
    Prueft, ob ein 8-stelliger Ziffernstring (oder von 7 auf 8 Stellen aufgefuellter String)
    eine plausible WAGO-Artikelnummer im 8-stelligen Langformat (ohne Bindestrich) ist.
    Schliesst interne 8-stellige SAP-Materialnummern (z. B. 51170648, 60295535, 99700236) aus.
    """
    if len(d8) != 8 or not d8.isdigit():
        return False
    series = int(d8[:4])
    item   = int(d8[4:])
    if item == 0 and series < 100:
        return False
    if was_7digit:
        return 100 <= series <= 999
    if d8.startswith("0"):
        return 100 <= series <= 999
    return (1000 <= series <= 2999) or (7000 <= series <= 8999)


def parse_article_number(value: Any, number_format: Optional[str] = None) -> Dict[str, Any]:
    """
    Analysiert und normalisiert eine WAGO-Artikelnummer in Kurz- und Langformat.

    Unterstuetzt:
    - Kurzformat: '221-412', '221-2411', '2000-115', '231-004/046-000'
    - Langformat: '0221-0412', '0221-2411', '2000-0115', '0231-0004/0046-0000'
    - Gemischte Formate / verlorene fuehrende Nullen: '221-0412', '0221-412', '85-101'
    - Durch Excel-Formatierung verlorene fuehrende Nullen bei Zahlenwerten:
      z. B. 2210412 -> '221-412' / '0221-0412', 2212411 -> '221-2411' / '0221-2411'
    - Varianten mit '/', 'K...', 'Z...', '_HI', '-JPE' etc.
    - Ungueltige oder mehrdeutige Werte bleiben unveraendert erhalten (is_valid=False).
    """
    if value is None:
        return {
            "original": "",
            "short": "",
            "long": "",
            "canonical": "",
            "is_valid": False,
            "format": "empty",
        }

    try:
        if pd.isna(value):
            return {
                "original": "",
                "short": "",
                "long": "",
                "canonical": "",
                "is_valid": False,
                "format": "empty",
            }
    except Exception:
        pass

    raw_str = str(value).strip()
    if not raw_str or raw_str.lower() in ("nan", "none", "<na>"):
        return {
            "original": "",
            "short": "",
            "long": "",
            "canonical": "",
            "is_valid": False,
            "format": "empty",
        }

    clean = raw_str[1:].strip() if raw_str.startswith("'") else raw_str
    if re.fullmatch(r"\d+\.0+", clean):
        clean = clean.split(".")[0]

    # Fall 1: Reiner Ziffernstring (z. B. durch Excel als Zahl gespeichert, fuehrende 0 verloren)
    if clean.isdigit():
        has_custom_fmt = bool(
            number_format
            and any(pat in str(number_format) for pat in ("0000-0000", '0000"-"0000', "000-000", "00000000"))
        )
        if len(clean) == 7 and _is_plausible_unhyphenated_wago_8digit(clean.zfill(8), was_7digit=True):
            d8 = clean.zfill(8)
            short_b, long_b, _ = _format_part_pair(d8[:4], d8[4:])
            return {
                "original": raw_str,
                "short": short_b,
                "long": long_b,
                "canonical": short_b,
                "is_valid": True,
                "format": "numeric",
            }
        if len(clean) == 8 and _is_plausible_unhyphenated_wago_8digit(clean, was_7digit=False):
            short_b, long_b, _ = _format_part_pair(clean[:4], clean[4:])
            return {
                "original": raw_str,
                "short": short_b,
                "long": long_b,
                "canonical": short_b,
                "is_valid": True,
                "format": "numeric",
            }
        if len(clean) == 6 and has_custom_fmt and int(clean[:3]) >= 100:
            short_b, long_b, _ = _format_part_pair(clean[:3], clean[3:])
            return {
                "original": raw_str,
                "short": short_b,
                "long": long_b,
                "canonical": short_b,
                "is_valid": True,
                "format": "numeric",
            }
        return {
            "original": raw_str,
            "short": raw_str,
            "long": raw_str,
            "canonical": raw_str,
            "is_valid": False,
            "format": "invalid",
        }

    # Fall 2: Artikelnummer mit Bindestrich (und optionalen '/'-Segmenten)
    segments = clean.split("/")
    if any(not seg for seg in segments):
        return {
            "original": raw_str,
            "short": raw_str,
            "long": raw_str,
            "canonical": raw_str,
            "is_valid": False,
            "format": "invalid",
        }

    short_segments: List[str] = []
    long_segments: List[str]  = []
    seg_formats: List[str]    = []

    seg0 = segments[0]
    m_num = _RE_PRIMARY_NUMERIC.match(seg0)
    m_alp = _RE_PRIMARY_ALPHANUM.match(seg0) if not m_num else None

    if m_num:
        p0, p1, suffix = m_num.group(1), m_num.group(2), m_num.group(3) or ""
        series_val = int(p0)
        if series_val < 85:
            return {
                "original": raw_str,
                "short": raw_str,
                "long": raw_str,
                "canonical": raw_str,
                "is_valid": False,
                "format": "invalid",
            }
        s_blk, l_blk, b_fmt = _format_part_pair(p0, p1, suffix)
        short_segments.append(s_blk)
        long_segments.append(l_blk)
        seg_formats.append(b_fmt)
    elif m_alp:
        p0, p1 = m_alp.group(1), m_alp.group(2)
        if int(p0) < 85:
            return {
                "original": raw_str,
                "short": raw_str,
                "long": raw_str,
                "canonical": raw_str,
                "is_valid": False,
                "format": "invalid",
            }
        p0_long  = p0.zfill(4)
        p0_short = p0_long[1:] if p0_long.startswith("0") else p0_long
        s_blk = f"{p0_short}-{p1}"
        l_blk = f"{p0_long}-{p1}"
        b_fmt = "both" if s_blk == l_blk else ("short" if seg0 == s_blk else ("long" if seg0 == l_blk else "mixed"))
        short_segments.append(s_blk)
        long_segments.append(l_blk)
        seg_formats.append(b_fmt)
    else:
        return {
            "original": raw_str,
            "short": raw_str,
            "long": raw_str,
            "canonical": raw_str,
            "is_valid": False,
            "format": "invalid",
        }

    for seg in segments[1:]:
        if _RE_VARIANT_KZ.match(seg) or _RE_VARIANT_ALPHA.match(seg):
            short_segments.append(seg)
            long_segments.append(seg)
            seg_formats.append("both")
            continue

        m_cplx = _RE_VARIANT_COMPLEX.match(seg)
        if m_cplx:
            cp0, cmid, cp2 = m_cplx.group(1), m_cplx.group(2), m_cplx.group(3)
            cp0_long  = cp0.zfill(4)
            cp0_short = cp0_long[1:] if cp0_long.startswith("0") else cp0_long
            cp2_long  = cp2.zfill(4)
            cp2_short = cp2_long[1:] if cp2_long.startswith("0") else cp2_long
            s_cplx = f"{cp0_short}-{cmid}-{cp2_short}"
            l_cplx = f"{cp0_long}-{cmid}-{cp2_long}"
            if seg == s_cplx and seg == l_cplx:
                c_fmt = "both"
            elif seg == s_cplx:
                c_fmt = "short"
            elif seg == l_cplx:
                c_fmt = "long"
            else:
                c_fmt = "mixed"
            short_segments.append(s_cplx)
            long_segments.append(l_cplx)
            seg_formats.append(c_fmt)
            continue

        m_var = _RE_VARIANT_NUMERIC.match(seg)
        if m_var:
            vp0, vp1, vsuffix = m_var.group(1), m_var.group(2), m_var.group(3) or ""
            if len(vp0) < 3 and len(vp1) < 3:
                return {
                    "original": raw_str,
                    "short": raw_str,
                    "long": raw_str,
                    "canonical": raw_str,
                    "is_valid": False,
                    "format": "invalid",
                }
            s_blk, l_blk, b_fmt = _format_part_pair(vp0, vp1, vsuffix)
            short_segments.append(s_blk)
            long_segments.append(l_blk)
            seg_formats.append(b_fmt)
            continue

        return {
            "original": raw_str,
            "short": raw_str,
            "long": raw_str,
            "canonical": raw_str,
            "is_valid": False,
            "format": "invalid",
        }

    short_full = "/".join(short_segments)
    long_full  = "/".join(long_segments)

    non_both = [f for f in seg_formats if f != "both"]
    if not non_both:
        overall_fmt = "both"
    elif all(f == "short" for f in non_both):
        overall_fmt = "short"
    elif all(f == "long" for f in non_both):
        overall_fmt = "long"
    else:
        overall_fmt = "mixed"

    return {
        "original": raw_str,
        "short": short_full,
        "long": long_full,
        "canonical": short_full,
        "is_valid": True,
        "format": overall_fmt,
    }


def is_valid_wago_article(value: Any) -> bool:
    """Gibt True zurueck, wenn der Wert eine eindeutige WAGO-Artikelnummer ist."""
    return parse_article_number(value)["is_valid"]


def to_short_article_number(value: Any) -> str:
    """Konvertiert eine WAGO-Artikelnummer in das Kurzformat (z. B. '0221-0412' -> '221-412')."""
    return parse_article_number(value)["short"]


def to_long_article_number(value: Any) -> str:
    """Konvertiert eine WAGO-Artikelnummer in das Langformat (z. B. '221-412' -> '0221-0412')."""
    return parse_article_number(value)["long"]


to_short_format = to_short_article_number
to_long_format  = to_long_article_number


def normalize_article_number(value: Any) -> str:
    """Normalisiert eine WAGO-Artikelnummer auf ihre kanonische Darstellung (Kurzformat)."""
    return parse_article_number(value)["canonical"]


def convert_article_number(value: Any) -> Tuple[str, str]:
    """Gibt (short_format, long_format) fuer eine Artikelnummer zurueck."""
    parsed = parse_article_number(value)
    return parsed["short"], parsed["long"]


# Isoliertes Offline-Datenmodul (Q4 2026 Snapshot) importieren
from offline_pcf_data import (  # noqa: E402
    DEFAULT_OFFLINE_PCF_UNIT,
    EMBEDDED_PCF_DATA,
    OFFLINE_DATA_SOURCE_LABEL,
    OFFLINE_EXCEL_HEADER_NOTE,
    OFFLINE_NOTICE_TEXT,
    OFFLINE_NOTICE_TITLE,
    OFFLINE_PORTAL_URL,
    OFFLINE_SNAPSHOT_VERSION,
    get_offline_pcf,
    get_offline_pcf_result,
    register_article_parser,
    validate_and_index_pcf_records,
)

register_article_parser(parse_article_number)

from portfolio_coverage import (  # noqa: E402
    ACTION_PUBLISHED,
    ACTION_REVIEW_CALCULATION,
    ACTION_REVIEW_PUBLICATION,
    ACTION_TECH_INVESTIGATION,
    CLASS_PCF_AVAILABLE_AAS,
    CLASS_PCF_MISSING_BOTH,
    CLASS_PCF_OFFLINE_ONLY,
    CLASS_TECHNICAL_ERROR,
    export_portfolio_coverage_excel,
    extract_product_series,
    run_portfolio_coverage_analysis,
)


# ── Struktur-agnostische Excel-Erkennung & -Verarbeitung (.xlsx & .xls) ──────

_PRIMARY_HEADER_PATTERNS = [
    re.compile(r"\barticle\s*(?:number|no\.?|nr\.?|id|code)?\b", re.I),
    re.compile(r"\bartikel(?:nummer|[\s\-]*nr\.?)?\b", re.I),
    re.compile(r"\bwago\s*(?:number|no\.?|nr\.?|id|article|artikel(?:nummer)?)?\b", re.I),
    re.compile(r"\bmaterial\s*lang\s*(?:kompr|ber)\.?\b", re.I),
    re.compile(r"\bmaterial\s*(?:number|no\.?|nr\.?|nummer)\b", re.I),
    re.compile(r"\bitem\s*(?:number|no\.?|nr\.?|id|code)\b", re.I),
    re.compile(r"\bproduct\s*(?:id|number|no\.?|nr\.?|code)\b", re.I),
    re.compile(r"\b(?:part|order|bestell)\s*(?:number|no\.?|nr\.?|nummer)\b", re.I),
    re.compile(r"\b(?:mat|art)\.?\s*(?:nr|no)\.?\b", re.I),
]

_SECONDARY_COMPONENT_PATTERNS = [
    re.compile(r"\b(?:komponente|component|sub[\s\-]*item|child|parent)\b", re.I),
]

_SHORT_HEADER_HINT = re.compile(r"\b(?:short|kurz|kompr|compressed)\b", re.I)
_LONG_HEADER_HINT  = re.compile(r"\b(?:long|lang|expanded|erweitert|ber\.?)\b", re.I)


def _score_column_header(col_name: Any) -> float:
    """Bewertet einen Spaltennamen als sekundaeres Signal (0.0 bis 0.15)."""
    if col_name is None:
        return 0.0
    s = str(col_name).strip()
    if not s or s.lower().startswith("unnamed"):
        return 0.0

    is_component = any(p.search(s) for p in _SECONDARY_COMPONENT_PATTERNS)
    if s.lower() in (
        "article number",
        "article number (short)",
        "article number (long)",
        "artikelnummer",
        "wago number",
        "material number",
        "materialnummer",
        "item number",
        "product id",
    ):
        return 0.03 if is_component else 0.15

    if any(p.search(s) for p in _PRIMARY_HEADER_PATTERNS):
        return 0.03 if is_component else 0.12

    if s.lower() in ("material", "item", "product", "part", "id", "code", "nr", "no"):
        return 0.05

    return 0.0


def _evaluate_dataframe_columns(df: pd.DataFrame) -> Dict[str, Any]:
    """Analysiert alle Spalten eines DataFrames inhaltsbasiert auf WAGO-Artikelnummern."""
    candidates: List[Dict[str, Any]] = []
    col_parsed_map: Dict[str, List[Dict[str, Any]]] = {}

    for col in df.columns:
        series = df[col]
        if isinstance(series, pd.DataFrame):
            series = series.iloc[:, 0]

        parsed_list: List[Dict[str, Any]] = []
        non_empty = 0
        valid_cnt = 0
        explicit_hyphen_cnt = 0
        fmt_counts: Dict[str, int] = {"short": 0, "long": 0, "both": 0, "mixed": 0, "numeric": 0}

        for val in series.tolist():
            p = parse_article_number(val)
            parsed_list.append(p)
            if p["format"] != "empty":
                non_empty += 1
                if p["is_valid"]:
                    valid_cnt += 1
                    fmt_counts[p["format"]] = fmt_counts.get(p["format"], 0) + 1
                    if p["format"] in ("short", "long", "both", "mixed"):
                        explicit_hyphen_cnt += 1

        col_parsed_map[str(col)] = parsed_list

        if non_empty == 0 or valid_cnt == 0:
            continue

        match_ratio = valid_cnt / non_empty
        hyphen_ratio = explicit_hyphen_cnt / valid_cnt if valid_cnt > 0 else 0.0

        short_compat = fmt_counts["short"] + fmt_counts["both"]
        long_compat  = fmt_counts["long"] + fmt_counts["both"]
        max_consistent = max(short_compat, long_compat, fmt_counts["numeric"], fmt_counts["mixed"])
        consistency = max_consistent / valid_cnt if valid_cnt > 0 else 0.0

        if fmt_counts["long"] > fmt_counts["short"] and fmt_counts["long"] >= fmt_counts["numeric"]:
            dom_fmt = "long"
        elif fmt_counts["short"] > 0 or fmt_counts["both"] > 0:
            dom_fmt = "short"
        elif fmt_counts["numeric"] > 0:
            dom_fmt = "numeric"
        else:
            dom_fmt = "mixed"

        header_score = _score_column_header(col)
        content_quality = 0.85 + 0.15 * hyphen_ratio
        confidence = (
            0.70 * match_ratio * content_quality
            + 0.15 * consistency
            + header_score
        )
        confidence = min(1.0, round(confidence, 4))

        candidates.append({
            "column": col,
            "confidence": confidence,
            "match_ratio": round(match_ratio, 4),
            "consistency": round(consistency, 4),
            "valid_count": valid_cnt,
            "non_empty_count": non_empty,
            "dominant_format": dom_fmt,
            "header_score": header_score,
        })

    candidates.sort(
        key=lambda c: (c["confidence"], c["valid_count"], c["match_ratio"], c["header_score"]),
        reverse=True,
    )

    if not candidates:
        return {
            "column": None,
            "short_column": None,
            "long_column": None,
            "has_both_formats": False,
            "confidence": 0.0,
            "is_ambiguous": False,
            "candidates": [],
        }

    best = candidates[0]
    best_col = best["column"]
    short_col: Optional[Any] = None
    long_col: Optional[Any]  = None
    has_both_formats = False
    is_ambiguous = False

    if len(candidates) >= 2:
        second = candidates[1]
        c1_name = str(best["column"])
        c2_name = str(second["column"])
        p1_list = col_parsed_map[c1_name]
        p2_list = col_parsed_map[c2_name]

        both_valid_rows = 0
        same_product_rows = 0
        for p1, p2 in zip(p1_list, p2_list):
            if p1["is_valid"] and p2["is_valid"]:
                both_valid_rows += 1
                if p1["canonical"] == p2["canonical"]:
                    same_product_rows += 1

        same_product_ratio = (same_product_rows / both_valid_rows) if both_valid_rows > 0 else 0.0

        if both_valid_rows > 0 and same_product_ratio >= 0.80:
            has_both_formats = True
            c1_is_long = (
                best["dominant_format"] == "long"
                or bool(_LONG_HEADER_HINT.search(c1_name))
                or bool(_SHORT_HEADER_HINT.search(c2_name))
            )
            if c1_is_long and not bool(_SHORT_HEADER_HINT.search(c1_name)):
                long_col  = best["column"]
                short_col = second["column"]
            else:
                short_col = best["column"]
                long_col  = second["column"]
        else:
            score_diff = best["confidence"] - second["confidence"]
            if second["match_ratio"] >= 0.50 and score_diff < 0.08:
                is_ambiguous = True

    if not has_both_formats:
        col_str = str(best_col)
        if col_str.lower() == COL_ARTICLE_SHORT.lower() or _SHORT_HEADER_HINT.search(col_str):
            short_col = best_col
        elif col_str.lower() == COL_ARTICLE_LONG.lower() or _LONG_HEADER_HINT.search(col_str):
            long_col = best_col

    return {
        "column": best_col,
        "short_column": short_col,
        "long_column": long_col,
        "has_both_formats": has_both_formats,
        "confidence": best["confidence"],
        "is_ambiguous": is_ambiguous,
        "candidates": candidates,
    }


def _extract_sheet_rows_with_formats(ws: openpyxl.worksheet.worksheet.Worksheet) -> List[List[str]]:
    """
    Liest alle Zeilen eines openpyxl-Worksheets (.xlsx) und stellt durch Excel-Zahlenformatierung
    verlorene fuehrende Nullen anhand von cell.number_format wieder her.
    """
    rows: List[List[str]] = []
    for row in ws.iter_rows(values_only=False):
        row_vals: List[str] = []
        for cell in row:
            val = cell.value
            if val is None:
                row_vals.append("")
                continue
            num_fmt = getattr(cell, "number_format", None)
            if isinstance(val, (int, float)):
                if isinstance(val, float) and val.is_integer():
                    val = int(val)
                if isinstance(val, int):
                    s_int = str(val)
                    p = parse_article_number(s_int, number_format=num_fmt)
                    if p["is_valid"] and p["format"] == "numeric":
                        if num_fmt and "0000" in str(num_fmt):
                            row_vals.append(p["long"])
                        else:
                            row_vals.append(s_int)
                    else:
                        row_vals.append(s_int)
                else:
                    row_vals.append(str(val))
            else:
                row_vals.append(str(val).strip())
        rows.append(row_vals)
    return rows


def _extract_xls_sheet_rows_with_formats(sheet: Any, book: Any) -> List[List[str]]:
    """
    Liest alle Zeilen eines Legacy-Excel-Worksheets (.xls via xlrd) und stellt
    durch Zahlenformatierung verlorene fuehrende Nullen wieder her.
    """
    rows: List[List[str]] = []
    for r_idx in range(sheet.nrows):
        row_vals: List[str] = []
        for c_idx in range(sheet.ncols):
            cell = sheet.cell(r_idx, c_idx)
            ctype = cell.ctype
            val = cell.value

            if ctype in (0, 6) or val is None:  # XL_CELL_EMPTY, XL_CELL_BLANK
                row_vals.append("")
                continue

            num_fmt = None
            if getattr(book, "formatting_info", False) and cell.xf_index is not None:
                try:
                    xf = book.xf_list[cell.xf_index]
                    fmt_obj = book.format_map.get(xf.format_key)
                    if fmt_obj:
                        num_fmt = fmt_obj.format_str
                except Exception:
                    num_fmt = None

            if ctype == 2:  # XL_CELL_NUMBER
                if isinstance(val, float) and val.is_integer():
                    s_int = str(int(val))
                    p = parse_article_number(s_int, number_format=num_fmt)
                    if p["is_valid"] and p["format"] == "numeric":
                        if num_fmt and "0000" in str(num_fmt):
                            row_vals.append(p["long"])
                        else:
                            row_vals.append(s_int)
                    else:
                        row_vals.append(s_int)
                else:
                    row_vals.append(str(val))
            elif ctype == 3 and xlrd is not None:  # XL_CELL_DATE
                try:
                    dt = xlrd.xldate_as_datetime(val, book.datemode)
                    row_vals.append(dt.strftime("%Y-%m-%d"))
                except Exception:
                    row_vals.append(str(val))
            elif ctype == 4:  # XL_CELL_BOOLEAN
                row_vals.append("TRUE" if val else "FALSE")
            else:
                row_vals.append(str(val).strip())
        rows.append(row_vals)
    return rows


def _parse_sheet_table(rows: List[List[str]]) -> Tuple[pd.DataFrame, Optional[int], bool]:
    """
    Erkennt automatisch:
    - Einleitende Zeilen vor der eigentlichen Tabelle (Introductory Rows)
    - Ob eine Kopfzeile (Header Row) existiert oder ob die Tabelle direkt mit Daten beginnt
    Gibt (df, header_row_index, has_header) zurueck.
    """
    rows = list(rows)
    while rows and all(c == "" for c in rows[-1]):
        rows.pop()

    if not rows:
        return pd.DataFrame(), None, True

    max_cols = max(len(r) for r in rows)
    norm_rows = [r + [""] * (max_cols - len(r)) for r in rows]

    first_article_row: Optional[int] = None
    for r_idx, r in enumerate(norm_rows):
        if any(is_valid_wago_article(cell) for cell in r):
            first_article_row = r_idx
            break

    if first_article_row is None:
        for r_idx, r in enumerate(norm_rows):
            if any(c != "" for c in r):
                headers = [c if c != "" else f"Unnamed: {i}" for i, c in enumerate(r)]
                data = norm_rows[r_idx + 1:]
                return pd.DataFrame(data, columns=headers), r_idx, True
        return pd.DataFrame(), None, True

    if first_article_row == 0:
        headers = [f"Column {i + 1}" for i in range(max_cols)]
        df = pd.DataFrame(norm_rows, columns=headers)
        return df, None, False

    prev_row = norm_rows[first_article_row - 1]
    prev_non_empty = [c for c in prev_row if c != ""]

    if prev_non_empty and not any(is_valid_wago_article(c) for c in prev_row):
        header_row_idx = first_article_row - 1
        raw_headers = norm_rows[header_row_idx]
        headers = []
        seen_counts: Dict[str, int] = {}
        for i, h in enumerate(raw_headers):
            base_h = h if h != "" else f"Unnamed: {i}"
            if base_h in seen_counts:
                seen_counts[base_h] += 1
                headers.append(f"{base_h}_{seen_counts[base_h]}")
            else:
                seen_counts[base_h] = 0
                headers.append(base_h)
        data = norm_rows[first_article_row:]
        df = pd.DataFrame(data, columns=headers)
        return df, header_row_idx, True
    else:
        headers = [f"Column {i + 1}" for i in range(max_cols)]
        data = norm_rows[first_article_row:]
        df = pd.DataFrame(data, columns=headers)
        return df, None, False


# ── Inhaltsbasierte Dateiformat-Erkennung & MHTML / HTML / XML Parser ────────

_META_CHARSET_RE = re.compile(
    rb"""<meta[^>]+charset\s*=\s*["']?\s*([A-Za-z0-9_\-]+)""",
    re.I,
)
_XML_ENCODING_RE = re.compile(
    rb"""<\?xml[^>]+encoding\s*=\s*["']([A-Za-z0-9_\-]+)["']""",
    re.I,
)


def _decode_bytes_smart(raw_bytes: bytes, declared_charset: Optional[str] = None) -> str:
    """
    Dekodiert ein Byte-Array zuversichtlich unter Beachtung von BOMs, deklariertem
    MIME/Meta/XML-Charset, UTF-8 und Windows-1252/ISO-8859-1-Fallback (fuer deutsche Umlaute).
    """
    if not raw_bytes:
        return ""
    if raw_bytes.startswith(b"\xef\xbb\xbf"):
        return raw_bytes[3:].decode("utf-8", errors="replace")
    if raw_bytes.startswith(b"\xff\xfe"):
        return raw_bytes[2:].decode("utf-16-le", errors="replace")
    if raw_bytes.startswith(b"\xfe\xff"):
        return raw_bytes[2:].decode("utf-16-be", errors="replace")
    if len(raw_bytes) >= 4 and raw_bytes[1:4:2] == b"\x00\x00" and raw_bytes[0] != 0:
        try:
            return raw_bytes.decode("utf-16-le")
        except Exception:
            pass

    candidates: List[str] = []
    if declared_charset:
        c_clean = declared_charset.strip().strip("\"'").lower()
        if c_clean:
            candidates.append(c_clean)

    m_meta = _META_CHARSET_RE.search(raw_bytes[:8192])
    if m_meta:
        try:
            candidates.append(m_meta.group(1).decode("ascii", errors="ignore").lower())
        except Exception:
            pass

    m_xml = _XML_ENCODING_RE.search(raw_bytes[:2048])
    if m_xml:
        try:
            candidates.append(m_xml.group(1).decode("ascii", errors="ignore").lower())
        except Exception:
            pass

    for enc in candidates:
        try:
            return raw_bytes.decode(enc)
        except (LookupError, UnicodeDecodeError):
            continue

    try:
        return raw_bytes.decode("utf-8")
    except UnicodeDecodeError:
        return raw_bytes.decode("cp1252", errors="replace")


def detect_excel_file_format(file_path: Union[str, Path]) -> str:
    """
    Erkennt das tatsaechliche interne Dateiformat anhand der Dateisignatur und des
    Inhalts (unabhaengig von der Dateiendung .xls / .xlsx):
      - 'xlsx'     : Natives Office Open XML Workbook (ZIP mit xl/)
      - 'xls_biff' : Natives binaeres Legacy-Excel-Workbook (OLE2 / BIFF)
      - 'mhtml'    : MHTML-basierter Excel-Export (MIME-Version / multipart/related, z. B. SAP)
      - 'xml_2003' : Excel 2003 XML Spreadsheet (SpreadsheetML)
      - 'html'     : HTML-basierter Excel-Export (<table>...</table>)
      - 'unknown'  : Unbekanntes oder beschaedigtes Format
    """
    p = Path(file_path)
    if not p.is_file():
        return "unknown"

    try:
        with open(p, "rb") as fp:
            head = fp.read(65536)
    except Exception:
        return "unknown"

    if not head:
        return "unknown"

    # 1. ZIP / Office Open XML (.xlsx)
    if head.startswith(b"PK\x03\x04"):
        try:
            with zipfile.ZipFile(p, "r") as zf:
                names = set(zf.namelist())
                if any(n.startswith("xl/") or n == "[Content_Types].xml" for n in names):
                    return "xlsx"
        except Exception:
            return "unknown"
        return "unknown"

    # 2. OLE2 Compound Document / BIFF2-BIFF8 (.xls)
    if head.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1") or (
        len(head) >= 8
        and head[:2] in (b"\x09\x00", b"\x09\x02", b"\x09\x04", b"\x09\x08")
        and head[2:4] in (b"\x04\x00", b"\x06\x00", b"\x08\x00", b"\x10\x00")
    ):
        return "xls_biff"

    # Dekodiere Text-Vorschau fuer textbasierte Enterprise-/SAP-Excel-Exporte
    text_sample = _decode_bytes_smart(head).lstrip()
    if not text_sample:
        return "unknown"

    text_lower = text_sample[:16384].lower()

    # 3. MHTML (MIME HTML, z. B. SAP-Export mit 'MIME-Version: 1.0')
    first_lines = text_sample[:4096]
    if (
        re.match(r"^(?:mime-version|from|subject|date|content-type|message-id)\s*:", first_lines, re.I | re.M)
        and (
            "mime-version:" in text_lower
            or "multipart/related" in text_lower
            or "content-transfer-encoding:" in text_lower
            or "content-location:" in text_lower
            or "application/vnd.ms-excel" in text_lower
            or "<table" in text_lower
            or "=3c" in text_lower
        )
    ):
        return "mhtml"

    # 4. Excel 2003 XML Spreadsheet (SpreadsheetML)
    if (
        "urn:schemas-microsoft-com:office:spreadsheet" in text_lower
        or 'progid="excel.sheet"' in text_lower
        or ("<workbook" in text_lower and ("<worksheet" in text_lower or "<ss:worksheet" in text_lower))
        or ("<ss:workbook" in text_lower)
    ):
        return "xml_2003"

    # 5. HTML-basierter Excel-Export
    if "<table" in text_lower and ("<tr" in text_lower or "<td" in text_lower or "<th" in text_lower):
        return "html"

    return "unknown"


class _HTMLTableExtractor(HTMLParser):
    """
    Robuster, verschachtelungssicherer HTML-Tabellen-Parser fuer MHTML- und HTML-Excel-Exporte.
    - Unterstuetzt verschachtelte Layout-Tabellen (extrahiert nur echte Datentabellen).
    - Unterstuetzt colspan und rowspan fuer exakte Spaltenausrichtung.
    - Ignoriert <style>, <script>, <xml>, <head>, <title> und Kommentare.
    - Bewahrt fuehrende Nullen und komplexe WAGO-Artikelnummern unveraendert als String.
    """

    _IGNORED_TAGS = {"style", "script", "xml", "head", "title", "noscript"}

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self._ignore_depth = 0
        self._table_stack: List[Dict[str, Any]] = []
        self.completed_tables: List[Dict[str, Any]] = []
        self._sheet_names_from_xml: List[Tuple[str, str]] = []

    def handle_starttag(self, tag: str, attrs: List[Tuple[str, Optional[str]]]) -> None:
        t = tag.lower()
        if t in self._IGNORED_TAGS:
            self._ignore_depth += 1
            return
        if self._ignore_depth > 0:
            return

        attr_map = {k.lower(): (v or "") for k, v in attrs if k}

        if t == "table":
            if self._table_stack:
                self._table_stack[-1]["has_child_table"] = True
            tbl_name = (
                attr_map.get("data-sheet-name")
                or attr_map.get("title")
                or attr_map.get("id")
                or ""
            ).strip()
            self._table_stack.append({
                "name": tbl_name,
                "caption": "",
                "in_caption": False,
                "has_child_table": False,
                "rows": [],
                "curr_row": None,
                "in_cell": False,
                "cell_buf": [],
                "colspan": 1,
                "rowspan": 1,
            })
        elif t == "caption" and self._table_stack:
            self._table_stack[-1]["in_caption"] = True
        elif t == "tr" and self._table_stack:
            self._table_stack[-1]["curr_row"] = []
        elif t in ("td", "th") and self._table_stack:
            top = self._table_stack[-1]
            if top["curr_row"] is None:
                top["curr_row"] = []
            try:
                cs = max(1, min(200, int(attr_map.get("colspan", "1") or "1")))
            except ValueError:
                cs = 1
            try:
                rs = max(1, min(10000, int(attr_map.get("rowspan", "1") or "1")))
            except ValueError:
                rs = 1
            top["in_cell"] = True
            top["cell_buf"] = []
            top["colspan"] = cs
            top["rowspan"] = rs
        elif t == "br" and self._table_stack and self._table_stack[-1]["in_cell"]:
            self._table_stack[-1]["cell_buf"].append(" ")

    def handle_endtag(self, tag: str) -> None:
        t = tag.lower()
        if t in self._IGNORED_TAGS:
            if self._ignore_depth > 0:
                self._ignore_depth -= 1
            return
        if self._ignore_depth > 0:
            return

        if t in ("td", "th") and self._table_stack:
            top = self._table_stack[-1]
            if top["in_cell"] and top["curr_row"] is not None:
                raw_txt = "".join(top["cell_buf"]).replace("\xa0", " ")
                clean_txt = re.sub(r"[ \t\r\n]+", " ", raw_txt).strip()
                top["curr_row"].append((clean_txt, top["colspan"], top["rowspan"]))
            top["in_cell"] = False
            top["cell_buf"] = []
        elif t == "tr" and self._table_stack:
            top = self._table_stack[-1]
            if top["curr_row"] is not None:
                top["rows"].append(top["curr_row"])
            top["curr_row"] = None
        elif t == "caption" and self._table_stack:
            self._table_stack[-1]["in_caption"] = False
        elif t == "table" and self._table_stack:
            finished = self._table_stack.pop()
            if finished["curr_row"]:
                finished["rows"].append(finished["curr_row"])
            grid = self._expand_spans_to_grid(finished["rows"])
            if grid and not (finished["has_child_table"] and len(grid) <= 2):
                self.completed_tables.append({
                    "name": finished["caption"].strip() or finished["name"],
                    "rows": grid,
                    "has_child_table": finished["has_child_table"],
                })

    def handle_data(self, data: str) -> None:
        if self._ignore_depth > 0:
            return
        if not self._table_stack:
            return
        top = self._table_stack[-1]
        if top["in_cell"]:
            top["cell_buf"].append(data)
        elif top["in_caption"]:
            top["caption"] += data

    @staticmethod
    def _expand_spans_to_grid(
        raw_rows: List[List[Tuple[str, int, int]]]
    ) -> List[List[str]]:
        if not raw_rows:
            return []
        grid_map: Dict[Tuple[int, int], str] = {}
        max_col = 0
        total_r = len(raw_rows)

        for r_idx, cell_tuples in enumerate(raw_rows):
            c_idx = 0
            for val, cs, rs in cell_tuples:
                while (r_idx, c_idx) in grid_map:
                    c_idx += 1
                for dr in range(rs):
                    for dc in range(cs):
                        target_r = r_idx + dr
                        target_c = c_idx + dc
                        if target_r < total_r + rs:
                            grid_map[(target_r, target_c)] = val if (dr == 0 and dc == 0) else ""
                c_idx += cs
                if c_idx > max_col:
                    max_col = c_idx

        if max_col == 0:
            return []

        out_rows: List[List[str]] = []
        for r_idx in range(total_r):
            row_vals = [grid_map.get((r_idx, c_idx), "") for c_idx in range(max_col)]
            out_rows.append(row_vals)

        while out_rows and all(c == "" for c in out_rows[-1]):
            out_rows.pop()
        return out_rows


def _extract_tables_from_html_text(
    html_text: str, default_sheet_prefix: str = "Sheet"
) -> List[Tuple[str, List[List[str]]]]:
    """Extrahiert alle relevanten Datentabellen aus einem HTML-Dokument als [(sheet_name, rows)]."""
    extractor = _HTMLTableExtractor()
    extractor.feed(html_text)
    extractor.close()

    tables = extractor.completed_tables
    if not tables:
        return []

    # Bevorzuge Blatt-Tabellen ohne verschachtelte Kind-Tabellen
    leaf_tables = [t for t in tables if not t["has_child_table"]]
    candidates = leaf_tables if leaf_tables else tables

    # Falls mehrere Tabellen vorhanden sind, ignoriere reine 1-Zellen-Banner/Layout-Tabellen
    if len(candidates) > 1:
        multi_cell = [
            t for t in candidates
            if len(t["rows"]) >= 2 or (t["rows"] and sum(1 for c in t["rows"][0] if c != "") >= 2)
        ]
        if multi_cell:
            candidates = multi_cell

    # Prüfe optionales <x:Name> oder <title> im HTML-Header als Fallback-Sheetname
    doc_sheet_names = [
        html.unescape(m.group(1)).strip()
        for m in re.finditer(r"<(?:x:)?Name>\s*(.*?)\s*</(?:x:)?Name>", html_text, re.I | re.S)
        if m.group(1).strip()
    ]
    doc_title = ""
    m_title = re.search(r"<title[^>]*>\s*(.*?)\s*</title>", html_text, re.I | re.S)
    if m_title:
        doc_title = html.unescape(m_title.group(1)).strip()

    result: List[Tuple[str, List[List[str]]]] = []
    used_names: Set[str] = set()
    for idx, tbl in enumerate(candidates, start=1):
        fallback_doc_name = (
            doc_sheet_names[idx - 1]
            if idx - 1 < len(doc_sheet_names)
            else (doc_title if len(candidates) == 1 and doc_title else "")
        )
        raw_name = tbl["name"] or fallback_doc_name or f"{default_sheet_prefix}{idx}"
        # Bereinige unzulässige Excel-Sheet-Zeichen (: \ / ? * [ ])
        clean_name = re.sub(r"[:\\/?*\[\]]+", "_", raw_name).strip()[:31] or f"{default_sheet_prefix}{idx}"
        if clean_name in used_names:
            clean_name = f"{clean_name[:27]}_{idx}"
        used_names.add(clean_name)
        result.append((clean_name, tbl["rows"]))

    return result


def _parse_html_sheets(file_path: str) -> List[Tuple[str, List[List[str]]]]:
    """Liest einen HTML-basierten Excel-Export (.xls / .html) ein."""
    raw_bytes = Path(file_path).read_bytes()
    html_text = _decode_bytes_smart(raw_bytes)
    sheets = _extract_tables_from_html_text(html_text)
    if not sheets:
        raise ValueError("No valid spreadsheet table found in HTML file.")
    return sheets


def _parse_mhtml_sheets(file_path: str) -> List[Tuple[str, List[List[str]]]]:
    """
    Liest einen MHTML-basierten Excel-Export (MIME / multipart/related, z. B. aus SAP)
    ein, dekodiert quoted-printable / Base64 sowie den Zeichensatz jedes MIME-Parts
    und extrahiert alle enthaltenen Worksheets in ihrer Originalreihenfolge.
    """
    raw_bytes = Path(file_path).read_bytes()

    # Falls die MHTML-Datei komplett in UTF-16 gespeichert wurde, zuerst nach UTF-8 konvertieren
    if raw_bytes.startswith((b"\xff\xfe", b"\xfe\xff")) or (
        len(raw_bytes) >= 4 and raw_bytes[1:4:2] == b"\x00\x00"
    ):
        text_u16 = _decode_bytes_smart(raw_bytes)
        raw_bytes = text_u16.encode("utf-8")
    elif raw_bytes.startswith(b"\xef\xbb\xbf"):
        raw_bytes = raw_bytes[3:]

    # Entferne etwaige fuehrende Leerzeilen vor dem ersten MIME-Header
    raw_bytes = raw_bytes.lstrip(b" \t\r\n")

    msg = email.message_from_bytes(raw_bytes, policy=email_policy.compat32)

    html_parts: List[Tuple[str, str]] = []  # [(content_location, decoded_html)]
    for part in msg.walk():
        if part.is_multipart():
            continue
        ctype = (part.get_content_type() or "").lower()
        loc = (part.get("Content-Location") or part.get("Content-ID") or "").strip()
        payload_bytes = part.get_payload(decode=True)
        if not payload_bytes:
            raw_payload = part.get_payload()
            if isinstance(raw_payload, str) and raw_payload.strip():
                payload_bytes = raw_payload.encode("utf-8", errors="replace")
        if not payload_bytes:
            continue

        if ctype in ("text/html", "application/vnd.ms-excel", "text/plain", "text/xml") or (
            b"<table" in payload_bytes.lower() or b"<x:excelworkbook" in payload_bytes.lower()
        ):
            declared_cs = part.get_content_charset()
            decoded_text = _decode_bytes_smart(payload_bytes, declared_charset=declared_cs)
            if "<table" in decoded_text.lower() or "<x:excelworksheet" in decoded_text.lower():
                html_parts.append((loc, decoded_text))

    # Fallback falls ein nicht-standardkonformer SAP-MHTML-Export ohne sauberen Boundary-Header vorliegt
    if not html_parts:
        raw_text = _decode_bytes_smart(raw_bytes)
        if "quoted-printable" in raw_text[:2048].lower() or "=3c" in raw_text.lower() or "=3d" in raw_text.lower():
            try:
                qp_bytes = quopri.decodestring(raw_bytes)
                decoded_qp = _decode_bytes_smart(qp_bytes)
                if "<table" in decoded_qp.lower():
                    html_parts.append(("", decoded_qp))
            except Exception:
                pass
        if not html_parts and "<table" in raw_text.lower():
            html_parts.append(("", raw_text))

    if not html_parts:
        raise ValueError("No HTML worksheet content found in MHTML file.")

    # Pruefe, ob im Root-Part <x:ExcelWorksheet>-Definitionen mit Namen und HRef vorhanden sind
    ws_href_to_name: List[Tuple[str, str]] = []
    for _loc, h_txt in html_parts:
        for m_ws in re.finditer(
            r"<x:ExcelWorksheet[^>]*>(.*?)</x:ExcelWorksheet>", h_txt, re.I | re.S
        ):
            block = m_ws.group(1)
            m_name = re.search(r"<x:Name[^>]*>(.*?)</x:Name>", block, re.I | re.S)
            m_href = re.search(
                r"""<x:WorksheetSource[^>]+HRef\s*=\s*["']?([^"'\s>]+)""", block, re.I
            )
            if m_name:
                s_name = html.unescape(m_name.group(1)).strip()
                s_href = m_href.group(1).strip() if m_href else ""
                ws_href_to_name.append((s_href, s_name))

    # Mappe Parts anhand des Dateinamens / Content-Location
    sheets_out: List[Tuple[str, List[List[str]]]] = []
    used_part_indices: Set[int] = set()

    if ws_href_to_name and len(html_parts) > 1:
        for idx_ws, (href, s_name) in enumerate(ws_href_to_name, start=1):
            href_base = href.replace("\\", "/").split("/")[-1].lower() if href else ""
            matched_idx: Optional[int] = None
            if href_base:
                for p_i, (loc, _ht) in enumerate(html_parts):
                    loc_base = loc.replace("\\", "/").split("/")[-1].lower()
                    if p_i not in used_part_indices and (loc_base == href_base or href_base in loc.lower()):
                        matched_idx = p_i
                        break
            if matched_idx is not None:
                used_part_indices.add(matched_idx)
                extracted = _extract_tables_from_html_text(html_parts[matched_idx][1], default_sheet_prefix=s_name)
                if extracted:
                    clean_s = re.sub(r"[:\\/?*\[\]]+", "_", s_name).strip()[:31] or f"Sheet{idx_ws}"
                    sheets_out.append((clean_s, extracted[0][1]))

    # Extrahiere alle uebrigen HTML-Parts mit Datentabellen
    for p_i, (_loc, h_txt) in enumerate(html_parts):
        if p_i in used_part_indices:
            continue
        extracted = _extract_tables_from_html_text(h_txt, default_sheet_prefix="Sheet")
        for s_name, rows in extracted:
            if ws_href_to_name and len(sheets_out) < len(ws_href_to_name):
                mapped_name = ws_href_to_name[len(sheets_out)][1]
                s_name = re.sub(r"[:\\/?*\[\]]+", "_", mapped_name).strip()[:31] or s_name
            sheets_out.append((s_name, rows))

    if not sheets_out:
        raise ValueError("MHTML file did not contain any extractable worksheet tables.")

    # Eindeutige Sheet-Namen sicherstellen
    final_sheets: List[Tuple[str, List[List[str]]]] = []
    seen_names: Set[str] = set()
    for idx, (s_name, rows) in enumerate(sheets_out, start=1):
        unique_name = s_name
        if unique_name in seen_names:
            unique_name = f"{s_name[:27]}_{idx}"
        seen_names.add(unique_name)
        final_sheets.append((unique_name, rows))

    return final_sheets


def _local_xml_tag(tag: str) -> str:
    """Entfernt den XML-Namespace-Teil aus einem ElementTree-Tag (z. B. '{urn:...}Worksheet' -> 'Worksheet')."""
    if "}" in tag:
        return tag.split("}", 1)[1]
    if ":" in tag:
        return tag.split(":", 1)[1]
    return tag


def _get_xml_attr(elem: ET.Element, local_name: str) -> Optional[str]:
    """Liest ein XML-Attribut unabhaengig vom Namespace-Praefix (z. B. ss:Name, ss:Index)."""
    target = local_name.lower()
    for k, v in elem.attrib.items():
        k_loc = k.split("}", 1)[1] if "}" in k else (k.split(":", 1)[1] if ":" in k else k)
        if k_loc.lower() == target:
            return v
    return None


def _parse_xml2003_sheets(file_path: str) -> List[Tuple[str, List[List[str]]]]:
    """
    Liest eine Excel 2003 XML Spreadsheet-Datei (SpreadsheetML, haeufig mit .xls-Endung)
    ein und rekonstruiert alle Worksheets inkl. ss:Index-Versatz und MergeAcross/MergeDown.
    """
    raw_bytes = Path(file_path).read_bytes()
    xml_text = _decode_bytes_smart(raw_bytes).lstrip()
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError as exc:
        raise ValueError(f"Invalid Excel 2003 XML structure: {exc}") from exc

    sheets_out: List[Tuple[str, List[List[str]]]] = []
    ws_Idx = 0

    for elem in root.iter():
        if _local_xml_tag(elem.tag).lower() != "worksheet":
            continue
        ws_Idx += 1
        raw_name = _get_xml_attr(elem, "Name") or f"Sheet{ws_Idx}"
        clean_name = re.sub(r"[:\\/?*\[\]]+", "_", raw_name).strip()[:31] or f"Sheet{ws_Idx}"

        # Suche <Table> innerhalb des <Worksheet>
        table_elem = None
        for child in elem.iter():
            if _local_xml_tag(child.tag).lower() == "table":
                table_elem = child
                break
        if table_elem is None:
            continue

        grid_map: Dict[Tuple[int, int], str] = {}
        curr_r = 0
        max_r = 0
        max_c = 0

        for row_el in table_elem:
            if _local_xml_tag(row_el.tag).lower() != "row":
                continue
            r_idx_attr = _get_xml_attr(row_el, "Index")
            if r_idx_attr and r_idx_attr.isdigit():
                curr_r = max(curr_r, int(r_idx_attr) - 1)

            curr_c = 0
            for cell_el in row_el:
                if _local_xml_tag(cell_el.tag).lower() != "cell":
                    continue
                c_idx_attr = _get_xml_attr(cell_el, "Index")
                if c_idx_attr and c_idx_attr.isdigit():
                    curr_c = int(c_idx_attr) - 1
                else:
                    while (curr_r, curr_c) in grid_map:
                        curr_c += 1

                merge_across = int(_get_xml_attr(cell_el, "MergeAcross") or "0")
                merge_down   = int(_get_xml_attr(cell_el, "MergeDown") or "0")

                cell_text = ""
                for sub in cell_el:
                    if _local_xml_tag(sub.tag).lower() in ("data", "ss:data"):
                        cell_text = "".join(sub.itertext()).strip()
                        break
                if not cell_text:
                    cell_text = "".join(cell_el.itertext()).strip()

                for dr in range(merge_down + 1):
                    for dc in range(merge_across + 1):
                        grid_map[(curr_r + dr, curr_c + dc)] = cell_text if (dr == 0 and dc == 0) else ""

                curr_c += merge_across + 1
                if curr_c > max_c:
                    max_c = curr_c

            curr_r += 1
            if curr_r > max_r:
                max_r = curr_r

        rows_list: List[List[str]] = []
        for r_i in range(max_r):
            rows_list.append([grid_map.get((r_i, c_i), "") for c_i in range(max_c)])
        while rows_list and all(c == "" for c in rows_list[-1]):
            rows_list.pop()

        if rows_list:
            sheets_out.append((clean_name, rows_list))

    if not sheets_out:
        raise ValueError("No valid worksheet data found in Excel 2003 XML file.")
    return sheets_out


def _open_xls_workbook(path_str: str) -> Any:
    """Oeffnet eine binaere .xls-Datei mit xlrd (mit Fallback falls formatting_info nicht unterstuetzt wird)."""
    if xlrd is None:
        raise ImportError("xlrd package is required to read legacy .xls files.")
    try:
        return xlrd.open_workbook(path_str, formatting_info=True)
    except Exception:
        return xlrd.open_workbook(path_str)


def _load_all_workbook_sheets_raw(
    path_str: str,
) -> Tuple[str, List[Tuple[str, List[List[str]]]]] :
    """
    Erkennt das interne Dateiformat inhaltsbasiert und extrahiert alle Worksheets als
    (detected_format, [(sheet_name, rows), ...]).
    Wirft bei nicht lesbaren oder beschaedigten Dateien einen klaren ValueError.
    """
    ext = os.path.splitext(path_str)[1].lower()
    if ext not in (".xlsx", ".xls", ".xml", ".mht", ".mhtml", ".htm", ".html"):
        raise ValueError(
            f"Unsupported file extension '{ext}'. Please select an Excel file (.xlsx or .xls)."
        )

    if not os.path.isfile(path_str):
        raise ValueError(f"File not found: {path_str}")

    detected_fmt = detect_excel_file_format(path_str)

    try:
        if detected_fmt == "xlsx":
            wb = openpyxl.load_workbook(path_str, data_only=True)
            try:
                sheets_data = [
                    (s_name, _extract_sheet_rows_with_formats(wb[s_name]))
                    for s_name in wb.sheetnames
                ]
            finally:
                wb.close()
            return "xlsx", sheets_data

        if detected_fmt == "xls_biff":
            book = _open_xls_workbook(path_str)
            try:
                sheets_data = [
                    (s_name, _extract_xls_sheet_rows_with_formats(book.sheet_by_name(s_name), book))
                    for s_name in book.sheet_names()
                ]
            finally:
                if hasattr(book, "release_resources"):
                    book.release_resources()
            return "xls", sheets_data

        if detected_fmt == "mhtml":
            return "mhtml", _parse_mhtml_sheets(path_str)

        if detected_fmt == "html":
            return "html", _parse_html_sheets(path_str)

        if detected_fmt == "xml_2003":
            return "xml_2003", _parse_xml2003_sheets(path_str)

    except Exception as exc:
        print(f"[Excel Import Diagnostic] Failed to parse '{path_str}' (detected={detected_fmt}): {exc}")
        raise ValueError(
            f"The file '{os.path.basename(path_str)}' could not be read as a valid Excel workbook "
            f"(detected format: {detected_fmt}). The file may be corrupted or in an unsupported format."
        ) from exc

    print(f"[Excel Import Diagnostic] Unrecognized file signature for '{path_str}'.")
    raise ValueError(
        f"The file '{os.path.basename(path_str)}' is not a valid or supported Excel workbook. "
        f"Supported formats are native .xlsx, binary .xls, MHTML (.xls), HTML (.xls), and Excel 2003 XML (.xls)."
    )


def load_excel_flexible(
    path_or_df: Union[str, Path, pd.DataFrame],
    sheet_name: Optional[str] = None,
    article_column: Optional[str] = None,
) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Laedt eine Excel-Datei (.xlsx, binaeres .xls, MHTML-.xls, HTML-.xls oder Excel 2003 XML)
    oder ein DataFrame struktur-agnostisch und erkennt automatisch das beste Worksheet,
    die Tabellen-Kopfzeile und die Artikelnummern-Spalte.
    """
    if isinstance(path_or_df, pd.DataFrame):
        df = path_or_df.copy()
        for col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
        info = _evaluate_dataframe_columns(df)
        if article_column and article_column in df.columns:
            info["column"] = article_column
            info["is_ambiguous"] = False
        info["sheet_name"] = sheet_name or "Sheet1"
        info["available_sheets"] = [info["sheet_name"]]
        info["header_row"] = 0
        info["has_header"] = True
        info["file_format"] = "dataframe"
        return df, info

    path_str = str(path_or_df)
    detected_fmt, all_sheets = _load_all_workbook_sheets_raw(path_str)
    sheet_names = [s_name for s_name, _ in all_sheets]
    sheet_map = dict(all_sheets)
    target_sheets = [sheet_name] if (sheet_name and sheet_name in sheet_map) else sheet_names

    best_df: Optional[pd.DataFrame] = None
    best_info: Optional[Dict[str, Any]] = None

    for s_name in target_sheets:
        rows = sheet_map[s_name]
        df_cand, header_row_idx, has_header = _parse_sheet_table(rows)
        if not df_cand.empty:
            keep_cols = [
                c for c in df_cand.columns
                if not (str(c).startswith(("Unnamed:", "Column ")) and (df_cand[c] == "").all())
            ]
            if keep_cols:
                df_cand = df_cand[keep_cols]

        info_cand = _evaluate_dataframe_columns(df_cand)
        info_cand["sheet_name"] = s_name
        info_cand["available_sheets"] = sheet_names
        info_cand["header_row"] = header_row_idx
        info_cand["has_header"] = has_header
        info_cand["file_format"] = detected_fmt

        if best_info is None or (
            info_cand["confidence"],
            info_cand["candidates"][0]["valid_count"] if info_cand["candidates"] else 0,
        ) > (
            best_info["confidence"],
            best_info["candidates"][0]["valid_count"] if best_info["candidates"] else 0,
        ):
            best_df = df_cand
            best_info = info_cand

    if best_df is None or best_info is None:
        return pd.DataFrame(), {
            "column": None,
            "short_column": None,
            "long_column": None,
            "has_both_formats": False,
            "confidence": 0.0,
            "is_ambiguous": False,
            "candidates": [],
            "sheet_name": sheet_name or (sheet_names[0] if sheet_names else "Sheet1"),
            "available_sheets": sheet_names,
            "header_row": None,
            "has_header": True,
            "file_format": detected_fmt,
        }

    if article_column and article_column in best_df.columns:
        best_info["column"] = article_column
        best_info["is_ambiguous"] = False

    return best_df, best_info


def detect_article_column(
    path_or_df: Union[str, Path, pd.DataFrame],
    sheet_name: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Erkennt die Artikelnummern-Spalte in einer Excel-Datei (.xlsx / .xls) oder einem DataFrame.
    """
    _, info = load_excel_flexible(path_or_df, sheet_name=sheet_name)
    return info


# ── HTTP-Setup ───────────────────────────────────────────────────────────────

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

_thread_local = threading.local()


def _get_session() -> requests.Session:
    """Gibt eine thread-lokale requests.Session mit Retry-Logik zurueck."""
    if not hasattr(_thread_local, "session") or _thread_local.session is None:
        from requests.adapters import HTTPAdapter, Retry
        s = requests.Session()
        retry = Retry(
            total=2,
            connect=2,
            read=1,
            backoff_factor=0.2,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=["GET"],
        )
        s.mount("https://", HTTPAdapter(max_retries=retry))
        headers = {
            "User-Agent": "WAGO-AAS-PCF-Query/3.0",
            "Accept": "application/json",
        }
        if API_AUTH_TOKEN:
            headers["Authorization"] = (
                API_AUTH_TOKEN if API_AUTH_TOKEN.lower().startswith("bearer ")
                else f"Bearer {API_AUTH_TOKEN}"
            )
        if API_KEY:
            headers["X-API-Key"] = API_KEY
        s.headers.update(headers)
        s.verify = CERT_VERIFY
        _thread_local.session = s
    return _thread_local.session


def _set_last_error(reason: str) -> None:
    _thread_local.last_error_reason = reason


def _get_last_error() -> str:
    return getattr(_thread_local, "last_error_reason", INTERNAL_STATUS_QUERY_ERROR)


def _jget(url: str, timeout: Tuple[int, int] = REQUEST_TIMEOUT) -> Any:
    """
    HTTP GET -> JSON-Dict, _NOT_FOUND (404), oder _ERROR (techn. Fehler).
    """
    _set_last_error(INTERNAL_STATUS_QUERY_ERROR)
    try:
        r = _get_session().get(url, timeout=timeout)
        if r.status_code == 404:
            return _NOT_FOUND
        if r.status_code in (401, 403):
            _set_last_error(INTERNAL_STATUS_AUTH_ERROR)
            return _ERROR
        if r.status_code >= 500:
            _set_last_error(INTERNAL_STATUS_SERVER_ERROR)
            return _ERROR
        r.raise_for_status()
        try:
            data = r.json()
        except (ValueError, TypeError):
            _set_last_error(INTERNAL_STATUS_INVALID_RESP)
            return _ERROR
        if not isinstance(data, dict):
            _set_last_error(INTERNAL_STATUS_INVALID_RESP)
            return _ERROR
        return data
    except requests.exceptions.Timeout:
        _set_last_error(INTERNAL_STATUS_TIMEOUT)
        return _ERROR
    except requests.exceptions.HTTPError as e:
        code = getattr(e.response, "status_code", None)
        if code in (401, 403):
            _set_last_error(INTERNAL_STATUS_AUTH_ERROR)
        elif code and code >= 500:
            _set_last_error(INTERNAL_STATUS_SERVER_ERROR)
        else:
            _set_last_error(INTERNAL_STATUS_QUERY_ERROR)
        return _ERROR
    except requests.RequestException:
        _set_last_error(INTERNAL_STATUS_QUERY_ERROR)
        return _ERROR
    except (ValueError, TypeError):
        _set_last_error(INTERNAL_STATUS_INVALID_RESP)
        return _ERROR


def _safe_jget(url: str, timeout: Tuple[int, int] = REQUEST_TIMEOUT) -> Any:
    """Wrapper um _jget() fuer konsistente Exception-Behandlung."""
    try:
        return _jget(url, timeout=timeout)
    except requests.exceptions.Timeout:
        _set_last_error(INTERNAL_STATUS_TIMEOUT)
        return _ERROR
    except requests.exceptions.HTTPError as e:
        code = getattr(getattr(e, "response", None), "status_code", None)
        if code in (401, 403):
            _set_last_error(INTERNAL_STATUS_AUTH_ERROR)
        elif code and code >= 500:
            _set_last_error(INTERNAL_STATUS_SERVER_ERROR)
        else:
            _set_last_error(INTERNAL_STATUS_QUERY_ERROR)
        return _ERROR
    except requests.RequestException:
        _set_last_error(INTERNAL_STATUS_QUERY_ERROR)
        return _ERROR
    except Exception:
        _set_last_error(INTERNAL_STATUS_INVALID_RESP)
        return _ERROR


# ── AAS-Hilfsfunktionen & PCF-Einheitenlogik ─────────────────────────────────

def _b64_encode(value: str) -> str:
    """Base64-Encoding fuer AAS-Identifier in Pfaden."""
    return base64.b64encode(value.encode("utf-8")).decode("ascii")


def _b64url_encode(value: str) -> str:
    """URL-safe Base64-Encoding ohne Padding fuer AAS v3.2 Query-Parameter."""
    return base64.urlsafe_b64encode(value.encode("utf-8")).decode("ascii").rstrip("=")


def _submodel_url_from_id(submodel_id: str) -> str:
    """Konstruiert die Submodel-Endpoint-URL fuer eine gegebene Submodel-ID."""
    encoded = _b64_encode(submodel_id)
    return f"{AAS_ENV_BASE}/submodels/{encoded}"


def _submodel_url(article_number: str) -> str:
    """Konstruiert die direkte AAS-Submodel-URL fuer das CarbonFootprint-Submodel."""
    canonical = normalize_article_number(article_number)
    submodel_id = SUBMODEL_ID_TEMPLATE.format(article=canonical)
    return _submodel_url_from_id(submodel_id)


def _registry_lookup_url(article_number: str) -> str:
    """Konstruiert die Shell-Registry-Lookup-URL ueber globalAssetId (AAS v3.2 Objektformat)."""
    canonical = normalize_article_number(article_number)
    asset_id = ASSET_ID_TEMPLATE.format(article=canonical)
    asset_id_obj = {"name": "globalAssetId", "value": asset_id}
    asset_id_encoded = _b64url_encode(json.dumps(asset_id_obj))
    return f"{AAS_REGISTRY_BASE}/shell-descriptors?assetIds={urllib.parse.quote(asset_id_encoded)}"


def _extract_all_values(elements: Any, path: str = "") -> Dict[str, Any]:
    """
    Traversiert rekursiv alle submodelElements (inkl. SubmodelElementList und
    SubmodelElementCollection) und gibt ein Dict {Pfad: Wert} zurueck.
    """
    result: Dict[str, Any] = {}
    if not isinstance(elements, list):
        return result

    for el in elements:
        if not isinstance(el, dict):
            continue
        idshort  = el.get("idShort") or ""
        new_path = f"{path}/{idshort}".lstrip("/") if path else idshort

        children = el.get("submodelElements")
        if children is None and isinstance(el.get("value"), list):
            children = el["value"]

        if isinstance(children, list) and children:
            result.update(_extract_all_values(children, new_path))
        else:
            result[new_path] = el.get("value")

    return result


def _parse_float(raw: Any) -> Optional[float]:
    """Parst einen Zahlenwert; akzeptiert Komma als Dezimaltrenner."""
    if raw is None or isinstance(raw, bool):
        return None
    s = str(raw).strip().replace(",", ".")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _derive_pcf_unit(
    all_values: Dict[str, Any],
) -> Tuple[str, str, Optional[float]]:
    """
    Bestimmt die korrekte PCF-Einheit und Bezugsgroesse aus den extrahierten
    Submodel-Elementen (IDTA 02023 CarbonFootprint Submodel).

    - Standard bei ReferenceImpactUnitForCalculation='piece' und Quantity=1:
      ('kg CO2e', 'piece', 1.0)
    - Falls eine abweichende Bezugseinheit (z. B. 'kg', 'm') oder Menge (!= 1)
      angegeben ist, wird diese explizit in der Einheit ausgewiesen, ohne den
      Zahlenwert stillschweigend umzurechnen (z. B. 'kg CO2e / kg', 'kg CO2e / 100 piece').
    """
    ref_unit_raw = ""
    ref_qty_val: Optional[float] = None
    explicit_unit = ""

    for path, value in all_values.items():
        pl = path.lower()
        if "referenceimpactunitforcalculation" in pl and value is not None:
            ref_unit_raw = str(value).strip()
        elif "quantityofmeasureforcalculation" in pl and value is not None:
            ref_qty_val = _parse_float(value)
        elif pl.endswith("/pcfunit") or pl == "pcfunit":
            if value is not None and str(value).strip():
                explicit_unit = str(value).strip()

    base_emission_unit = explicit_unit if explicit_unit else DEFAULT_PCF_UNIT
    ref_unit_norm = ref_unit_raw.lower()
    is_single_qty = ref_qty_val is None or abs(ref_qty_val - 1.0) < 1e-9
    is_piece_unit = ref_unit_norm in ("", "piece", "pieces", "pc", "pcs", "item", "items", "stueck", "stück", "stk")

    if is_piece_unit and is_single_qty:
        return base_emission_unit, (ref_unit_raw or "piece"), (ref_qty_val if ref_qty_val is not None else 1.0)

    if not is_single_qty and ref_qty_val is not None:
        qty_str = str(int(ref_qty_val)) if float(ref_qty_val).is_integer() else str(ref_qty_val)
        u_str = ref_unit_raw or "piece"
        return f"{base_emission_unit} / {qty_str} {u_str}", u_str, ref_qty_val

    return f"{base_emission_unit} / {ref_unit_raw}", ref_unit_raw, (ref_qty_val if ref_qty_val is not None else 1.0)


def _evaluate_pcf_and_unit_detailed(
    result: Any,
) -> Tuple[Optional[float], str, str, str]:
    """
    Wertet eine Submodel-API-Antwort aus und gibt
    (pcf_value, pcf_unit, customer_status, internal_status) zurueck.
    """
    if result is _NOT_FOUND:
        return None, "", STATUS_NOT_YET_AVAILABLE, INTERNAL_STATUS_NOT_FOUND

    if result is _ERROR:
        return None, "", STATUS_QUERY_ERROR, _get_last_error()

    if not isinstance(result, dict):
        return None, "", STATUS_QUERY_ERROR, INTERNAL_STATUS_INVALID_RESP

    has_sm_elements = "submodelElements" in result and isinstance(result.get("submodelElements"), list)
    is_submodel_type = result.get("modelType") == "Submodel" and (
        result.get("submodelElements") is None or isinstance(result.get("submodelElements"), list)
    )
    if not has_sm_elements and not is_submodel_type:
        return None, "", STATUS_QUERY_ERROR, INTERNAL_STATUS_INVALID_RESP

    all_values = _extract_all_values(result.get("submodelElements") or [])
    for path, value in all_values.items():
        if "pcfco2eq" in path.lower():
            parsed = _parse_float(value)
            if parsed is not None:
                unit_str, _, _ = _derive_pcf_unit(all_values)
                return parsed, unit_str, STATUS_PCF_AVAILABLE, INTERNAL_STATUS_PCF_AVAILABLE

    return None, "", STATUS_NOT_YET_AVAILABLE, INTERNAL_STATUS_AAS_NO_PCF


def _evaluate_pcf_result_detailed(result: Any) -> Tuple[Optional[float], str, str]:
    """Gibt (pcf_value, customer_status, internal_status) zurueck."""
    pcf_val, _, cust_status, int_status = _evaluate_pcf_and_unit_detailed(result)
    return pcf_val, cust_status, int_status


def _evaluate_pcf_result(result: Any) -> Tuple[Optional[float], str]:
    """Gibt (pcf_value, customer_status) fuer eine Submodel-Antwort zurueck."""
    pcf_val, _, cust_status, _ = _evaluate_pcf_and_unit_detailed(result)
    return pcf_val, cust_status


def get_pcf_result(article_number: Any) -> Dict[str, Any]:
    """
    Gibt ein Dict mit PCF-Wert, PCF-Einheit, kundenfreundlichem Status,
    internem Diagnosestatus sowie Kurz-/Langformat der Artikelnummer zurueck.
    """
    parsed_art = parse_article_number(article_number)
    if not parsed_art["original"]:
        return {
            "pcf": None,
            "unit": "",
            "status": STATUS_NOT_YET_AVAILABLE,
            "internal_status": INTERNAL_STATUS_NOT_FOUND,
            "article_short": "",
            "article_long": "",
            "source": "WAGO AAS API",
        }

    query_article = parsed_art["canonical"]
    short_art     = parsed_art["short"]
    long_art      = parsed_art["long"]

    # 1. Fast Path: Direkter Abruf des CarbonFootprint-Submodels
    url = _submodel_url(query_article)
    fast_result = _safe_jget(url)

    pcf_val, pcf_unit, pcf_status, internal_status = _evaluate_pcf_and_unit_detailed(fast_result)
    if pcf_status == STATUS_PCF_AVAILABLE:
        return {
            "pcf": pcf_val,
            "unit": pcf_unit,
            "status": STATUS_PCF_AVAILABLE,
            "internal_status": INTERNAL_STATUS_PCF_AVAILABLE,
            "article_short": short_art,
            "article_long": long_art,
            "source": "WAGO AAS API",
        }

    if pcf_status == STATUS_QUERY_ERROR:
        return {
            "pcf": None,
            "unit": "",
            "status": STATUS_QUERY_ERROR,
            "internal_status": internal_status,
            "article_short": short_art,
            "article_long": long_art,
            "source": "WAGO AAS API",
        }

    # 2. Safe Fallback: Shell-Descriptor-Lookup via globalAssetId
    reg_url = _registry_lookup_url(query_article)
    reg_res = _safe_jget(reg_url, timeout=(10, 30))

    if reg_res is _ERROR:
        return {
            "pcf": None,
            "unit": "",
            "status": STATUS_QUERY_ERROR,
            "internal_status": _get_last_error(),
            "article_short": short_art,
            "article_long": long_art,
            "source": "WAGO AAS API",
        }

    if reg_res is _NOT_FOUND:
        return {
            "pcf": None,
            "unit": "",
            "status": STATUS_NOT_YET_AVAILABLE,
            "internal_status": INTERNAL_STATUS_NOT_FOUND,
            "article_short": short_art,
            "article_long": long_art,
            "source": "WAGO AAS API",
        }

    if not isinstance(reg_res, dict):
        return {
            "pcf": None,
            "unit": "",
            "status": STATUS_QUERY_ERROR,
            "internal_status": INTERNAL_STATUS_INVALID_RESP,
            "article_short": short_art,
            "article_long": long_art,
            "source": "WAGO AAS API",
        }

    if isinstance(reg_res.get("result"), list):
        shells = reg_res["result"]
    elif "submodelDescriptors" in reg_res or "submodels" in reg_res:
        shells = [reg_res]
    else:
        return {
            "pcf": None,
            "unit": "",
            "status": STATUS_QUERY_ERROR,
            "internal_status": INTERNAL_STATUS_INVALID_RESP,
            "article_short": short_art,
            "article_long": long_art,
            "source": "WAGO AAS API",
        }

    if not shells:
        return {
            "pcf": None,
            "unit": "",
            "status": STATUS_NOT_YET_AVAILABLE,
            "internal_status": internal_status,
            "article_short": short_art,
            "article_long": long_art,
            "source": "WAGO AAS API",
        }

    for shell in shells:
        if not isinstance(shell, dict):
            continue
        submodels = shell.get("submodelDescriptors") or shell.get("submodels") or []
        if not isinstance(submodels, list):
            continue
        for sm in submodels:
            if not isinstance(sm, dict):
                continue
            sm_id = sm.get("id")
            if not sm_id and isinstance(sm.get("keys"), list) and sm["keys"]:
                first_key = sm["keys"][0]
                if isinstance(first_key, dict):
                    sm_id = first_key.get("value")
            id_short = str(sm.get("idShort") or "").lower()
            sm_id_lower = str(sm_id or "").lower()

            if ("pcf" in id_short or "carbonfootprint" in id_short
                    or "/pcf" in sm_id_lower or "carbonfootprint" in sm_id_lower):
                if sm_id:
                    sm_url = _submodel_url_from_id(sm_id)
                    sm_res = _safe_jget(sm_url, timeout=(10, 30))
                    fb_val, fb_unit, fb_status, fb_internal = _evaluate_pcf_and_unit_detailed(sm_res)
                    if fb_status == STATUS_PCF_AVAILABLE:
                        return {
                            "pcf": fb_val,
                            "unit": fb_unit,
                            "status": STATUS_PCF_AVAILABLE,
                            "internal_status": INTERNAL_STATUS_PCF_AVAILABLE,
                            "article_short": short_art,
                            "article_long": long_art,
                            "source": "WAGO AAS API",
                        }
                    if fb_status == STATUS_QUERY_ERROR:
                        return {
                            "pcf": None,
                            "unit": "",
                            "status": STATUS_QUERY_ERROR,
                            "internal_status": fb_internal,
                            "article_short": short_art,
                            "article_long": long_art,
                            "source": "WAGO AAS API",
                        }

    return {
        "pcf": None,
        "unit": "",
        "status": STATUS_NOT_YET_AVAILABLE,
        "internal_status": INTERNAL_STATUS_AAS_NO_PCF,
        "article_short": short_art,
        "article_long": long_art,
        "source": "WAGO AAS API",
    }


def get_pcf(article_number: Any) -> Optional[float]:
    """Rueckwaerts-kompatibler Wrapper: gibt nur den PCF-Wert (float | None) zurueck."""
    return get_pcf_result(article_number)["pcf"]


def _query_with_timing(article_number: str, mode: str = MODE_ONLINE) -> Dict[str, Any]:
    """
    Fuehrt die PCF-Abfrage im gewaehlten Modus (Online oder Offline) aus
    und loggt langsame Requests. Geeignet fuer Worker-Threads.
    """
    if mode == MODE_OFFLINE:
        return get_offline_pcf_result(article_number)

    t0 = time.perf_counter()
    try:
        res = get_pcf_result(article_number)
    except Exception as e:
        print(f"  Unerwarteter Fehler fuer {article_number}: {e}")
        parsed = parse_article_number(article_number)
        res = {
            "pcf": None,
            "unit": "",
            "status": STATUS_QUERY_ERROR,
            "internal_status": INTERNAL_STATUS_QUERY_ERROR,
            "article_short": parsed["short"],
            "article_long": parsed["long"],
            "source": "WAGO AAS API",
        }
    dur = time.perf_counter() - t0
    if dur >= SLOW_REQUEST_THRESHOLD:
        diag = res.get("internal_status") or res.get("status")
        print(f"  Slow query: {article_number} - {dur:.1f}s [{diag}]")
    return res


# ── Single Query Workflow (Section 6.1) ──────────────────────────────────────

def query_single_article(article_number: Any, mode: str = MODE_ONLINE) -> Dict[str, Any]:
    """
    Fuehrt eine Einzelabfrage fuer eine WAGO-Artikelnummer im Online- oder Offline-Modus
    durch und gibt ein strukturiertes Ergebnis-Dict fuer GUI und Zwischenablage zurueck.
    """
    active_mode = MODE_OFFLINE if str(mode).strip().lower() == MODE_OFFLINE else MODE_ONLINE
    parsed = parse_article_number(article_number)
    raw_input = parsed["original"]

    if not raw_input:
        return {
            "article_input": "",
            "article_short": "",
            "article_long": "",
            "pcf": None,
            "pcf_display": STATUS_NOT_YET_AVAILABLE,
            "unit": "",
            "status": STATUS_NOT_YET_AVAILABLE,
            "internal_status": INTERNAL_STATUS_INVALID_ART,
            "data_source": (
                "WAGO AAS API (Online)"
                if active_mode == MODE_ONLINE
                else f"Embedded Snapshot ({OFFLINE_SNAPSHOT_VERSION} Offline)"
            ),
            "mode": active_mode,
        }

    res = _query_with_timing(raw_input, mode=active_mode)
    pcf_val = res.get("pcf")
    status  = res.get("status", STATUS_NOT_YET_AVAILABLE)

    if pcf_val is not None:
        pcf_display = str(pcf_val)
        unit_str = res.get("unit") or DEFAULT_PCF_UNIT
    elif status == STATUS_QUERY_ERROR:
        pcf_display = STATUS_QUERY_ERROR
        unit_str = ""
    else:
        pcf_display = STATUS_NOT_YET_AVAILABLE
        unit_str = ""

    return {
        "article_input": raw_input,
        "article_short": res.get("article_short") or parsed["short"],
        "article_long": res.get("article_long") or parsed["long"],
        "pcf": pcf_val,
        "pcf_display": pcf_display,
        "unit": unit_str,
        "status": status,
        "internal_status": res.get("internal_status", status),
        "data_source": (
            "WAGO AAS API (Online)"
            if active_mode == MODE_ONLINE
            else f"Embedded Snapshot ({OFFLINE_SNAPSHOT_VERSION} Offline)"
        ),
        "mode": active_mode,
    }


def format_single_result_for_clipboard(result: Dict[str, Any], value_only: bool = False) -> str:
    """Formatiert ein Single-Query-Ergebnis fuer die Zwischenablage."""
    if value_only:
        if result.get("pcf") is not None:
            return str(result["pcf"])
        return str(result.get("pcf_display") or STATUS_NOT_YET_AVAILABLE)

    lines = [
        f"Article Number (Short): {result.get('article_short', '')}",
        f"Article Number (Long):  {result.get('article_long', '')}",
        f"PCF:                    {result.get('pcf_display', STATUS_NOT_YET_AVAILABLE)}",
        f"PCF Unit:               {result.get('unit', '')}",
        f"Data Source:            {result.get('data_source', '')}",
    ]
    return "\n".join(lines)


# ── Excel-Aufbereitung & Standardisierter Export ('PCF Results' Worksheet) ───

def _extract_row_articles(
    df: pd.DataFrame,
    detection_info: Dict[str, Any],
    article_column: Optional[str] = None,
) -> Tuple[List[str], List[str], List[str]]:
    """
    Extrahiert fuer jede Zeile des DataFrames:
        (canonical_articles, short_articles, long_articles)
    """
    primary_col = article_column or detection_info.get("column")
    short_col   = detection_info.get("short_column")
    long_col    = detection_info.get("long_column")
    has_both    = bool(detection_info.get("has_both_formats")) and not article_column

    if not has_both:
        col_lookup = {str(c).strip().lower(): c for c in df.columns}
        if COL_ARTICLE_SHORT.lower() in col_lookup and COL_ARTICLE_LONG.lower() in col_lookup:
            short_col = col_lookup[COL_ARTICLE_SHORT.lower()]
            long_col  = col_lookup[COL_ARTICLE_LONG.lower()]
            has_both  = True

    canonical_list: List[str] = []
    short_list: List[str]     = []
    long_list: List[str]      = []

    if has_both and short_col in df.columns and long_col in df.columns:
        s_vals = df[short_col].fillna("").astype(str).str.strip().tolist()
        l_vals = df[long_col].fillna("").astype(str).str.strip().tolist()
        for sv, lv in zip(s_vals, l_vals):
            ps = parse_article_number(sv)
            pl = parse_article_number(lv)
            if ps["is_valid"] and pl["is_valid"]:
                if ps["canonical"] == pl["canonical"]:
                    canonical_list.append(ps["canonical"])
                    short_list.append(ps["short"])
                    long_list.append(pl["long"])
                else:
                    p_main = ps if primary_col == short_col else pl
                    canonical_list.append(p_main["canonical"])
                    short_list.append(p_main["short"])
                    long_list.append(p_main["long"])
            elif ps["is_valid"]:
                canonical_list.append(ps["canonical"])
                short_list.append(ps["short"])
                long_list.append(ps["long"])
            elif pl["is_valid"]:
                canonical_list.append(pl["canonical"])
                short_list.append(pl["short"])
                long_list.append(pl["long"])
            else:
                raw = sv or lv
                canonical_list.append(raw)
                short_list.append(sv if sv else lv)
                long_list.append(lv if lv else sv)
        return canonical_list, short_list, long_list

    if primary_col is None or primary_col not in df.columns:
        return [], [], []

    raw_vals = df[primary_col].fillna("").astype(str).str.strip().tolist()
    for rv in raw_vals:
        p = parse_article_number(rv)
        if not p["original"]:
            canonical_list.append("")
            short_list.append("")
            long_list.append("")
        else:
            canonical_list.append(p["canonical"])
            short_list.append(p["short"])
            long_list.append(p["long"])

    return canonical_list, short_list, long_list


def prepare_output_dataframe(
    df: pd.DataFrame,
    detection_info: Dict[str, Any],
    cache: Dict[str, Dict[str, Any]],
    article_column: Optional[str] = None,
) -> pd.DataFrame:
    """
    Erzeugt das standardisierte 4-spaltige Ergebnis-DataFrame fuer das Worksheet 'PCF Results':
        | Article Number (Short) | Article Number (Long) | PCF | PCF Unit |

    Hinweis:
    - Keine separate 'PCF Status'-Spalte mehr (Section 2.2 & Section 4).
    - Bei fehlendem PCF steht 'Not yet available' direkt in der PCF-Zelle und PCF Unit bleibt leer.
    """
    canonical_list, short_list, long_list = _extract_row_articles(
        df, detection_info, article_column=article_column
    )

    rows_out: List[Dict[str, Any]] = []
    for canonical, short_val, long_val in zip(canonical_list, short_list, long_list):
        if not canonical:
            continue
        res = cache.get(canonical, {"pcf": None, "unit": "", "status": STATUS_NOT_YET_AVAILABLE})
        pcf_val = res.get("pcf")
        if pcf_val is not None:
            pcf_cell: Any = pcf_val
            unit_cell = res.get("unit") or DEFAULT_PCF_UNIT
        else:
            pcf_cell  = STATUS_NOT_YET_AVAILABLE
            unit_cell = ""

        rows_out.append({
            COL_ARTICLE_SHORT: short_val,
            COL_ARTICLE_LONG:  long_val,
            COL_PCF:           pcf_cell,
            COL_PCF_UNIT:      unit_cell,
        })

    return pd.DataFrame(
        rows_out,
        columns=[COL_ARTICLE_SHORT, COL_ARTICLE_LONG, COL_PCF, COL_PCF_UNIT],
    )


def _populate_pcf_results_worksheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    results_df: pd.DataFrame,
    mode: str = MODE_ONLINE,
    export_timestamp: Optional[str] = None,
) -> None:
    """
    Befuellt und formatiert das Worksheet 'PCF Results' im WAGO Corporate Design
    inkl. Information Header (Zeilen 1-9) und Datentabelle (ab Zeile 11).
    """
    active_mode = MODE_OFFLINE if str(mode).strip().lower() == MODE_OFFLINE else MODE_ONLINE
    ts_str = export_timestamp or datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    total_queried = len(results_df)
    available_cnt = sum(
        1 for v in results_df[COL_PCF].tolist() if isinstance(v, (int, float))
    ) if not results_df.empty else 0
    missing_cnt = total_queried - available_cnt

    if active_mode == MODE_OFFLINE:
        source_str   = "Embedded Offline Dataset (WAGO PCF Query Tool)"
        currency_str = f"Data snapshot: {OFFLINE_SNAPSHOT_VERSION}"
        notice_str   = OFFLINE_EXCEL_HEADER_NOTE
    else:
        source_str   = "WAGO AAS API (retrieved from WAGO Asset Administration Shell at query time)"
        currency_str = "Live query via WAGO SmartService AAS API (export timestamp is not the calculation date)"
        notice_str   = "PCF values were retrieved from the WAGO AAS API at the time of the query. Current PCF values are available in the WAGO AAS Portal."

    # Styles (WAGO Corporate Design)
    c_anthracite    = "1F2837"
    c_green         = "6EC800"
    c_light_grey    = "EFF0F1"
    c_anthracite_15 = "DEDFE1"
    c_anthracite_70 = "616A73"
    c_white         = "FFFFFF"

    font_title = Font(name="Segoe UI", size=14, bold=True, color=c_white)
    font_lbl   = Font(name="Segoe UI", size=10, bold=True, color=c_anthracite)
    font_val   = Font(name="Segoe UI", size=10, color=c_anthracite)
    font_note  = Font(name="Segoe UI", size=9, italic=True, color=c_anthracite_70)
    font_link  = Font(name="Segoe UI", size=10, underline="single", color="0563C1")
    font_thdr  = Font(name="Segoe UI", size=10, bold=True, color=c_white)
    font_tcell = Font(name="Segoe UI", size=10, color=c_anthracite)

    fill_title = PatternFill(start_color=c_anthracite, end_color=c_anthracite, fill_type="solid")
    fill_info  = PatternFill(start_color=c_light_grey, end_color=c_light_grey, fill_type="solid")
    fill_thdr  = PatternFill(start_color=c_green, end_color=c_green, fill_type="solid")
    fill_zebra = PatternFill(start_color=c_light_grey, end_color=c_light_grey, fill_type="solid")

    thin_border = Border(
        left=Side(style="thin", color=c_anthracite_15),
        right=Side(style="thin", color=c_anthracite_15),
        top=Side(style="thin", color=c_anthracite_15),
        bottom=Side(style="thin", color=c_anthracite_15),
    )

    # Row 1: Title Banner
    ws.merge_cells("A1:D1")
    c_banner = ws["A1"]
    c_banner.value = "WAGO Product Carbon Footprint Data"
    c_banner.font = font_title
    c_banner.fill = fill_title
    c_banner.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 28

    # Rows 2-8: Information Header Fields
    info_fields = [
        ("Export Date & Time", ts_str),
        ("Data Source", source_str),
        ("Data Currency", currency_str),
        ("Number of Queried Products", total_queried),
        ("Products with Available PCF", available_cnt),
        ("Products without Available PCF", missing_cnt),
        ("WAGO AAS Portal", PORTAL_URL),
    ]

    for idx, (label, val) in enumerate(info_fields, start=2):
        ws.merge_cells(start_row=idx, start_column=2, end_row=idx, end_column=4)
        c_l = ws.cell(idx, 1, label)
        c_v = ws.cell(idx, 2, val)
        c_l.font = font_lbl
        c_l.fill = fill_info
        c_v.font = font_val
        c_v.alignment = Alignment(horizontal="left", vertical="center")
        if label == "WAGO AAS Portal":
            c_v.hyperlink = PORTAL_URL
            c_v.font = font_link
        for col_i in range(1, 5):
            ws.cell(idx, col_i).border = thin_border
        ws.row_dimensions[idx].height = 19

    # Row 9: Data Currency / Snapshot Notice
    ws.merge_cells("A9:D9")
    c_notice = ws["A9"]
    c_notice.value = notice_str
    c_notice.font = font_note
    c_notice.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[9].height = 22

    # Row 11: Table Headers
    table_header_row = 11
    headers = [COL_ARTICLE_SHORT, COL_ARTICLE_LONG, COL_PCF, COL_PCF_UNIT]
    ws.row_dimensions[table_header_row].height = 24
    for c_idx, h_name in enumerate(headers, start=1):
        c_h = ws.cell(table_header_row, c_idx, h_name)
        c_h.font = font_thdr
        c_h.fill = fill_thdr
        c_h.alignment = Alignment(horizontal="left", vertical="center")
        c_h.border = thin_border

    # Rows 12+: Table Data
    for r_offset, (_, row_data) in enumerate(results_df.iterrows()):
        r_idx = table_header_row + 1 + r_offset
        is_even = r_offset % 2 == 1

        # Col 1: Article Number (Short) - explicitly text
        c1 = ws.cell(r_idx, 1, str(row_data[COL_ARTICLE_SHORT]))
        c1.data_type = "s"
        c1.number_format = "@"

        # Col 2: Article Number (Long) - explicitly text
        c2 = ws.cell(r_idx, 2, str(row_data[COL_ARTICLE_LONG]))
        c2.data_type = "s"
        c2.number_format = "@"

        # Col 3: PCF - numeric float if available, else 'Not yet available'
        pcf_val = row_data[COL_PCF]
        c3 = ws.cell(r_idx, 3)
        if isinstance(pcf_val, (int, float)) and not isinstance(pcf_val, bool):
            c3.value = float(pcf_val)
            c3.alignment = Alignment(horizontal="right", vertical="center")
        else:
            c3.value = str(pcf_val)
            c3.data_type = "s"

        # Col 4: PCF Unit - string unit if available, else empty
        unit_val = row_data[COL_PCF_UNIT]
        c4 = ws.cell(r_idx, 4, str(unit_val) if unit_val else "")
        c4.data_type = "s"

        for cell in (c1, c2, c3, c4):
            cell.font = font_tcell
            if is_even:
                cell.fill = fill_zebra
            cell.border = thin_border

    if results_df.empty:
        ws.cell(table_header_row + 1, 1, None)
    last_row = max(table_header_row + len(results_df), table_header_row + 1)
    ws.auto_filter.ref = f"A{table_header_row}:D{last_row}"
    ws.freeze_panes = f"A{table_header_row + 1}"
    if ws.views and ws.views.sheetView:
        sv = ws.views.sheetView[0]
        sv.tabSelected = True
        if sv.selection:
            sv.selection[0].activeCell = f"A{table_header_row + 1}"
            sv.selection[0].sqref = f"A{table_header_row + 1}"

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 18


def _copy_xls_sheets_to_openpyxl_workbook(
    xls_path: str,
    target_wb: openpyxl.Workbook,
) -> None:
    """
    Kopiert alle Original-Worksheets einer Legacy-.xls-Datei inhalts- und strukturgetreu
    in ein openpyxl-Workbook (Werte, Datentypen, zusammengefasste Zellen, Spaltenbreiten).
    """
    book = _open_xls_workbook(xls_path)
    try:
        for s_name in book.sheet_names():
            sh = book.sheet_by_name(s_name)
            # Falls der Name 'PCF Results' bereits vergeben ist, Original-Sheet-Namen bewahren
            ws_title = s_name if s_name not in target_wb.sheetnames else f"{s_name} (Original)"
            ws_out = target_wb.create_sheet(title=ws_title)

            for r_idx in range(sh.nrows):
                for c_idx in range(sh.ncols):
                    cell = sh.cell(r_idx, c_idx)
                    ctype = cell.ctype
                    val = cell.value
                    out_cell = ws_out.cell(r_idx + 1, c_idx + 1)

                    if ctype in (0, 6) or val is None:
                        out_cell.value = None
                    elif ctype == 1:  # XL_CELL_TEXT
                        out_cell.value = str(val)
                        out_cell.data_type = "s"
                        out_cell.number_format = "@"
                    elif ctype == 2:  # XL_CELL_NUMBER
                        if isinstance(val, float) and val.is_integer():
                            out_cell.value = int(val)
                        else:
                            out_cell.value = float(val)
                    elif ctype == 3 and xlrd is not None:  # XL_CELL_DATE
                        try:
                            out_cell.value = xlrd.xldate_as_datetime(val, book.datemode)
                            out_cell.number_format = "yyyy-mm-dd"
                        except Exception:
                            out_cell.value = val
                    elif ctype == 4:  # XL_CELL_BOOLEAN
                        out_cell.value = bool(val)
                    else:
                        out_cell.value = str(val)

            # Zusammengefasste Zellen uebernehmen (falls verfuegbar)
            for rlo, rhi, clo, chi in getattr(sh, "merged_cells", []):
                try:
                    ws_out.merge_cells(
                        start_row=rlo + 1,
                        start_column=clo + 1,
                        end_row=rhi,
                        end_column=chi,
                    )
                except Exception:
                    pass
    finally:
        if hasattr(book, "release_resources"):
            book.release_resources()


def _copy_parsed_text_sheets_to_openpyxl_workbook(
    sheets_data: List[Tuple[str, List[List[str]]]],
    target_wb: openpyxl.Workbook,
) -> None:
    """
    Rekonstruiert die Original-Worksheets eines MHTML-, HTML- oder Excel-2003-XML-Exports
    in einem openpyxl-Workbook (inkl. Erhalt fuehrender Nullen, Artikelnummern-Textformat '@',
    Reihenfolge der Tabellenblaetter und Spaltenbreiten).
    """
    for s_name, rows in sheets_data:
        ws_title = s_name if s_name not in target_wb.sheetnames else f"{s_name} (Original)"
        ws_out = target_wb.create_sheet(title=ws_title[:31])
        col_max_len: Dict[int, int] = {}

        for r_idx, row_vals in enumerate(rows, start=1):
            for c_idx, val in enumerate(row_vals, start=1):
                s_val = str(val).strip() if val is not None else ""
                out_cell = ws_out.cell(r_idx, c_idx)
                if not s_val:
                    out_cell.value = None
                    continue

                col_max_len[c_idx] = max(col_max_len.get(c_idx, 8), min(45, len(s_val) + 3))

                # Bewahre Werte mit fuehrenden Nullen (z. B. '0221-0412', '0010') oder
                # WAGO-Artikelnummern strikt als Text ('@'), damit keine Nullen verloren gehen
                if (
                    (len(s_val) > 1 and s_val.startswith("0") and not s_val.startswith(("0.", "0,")))
                    or "-" in s_val
                    or "/" in s_val
                    or is_valid_wago_article(s_val)
                ):
                    out_cell.value = s_val
                    out_cell.data_type = "s"
                    out_cell.number_format = "@"
                elif re.fullmatch(r"[+-]?\d+", s_val):
                    try:
                        out_cell.value = int(s_val)
                    except ValueError:
                        out_cell.value = s_val
                        out_cell.data_type = "s"
                elif re.fullmatch(r"[+-]?\d+\.\d+", s_val):
                    try:
                        out_cell.value = float(s_val)
                    except ValueError:
                        out_cell.value = s_val
                        out_cell.data_type = "s"
                else:
                    out_cell.value = s_val
                    out_cell.data_type = "s"
                    out_cell.number_format = "@"

        for c_idx, width in col_max_len.items():
            ws_out.column_dimensions[get_column_letter(c_idx)].width = width


def export_batch_pcf_excel(
    input_file: Union[str, Path],
    output_file: Union[str, Path],
    results_df: pd.DataFrame,
    mode: str = MODE_ONLINE,
) -> str:
    """
    Schreibt das Ergebnis einer Batch-Abfrage als neue .xlsx-Datei:
    1. 'PCF Results' steht immer als erstes Worksheet an Position 0.
    2. Alle bestehenden Worksheets der Eingabedatei (.xlsx, binaeres .xls, MHTML-.xls,
       HTML-.xls oder Excel 2003 XML) bleiben unveraendert bzw. inhaltsgetreu rekonstruiert
       und in ihrer relativen Originalreihenfolge erhalten.
    3. Ueberschreibt niemals die Original-Eingabedatei.
    """
    in_path = Path(input_file).resolve()
    out_path = Path(output_file).resolve()

    if out_path.suffix.lower() != ".xlsx":
        out_path = out_path.with_suffix(".xlsx")

    if out_path == in_path:
        out_path = in_path.with_name(f"{in_path.stem}_PCF.xlsx")

    detected_fmt = detect_excel_file_format(in_path) if in_path.exists() else "unknown"

    if detected_fmt == "xlsx":
        wb = openpyxl.load_workbook(str(in_path), data_only=False)
        if PCF_RESULTS_SHEET_NAME in wb.sheetnames:
            old_ws = wb[PCF_RESULTS_SHEET_NAME]
            old_ws.title = f"{PCF_RESULTS_SHEET_NAME} (Input)"
        for existing_ws in wb.worksheets:
            if existing_ws.views and existing_ws.views.sheetView:
                for sv in existing_ws.views.sheetView:
                    sv.tabSelected = False
        try:
            for defn in wb.defined_names.values():
                if getattr(defn, "localSheetId", None) is not None:
                    defn.localSheetId = int(defn.localSheetId) + 1
        except Exception:
            pass
        ws_pcf = wb.create_sheet(title=PCF_RESULTS_SHEET_NAME, index=0)
        wb.active = 0
        _populate_pcf_results_worksheet(ws_pcf, results_df, mode=mode)
        wb.save(str(out_path))
        wb.close()
    elif detected_fmt == "xls_biff":
        wb = openpyxl.Workbook()
        ws_pcf = wb.active
        ws_pcf.title = PCF_RESULTS_SHEET_NAME
        _populate_pcf_results_worksheet(ws_pcf, results_df, mode=mode)
        _copy_xls_sheets_to_openpyxl_workbook(str(in_path), wb)
        for other_ws in wb.worksheets[1:]:
            if other_ws.views and other_ws.views.sheetView:
                for sv in other_ws.views.sheetView:
                    sv.tabSelected = False
        wb.active = 0
        wb.save(str(out_path))
        wb.close()
    elif detected_fmt in ("mhtml", "html", "xml_2003"):
        wb = openpyxl.Workbook()
        ws_pcf = wb.active
        ws_pcf.title = PCF_RESULTS_SHEET_NAME
        _populate_pcf_results_worksheet(ws_pcf, results_df, mode=mode)
        try:
            _, sheets_data = _load_all_workbook_sheets_raw(str(in_path))
            _copy_parsed_text_sheets_to_openpyxl_workbook(sheets_data, wb)
        except Exception as exc:
            print(f"[Export Diagnostic] Could not reconstruct text-based input sheets: {exc}")
        for other_ws in wb.worksheets[1:]:
            if other_ws.views and other_ws.views.sheetView:
                for sv in other_ws.views.sheetView:
                    sv.tabSelected = False
        wb.active = 0
        wb.save(str(out_path))
        wb.close()
    else:
        wb = openpyxl.Workbook()
        ws_pcf = wb.active
        ws_pcf.title = PCF_RESULTS_SHEET_NAME
        _populate_pcf_results_worksheet(ws_pcf, results_df, mode=mode)
        wb.active = 0
        wb.save(str(out_path))
        wb.close()

    return str(out_path)


def read_pcf_results_table(excel_path: Union[str, Path]) -> pd.DataFrame:
    """
    Liest die standardisierte Ergebnistabelle ('Article Number (Short)',
    'Article Number (Long)', 'PCF', 'PCF Unit') aus dem Worksheet 'PCF Results'
    einer exportierten .xlsx-Datei ein.
    """
    wb = openpyxl.load_workbook(str(excel_path), data_only=True)
    try:
        ws = wb[PCF_RESULTS_SHEET_NAME] if PCF_RESULTS_SHEET_NAME in wb.sheetnames else wb.worksheets[0]
        header_row_idx = None
        for r_idx in range(1, min(ws.max_row + 1, 25)):
            c1 = str(ws.cell(r_idx, 1).value or "").strip()
            if c1 == COL_ARTICLE_SHORT:
                header_row_idx = r_idx
                break

        if header_row_idx is None:
            return pd.DataFrame(columns=[COL_ARTICLE_SHORT, COL_ARTICLE_LONG, COL_PCF, COL_PCF_UNIT])

        headers = [
            str(ws.cell(header_row_idx, c).value or "").strip()
            for c in range(1, 5)
        ]
        data_rows: List[List[Any]] = []
        for r_idx in range(header_row_idx + 1, ws.max_row + 1):
            row_vals = [ws.cell(r_idx, c).value for c in range(1, 5)]
            if all(v is None or str(v).strip() == "" for v in row_vals):
                continue
            data_rows.append([
                "" if row_vals[0] is None else str(row_vals[0]),
                "" if row_vals[1] is None else str(row_vals[1]),
                row_vals[2],
                "" if row_vals[3] is None else str(row_vals[3]),
            ])
        return pd.DataFrame(data_rows, columns=headers)
    finally:
        wb.close()


# ── Hauptprogramm (Batch Query, ohne GUI) ────────────────────────────────────

def main(
    input_file: str,
    output_file: str,
    progress_callback: Optional[Callable[[int, int, str, Dict[str, Any]], None]] = None,
    article_column: Optional[str] = None,
    sheet_name: Optional[str] = None,
    mode: str = MODE_ONLINE,
) -> Optional[Dict[str, Any]]:
    """
    Verarbeitet eine Excel-Datei (.xlsx oder .xls) mit WAGO-Artikelnummern und
    schreibt das Ergebnis in eine neue .xlsx-Datei mit dem ersten Worksheet
    'PCF Results' sowie allen unveraenderten Original-Worksheets.
    """
    active_mode = MODE_OFFLINE if str(mode).strip().lower() == MODE_OFFLINE else MODE_ONLINE
    print(f"Lese {input_file} (Modus: {MODE_LABELS.get(active_mode, active_mode)}) ...")
    if not os.path.exists(input_file):
        print(f"Fehler: Datei nicht gefunden: {input_file}")
        return None

    try:
        df, detection_info = load_excel_flexible(
            input_file, sheet_name=sheet_name, article_column=article_column
        )
    except Exception as e:
        print(f"Fehler beim Lesen von {input_file}: {e}")
        return None

    detected_col = article_column or detection_info.get("column")
    if df.empty or not detected_col or detected_col not in df.columns:
        print(f"Fehler: Keine gueltige Artikelnummern-Spalte in {input_file} gefunden.")
        return None

    canonical_all, _, _ = _extract_row_articles(
        df, detection_info, article_column=article_column
    )
    valid_canonical = [a for a in canonical_all if a]
    unique_articles = list(dict.fromkeys(valid_canonical))
    unique_total    = len(unique_articles)
    total           = len(canonical_all)

    print(
        f"Erkannte Spalte: '{detected_col}' "
        f"(Sheet: '{detection_info.get('sheet_name')}', Confidence: {detection_info.get('confidence', 0):.0%})"
    )
    print(f"Gefunden: {total} Zeilen, {unique_total} eindeutige Produkte (normalisiert)")
    if active_mode == MODE_OFFLINE:
        print(f"{OFFLINE_NOTICE_TITLE}: {OFFLINE_NOTICE_TEXT}\n")
    else:
        print(f"Starte Abfrage mit {WORKERS} parallelen Threads ...\n")

    cache: Dict[str, Dict[str, Any]] = {}
    cache_lock = threading.Lock()
    done_count = [0]

    def _worker(article: str) -> Tuple[str, Dict[str, Any]]:
        res = _query_with_timing(article, mode=active_mode)
        with cache_lock:
            cache[article] = res
            done_count[0] += 1
            done = done_count[0]
        if progress_callback:
            progress_callback(done, unique_total, article, res)
        return article, res

    if unique_total > 0:
        with ThreadPoolExecutor(max_workers=WORKERS) as executor:
            futures = {executor.submit(_worker, a): a for a in unique_articles}
            for future in as_completed(futures):
                article, res = future.result()
                if res["pcf"] is not None:
                    u_str = res.get("unit") or DEFAULT_PCF_UNIT
                    pcf_str = f"{res['pcf']} {u_str}"
                elif res["status"] == STATUS_QUERY_ERROR:
                    pcf_str = STATUS_QUERY_ERROR
                else:
                    pcf_str = STATUS_NOT_YET_AVAILABLE
                print(f"  {article:<25} {pcf_str:<22} [{res['status']}]")

    results_df = prepare_output_dataframe(
        df, detection_info, cache, article_column=article_column
    )

    try:
        final_output_path = export_batch_pcf_excel(
            input_file, output_file, results_df, mode=active_mode
        )
    except Exception as e:
        print(f"\nFehler beim Schreiben von {output_file}: {e}")
        return None

    cust_counts = {
        STATUS_PCF_AVAILABLE:     0,
        STATUS_NOT_YET_AVAILABLE: 0,
        STATUS_QUERY_ERROR:       0,
    }
    internal_counts = {
        INTERNAL_STATUS_PCF_AVAILABLE: 0,
        INTERNAL_STATUS_AAS_NO_PCF:    0,
        INTERNAL_STATUS_NOT_FOUND:     0,
        INTERNAL_STATUS_QUERY_ERROR:   0,
    }

    for canonical in valid_canonical:
        res = cache.get(canonical, {"status": STATUS_QUERY_ERROR, "internal_status": INTERNAL_STATUS_QUERY_ERROR})
        st = res.get("status", STATUS_QUERY_ERROR)
        ist = res.get("internal_status", st)
        if st in cust_counts:
            cust_counts[st] += 1
        if ist in internal_counts:
            internal_counts[ist] += 1
        elif st == STATUS_QUERY_ERROR:
            internal_counts[INTERNAL_STATUS_QUERY_ERROR] += 1

    summary = {
        "total_rows":        total,
        "queried_products":  len(valid_canonical),
        "unique_queried":    unique_total,
        "pcf_available":     cust_counts[STATUS_PCF_AVAILABLE],
        "not_yet_available": cust_counts[STATUS_NOT_YET_AVAILABLE],
        "query_errors":      cust_counts[STATUS_QUERY_ERROR],
        "aas_no_pcf":        internal_counts[INTERNAL_STATUS_AAS_NO_PCF],
        "not_found":         internal_counts[INTERNAL_STATUS_NOT_FOUND],
        "detected_column":   detected_col,
        "sheet_name":        detection_info.get("sheet_name"),
        "mode":              active_mode,
        "output_file":       final_output_path,
    }

    print(f"""
Finished.

Rows total        : {total}
Queried products  : {len(valid_canonical)}
Unique queried    : {unique_total}
PCF available     : {summary['pcf_available']}
Not yet available : {summary['not_yet_available']}
Query errors      : {summary['query_errors']}

Output written to:
{final_output_path}
""")

    return summary


# ── Native Windows Explorer Drag-and-Drop Helper (WM_DROPFILES via ctypes) ──

WM_DROPFILES      = 0x0233
WM_COPYGLOBALDATA = 0x0049
MSGFLT_ALLOW      = 1
GWLP_WNDPROC      = -4


def parse_dropped_file_paths(raw_input: Union[str, List[str], Tuple[str, ...], Any]) -> List[str]:
    """
    Parst gedroppte Dateipfade aus:
    - einer Python-Liste/Tupel von Pfaden (aus Win32 DragQueryFileW),
    - einem einzelnen Pfadstring (auch mit Leerzeichen, Umlauten oder Anfuehrungszeichen),
    - oder einem Tcl/TkDND-formatierten Event-String (z. B. '{C:/Pfad mit Leerzeichen/a.xlsx} C:/b.xlsx').
    """
    if raw_input is None:
        return []

    if isinstance(raw_input, (list, tuple)):
        out: List[str] = []
        for item in raw_input:
            if item is not None:
                s = str(item).strip().strip('"')
                if s:
                    out.append(s)
        return out

    s_raw = str(raw_input).strip()
    if not s_raw:
        return []

    # Falls der gesamte String bereits als existierender Dateipfad (oder einfacher Pfad ohne '{') vorliegt:
    clean_single = s_raw.strip('"')
    if "{" not in s_raw and ("\n" not in s_raw):
        if os.path.exists(clean_single) or len(s_raw.split()) == 1 or re.match(r"^[A-Za-z]:[\\/]", clean_single):
            # Pruefe, ob mehrere durch Zeilenumbruch oder Semikolon getrennte Pfade vorliegen
            if ";" in clean_single and not os.path.exists(clean_single):
                return [p.strip().strip('"') for p in clean_single.split(";") if p.strip().strip('"')]
            return [clean_single]

    # Tcl/TkDND-Listenformat: '{Pfad 1 mit Leerzeichen} {Pfad 2} Pfad3'
    results: List[str] = []
    i = 0
    n = len(s_raw)
    while i < n:
        while i < n and s_raw[i].isspace():
            i += 1
        if i >= n:
            break
        if s_raw[i] == "{":
            end_brace = s_raw.find("}", i + 1)
            if end_brace == -1:
                token = s_raw[i + 1 :].strip().strip('"')
                if token:
                    results.append(token)
                break
            token = s_raw[i + 1 : end_brace].strip().strip('"')
            if token:
                results.append(token)
            i = end_brace + 1
        elif s_raw[i] == '"':
            end_quote = s_raw.find('"', i + 1)
            if end_quote == -1:
                token = s_raw[i + 1 :].strip()
                if token:
                    results.append(token)
                break
            token = s_raw[i + 1 : end_quote].strip()
            if token:
                results.append(token)
            i = end_quote + 1
        else:
            start = i
            while i < n and not s_raw[i].isspace():
                i += 1
            token = s_raw[start:i].strip().strip('"')
            if token:
                results.append(token)
    return results


class _Win32FileDropTarget:
    """
    Bindet natives Windows-Explorer Drag-and-Drop (WM_DROPFILES) ohne externe
    C-Extensions an ein Tkinter-Fenster an.

    WICHTIG fuer GIL- und Thread-Sicherheit in Python/Tkinter:
    Innerhalb der C-WNDPROC (_py_wndproc) darf NIEMALS direkt in den Tcl/Tk-Interpreter
    (z. B. via widget.after) aufgerufen werden, da Tcl_DoOneEvent waehrend DispatchMessageW
    das Python-GIL per PyEval_SaveThread freigibt (andernfalls droht ein Fatal Python Error
    in PyEval_RestoreThread). Stattdessen legt _py_wndproc die extrahierten Dateipfade
    ausschliesslich in eine reine Python-Queue (_drop_queue), die vom regulären Tkinter-
    Event-Loop sicher abgefragt wird.
    """

    def __init__(self, widget: tk.Misc, on_drop_callback: Callable[[List[str]], Any]):
        self.widget = widget
        self.on_drop_callback = on_drop_callback
        self._drop_queue: "queue.SimpleQueue[List[str]]" = queue.SimpleQueue()
        self._hooked_hwnds: List[Tuple[int, int, Any]] = []
        self._after_id: Optional[str] = None
        self._uninstalled = False
        self.enabled = False

        if sys.platform == "win32":
            self._install()

    def _install(self) -> None:
        try:
            self.widget.update_idletasks()
            hwnd = int(self.widget.winfo_id())
            if not hwnd:
                return

            shell32 = ctypes.windll.shell32
            user32  = ctypes.windll.user32

            WNDPROC = ctypes.WINFUNCTYPE(
                ctypes.c_ssize_t,
                wintypes.HWND,
                wintypes.UINT,
                wintypes.WPARAM,
                wintypes.LPARAM,
            )

            shell32.DragAcceptFiles.argtypes = [wintypes.HWND, wintypes.BOOL]
            shell32.DragAcceptFiles.restype  = None

            shell32.DragQueryFileW.argtypes = [
                ctypes.c_void_p,
                wintypes.UINT,
                wintypes.LPWSTR,
                wintypes.UINT,
            ]
            shell32.DragQueryFileW.restype = wintypes.UINT

            shell32.DragFinish.argtypes = [ctypes.c_void_p]
            shell32.DragFinish.restype  = None

            user32.GetParent.argtypes = [wintypes.HWND]
            user32.GetParent.restype  = wintypes.HWND

            set_wnd_long = getattr(user32, "SetWindowLongPtrW", user32.SetWindowLongW)
            set_wnd_long.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_void_p]
            set_wnd_long.restype  = ctypes.c_void_p

            call_wnd_proc = user32.CallWindowProcW
            call_wnd_proc.argtypes = [
                ctypes.c_void_p,
                wintypes.HWND,
                wintypes.UINT,
                wintypes.WPARAM,
                wintypes.LPARAM,
            ]
            call_wnd_proc.restype = ctypes.c_ssize_t

            change_filter_ex = getattr(user32, "ChangeWindowMessageFilterEx", None)
            if change_filter_ex is not None:
                change_filter_ex.argtypes = [
                    wintypes.HWND,
                    wintypes.UINT,
                    wintypes.DWORD,
                    ctypes.c_void_p,
                ]
                change_filter_ex.restype = wintypes.BOOL

            target_hwnds = [hwnd]

            for target_h in target_hwnds:
                old_proc_holder = [0]

                def _make_wndproc(holder_ref):
                    def _py_wndproc(h_wnd, msg, w_param, l_param):
                        if msg == WM_DROPFILES:
                            hdrop = w_param
                            paths: List[str] = []
                            try:
                                count = shell32.DragQueryFileW(hdrop, 0xFFFFFFFF, None, 0)
                                for idx in range(count):
                                    req_len = shell32.DragQueryFileW(hdrop, idx, None, 0)
                                    if req_len > 0:
                                        buf = ctypes.create_unicode_buffer(req_len + 2)
                                        shell32.DragQueryFileW(hdrop, idx, buf, req_len + 2)
                                        if buf.value:
                                            paths.append(buf.value)
                            except BaseException:
                                pass
                            finally:
                                try:
                                    shell32.DragFinish(hdrop)
                                except BaseException:
                                    pass
                            if paths:
                                try:
                                    self._drop_queue.put(paths)
                                except BaseException:
                                    pass
                            return 0
                        try:
                            return call_wnd_proc(holder_ref[0], h_wnd, msg, w_param, l_param)
                        except BaseException:
                            return 0
                    return WNDPROC(_py_wndproc)

                cb = _make_wndproc(old_proc_holder)
                cb_ptr = ctypes.cast(cb, ctypes.c_void_p).value
                old_proc = set_wnd_long(target_h, GWLP_WNDPROC, cb_ptr)
                if old_proc:
                    old_proc_holder[0] = old_proc
                    if change_filter_ex is not None:
                        try:
                            change_filter_ex(target_h, WM_DROPFILES, MSGFLT_ALLOW, None)
                            change_filter_ex(target_h, WM_COPYGLOBALDATA, MSGFLT_ALLOW, None)
                        except Exception:
                            pass
                    shell32.DragAcceptFiles(target_h, True)
                    self._hooked_hwnds.append((target_h, old_proc, cb))

            if self._hooked_hwnds:
                self.enabled = True
                self._after_id = self.widget.after(50, self._poll_drop_queue)
        except Exception:
            self.enabled = False

    def _poll_drop_queue(self) -> None:
        self._after_id = None
        if self._uninstalled:
            return
        latest_batch: Optional[List[str]] = None
        while not self._drop_queue.empty():
            try:
                latest_batch = self._drop_queue.get_nowait()
            except Exception:
                break
        if latest_batch:
            try:
                self.on_drop_callback(latest_batch)
            except Exception as e:
                print(f"Drag-and-drop callback error: {e}")
        if not self._uninstalled:
            try:
                self._after_id = self.widget.after(50, self._poll_drop_queue)
            except Exception:
                pass

    def uninstall(self) -> None:
        self._uninstalled = True
        if self._after_id is not None:
            try:
                self.widget.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None
        if self.enabled and self._hooked_hwnds and sys.platform == "win32":
            try:
                user32  = ctypes.windll.user32
                shell32 = ctypes.windll.shell32
                set_wnd_long = getattr(user32, "SetWindowLongPtrW", user32.SetWindowLongW)
                set_wnd_long.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_void_p]
                set_wnd_long.restype  = ctypes.c_void_p
                shell32.DragAcceptFiles.argtypes = [wintypes.HWND, wintypes.BOOL]
                shell32.DragAcceptFiles.restype  = None
                for target_h, old_proc, _cb in reversed(self._hooked_hwnds):
                    try:
                        shell32.DragAcceptFiles(target_h, False)
                        set_wnd_long(target_h, GWLP_WNDPROC, old_proc)
                    except Exception:
                        pass
            except Exception:
                pass
        self._hooked_hwnds.clear()
        self.enabled = False


# ── Application Icon, Windows Taskbar Identity & Responsive Geometry ─────────

APP_USER_MODEL_ID = "WAGO.Sustainability.AASPCFQuery.3.0"
ICON_FILENAME     = "wago_pcf_icon.ico"
DEFAULT_WIN_WIDTH  = 720
DEFAULT_WIN_HEIGHT = 820
MIN_WIN_WIDTH      = 620
MIN_WIN_HEIGHT     = 560


def get_resource_path(relative_path: str) -> str:
    """Loest Ressourcenpfade sowohl in der Entwicklungsumgebung als auch in PyInstaller (_MEIPASS) auf."""
    if hasattr(sys, "_MEIPASS"):
        candidate = os.path.join(getattr(sys, "_MEIPASS"), relative_path)
        if os.path.exists(candidate):
            return candidate
    return os.path.join(os.path.abspath(os.path.dirname(__file__)), relative_path)


def _set_windows_app_user_model_id() -> None:
    """Setzt die explizite Windows AppUserModelID, damit die Taskleiste das App-Icon korrekt anzeigt."""
    if sys.platform == "win32":
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
        except Exception:
            pass


def _generate_fallback_ico_bytes() -> bytes:
    """
    Erzeugt zur Laufzeit ein Multi-Resolution-ICO im WAGO Corporate Design als Fallback,
    falls wago_pcf_icon.ico in einer isolierten Testumgebung nicht auf der Festplatte liegt.
    """
    import struct
    from PIL import Image, ImageDraw

    sizes = [16, 24, 32, 48, 64, 128, 256]
    payloads: List[Tuple[int, bytes]] = []
    for sz in sizes:
        im = Image.new("RGBA", (sz, sz), (0x1F, 0x28, 0x37, 255))
        dr = ImageDraw.Draw(im)
        pad = max(1, sz // 8)
        dr.rectangle([pad, pad, sz - 1 - pad, sz - 1 - pad], outline=(0x6E, 0xC8, 0x00, 255), width=max(1, sz // 12))
        buf = io.BytesIO()
        im.save(buf, format="PNG")
        payloads.append((sz, buf.getvalue()))

    header = struct.pack("<HHH", 0, 1, len(payloads))
    offset = 6 + 16 * len(payloads)
    entries: List[bytes] = []
    for sz, data in payloads:
        wb_sz = 0 if sz == 256 else sz
        entries.append(struct.pack("<BBBBHHII", wb_sz, wb_sz, 0, 0, 1, 32, len(data), offset))
        offset += len(data)
    return header + b"".join(entries) + b"".join(d for _, d in payloads)


def _ensure_application_icon_path() -> Optional[str]:
    """
    Gibt den Pfad zur eingebetteten wago_pcf_icon.ico zurueck (oder erzeugt eine
    temporaere Fallback-Datei, sodass zur Laufzeit niemals eine externe Datei erforderlich ist).
    """
    primary = get_resource_path(ICON_FILENAME)
    if os.path.isfile(primary):
        return primary
    try:
        import tempfile
        fallback_path = os.path.join(tempfile.gettempdir(), "wago_pcf_icon_runtime.ico")
        if not os.path.isfile(fallback_path):
            Path(fallback_path).write_bytes(_generate_fallback_ico_bytes())
        return fallback_path
    except Exception:
        return None


def get_available_desktop_work_area(widget: Optional[tk.Misc] = None) -> Tuple[int, int, int, int]:
    """
    Ermittelt den verfuegbaren Desktop-Arbeitsbereich (ohne Windows-Taskleiste) als
    (work_x, work_y, work_width, work_height).
    """
    if sys.platform == "win32":
        try:
            SPI_GETWORKAREA = 0x0030
            rect = wintypes.RECT()
            ok = ctypes.windll.user32.SystemParametersInfoW(
                SPI_GETWORKAREA, 0, ctypes.byref(rect), 0
            )
            if ok:
                w = int(rect.right - rect.left)
                h = int(rect.bottom - rect.top)
                if w >= 400 and h >= 300:
                    return int(rect.left), int(rect.top), w, h
        except Exception:
            pass

    if widget is not None:
        try:
            sw = int(widget.winfo_screenwidth())
            sh = int(widget.winfo_screenheight())
            return 0, 0, max(640, sw), max(480, sh - 48)
        except Exception:
            pass
    return 0, 0, 1920, 1032


def compute_responsive_window_geometry(
    req_w: int,
    req_h: int,
    work_w: int,
    work_h: int,
    work_x: int = 0,
    work_y: int = 0,
) -> Tuple[int, int, int, int, int, int]:
    """
    Berechnet (width, height, pos_x, pos_y, min_width, min_height) unter
    Beruecksichtigung der tatsaechlichen Inhaltsgroesse, der Bildschirmaufloesung
    (z. B. 1366x768, 1920x1080, 2560x1440, 3840x2160) und der Windows-DPI-Skalierung
    (100%, 125%, 150%, 200%), sodass das Fenster niemals ueber den verfuegbaren
    Desktop-Arbeitsbereich hinausragt und zentriert geoeffnet wird.
    """
    margin_w = 24 if work_w >= 800 else 8
    margin_h = 28 if work_h >= 700 else 12

    max_avail_w = max(480, work_w - margin_w)
    max_avail_h = max(420, work_h - margin_h)

    min_w = min(MIN_WIN_WIDTH, max_avail_w)
    min_h = min(MIN_WIN_HEIGHT, max_avail_h)

    desired_w = max(req_w, DEFAULT_WIN_WIDTH, min_w)
    desired_h = max(req_h, DEFAULT_WIN_HEIGHT, min_h)

    final_w = min(desired_w, max_avail_w)
    final_h = min(desired_h, max_avail_h)

    pos_x = work_x + max(0, (work_w - final_w) // 2)
    pos_y = work_y + max(0, (work_h - final_h) // 2)

    return final_w, final_h, pos_x, pos_y, min_w, min_h


# ── GUI (WAGO Corporate Design, 3 Workflows) ─────────────────────────────────

class PCFQueryApp(tk.Tk):
    """
    WAGO AAS PCF Query Tool GUI im offiziellen WAGO Corporate Design.
    Unterstuetzt drei Workflows:
    1. Single Query
    2. Batch Query (inkl. .xlsx/.xls & Drag-and-Drop)
    3. PCF Portfolio Coverage (inkl. Hintergrund-Excel-Export & Fortschrittsanzeige)
    """

    # Official WAGO Corporate Design Colors (Section 8.2)
    C_BG            = WAGO_LIGHT_GREY     # #EFF0F1
    C_PANEL         = WAGO_WHITE          # #FFFFFF
    C_HEADER_BG     = WAGO_ANTHRACITE     # #1F2837
    C_BORDER        = WAGO_ANTHRACITE_15  # #DEDFE1
    C_BORDER_STRONG = WAGO_ANTHRACITE_40  # #A5A8AF
    C_ACCENT        = WAGO_GREEN          # #6EC800
    C_ACCENT_HOVER  = WAGO_GREEN_HOVER    # #5EB000
    C_TEXT          = WAGO_ANTHRACITE     # #1F2837
    C_MUTED         = WAGO_ANTHRACITE_70  # #616A73
    C_BTN_FG        = WAGO_ANTHRACITE     # #1F2837 (High contrast on WAGO Green & light buttons)
    C_SUCCESS       = "#2E8540"
    C_WARN          = "#B45309"
    C_ERROR         = "#C81E1E"
    C_INFO          = "#1D4ED8"
    C_INFO_BG       = "#EBF3FE"
    C_INFO_BORDER   = "#93C5FD"
    C_DROP_BG       = "#F7FAF2"
    C_DROP_ACTIVE   = "#E6F4D2"

    FONT_HEAD  = ("Segoe UI", 13, "bold")
    FONT_LABEL = ("Segoe UI", 9)
    FONT_BOLD  = ("Segoe UI", 9, "bold")
    FONT_MONO  = ("Consolas", 9)
    FONT_SMALL = ("Segoe UI", 8)

    def __init__(self):
        import gc
        gc.collect()
        _set_windows_app_user_model_id()
        super().__init__()
        self.title("WAGO AAS PCF Query")
        self.configure(bg=self.C_BG)
        self.resizable(True, True)

        self._icon_path: Optional[str]                = None
        self._app_icon_photos: List[Any]              = []
        self._win32_hicons: List[int]                 = []

        self._active_workflow: str                    = WORKFLOW_BATCH
        self._mode_var                                = tk.StringVar(value=MODE_ONLINE)
        self._offline_notice_visible: bool            = False
        self._vscroll_visible: bool                   = False

        # Batch Query State
        self._input_path: Optional[str]               = None
        self._output_path: Optional[str]              = None
        self._df: Optional[pd.DataFrame]              = None
        self._detection_info: Dict[str, Any]          = {}
        self._selected_sheet: Optional[str]           = None
        self._selected_column: Optional[str]          = None
        self._canonical_articles: List[str]           = []
        self._query_thread: Optional[threading.Thread] = None
        self._stop_event                              = threading.Event()

        # Single Query State
        self._single_article_var                      = tk.StringVar(value="")
        self._single_last_result: Optional[Dict[str, Any]] = None
        self._single_thread: Optional[threading.Thread] = None

        # Portfolio Coverage & Background Export State
        self._portfolio_thread: Optional[threading.Thread] = None
        self._portfolio_stop_event                    = threading.Event()
        self._portfolio_result: Optional[Dict[str, Any]] = None
        self._portfolio_output_path: Optional[str]    = None
        self._portfolio_export_thread: Optional[threading.Thread] = None
        self._portfolio_export_stop_event             = threading.Event()
        self._portfolio_export_running: bool          = False
        self._portfolio_export_queue: "queue.Queue[Tuple[str, Any]]" = queue.Queue()

        self._drop_target: Optional[_Win32FileDropTarget] = None

        self._apply_window_and_taskbar_icon()
        self._build_ui()
        self._center_window()
        self._apply_win32_taskbar_icons()
        self._drop_target = _Win32FileDropTarget(self, self.handle_dropped_files)

    def destroy(self):
        if self._drop_target is not None:
            self._drop_target.uninstall()
            self._drop_target = None
        if sys.platform == "win32" and self._win32_hicons:
            user32 = ctypes.windll.user32
            user32.DestroyIcon.argtypes = [ctypes.c_void_p]
            user32.DestroyIcon.restype  = wintypes.BOOL
            for hicon in self._win32_hicons:
                try:
                    user32.DestroyIcon(hicon)
                except Exception:
                    pass
            self._win32_hicons.clear()
        # Clear ImageTk.PhotoImage references while this Tcl interpreter is still alive
        # so PhotoImage.__del__ never fires during garbage collection on a later Tcl interpreter.
        try:
            self.iconphoto(False)
        except Exception:
            pass
        self._app_icon_photos.clear()
        self._logo_img = None
        super().destroy()

    # ── Icon-Konfiguration (Window, Taskbar, Alt+Tab, Explorer) ───────────────

    def _apply_window_and_taskbar_icon(self) -> None:
        icon_path = _ensure_application_icon_path()
        self._icon_path = icon_path
        if not icon_path or not os.path.isfile(icon_path):
            return

        try:
            self.iconbitmap(default=icon_path)
            self.iconbitmap(icon_path)
        except Exception:
            pass

        try:
            from PIL import Image, ImageTk
            with Image.open(icon_path) as ico_img:
                photos = []
                for sz in (48, 32, 24, 16):
                    try:
                        ico_img.size = (sz, sz)
                        frame = ico_img.copy().convert("RGBA")
                    except Exception:
                        frame = ico_img.resize((sz, sz), Image.Resampling.LANCZOS).convert("RGBA")
                    photos.append(ImageTk.PhotoImage(frame))
                if photos:
                    self._app_icon_photos = photos
                    self.iconphoto(True, *photos)
        except Exception:
            pass

    def _apply_win32_taskbar_icons(self) -> None:
        """Setzt ICON_SMALL (16x16) und ICON_BIG (32x32) explizit via Win32 WM_SETICON."""
        if sys.platform != "win32" or not self._icon_path or not os.path.isfile(self._icon_path):
            return
        try:
            self.update_idletasks()
            user32 = ctypes.windll.user32
            WM_SETICON = 0x0080
            ICON_SMALL = 0
            ICON_BIG   = 1
            IMAGE_ICON = 1
            LR_LOADFROMFILE = 0x0010

            user32.LoadImageW.argtypes = [
                wintypes.HINSTANCE,
                wintypes.LPCWSTR,
                wintypes.UINT,
                ctypes.c_int,
                ctypes.c_int,
                wintypes.UINT,
            ]
            user32.LoadImageW.restype = ctypes.c_void_p

            user32.SendMessageW.argtypes = [
                wintypes.HWND,
                wintypes.UINT,
                wintypes.WPARAM,
                wintypes.LPARAM,
            ]
            user32.SendMessageW.restype = ctypes.c_ssize_t

            user32.GetParent.argtypes = [wintypes.HWND]
            user32.GetParent.restype  = wintypes.HWND

            h_small = user32.LoadImageW(None, self._icon_path, IMAGE_ICON, 16, 16, LR_LOADFROMFILE)
            h_big   = user32.LoadImageW(None, self._icon_path, IMAGE_ICON, 32, 32, LR_LOADFROMFILE)

            hwnd = self.winfo_id()
            parent_hwnd = user32.GetParent(hwnd)
            for target_hwnd in (hwnd, parent_hwnd):
                if target_hwnd:
                    if h_small:
                        user32.SendMessageW(target_hwnd, WM_SETICON, ICON_SMALL, h_small)
                    if h_big:
                        user32.SendMessageW(target_hwnd, WM_SETICON, ICON_BIG, h_big)

            if h_small:
                self._win32_hicons.append(h_small)
            if h_big:
                self._win32_hicons.append(h_big)
        except Exception:
            pass

    # ── UI-Aufbau (Responsiv + Scrollbar bei kleinen Bildschirmen) ────────────

    def _build_ui(self):
        # 1. Fixed Top Header & Navigation
        tk.Frame(self, bg=self.C_ACCENT, height=5).pack(side="top", fill="x")

        header_frame = tk.Frame(self, bg=self.C_HEADER_BG)
        header_frame.pack(side="top", fill="x")

        title_inner = tk.Frame(header_frame, bg=self.C_HEADER_BG)
        title_inner.pack(fill="x", padx=18, pady=10)

        text_frame = tk.Frame(title_inner, bg=self.C_HEADER_BG)
        text_frame.pack(side="left")

        tk.Label(
            text_frame, text="WAGO AAS PCF Query",
            font=("Segoe UI", 15, "bold"),
            fg=WAGO_WHITE, bg=self.C_HEADER_BG,
        ).pack(anchor="w")
        tk.Label(
            text_frame, text="Product Carbon Footprint – AAS Portal & Portfolio Tool",
            font=self.FONT_SMALL, fg=WAGO_ANTHRACITE_15, bg=self.C_HEADER_BG,
        ).pack(anchor="w")

        try:
            from PIL import Image, ImageTk
            logo_path = get_resource_path("WAGO_LOGO.webp")
            if os.path.exists(logo_path):
                img = Image.open(logo_path)
                h_new = 34
                w_new = int(img.width * (h_new / img.height))
                img = img.resize((w_new, h_new), Image.Resampling.LANCZOS)
                self._logo_img = ImageTk.PhotoImage(img)
                tk.Label(title_inner, image=self._logo_img, bg=self.C_HEADER_BG).pack(side="right", anchor="e")
        except Exception as e:
            print(f"Logo not loaded: {e}")

        # Workflow Navigation Bar (3 Tabs)
        nav_bar = tk.Frame(self, bg=WAGO_ANTHRACITE_15, height=36)
        nav_bar.pack(side="top", fill="x")

        self._tab_buttons: Dict[str, tk.Button] = {}
        tab_defs = [
            (WORKFLOW_SINGLE,    "  Single Query  "),
            (WORKFLOW_BATCH,     "  Batch Query (Excel)  "),
            (WORKFLOW_PORTFOLIO, "  Portfolio Coverage  "),
        ]
        for wf_key, wf_label in tab_defs:
            btn = tk.Button(
                nav_bar,
                text=wf_label,
                command=lambda k=wf_key: self.select_workflow(k),
                font=self.FONT_BOLD,
                relief="flat",
                bd=0,
                padx=14,
                pady=7,
                cursor="hand2",
            )
            btn.pack(side="left", padx=(2, 0), pady=(2, 0))
            self._tab_buttons[wf_key] = btn

        # 2. Fixed Bottom Footer (Docked at side='bottom' BEFORE scrollable body so it is never cut off)
        self._frm_footer_container = tk.Frame(self, bg=self.C_BG)
        self._frm_footer_container.pack(side="bottom", fill="x")

        tk.Frame(self._frm_footer_container, bg=self.C_BORDER, height=1).pack(fill="x", padx=8, pady=(2, 4))
        frm_footer = tk.Frame(self._frm_footer_container, bg=self.C_BG)
        frm_footer.pack(fill="x", padx=16, pady=(0, 6))

        tk.Label(
            frm_footer, text="Developed by Alexander Flekler · WAGO Corporate Sustainability",
            font=self.FONT_SMALL, fg=self.C_MUTED, bg=self.C_BG,
        ).pack(side="left")

        btn_about = tk.Button(
            frm_footer, text="About", command=self._on_about,
            font=self.FONT_SMALL, fg=self.C_MUTED, bg=self.C_BG,
            relief="flat", cursor="hand2", bd=0,
            activebackground=self.C_BG, activeforeground=self.C_TEXT,
        )
        btn_about.pack(side="right")

        # 3. Scrollable Main Body Area (Canvas + Vertical Scrollbar)
        self._frm_body = tk.Frame(self, bg=self.C_BG)
        self._frm_body.pack(side="top", fill="both", expand=True)

        self._body_canvas = tk.Canvas(
            self._frm_body, bg=self.C_BG, highlightthickness=0, bd=0
        )
        self._vscroll = ttk.Scrollbar(
            self._frm_body, orient="vertical", command=self._body_canvas.yview
        )
        self._body_canvas.configure(yscrollcommand=self._vscroll.set)
        self._body_canvas.pack(side="left", fill="both", expand=True)

        self._scroll_content = tk.Frame(self._body_canvas, bg=self.C_BG)
        self._canvas_win_id = self._body_canvas.create_window(
            (0, 0), window=self._scroll_content, anchor="nw"
        )

        self._scroll_content.bind("<Configure>", self._on_scroll_content_configure)
        self._body_canvas.bind("<Configure>", self._on_body_canvas_configure)
        self.bind_all("<MouseWheel>", self._on_mousewheel, add="+")

        # ── Shared Data Source Mode Selector & Q4 2026 Notice ────────────────
        self._frm_mode_section = tk.Frame(self._scroll_content, bg=self.C_BG)
        self._frm_mode_section.pack(fill="x", padx=16, pady=(6, 2))

        tk.Label(
            self._frm_mode_section, text="DATA SOURCE MODE",
            font=("Segoe UI", 8, "bold"), fg=self.C_MUTED, bg=self.C_BG,
        ).pack(anchor="w")

        self._frm_mode_container = tk.Frame(
            self._frm_mode_section, bg=self.C_PANEL,
            highlightbackground=self.C_BORDER, highlightthickness=1,
        )
        self._frm_mode_container.pack(fill="x", pady=(2, 2))

        frm_radios = tk.Frame(self._frm_mode_container, bg=self.C_PANEL)
        frm_radios.pack(fill="x", padx=10, pady=5)

        self._rb_online = tk.Radiobutton(
            frm_radios,
            text=MODE_LABELS[MODE_ONLINE],
            variable=self._mode_var,
            value=MODE_ONLINE,
            command=self._on_mode_changed,
            font=self.FONT_BOLD,
            fg=self.C_TEXT,
            bg=self.C_PANEL,
            selectcolor=self.C_PANEL,
            activebackground=self.C_PANEL,
            activeforeground=self.C_TEXT,
            cursor="hand2",
        )
        self._rb_online.pack(side="left", padx=(0, 18))

        self._rb_offline = tk.Radiobutton(
            frm_radios,
            text=MODE_LABELS[MODE_OFFLINE],
            variable=self._mode_var,
            value=MODE_OFFLINE,
            command=self._on_mode_changed,
            font=self.FONT_BOLD,
            fg=self.C_TEXT,
            bg=self.C_PANEL,
            selectcolor=self.C_PANEL,
            activebackground=self.C_PANEL,
            activeforeground=self.C_TEXT,
            cursor="hand2",
        )
        self._rb_offline.pack(side="left")

        self._lbl_portfolio_mode_hint = tk.Label(
            frm_radios,
            text="Live WAGO AAS API required for Portfolio Coverage (Q4 2026 used for comparison only)",
            font=self.FONT_SMALL,
            fg=self.C_MUTED,
            bg=self.C_PANEL,
        )

        # Informations-Banner fuer den Offline-Modus (standardmaessig ausgeblendet)
        self._frm_offline_notice = tk.Frame(
            self._frm_mode_section,
            bg=self.C_INFO_BG,
            highlightbackground=self.C_INFO_BORDER,
            highlightthickness=1,
        )
        self._lbl_offline_title = tk.Label(
            self._frm_offline_notice,
            text=OFFLINE_NOTICE_TITLE,
            font=self.FONT_BOLD,
            fg=self.C_INFO,
            bg=self.C_INFO_BG,
            anchor="w",
            justify="left",
        )
        self._lbl_offline_title.pack(fill="x", padx=10, pady=(5, 1))

        self._lbl_offline_text = tk.Label(
            self._frm_offline_notice,
            text=OFFLINE_NOTICE_TEXT,
            font=self.FONT_SMALL,
            fg=self.C_TEXT,
            bg=self.C_INFO_BG,
            wraplength=640,
            anchor="w",
            justify="left",
        )
        self._lbl_offline_text.pack(fill="x", padx=10, pady=(0, 1))

        self._lbl_offline_link = tk.Label(
            self._frm_offline_notice,
            text=OFFLINE_PORTAL_URL,
            font=("Segoe UI", 8, "underline"),
            fg=self.C_INFO,
            bg=self.C_INFO_BG,
            cursor="hand2",
            anchor="w",
        )
        self._lbl_offline_link.pack(anchor="w", padx=10, pady=(0, 5))
        self._lbl_offline_link.bind("<Button-1>", lambda _e: webbrowser.open(OFFLINE_PORTAL_URL))

        # ── Main Content Container for the 3 Workflow Frames ─────────────────
        self._frm_content = tk.Frame(self._scroll_content, bg=self.C_BG)
        self._frm_content.pack(fill="both", expand=True)

        self._workflow_frames: Dict[str, tk.Frame] = {
            WORKFLOW_SINGLE:    tk.Frame(self._frm_content, bg=self.C_BG),
            WORKFLOW_BATCH:     tk.Frame(self._frm_content, bg=self.C_BG),
            WORKFLOW_PORTFOLIO: tk.Frame(self._frm_content, bg=self.C_BG),
        }

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

        self._wrap_labels: List[tk.Label] = [self._lbl_offline_text]

        self._build_single_query_tab(self._workflow_frames[WORKFLOW_SINGLE])
        self._build_batch_query_tab(self._workflow_frames[WORKFLOW_BATCH])
        self._build_portfolio_coverage_tab(self._workflow_frames[WORKFLOW_PORTFOLIO])

        self.select_workflow(WORKFLOW_BATCH)

    def _on_scroll_content_configure(self, _event=None):
        try:
            self._body_canvas.configure(scrollregion=self._body_canvas.bbox("all"))
            self._update_scrollbar_visibility()
        except Exception:
            pass

    def _on_body_canvas_configure(self, event=None):
        try:
            canvas_w = event.width if (event and event.width > 100) else self._body_canvas.winfo_width()
            if canvas_w > 100:
                self._body_canvas.itemconfigure(self._canvas_win_id, width=canvas_w)
                wrap_px = max(320, canvas_w - 56)
                for lbl in self._wrap_labels:
                    try:
                        lbl.configure(wraplength=wrap_px)
                    except Exception:
                        pass
            self._update_scrollbar_visibility()
        except Exception:
            pass

    def _update_scrollbar_visibility(self):
        try:
            req_h = self._scroll_content.winfo_reqheight()
            can_h = self._body_canvas.winfo_height()
            need_scroll = can_h > 50 and req_h > (can_h + 4)
            if need_scroll and not self._vscroll_visible:
                self._vscroll.pack(side="right", fill="y")
                self._vscroll_visible = True
            elif not need_scroll and self._vscroll_visible:
                self._vscroll.pack_forget()
                self._vscroll_visible = False
                self._body_canvas.yview_moveto(0.0)
        except Exception:
            pass

    def _on_mousewheel(self, event):
        try:
            if self._vscroll_visible and self._body_canvas.winfo_exists():
                delta = int(-1 * (event.delta / 120))
                if delta != 0:
                    self._body_canvas.yview_scroll(delta, "units")
        except Exception:
            pass

    # ── Tab 1: Single Query UI ────────────────────────────────────────────────

    def _build_single_query_tab(self, parent: tk.Frame):
        self._section_in(parent, "Single Product PCF Lookup")
        frm_in = self._make_panel_in(parent)
        frm_in.pack(fill="x", padx=16, pady=(2, 6))

        lbl_prompt = tk.Label(
            frm_in,
            text="Enter a WAGO article number (Short or Long format, e.g. 221-412 or 0221-0412):",
            font=self.FONT_LABEL, fg=self.C_TEXT, bg=self.C_PANEL,
            anchor="w", justify="left", wraplength=640,
        )
        lbl_prompt.pack(fill="x", padx=12, pady=(10, 4))
        self._wrap_labels.append(lbl_prompt)

        entry_row = tk.Frame(frm_in, bg=self.C_PANEL)
        entry_row.pack(fill="x", padx=12, pady=(0, 8))

        self._entry_single = tk.Entry(
            entry_row,
            textvariable=self._single_article_var,
            font=("Consolas", 11),
            fg=self.C_TEXT,
            bg=WAGO_LIGHT_GREY,
            relief="solid",
            bd=1,
            width=28,
        )
        self._entry_single.pack(side="left", padx=(0, 10), ipady=4)
        self._entry_single.bind("<Return>", lambda _e: self._on_single_query())

        self._btn_single_query = self._make_button(
            entry_row, "  Query PCF  ", self._on_single_query, accent=True
        )
        self._btn_single_query.pack(side="left")

        self._lbl_single_status = tk.Label(
            frm_in, text="Ready.", font=self.FONT_SMALL, fg=self.C_MUTED, bg=self.C_PANEL,
            anchor="w", justify="left", wraplength=640,
        )
        self._lbl_single_status.pack(fill="x", padx=12, pady=(0, 8))
        self._wrap_labels.append(self._lbl_single_status)

        # Result Card
        self._section_in(parent, "Query Result")
        frm_res = self._make_panel_in(parent)
        frm_res.pack(fill="x", padx=16, pady=(2, 6))

        self._single_fields: Dict[str, tk.Label] = {}
        field_defs = [
            ("article_short", "Article Number (Short):"),
            ("article_long",  "Article Number (Long):"),
            ("pcf_display",   "PCF:"),
            ("unit",          "PCF Unit:"),
            ("data_source",   "Data Source:"),
        ]
        for f_key, f_lbl in field_defs:
            row = tk.Frame(frm_res, bg=self.C_PANEL)
            row.pack(fill="x", padx=12, pady=4)
            tk.Label(
                row, text=f_lbl, font=self.FONT_BOLD,
                fg=self.C_MUTED, bg=self.C_PANEL, width=24, anchor="w",
            ).pack(side="left")
            v_lbl = tk.Label(
                row, text="—", font=self.FONT_MONO,
                fg=self.C_TEXT, bg=self.C_PANEL, anchor="w", justify="left",
            )
            v_lbl.pack(side="left", fill="x", expand=True)
            self._single_fields[f_key] = v_lbl

        # Copy Buttons
        frm_copy = tk.Frame(parent, bg=self.C_BG)
        frm_copy.pack(fill="x", padx=16, pady=(4, 8))

        self._btn_copy_val = self._make_button(
            frm_copy, "  Copy PCF Value  ", lambda: self._copy_single_result(value_only=True)
        )
        self._btn_copy_val.pack(side="left", padx=(0, 8))
        self._btn_copy_val.config(state="disabled")

        self._btn_copy_all = self._make_button(
            frm_copy, "  Copy Complete Result  ", lambda: self._copy_single_result(value_only=False)
        )
        self._btn_copy_all.pack(side="left")
        self._btn_copy_all.config(state="disabled")

    # ── Tab 2: Batch Query UI ─────────────────────────────────────────────────

    def _build_batch_query_tab(self, parent: tk.Frame):
        self._section_in(parent, "Excel File (.xlsx / .xls) — Drag & Drop or Browse")

        # Drag-and-Drop Zone
        self._frm_drop_zone = tk.Frame(
            parent,
            bg=self.C_DROP_BG,
            highlightbackground=self.C_ACCENT,
            highlightthickness=2,
            cursor="hand2",
        )
        self._frm_drop_zone.pack(fill="x", padx=16, pady=(2, 4))

        self._lbl_drop_hint = tk.Label(
            self._frm_drop_zone,
            text="Drag & Drop Excel file (.xlsx, .xls) here or click 'Select Excel File'",
            font=self.FONT_BOLD,
            fg=self.C_TEXT,
            bg=self.C_DROP_BG,
            pady=7,
            wraplength=640,
        )
        self._lbl_drop_hint.pack(fill="x", padx=10)
        self._lbl_drop_hint.bind("<Button-1>", lambda _e: self._on_select_file())
        self._wrap_labels.append(self._lbl_drop_hint)

        btn_row = tk.Frame(parent, bg=self.C_BG)
        btn_row.pack(fill="x", padx=16, pady=(2, 4))

        self._btn_select = self._make_button(
            btn_row, "  Select Excel File", self._on_select_file
        )
        self._btn_select.pack(side="left", padx=(0, 8))

        self._btn_clear_file = self._make_button(
            btn_row, "  Remove File", self.clear_selected_file
        )
        self._btn_clear_file.pack(side="left")
        self._btn_clear_file.config(state="disabled")

        self._lbl_path = tk.Label(
            parent, text="No file selected.",
            font=self.FONT_MONO, fg=self.C_MUTED, bg=self.C_BG,
            wraplength=640, justify="left", anchor="w",
        )
        self._lbl_path.pack(fill="x", padx=16, pady=(0, 3))
        self._wrap_labels.append(self._lbl_path)

        self._frm_selectors = tk.Frame(parent, bg=self.C_BG)
        self._frm_selectors.pack(fill="x", padx=16, pady=(0, 4))

        tk.Label(
            self._frm_selectors, text="Sheet:", font=self.FONT_SMALL,
            fg=self.C_MUTED, bg=self.C_BG,
        ).pack(side="left", padx=(0, 4))
        self._sheet_var = tk.StringVar(value="")
        self._cmb_sheet = ttk.Combobox(
            self._frm_selectors, textvariable=self._sheet_var,
            state="disabled", width=20, font=self.FONT_SMALL,
        )
        self._cmb_sheet.pack(side="left", padx=(0, 12))
        self._cmb_sheet.bind("<<ComboboxSelected>>", self._on_sheet_changed)

        tk.Label(
            self._frm_selectors, text="Article Column:", font=self.FONT_SMALL,
            fg=self.C_MUTED, bg=self.C_BG,
        ).pack(side="left", padx=(0, 4))
        self._col_var = tk.StringVar(value="")
        self._cmb_col = ttk.Combobox(
            self._frm_selectors, textvariable=self._col_var,
            state="disabled", width=26, font=self.FONT_SMALL,
        )
        self._cmb_col.pack(side="left", fill="x", expand=True)
        self._cmb_col.bind("<<ComboboxSelected>>", self._on_column_changed)

        self._section_in(parent, "File Summary")
        frm_summary = self._make_panel_in(parent)
        frm_summary.pack(fill="x", padx=16, pady=(2, 3))
        self._lbl_summary = tk.Label(
            frm_summary, text="—",
            font=self.FONT_MONO, fg=self.C_MUTED, bg=self.C_PANEL,
            justify="left", anchor="w", wraplength=640,
        )
        self._lbl_summary.pack(fill="x", padx=10, pady=5)
        self._wrap_labels.append(self._lbl_summary)

        self._section_in(parent, "Preview  (first 8 unique products: Short / Long)")
        frm_prev = self._make_panel_in(parent)
        frm_prev.pack(fill="x", padx=16, pady=(2, 3))
        self._lbl_preview = tk.Label(
            frm_prev, text="—",
            font=self.FONT_MONO, fg=self.C_MUTED, bg=self.C_PANEL,
            justify="left", anchor="w",
        )
        self._lbl_preview.pack(fill="x", padx=10, pady=5)

        btn_start_frame = tk.Frame(parent, bg=self.C_BG)
        btn_start_frame.pack(fill="x", padx=16, pady=4)
        self._btn_start = self._make_button(
            btn_start_frame, "  Start PCF Query", self._on_start_query, accent=True
        )
        self._btn_start.pack(side="left")
        self._btn_start.config(state="disabled")

        self._section_in(parent, "Progress & Results")
        frm_prog = self._make_panel_in(parent)
        frm_prog.pack(fill="x", padx=16, pady=(2, 4))

        self._progress_var = tk.DoubleVar(value=0)
        ttk.Progressbar(
            frm_prog,
            variable=self._progress_var,
            maximum=100,
            style="WAGO.Horizontal.TProgressbar",
        ).pack(padx=10, pady=(7, 3), fill="x")

        self._lbl_prog_count = tk.Label(
            frm_prog, text="—",
            font=self.FONT_LABEL, fg=self.C_MUTED, bg=self.C_PANEL,
            anchor="w", justify="left", wraplength=640,
        )
        self._lbl_prog_count.pack(fill="x", padx=10)
        self._wrap_labels.append(self._lbl_prog_count)

        self._lbl_prog_article = tk.Label(
            frm_prog, text="",
            font=self.FONT_MONO, fg=self.C_TEXT, bg=self.C_PANEL, anchor="w",
        )
        self._lbl_prog_article.pack(fill="x", padx=10)

        self._lbl_eta = tk.Label(
            frm_prog, text="",
            font=self.FONT_SMALL, fg=self.C_MUTED, bg=self.C_PANEL, anchor="w",
        )
        self._lbl_eta.pack(fill="x", padx=10, pady=(0, 3))

        self._result_rows: Dict[str, tk.Label] = {}
        result_defs = [
            ("pcf_available",     "PCF available:",     self.C_SUCCESS),
            ("not_yet_available", "Not yet available:", self.C_WARN),
            ("query_errors",      "Query errors:",      self.C_ERROR),
        ]
        for key, label, color in result_defs:
            row = tk.Frame(frm_prog, bg=self.C_PANEL)
            row.pack(fill="x", padx=10, pady=1)
            tk.Label(
                row, text=label, font=self.FONT_LABEL,
                fg=self.C_MUTED, bg=self.C_PANEL, width=24, anchor="w",
            ).pack(side="left")
            val_lbl = tk.Label(
                row, text="—", font=self.FONT_BOLD,
                fg=color, bg=self.C_PANEL, width=10, anchor="e",
            )
            val_lbl.pack(side="left")
            self._result_rows[key] = val_lbl
        tk.Frame(frm_prog, bg=self.C_PANEL, height=3).pack()

        frm_out = tk.Frame(parent, bg=self.C_BG)
        frm_out.pack(fill="x", padx=16, pady=(4, 8))

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

    # ── Tab 3: PCF Portfolio Coverage UI (inkl. Background Export Progress) ──

    def _build_portfolio_coverage_tab(self, parent: tk.Frame):
        self._section_in(parent, "Internal WAGO AAS PCF Portfolio Coverage Analysis")
        frm_desc = self._make_panel_in(parent)
        frm_desc.pack(fill="x", padx=16, pady=(2, 4))

        self._lbl_port_desc = tk.Label(
            frm_desc,
            text=(
                "Analyze PCF availability across the WAGO AAS product portfolio and identify "
                "products without published PCF data.\n"
                "Uses the live WAGO AAS API as primary source and compares missing PCFs against "
                "the embedded Q4 2026 dataset to identify publication and calculation candidates."
            ),
            font=self.FONT_SMALL,
            fg=self.C_TEXT,
            bg=self.C_PANEL,
            wraplength=640,
            justify="left",
            anchor="w",
        )
        self._lbl_port_desc.pack(fill="x", padx=10, pady=6)
        self._wrap_labels.append(self._lbl_port_desc)

        frm_actions = tk.Frame(parent, bg=self.C_BG)
        frm_actions.pack(fill="x", padx=16, pady=4)

        self._btn_port_start = self._make_button(
            frm_actions, "  Start Portfolio Analysis  ", self._on_start_portfolio_analysis, accent=True
        )
        self._btn_port_start.pack(side="left", padx=(0, 8))

        self._btn_port_cancel = self._make_button(
            frm_actions, "  Cancel Analysis  ", self._on_cancel_portfolio_analysis
        )
        self._btn_port_cancel.pack(side="left")
        self._btn_port_cancel.config(state="disabled")

        self._section_in(parent, "Analysis Progress & PCF Coverage Visualization")
        frm_pprog = self._make_panel_in(parent)
        frm_pprog.pack(fill="x", padx=16, pady=(2, 4))

        self._port_progress_var = tk.DoubleVar(value=0)
        ttk.Progressbar(
            frm_pprog,
            variable=self._port_progress_var,
            maximum=100,
            style="WAGO.Horizontal.TProgressbar",
        ).pack(padx=10, pady=(7, 3), fill="x")

        self._lbl_port_status = tk.Label(
            frm_pprog, text="Ready to start portfolio analysis.",
            font=self.FONT_LABEL, fg=self.C_MUTED, bg=self.C_PANEL,
            anchor="w", justify="left", wraplength=640,
        )
        self._lbl_port_status.pack(fill="x", padx=10)
        self._wrap_labels.append(self._lbl_port_status)

        self._lbl_port_timing = tk.Label(
            frm_pprog, text="",
            font=self.FONT_SMALL, fg=self.C_MUTED, bg=self.C_PANEL, anchor="w",
        )
        self._lbl_port_timing.pack(fill="x", padx=10, pady=(0, 3))

        tk.Label(
            frm_pprog, text="Portfolio PCF Coverage (% of evaluated products):",
            font=self.FONT_SMALL, fg=self.C_MUTED, bg=self.C_PANEL, anchor="w",
        ).pack(fill="x", padx=10)

        self._port_cov_bar_var = tk.DoubleVar(value=0)
        ttk.Progressbar(
            frm_pprog,
            variable=self._port_cov_bar_var,
            maximum=100,
            style="WAGO.Horizontal.TProgressbar",
        ).pack(padx=10, pady=(2, 6), fill="x")

        self._section_in(parent, "Portfolio Summary")
        frm_psum = self._make_panel_in(parent)
        frm_psum.pack(fill="x", padx=16, pady=(2, 4))

        self._portfolio_fields: Dict[str, tk.Label] = {}
        port_defs = [
            ("total_shells_retrieved",          "Total AAS shells retrieved:"),
            ("total_unique_products",           "Total unique products identified:"),
            ("successfully_evaluated_products", "Successfully evaluated products:"),
            ("available_pcfs",                  "Available PCFs (in AAS):"),
            ("missing_pcfs",                    "Confirmed missing PCFs (in AAS):"),
            ("publication_candidates_count",    "PCFs available offline (Q4 2026) / missing in AAS:"),
            ("calculation_candidates_count",    "PCFs not found in either data source:"),
            ("technical_errors",                "Technical errors:"),
            ("pcf_coverage_pct",                "PCF coverage:"),
            ("analysis_completeness_pct",       "Analysis completeness:"),
        ]
        for p_key, p_lbl in port_defs:
            row = tk.Frame(frm_psum, bg=self.C_PANEL)
            row.pack(fill="x", padx=10, pady=1)
            tk.Label(
                row, text=p_lbl, font=self.FONT_LABEL,
                fg=self.C_MUTED, bg=self.C_PANEL, width=44, anchor="w",
            ).pack(side="left")
            v_lbl = tk.Label(
                row, text="—", font=self.FONT_BOLD,
                fg=self.C_TEXT, bg=self.C_PANEL, width=14, anchor="e",
            )
            v_lbl.pack(side="left")
            self._portfolio_fields[p_key] = v_lbl
        tk.Frame(frm_psum, bg=self.C_PANEL, height=3).pack()

        # ── Dedicated Portfolio Excel Export & Progress Section ───────────────
        self._section_in(parent, "Portfolio Excel Report Export")
        frm_pexp_card = self._make_panel_in(parent)
        frm_pexp_card.pack(fill="x", padx=16, pady=(2, 8))

        frm_pexp_btns = tk.Frame(frm_pexp_card, bg=self.C_PANEL)
        frm_pexp_btns.pack(fill="x", padx=10, pady=(7, 4))

        self._btn_port_export = self._make_button(
            frm_pexp_btns, "  Export Portfolio Report (.xlsx)  ", self._on_export_portfolio_report, accent=True
        )
        self._btn_port_export.pack(side="left", padx=(0, 8))
        self._btn_port_export.config(state="disabled")

        self._btn_port_export_cancel = self._make_button(
            frm_pexp_btns, "  Cancel Export  ", self.cancel_portfolio_export
        )
        self._btn_port_export_cancel.pack(side="left", padx=(0, 8))
        self._btn_port_export_cancel.config(state="disabled")

        self._btn_port_open_excel = self._make_button(
            frm_pexp_btns, "  Open Report (.xlsx)  ", self._on_open_portfolio_excel
        )
        self._btn_port_open_excel.pack(side="left", padx=(0, 8))
        self._btn_port_open_excel.config(state="disabled")

        self._btn_port_open_folder = self._make_button(
            frm_pexp_btns, "  Open Folder  ", self._on_open_portfolio_folder
        )
        self._btn_port_open_folder.pack(side="left")
        self._btn_port_open_folder.config(state="disabled")

        self._port_export_progress_var = tk.DoubleVar(value=0)
        self._bar_port_export = ttk.Progressbar(
            frm_pexp_card,
            variable=self._port_export_progress_var,
            maximum=100,
            style="WAGO.Horizontal.TProgressbar",
        )
        self._bar_port_export.pack(padx=10, pady=(2, 3), fill="x")

        self._lbl_port_export_stage = tk.Label(
            frm_pexp_card,
            text="Run or complete a portfolio analysis to export the 7-worksheet Excel report.",
            font=self.FONT_SMALL,
            fg=self.C_MUTED,
            bg=self.C_PANEL,
            anchor="w",
            justify="left",
            wraplength=640,
        )
        self._lbl_port_export_stage.pack(fill="x", padx=10)
        self._wrap_labels.append(self._lbl_port_export_stage)

        self._lbl_port_export_timing = tk.Label(
            frm_pexp_card,
            text="",
            font=self.FONT_SMALL,
            fg=self.C_MUTED,
            bg=self.C_PANEL,
            anchor="w",
        )
        self._lbl_port_export_timing.pack(fill="x", padx=10, pady=(0, 6))

    # ── Widget-Helfer & Responsive Fenster-Geometrie ──────────────────────────

    def _section_in(self, parent: tk.Widget, text: str):
        tk.Label(
            parent, text=text.upper(),
            font=("Segoe UI", 8, "bold"), fg=self.C_MUTED, bg=self.C_BG,
        ).pack(anchor="w", padx=16, pady=(5, 0))

    def _make_panel_in(self, parent: tk.Widget) -> tk.Frame:
        return tk.Frame(
            parent, bg=self.C_PANEL, relief="flat",
            highlightbackground=self.C_BORDER, highlightthickness=1,
        )

    def _make_button(
        self, parent: tk.Widget, text: str, command: Callable, accent: bool = False
    ) -> tk.Button:
        bg = self.C_ACCENT if accent else self.C_PANEL
        fg = WAGO_ANTHRACITE
        btn = tk.Button(
            parent, text=text, command=command,
            font=self.FONT_BOLD,
            bg=bg, fg=fg,
            activebackground=self.C_ACCENT_HOVER if accent else self.C_BORDER,
            activeforeground=fg,
            relief="solid" if not accent else "flat",
            cursor="hand2",
            padx=12, pady=5,
            bd=1 if not accent else 0,
        )
        btn.bind("<Enter>", lambda _e: btn.config(
            bg=self.C_ACCENT_HOVER if accent else self.C_BORDER))
        btn.bind("<Leave>", lambda _e: btn.config(bg=bg))
        return btn

    def _center_window(self):
        self.update_idletasks()
        work_x, work_y, work_w, work_h = get_available_desktop_work_area(self)

        # Measure content requirement across header + scrollable content + footer
        req_w = max(self.winfo_reqwidth(), self._scroll_content.winfo_reqwidth() + 24, DEFAULT_WIN_WIDTH)
        chrome_h = (
            self.winfo_reqheight() - self._frm_body.winfo_reqheight()
            + self._scroll_content.winfo_reqheight() + 12
        )
        req_h = max(chrome_h, DEFAULT_WIN_HEIGHT)

        w, h, px, py, min_w, min_h = compute_responsive_window_geometry(
            req_w, req_h, work_w, work_h, work_x=work_x, work_y=work_y
        )
        self.minsize(min_w, min_h)
        self.geometry(f"{w}x{h}+{px}+{py}")
        self.update_idletasks()
        self._update_scrollbar_visibility()

    # ── Workflow & Mode Management ────────────────────────────────────────────

    def get_active_workflow(self) -> str:
        return self._active_workflow

    def select_workflow(self, workflow: str) -> None:
        if workflow not in self._workflow_frames:
            return
        self._active_workflow = workflow

        for wf_key, frm in self._workflow_frames.items():
            if wf_key == workflow:
                frm.pack(fill="both", expand=True)
            else:
                frm.pack_forget()

        for wf_key, btn in self._tab_buttons.items():
            if wf_key == workflow:
                btn.config(bg=self.C_ACCENT, fg=WAGO_ANTHRACITE)
            else:
                btn.config(bg=WAGO_WHITE, fg=self.C_MUTED)

        # In Portfolio Coverage ist Live-API-Zugriff erforderlich
        if workflow == WORKFLOW_PORTFOLIO:
            self._rb_online.config(state="disabled")
            self._rb_offline.config(state="disabled")
            self._lbl_portfolio_mode_hint.pack(side="left", padx=(12, 0))
            if self._offline_notice_visible:
                self._frm_offline_notice.pack_forget()
                self._offline_notice_visible = False
        else:
            self._lbl_portfolio_mode_hint.pack_forget()
            self._rb_online.config(state="normal")
            self._rb_offline.config(state="normal")
            self._on_mode_changed()

        self.update_idletasks()
        self._on_scroll_content_configure()

    def get_mode(self) -> str:
        """Gibt den aktuell ausgewaehlten Betriebsmodus ('online' oder 'offline') zurueck."""
        return self._mode_var.get()

    def set_mode(self, mode: str) -> None:
        """Setzt den Betriebsmodus und blendet den Q4-2026-Hinweis entsprechend ein/aus."""
        target = MODE_OFFLINE if str(mode).strip().lower() == MODE_OFFLINE else MODE_ONLINE
        self._mode_var.set(target)
        self._on_mode_changed()

    def _on_mode_changed(self):
        is_offline = (
            self._mode_var.get() == MODE_OFFLINE
            and self._active_workflow != WORKFLOW_PORTFOLIO
        )
        if is_offline and not self._offline_notice_visible:
            self._frm_offline_notice.pack(fill="x", pady=(4, 2))
            self._offline_notice_visible = True
        elif not is_offline and self._offline_notice_visible:
            self._frm_offline_notice.pack_forget()
            self._offline_notice_visible = False
        self.update_idletasks()
        self._on_scroll_content_configure()

    # ── Single Query Handlers ─────────────────────────────────────────────────

    def _on_single_query(self):
        raw_art = self._single_article_var.get().strip()
        if not raw_art:
            self._lbl_single_status.config(
                text="Please enter a WAGO article number.", fg=self.C_WARN
            )
            return

        self._btn_single_query.config(state="disabled")
        self._lbl_single_status.config(
            text=f"Querying '{raw_art}'...", fg=self.C_MUTED
        )
        self.update_idletasks()

        mode = self._mode_var.get()
        res = query_single_article(raw_art, mode=mode)
        self._apply_single_query_result(res)

    def _apply_single_query_result(self, res: Dict[str, Any]):
        self._single_last_result = res
        self._btn_single_query.config(state="normal")

        self._single_fields["article_short"].config(text=res.get("article_short") or "—")
        self._single_fields["article_long"].config(text=res.get("article_long") or "—")
        self._single_fields["pcf_display"].config(
            text=res.get("pcf_display") or STATUS_NOT_YET_AVAILABLE,
            fg=self.C_SUCCESS if res.get("pcf") is not None else (
                self.C_ERROR if res.get("status") == STATUS_QUERY_ERROR else self.C_WARN
            ),
        )
        self._single_fields["unit"].config(text=res.get("unit") or "—")
        self._single_fields["data_source"].config(text=res.get("data_source") or "—")

        if res.get("status") == STATUS_QUERY_ERROR:
            self._lbl_single_status.config(
                text="Query could not be completed due to a technical communication error.",
                fg=self.C_ERROR,
            )
        else:
            self._lbl_single_status.config(text="Query completed.", fg=self.C_SUCCESS)

        self._btn_copy_val.config(state="normal")
        self._btn_copy_all.config(state="normal")
        self.update_idletasks()

    def _copy_single_result(self, value_only: bool = False) -> str:
        if not self._single_last_result:
            return ""
        text = format_single_result_for_clipboard(self._single_last_result, value_only=value_only)
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update_idletasks()
        self._lbl_single_status.config(
            text="Copied to clipboard.", fg=self.C_INFO
        )
        return text

    # ── Batch Query & Drag-and-Drop Handlers ──────────────────────────────────

    def _on_about(self):
        messagebox.showinfo(
            "About",
            "WAGO AAS PCF Query\nDeveloped by Alexander Flekler\nCorporate Sustainability\nVersion 3.0",
        )

    def handle_dropped_files(self, raw_paths: Union[str, List[str], Tuple[str, ...], Any]) -> bool:
        """
        Verarbeitet eine oder mehrere per Drag-and-Drop abgelegte Dateien sicher.
        - Unterstuetzt .xlsx und .xls sowie Pfade mit Leerzeichen, Umlauten und Sonderzeichen.
        - Behandelt mehrere gleichzeitig gedroppte Dateien fehlertolerant (waehlt die erste
          gueltige Excel-Datei aus und informiert den Nutzer).
        - Startet die PCF-Abfrage nicht automatisch.
        - Faengt alle Ausnahmen ab, damit fehlerhafte Drops niemals zum Programmabsturz fuehren.
        """
        try:
            paths = parse_dropped_file_paths(raw_paths)
            if not paths:
                return False

            valid_excel_paths: List[str] = []
            for p in paths:
                ext = os.path.splitext(p)[1].lower()
                if ext in SUPPORTED_EXCEL_EXTENSIONS and os.path.isfile(p):
                    valid_excel_paths.append(p)

            if not valid_excel_paths:
                first_name = os.path.basename(paths[0]) or str(paths[0])
                self._frm_drop_zone.config(bg="#FEE2E2", highlightbackground=self.C_ERROR)
                self._lbl_drop_hint.config(
                    text=f"Unsupported file: '{first_name}'. Please drop a valid .xlsx or .xls file.",
                    fg=self.C_ERROR,
                    bg="#FEE2E2",
                )
                return False

            chosen_path = valid_excel_paths[0]
            self.select_workflow(WORKFLOW_BATCH)
            self._frm_drop_zone.config(bg=self.C_DROP_ACTIVE, highlightbackground=self.C_ACCENT)

            if len(paths) > 1:
                hint_txt = (
                    f"Loaded: {os.path.basename(chosen_path)} "
                    f"(1 of {len(paths)} dropped files selected — drop another file to replace)"
                )
            else:
                hint_txt = f"Loaded: {os.path.basename(chosen_path)} (Drop another file to replace)"

            self._lbl_drop_hint.config(
                text=hint_txt,
                fg=self.C_TEXT,
                bg=self.C_DROP_ACTIVE,
            )
            self._load_file(chosen_path)
            return True
        except Exception as exc:
            print(f"Drag-and-drop processing error: {exc}")
            try:
                self._frm_drop_zone.config(bg="#FEE2E2", highlightbackground=self.C_ERROR)
                self._lbl_drop_hint.config(
                    text="Could not process dropped file. Please select a valid .xlsx or .xls file.",
                    fg=self.C_ERROR,
                    bg="#FEE2E2",
                )
            except Exception:
                pass
            return False

    def handle_dropped_file(self, path: Union[str, List[str], Tuple[str, ...]]) -> bool:
        """Abwaertskompatibler Wrapper fuer einzelne oder mehrere gedroppte Dateipfade."""
        return self.handle_dropped_files(path)

    def clear_selected_file(self) -> None:
        """Entfernt die aktuell ausgewaehlte Excel-Datei und setzt die Batch-Ansicht zurueck."""
        self._input_path = None
        self._output_path = None
        self._df = None
        self._detection_info = {}
        self._selected_sheet = None
        self._selected_column = None
        self._canonical_articles = []

        self._lbl_path.config(text="No file selected.", fg=self.C_MUTED)
        self._frm_drop_zone.config(bg=self.C_DROP_BG, highlightbackground=self.C_ACCENT)
        self._lbl_drop_hint.config(
            text="Drag & Drop Excel file (.xlsx, .xls) here or click 'Select Excel File'",
            fg=self.C_TEXT,
            bg=self.C_DROP_BG,
        )
        self._sheet_var.set("")
        self._col_var.set("")
        self._cmb_sheet.config(state="disabled", values=[])
        self._cmb_col.config(state="disabled", values=[])
        self._lbl_summary.config(text="—", fg=self.C_MUTED)
        self._lbl_preview.config(text="—", fg=self.C_MUTED)
        self._btn_start.config(state="disabled")
        self._btn_clear_file.config(state="disabled")
        self.update_idletasks()

    def _on_select_file(self):
        path = filedialog.askopenfilename(
            title="WAGO PCF Query – Select Excel File",
            filetypes=[
                ("Excel Files", "*.xlsx *.xls"),
                ("Excel Workbook (*.xlsx)", "*.xlsx"),
                ("Legacy Excel 97-2003 (*.xls)", "*.xls"),
                ("All Files", "*.*"),
            ],
        )
        if path:
            self.handle_dropped_file(path)

    def _on_sheet_changed(self, _event=None):
        if self._input_path and self._sheet_var.get():
            self._load_file(self._input_path, sheet_name=self._sheet_var.get(), article_column=None)

    def _on_column_changed(self, _event=None):
        if self._df is not None and self._col_var.get():
            self._selected_column = self._col_var.get()
            self._refresh_file_summary(article_column=self._selected_column)

    def _load_file(
        self,
        path: str,
        sheet_name: Optional[str] = None,
        article_column: Optional[str] = None,
    ):
        self._input_path = path
        self._btn_start.config(state="disabled")
        self._btn_clear_file.config(state="normal")
        self._lbl_path.config(text=path, fg=self.C_TEXT)
        self._df = None

        try:
            df, info = load_excel_flexible(
                path, sheet_name=sheet_name, article_column=article_column
            )
        except Exception as e:
            self._frm_drop_zone.config(bg="#FEE2E2", highlightbackground=self.C_ERROR)
            self._lbl_drop_hint.config(
                text=f"Could not read '{os.path.basename(path)}'. Select or drop another Excel file.",
                fg=self.C_ERROR,
                bg="#FEE2E2",
            )
            self._lbl_summary.config(
                text=f"Could not read file:\n{e}", fg=self.C_ERROR,
            )
            self._lbl_preview.config(text="—", fg=self.C_MUTED)
            self._sheet_var.set("")
            self._col_var.set("")
            self._cmb_sheet.config(state="disabled", values=[])
            self._cmb_col.config(state="disabled", values=[])
            self.update_idletasks()
            return

        self._df = df
        self._detection_info = info
        self._selected_sheet = info.get("sheet_name")
        self._selected_column = article_column or info.get("column")

        sheets = info.get("available_sheets") or []
        self._cmb_sheet["values"] = sheets
        self._sheet_var.set(self._selected_sheet or "")
        self._cmb_sheet.config(state="readonly" if len(sheets) > 1 else "disabled")

        cols = [str(c) for c in df.columns.tolist()]
        self._cmb_col["values"] = cols
        self._col_var.set(str(self._selected_column) if self._selected_column else "")
        self._cmb_col.config(state="readonly" if len(cols) > 1 else "disabled")

        stem = os.path.splitext(os.path.basename(path))[0]
        self._output_path = os.path.join(os.path.dirname(path), f"{stem}_PCF.xlsx")

        self._refresh_file_summary(article_column=article_column)

    def _refresh_file_summary(self, article_column: Optional[str] = None):
        if self._df is None or self._df.empty:
            self._lbl_summary.config(
                text="No table data found in this file.", fg=self.C_ERROR,
            )
            self._lbl_preview.config(text="—", fg=self.C_MUTED)
            self._btn_start.config(state="disabled")
            return

        info = self._detection_info
        target_col = article_column or info.get("column")

        if not target_col or target_col not in self._df.columns:
            cols_str = ", ".join(str(c) for c in self._df.columns.tolist())
            self._lbl_summary.config(
                text=(
                    "No WAGO article numbers detected automatically.\n"
                    f"Please select a column above. Available: {cols_str}"
                ),
                fg=self.C_WARN,
            )
            self._lbl_preview.config(text="—", fg=self.C_MUTED)
            self._btn_start.config(state="disabled")
            return

        canonical_all, short_all, long_all = _extract_row_articles(
            self._df, info, article_column=article_column
        )
        valid_pairs = [
            (c, s, l)
            for c, s, l in zip(canonical_all, short_all, long_all)
            if c
        ]
        valid_canonical = [p[0] for p in valid_pairs]
        unique_canonical = list(dict.fromkeys(valid_canonical))
        duplicates = len(valid_canonical) - len(unique_canonical)
        empty_count = len(canonical_all) - len(valid_canonical)

        self._canonical_articles = canonical_all

        if info.get("has_both_formats") and not article_column:
            col_desc = f"{info.get('short_column')} + {info.get('long_column')}"
        else:
            col_desc = str(target_col)

        summary_text = (
            f"  Detected column :  {col_desc}\n"
            f"  Rows            :  {len(canonical_all)}\n"
            f"  Valid entries   :  {len(valid_canonical)}\n"
            f"  Unique products :  {len(unique_canonical)}\n"
            f"  Duplicates      :  {duplicates}\n"
            f"  Empty values    :  {empty_count}"
        )

        if info.get("is_ambiguous") and not article_column:
            cand_names = ", ".join(str(c["column"]) for c in info.get("candidates", [])[:3])
            summary_text += (
                f"\n  Note            :  Multiple plausible columns ({cand_names}).\n"
                f"                     Verify selection in 'Article Column' above."
            )
            self._lbl_summary.config(text=summary_text, fg=self.C_WARN)
        else:
            self._lbl_summary.config(text=summary_text, fg=self.C_TEXT)

        seen_prev = set()
        preview_lines: List[str] = []
        for c, s, l in valid_pairs:
            if c not in seen_prev:
                seen_prev.add(c)
                if len(preview_lines) < 8:
                    preview_lines.append(f"  {s:<22} ->  {l}")

        if len(unique_canonical) > 8:
            preview_text = "\n".join(preview_lines) + f"\n  ... ({len(unique_canonical) - 8} more)"
        else:
            preview_text = "\n".join(preview_lines)

        self._lbl_preview.config(
            text=preview_text if preview_lines else "  (no articles found)",
            fg=self.C_TEXT if preview_lines else self.C_MUTED,
        )

        if valid_canonical:
            self._btn_start.config(state="normal")
        else:
            self._lbl_summary.config(
                text=summary_text + "\n\n  No valid article numbers found.",
                fg=self.C_WARN,
            )
            self._btn_start.config(state="disabled")

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
        start_time = time.time()

        counts = {
            STATUS_PCF_AVAILABLE:     0,
            STATUS_NOT_YET_AVAILABLE: 0,
            STATUS_QUERY_ERROR:       0,
        }

        articles_all    = self._canonical_articles
        valid_articles  = [a for a in articles_all if a]
        unique_articles = list(dict.fromkeys(valid_articles))
        unique_total    = len(unique_articles)
        active_mode     = self._mode_var.get()

        cache: Dict[str, Dict[str, Any]] = {}
        cache_lock = threading.Lock()
        done_count = [0]

        def _worker(article: str) -> Tuple[str, Dict[str, Any]]:
            if self._stop_event.is_set():
                return article, {"pcf": None, "unit": "", "status": STATUS_QUERY_ERROR}

            res = _query_with_timing(article, mode=active_mode)

            with cache_lock:
                cache[article] = res
                done_count[0] += 1
                done = done_count[0]
                if res["status"] in counts:
                    counts[res["status"]] += 1

            elapsed = time.time() - start_time
            if 3 <= done < unique_total:
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
                pass

        if self._stop_event.is_set():
            self.after(0, lambda: self._on_query_done(None))
            return

        override_col = (
            self._selected_column
            if self._selected_column != self._detection_info.get("column")
            else None
        )
        results_df = prepare_output_dataframe(
            self._df, self._detection_info, cache, article_column=override_col
        )

        summary = None
        try:
            final_out = export_batch_pcf_excel(
                self._input_path or "input.xlsx",
                self._output_path or "output_PCF.xlsx",
                results_df,
                mode=active_mode,
            )
            self._output_path = final_out
            summary = {
                "unique_queried":    unique_total,
                "pcf_available":     counts[STATUS_PCF_AVAILABLE],
                "not_yet_available": counts[STATUS_NOT_YET_AVAILABLE],
                "query_errors":      counts[STATUS_QUERY_ERROR],
                "output_file":       final_out,
            }
        except Exception as e:
            print(f"Fehler beim Schreiben der Ausgabedatei: {e}")
            err = str(e)
            self.after(0, lambda msg=err: messagebox.showerror(
                "Write Error", f"Could not write output file:\n{msg}",
            ))

        self.after(0, lambda s=summary: self._on_query_done(s))

    def _update_progress(
        self,
        pct: float,
        unique_done: int,
        unique_total: int,
        article: str,
        eta_str: str,
        counts: Dict[str, int],
    ):
        self._progress_var.set(pct)
        self._lbl_prog_count.config(
            text=f"Processing {unique_done} / {unique_total} unique products",
            fg=self.C_TEXT,
        )
        self._lbl_prog_article.config(
            text=f"Current article: {article}", fg=self.C_TEXT,
        )
        self._lbl_eta.config(text=eta_str, fg=self.C_MUTED)

        self._result_rows["pcf_available"].config(
            text=str(counts.get(STATUS_PCF_AVAILABLE, 0)))
        self._result_rows["not_yet_available"].config(
            text=str(counts.get(STATUS_NOT_YET_AVAILABLE, 0)))
        self._result_rows["query_errors"].config(
            text=str(counts.get(STATUS_QUERY_ERROR, 0)))

    def _on_query_done(self, summary: Optional[Dict[str, Any]]):
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

        if summary["query_errors"] > 0:
            self._lbl_prog_count.config(
                text=(
                    f"Query completed with {summary['query_errors']} technical error(s) — "
                    f"{summary['unique_queried']} unique products processed"
                ),
                fg=self.C_WARN,
            )
        else:
            self._lbl_prog_count.config(
                text=(
                    f"Query completed — "
                    f"{summary['unique_queried']} unique products processed"
                ),
                fg=self.C_SUCCESS,
            )

        self._result_rows["pcf_available"].config(text=str(summary["pcf_available"]))
        self._result_rows["not_yet_available"].config(text=str(summary["not_yet_available"]))
        self._result_rows["query_errors"].config(text=str(summary["query_errors"]))

        self._btn_open_folder.config(state="normal")
        self._btn_open_excel.config(state="normal")

    def _set_ui_running(self, running: bool):
        state = "disabled" if running else "normal"
        self._btn_select.config(state=state)
        self._btn_clear_file.config(state=state if self._input_path else "disabled")
        self._btn_start.config(state=state)
        self._rb_online.config(state=state)
        self._rb_offline.config(state=state)
        if running:
            self._cmb_sheet.config(state="disabled")
            self._cmb_col.config(state="disabled")
        else:
            sheets = self._detection_info.get("available_sheets") or []
            self._cmb_sheet.config(state="readonly" if len(sheets) > 1 else "disabled")
            cols = list(self._df.columns) if self._df is not None else []
            self._cmb_col.config(state="readonly" if len(cols) > 1 else "disabled")

    # ── Portfolio Coverage & Background Export Handlers ──────────────────────

    def _on_start_portfolio_analysis(self):
        if self._portfolio_export_running:
            return
        self._btn_port_start.config(state="disabled")
        self._btn_port_cancel.config(state="normal")
        self._btn_port_export.config(state="disabled")
        self._port_progress_var.set(0)
        self._port_cov_bar_var.set(0)
        self._lbl_port_status.config(text="Enumerating WAGO AAS shells...", fg=self.C_TEXT)
        self._lbl_port_timing.config(text="Elapsed: 0s")
        for lbl in self._portfolio_fields.values():
            lbl.config(text="—")

        self._portfolio_stop_event.clear()
        self._portfolio_thread = threading.Thread(
            target=self._run_portfolio_worker, daemon=True
        )
        self._portfolio_thread.start()

    def _on_cancel_portfolio_analysis(self):
        self._portfolio_stop_event.set()
        self._lbl_port_status.config(
            text="Cancelling analysis... preparing partial results...", fg=self.C_WARN
        )

    def _run_portfolio_worker(self):
        def _prog_cb(p_info: Dict[str, Any]):
            self.after(0, lambda info=dict(p_info): self._update_portfolio_progress(info))

        res = run_portfolio_coverage_analysis(
            progress_callback=_prog_cb,
            stop_event=self._portfolio_stop_event,
        )
        self.after(0, lambda r=res: self._apply_portfolio_result(r))

    def _update_portfolio_progress(self, info: Dict[str, Any]):
        phase = info.get("phase", "enumerating")
        shells = info.get("shells_retrieved", 0)
        uniq = info.get("unique_products", 0)
        proc = info.get("processed_products", 0)
        tot_v = info.get("total_to_verify", 0)
        elapsed = int(info.get("elapsed_seconds", 0))
        eta = info.get("eta_seconds")

        if phase == "enumerating":
            self._lbl_port_status.config(
                text=f"Enumerating AAS portfolio: {shells} shells retrieved ({uniq} unique products)...",
                fg=self.C_TEXT,
            )
            self._port_progress_var.set(min(45.0, (shells / 1000.0) * 10.0))
        else:
            pct = 50.0 + (proc / max(tot_v, 1)) * 50.0
            self._port_progress_var.set(min(99.0, pct))
            self._lbl_port_status.config(
                text=f"Verifying PCF submodels: {proc} / {tot_v} products ({uniq} total unique)...",
                fg=self.C_TEXT,
            )

        eta_txt = f" | ETA: {int(eta // 60)}m {int(eta % 60):02d}s" if eta is not None else ""
        self._lbl_port_timing.config(text=f"Elapsed: {elapsed // 60}m {elapsed % 60:02d}s{eta_txt}")

    def _apply_portfolio_result(self, res: Dict[str, Any]):
        self._portfolio_result = res
        self._btn_port_start.config(state="normal")
        self._btn_port_cancel.config(state="disabled")
        self._port_progress_var.set(100)

        cov_pct = float(res.get("pcf_coverage_pct", 0.0))
        comp_pct = float(res.get("analysis_completeness_pct", 0.0))
        self._port_cov_bar_var.set(cov_pct)

        if res.get("was_cancelled"):
            self._lbl_port_status.config(
                text="Portfolio analysis cancelled — displaying partial results.",
                fg=self.C_WARN,
            )
        elif res.get("is_partial"):
            self._lbl_port_status.config(
                text="Portfolio analysis completed (partial/bounded scope).",
                fg=self.C_WARN,
            )
        else:
            self._lbl_port_status.config(
                text="Portfolio analysis completed.",
                fg=self.C_SUCCESS,
            )

        self._lbl_port_timing.config(text=f"Total elapsed: {res.get('elapsed_seconds', 0):.1f}s")

        for k in (
            "total_shells_retrieved",
            "total_unique_products",
            "successfully_evaluated_products",
            "available_pcfs",
            "missing_pcfs",
            "publication_candidates_count",
            "calculation_candidates_count",
            "technical_errors",
        ):
            if k in self._portfolio_fields:
                self._portfolio_fields[k].config(text=str(res.get(k, 0)))

        self._portfolio_fields["pcf_coverage_pct"].config(text=f"{cov_pct:.2f}%")
        self._portfolio_fields["analysis_completeness_pct"].config(text=f"{comp_pct:.2f}%")

        if not self._portfolio_export_running:
            self._btn_port_export.config(state="normal")
            self._lbl_port_export_stage.config(
                text="Ready to export portfolio report (.xlsx).",
                fg=self.C_TEXT,
            )
        self.update_idletasks()
        self._on_scroll_content_configure()

    def _on_export_portfolio_report(self):
        if not self._portfolio_result or self._portfolio_export_running:
            return
        default_name = f"WAGO_PCF_Portfolio_Coverage_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path = filedialog.asksaveasfilename(
            title="Save WAGO PCF Portfolio Coverage Report",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel Workbook (*.xlsx)", "*.xlsx")],
        )
        if path:
            self.start_portfolio_export(path)

    def start_portfolio_export(
        self, output_path: str, show_error_dialog: bool = True
    ) -> bool:
        """
        Startet den Portfolio-Coverage-Excel-Export in einem Hintergrund-Worker-Thread,
        damit die GUI jederzeit responsiv bleibt und keine mehrfachen Exporte parallel laufen.
        """
        if not self._portfolio_result or self._portfolio_export_running:
            return False

        self._portfolio_export_running = True
        self._portfolio_export_stop_event.clear()
        while not self._portfolio_export_queue.empty():
            try:
                self._portfolio_export_queue.get_nowait()
            except queue.Empty:
                break

        self._btn_port_export.config(state="disabled")
        self._btn_port_start.config(state="disabled")
        self._btn_port_export_cancel.config(state="normal")
        self._btn_port_open_excel.config(state="disabled")
        self._btn_port_open_folder.config(state="disabled")

        self._port_export_progress_var.set(1.0)
        self._lbl_port_export_stage.config(
            text="Preparing portfolio data...",
            fg=self.C_TEXT,
        )
        self._lbl_port_export_timing.config(text="Elapsed: 0.0s")
        self.update_idletasks()

        self._portfolio_export_thread = threading.Thread(
            target=self._run_portfolio_export_worker,
            args=(output_path, show_error_dialog),
            daemon=True,
        )
        self._portfolio_export_thread.start()
        self.after(30, self._poll_portfolio_export_queue)
        return True

    def cancel_portfolio_export(self) -> None:
        """Bricht einen laufenden Portfolio-Excel-Export kontrolliert ab."""
        if not self._portfolio_export_running:
            return
        self._portfolio_export_stop_event.set()
        self._btn_port_export_cancel.config(state="disabled")
        self._lbl_port_export_stage.config(
            text="Cancelling Excel export...",
            fg=self.C_WARN,
        )
        self.update_idletasks()

    def _run_portfolio_export_worker(self, output_path: str, show_error_dialog: bool = True):
        def _export_prog_cb(p_info: Dict[str, Any]):
            self._portfolio_export_queue.put(("progress", dict(p_info)))

        out_p: Optional[str] = None
        err_msg: Optional[str] = None
        cancelled = False
        t0 = time.perf_counter()

        try:
            out_p = export_portfolio_coverage_excel(
                self._portfolio_result,
                output_path,
                progress_callback=_export_prog_cb,
                stop_event=self._portfolio_export_stop_event,
            )
        except Exception as e:
            if self._portfolio_export_stop_event.is_set() or "cancelled" in str(e).lower():
                cancelled = True
            else:
                err_msg = str(e)

        dur = time.perf_counter() - t0
        self._portfolio_export_queue.put(
            ("done", (out_p, err_msg, cancelled, dur, show_error_dialog))
        )

    def _poll_portfolio_export_queue(self) -> None:
        """Verarbeitet Fortschritts- und Abschlussmeldungen des Hintergrund-Exports im Haupt-GUI-Thread."""
        last_progress: Optional[Dict[str, Any]] = None
        done_payload: Optional[Tuple[Optional[str], Optional[str], bool, float, bool]] = None

        while True:
            try:
                msg_type, payload = self._portfolio_export_queue.get_nowait()
            except queue.Empty:
                break
            if msg_type == "progress":
                last_progress = payload
            elif msg_type == "done":
                done_payload = payload

        if last_progress is not None and done_payload is None:
            self._update_portfolio_export_progress(last_progress)

        if done_payload is not None:
            out_p, err_msg, cancelled, dur, show_dlg = done_payload
            self._on_portfolio_export_finished(out_p, err_msg, cancelled, dur, show_dlg)
            return

        if self._portfolio_export_running:
            self.after(30, self._poll_portfolio_export_queue)

    def _update_portfolio_export_progress(self, info: Dict[str, Any]):
        if not self._portfolio_export_running or not self.winfo_exists():
            return
        stage = info.get("stage", "Exporting portfolio report...")
        ws_name = info.get("worksheet", "")
        proc = int(info.get("processed_records", 0))
        tot = int(info.get("total_records", 0))
        pct = min(99.0, float(info.get("percent", 0.0)))
        elapsed = float(info.get("elapsed_seconds", 0.0))
        eta = info.get("eta_seconds")

        # Do not show 100% until file save is completely finished
        self._port_export_progress_var.set(pct)
        self._lbl_port_export_stage.config(text=stage, fg=self.C_TEXT)

        parts = []
        if ws_name:
            parts.append(f"Worksheet: {ws_name}")
        if tot > 0:
            parts.append(f"Records: {proc:,} / {tot:,}")
        parts.append(f"Elapsed: {elapsed:.1f}s")
        if eta is not None:
            parts.append(f"ETA: {float(eta):.1f}s")
        self._lbl_port_export_timing.config(text=" | ".join(parts))

    def _on_portfolio_export_finished(
        self,
        out_path: Optional[str],
        err_msg: Optional[str],
        cancelled: bool,
        duration: float,
        show_error_dialog: bool = True,
    ):
        self._portfolio_export_running = False
        self._btn_port_export_cancel.config(state="disabled")
        self._btn_port_start.config(state="normal")
        if self._portfolio_result is not None:
            self._btn_port_export.config(state="normal")

        if cancelled:
            self._port_export_progress_var.set(0.0)
            self._lbl_port_export_stage.config(
                text="Export cancelled — no incomplete file was written. Analysis results remain available.",
                fg=self.C_WARN,
            )
            self._lbl_port_export_timing.config(text=f"Cancelled after {duration:.1f}s")
            return

        if err_msg or not out_path or not os.path.isfile(out_path):
            self._port_export_progress_var.set(0.0)
            msg = err_msg or "File could not be saved."
            self._lbl_port_export_stage.config(
                text=f"Export failed: {msg} (Analysis results preserved — you can retry.)",
                fg=self.C_ERROR,
            )
            self._lbl_port_export_timing.config(text="")
            if show_error_dialog:
                messagebox.showerror("Export Error", f"Could not write Portfolio Report:\n{msg}")
            return

        # 100% is ONLY displayed after the workbook has been successfully saved!
        self._portfolio_output_path = out_path
        self._port_export_progress_var.set(100.0)
        fname = os.path.basename(out_path)
        self._lbl_port_export_stage.config(
            text=f"Report exported successfully: {fname}",
            fg=self.C_SUCCESS,
        )
        self._lbl_port_export_timing.config(
            text=f"Completed in {duration:.2f}s — saved to {out_path}"
        )
        self._lbl_port_status.config(
            text=f"Report exported to: {fname}",
            fg=self.C_SUCCESS,
        )
        self._btn_port_open_excel.config(state="normal")
        self._btn_port_open_folder.config(state="normal")
        self.update_idletasks()

    def _on_open_portfolio_excel(self):
        if self._portfolio_output_path and os.path.isfile(self._portfolio_output_path):
            os.startfile(self._portfolio_output_path)
        else:
            messagebox.showwarning(
                "File not found",
                f"Portfolio report file not found:\n{self._portfolio_output_path}",
            )

    def _on_open_portfolio_folder(self):
        if self._portfolio_output_path:
            import platform
            folder = Path(self._portfolio_output_path).parent.resolve()
            try:
                if platform.system() == "Windows":
                    os.startfile(str(folder))
                else:
                    subprocess.Popen(["xdg-open", str(folder)])
            except Exception as e:
                messagebox.showerror("Error", f"Could not open folder:\n{e}")

    # ── Output-Aktionen ───────────────────────────────────────────────────────

    def _on_open_folder(self):
        if self._output_path:
            import platform
            folder = Path(self._output_path).parent.resolve()
            try:
                if platform.system() == "Windows":
                    os.startfile(str(folder))
                else:
                    subprocess.Popen(["xdg-open", str(folder)])
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


def _post_win32_hdrop_message(hwnd: int, paths: List[str]) -> bool:
    """
    Erzeugt einen echten Win32 HDROP-Speicherblock (DROPFILES-Header + UTF-16LE-Pfadliste)
    und sendet WM_DROPFILES per PostMessageW an das angegebene Fenster-Handle (wie Windows Explorer).
    """
    if sys.platform != "win32" or not paths:
        return False
    import struct
    encoded_paths = b"".join(str(p).encode("utf-16le") + b"\x00\x00" for p in paths) + b"\x00\x00"
    dropfiles_hdr = struct.pack("<IiiII", 20, 10, 10, 0, 1)
    payload = dropfiles_hdr + encoded_paths

    kernel32 = ctypes.windll.kernel32
    kernel32.GlobalAlloc.argtypes  = [wintypes.UINT, ctypes.c_size_t]
    kernel32.GlobalAlloc.restype   = ctypes.c_void_p
    kernel32.GlobalLock.argtypes   = [ctypes.c_void_p]
    kernel32.GlobalLock.restype    = ctypes.c_void_p
    kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]

    hglobal = kernel32.GlobalAlloc(0x0042, len(payload))
    if not hglobal:
        return False
    ptr = kernel32.GlobalLock(hglobal)
    ctypes.memmove(ptr, payload, len(payload))
    kernel32.GlobalUnlock(hglobal)

    user32 = ctypes.windll.user32
    user32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
    user32.PostMessageW.restype  = wintypes.BOOL
    return bool(user32.PostMessageW(hwnd, WM_DROPFILES, hglobal, 0))


# ── Einstiegspunkt ────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] in ("--batch", "--single", "--portfolio", "--test-dnd-gui"):
        import argparse
        parser = argparse.ArgumentParser(description="WAGO AAS PCF Query CLI")
        parser.add_argument("--batch", nargs=2, metavar=("INPUT", "OUTPUT"))
        parser.add_argument("--single", metavar="ARTICLE")
        parser.add_argument("--portfolio", metavar="OUTPUT")
        parser.add_argument("--test-dnd-gui", nargs=2, metavar=("SCENARIOS_JSON", "REPORT_JSON"))
        parser.add_argument("--max-shells", type=int, default=None)
        parser.add_argument("--mode", choices=[MODE_ONLINE, MODE_OFFLINE], default=MODE_ONLINE)
        parser.add_argument("--column", default=None)
        parser.add_argument("--sheet", default=None)
        args = parser.parse_args()

        if args.batch:
            res_summary = main(
                args.batch[0],
                args.batch[1],
                article_column=args.column,
                sheet_name=args.sheet,
                mode=args.mode,
            )
            sys.exit(0 if res_summary is not None else 1)
        elif args.single:
            res_single = query_single_article(args.single, mode=args.mode)
            print(json.dumps(res_single))
            sys.exit(0)
        elif args.portfolio:
            res_port = run_portfolio_coverage_analysis(max_shells=args.max_shells)
            export_portfolio_coverage_excel(res_port, args.portfolio)
            print(json.dumps({
                "total_shells_retrieved": res_port["total_shells_retrieved"],
                "total_unique_products": res_port["total_unique_products"],
                "pcf_coverage_pct": res_port["pcf_coverage_pct"],
                "output_file": args.portfolio,
            }))
            sys.exit(0)
        elif args.test_dnd_gui:
            scenarios_path, report_path = args.test_dnd_gui
            scenarios = json.loads(Path(scenarios_path).read_text(encoding="utf-8"))
            app = PCFQueryApp()
            results_out: List[Dict[str, Any]] = []
            step_idx = [0]

            def _run_next_step():
                idx = step_idx[0]
                if idx >= len(scenarios):
                    Path(report_path).write_text(
                        json.dumps({"status": "ok", "results": results_out}, indent=2),
                        encoding="utf-8",
                    )
                    app.destroy()
                    return
                scen = scenarios[idx]
                _post_win32_hdrop_message(app.winfo_id(), scen["paths"])
                app.after(140, _record_step)

            def _record_step():
                idx = step_idx[0]
                scen = scenarios[idx]
                results_out.append({
                    "name": scen["name"],
                    "input_path": app._input_path,
                    "hint_text": app._lbl_drop_hint.cget("text"),
                    "query_running": bool(app._query_thread and app._query_thread.is_alive()),
                })
                step_idx[0] += 1
                app.after(20, _run_next_step)

            app.after(80, _run_next_step)
            app.mainloop()
            sys.exit(0)
    else:
        app = PCFQueryApp()
        app.mainloop()
