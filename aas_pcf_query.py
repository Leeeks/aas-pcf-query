"""
aas_pcf_query.py
================
Liest WAGO-Artikelnummern aus einer Excel-Datei, fragt fuer jede Nummer den
PCF-Wert (Product Carbon Footprint, kg CO2e) ueber die WAGO AAS API ab
und schreibt die Ergebnisse kundenfreundlich in eine neue Excel-Datei.

Ausfuehren:
    python aas_pcf_query.py

Technische Grundlage (WAGO SmartService AAS API v3.2):
    Basis-URL       : https://c1.api.wago.com/smartservice-aas
    Submodel-ID     : https://www.wago.com/submodel/{article_short}/CarbonFootprint
    Shell-ID        : https://wago.com/ids/aas/{article_short}
    GlobalAsset-ID  : https://wago.com/ids/asset/{article_short}
    Fast-Path-URL   : {AAS_BASE_URL}/submodels/{base64(Submodel-ID)}
    Fallback-URL    : {AAS_REGISTRY_BASE}/shell-descriptors?assetIds={base64url(AssetIdObj)}

    Der PCF-Wert befindet sich im Feld PcfCO2eq innerhalb der
    SubmodelElementList ProductCarbonFootprints.
"""

from __future__ import annotations

import base64
import json
import os
import re
import subprocess
import sys
import threading
import time
import tkinter as tk
import urllib.parse
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable, Dict, List, Optional, Tuple, Union

import openpyxl
import pandas as pd
import requests
import urllib3

# ── Konfiguration (Zentralisiert) ────────────────────────────────────────────

ARTICLE_COLUMN = "Article Number"   # Standard-Spaltenname (Fallback/Default)

COL_ARTICLE_SHORT = "Article Number (Short)"
COL_ARTICLE_LONG  = "Article Number (Long)"
COL_PCF           = "PCF"
COL_PCF_STATUS    = "PCF Status"

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

# ── Status-Konstanten ────────────────────────────────────────────────────────

# Kundenfreundliche Statuswerte (fuer Excel-Export und UI)
STATUS_PCF_AVAILABLE     = "PCF available"
STATUS_NOT_YET_AVAILABLE = "Not yet available"
STATUS_NOT_AVAILABLE     = STATUS_NOT_YET_AVAILABLE  # Alias
STATUS_QUERY_ERROR       = "Query error"

# Interne Diagnose-Statuswerte (fuer Logging, Troubleshooting & interne Auswertung)
INTERNAL_STATUS_PCF_AVAILABLE = "PCF available"
INTERNAL_STATUS_AAS_NO_PCF    = "AAS available - PCF not available"
INTERNAL_STATUS_NOT_FOUND     = "Article not found"
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

# ── WAGO Artikelnummern-Normalisierung & Konvertierung ───────────────────────

# Primärer Block: <Serie>-<Artikel>[Suffix], z. B. 221-412, 0221-0412, 2000-115, 770-103-JPE
_RE_PRIMARY_NUMERIC = re.compile(
    r"^(\d{2,4})-(\d{3,4})((?:_[A-Za-z0-9]+|-[A-Za-z][A-Za-z0-9]*))?$"
)
# Primärer Block mit alphanumerischer Artikelkennung, z. B. 835-0FAR, 0835-0RSC
_RE_PRIMARY_ALPHANUM = re.compile(
    r"^(\d{3,4})-([0-9][A-Za-z0-9]{3})$"
)
# Varianten-Block (nach '/'): numerisch, z. B. 046-000, 0046-0000, 000-100, 281-413_HI
_RE_VARIANT_NUMERIC = re.compile(
    r"^(\d{1,4})-(\d{1,4})((?:_[A-Za-z0-9]+|-[A-Za-z][A-Za-z0-9]*))?$"
)
# Kunden-/Sonder-Varianten-Block (nach '/'): z. B. K010-2937, K000-0001, Z000-1809
_RE_VARIANT_KZ = re.compile(
    r"^[A-Za-z]\d{3}-\d{4}(?:_[A-Za-z0-9]+)?$"
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
        # 7-stellige Zahl entstand durch Wegfall der fuehrenden '0' in Excel (0100..0999)
        return 100 <= series <= 999
    if d8.startswith("0"):
        return 100 <= series <= 999
    # 4-stellige WAGO-Serien ohne fuehrende Null (z. B. 2000..2999, 7000..8999, 1000..1999)
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

    # Excel-Textpraefix (fuehrendes Apostroph) und Float-Suffix (.0) bereinigen
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

    # 2a. Erster Block (Hauptartikelnummer)
    seg0 = segments[0]
    m_num = _RE_PRIMARY_NUMERIC.match(seg0)
    m_alp = _RE_PRIMARY_ALPHANUM.match(seg0) if not m_num else None

    if m_num:
        p0, p1, suffix = m_num.group(1), m_num.group(2), m_num.group(3) or ""
        series_val = int(p0)
        # WAGO-Seriennummern liegen im Bereich >= 85 (verhindert Datums-/Bereichsangaben wie 12-345)
        if series_val < 85:
            return {
                "original": raw_str,
                "short": raw_str,
                "long": raw_str,
                "canonical": raw_str,
                "is_valid": False,
                "format": "invalid",
            }
        # 4-stellige Teilnummern, die bereits >= 1000 sind, duerfen nicht als 5-stellig interpretiert werden
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

    # 2b. Optionale Varianten-Segmente nach '/'
    for seg in segments[1:]:
        if _RE_VARIANT_KZ.match(seg) or _RE_VARIANT_ALPHA.match(seg):
            short_segments.append(seg)
            long_segments.append(seg)
            seg_formats.append("both")
            continue

        m_var = _RE_VARIANT_NUMERIC.match(seg)
        if m_var:
            vp0, vp1, vsuffix = m_var.group(1), m_var.group(2), m_var.group(3) or ""
            # Mindestens ein Teil sollte >= 3 Stellen haben oder eine Null-Variante sein
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

        # Unbekanntes/ungueltiges Segment
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
    """
    Konvertiert eine WAGO-Artikelnummer in das Kurzformat (z. B. '0221-0412' -> '221-412').
    Gibt bei nicht eindeutig konvertierbaren Werten den Originalwert zurueck.
    """
    return parse_article_number(value)["short"]


def to_long_article_number(value: Any) -> str:
    """
    Konvertiert eine WAGO-Artikelnummer in das Langformat (z. B. '221-412' -> '0221-0412').
    Gibt bei nicht eindeutig konvertierbaren Werten den Originalwert zurueck.
    """
    return parse_article_number(value)["long"]


# Kompakte Funktions-Aliase
to_short_format = to_short_article_number
to_long_format  = to_long_article_number


def normalize_article_number(value: Any) -> str:
    """
    Normalisiert eine WAGO-Artikelnummer auf ihre kanonische Darstellung (Kurzformat)
    fuer zuverlässige API-Abfragen und Deduplizierung.
    """
    return parse_article_number(value)["canonical"]


def convert_article_number(value: Any) -> Tuple[str, str]:
    """Gibt (short_format, long_format) fuer eine Artikelnummer zurueck."""
    parsed = parse_article_number(value)
    return parsed["short"], parsed["long"]


# ── Struktur-agnostische Excel-Erkennung & -Verarbeitung ─────────────────────

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
    """
    Bewertet einen Spaltennamen als sekundaeres Signal (0.0 bis 0.15).
    Komponenten-Spalten erhalten einen kleinen Abzug gegenueber Hauptartikel-Spalten.
    """
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
    """
    Analysiert alle Spalten eines DataFrames inhaltsbasiert auf WAGO-Artikelnummern.
    """
    candidates: List[Dict[str, Any]] = []
    col_parsed_map: Dict[str, List[Dict[str, Any]]] = {}

    for col in df.columns:
        series = df[col]
        # Falls doppelte Spaltennamen existieren und df[col] ein DataFrame ist, erste Spalte nehmen
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

        # Konsistenz der Formate über die Zeilen hinweg
        short_compat = fmt_counts["short"] + fmt_counts["both"]
        long_compat  = fmt_counts["long"] + fmt_counts["both"]
        max_consistent = max(short_compat, long_compat, fmt_counts["numeric"], fmt_counts["mixed"])
        consistency = max_consistent / valid_cnt if valid_cnt > 0 else 0.0

        # Dominantes Format bestimmen
        if fmt_counts["long"] > fmt_counts["short"] and fmt_counts["long"] >= fmt_counts["numeric"]:
            dom_fmt = "long"
        elif fmt_counts["short"] > 0 or fmt_counts["both"] > 0:
            dom_fmt = "short"
        elif fmt_counts["numeric"] > 0:
            dom_fmt = "numeric"
        else:
            dom_fmt = "mixed"

        header_score = _score_column_header(col)

        # Inhaltsbasierter Gesamt-Confidence-Score:
        # - match_ratio ist das Hauptkriterium (bis 0.70)
        # - Konsistenz + explizite Bindestrich-Struktur (bis 0.15)
        # - Header-Signal als sekundaerer Tie-Breaker (bis 0.15)
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

    # Sortiere Kandidaten absteigend nach Confidence, dann valid_count, dann match_ratio
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

    # Pruefe, ob die Top-Kandidaten ein zusammengehoeriges (Short, Long)-Spaltenpaar
    # fuer dieselben Produkte bilden oder ob konkurrierende Spalten vorliegen.
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

        # Wenn zwei Spalten dieselben Produkte darstellen (eine Short, eine Long)
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
            # Verschiedene Artikelnummern in mehreren plausiblen Spalten:
            # Ambiguitaet liegt vor, wenn beide hohe Match-Ratio haben und die Scores sehr nah sind.
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
    Liest alle Zeilen eines openpyxl-Worksheets und stellt durch Excel-Zahlenformatierung
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
                # Integer-Wert (auch 2210412.0)
                if isinstance(val, float) and val.is_integer():
                    val = int(val)
                if isinstance(val, int):
                    s_int = str(val)
                    p = parse_article_number(s_int, number_format=num_fmt)
                    if p["is_valid"] and p["format"] == "numeric":
                        # Wenn das Zahlenformat 0000-0000 ist, Langformat, sonst Kurzformat/Original
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


def _parse_sheet_table(rows: List[List[str]]) -> Tuple[pd.DataFrame, Optional[int], bool]:
    """
    Erkennt automatisch:
    - Einleitende Zeilen vor der eigentlichen Tabelle (Introductory Rows)
    - Ob eine Kopfzeile (Header Row) existiert oder ob die Tabelle direkt mit Daten beginnt
    Gibt (df, header_row_index, has_header) zurueck.
    """
    # Entferne komplett leere Zeilen am Ende
    while rows and all(c == "" for c in rows[-1]):
        rows.pop()

    if not rows:
        return pd.DataFrame(), None, True

    max_cols = max(len(r) for r in rows)
    norm_rows = [r + [""] * (max_cols - len(r)) for r in rows]

    # Finde die erste Zeile, die mindestens eine gueltige WAGO-Artikelnummer enthaelt
    first_article_row: Optional[int] = None
    for r_idx, r in enumerate(norm_rows):
        if any(is_valid_wago_article(cell) for cell in r):
            first_article_row = r_idx
            break

    if first_article_row is None:
        # Keine WAGO-Artikelnummer im gesamten Sheet gefunden -> erste nicht-leere Zeile als Header
        for r_idx, r in enumerate(norm_rows):
            if any(c != "" for c in r):
                headers = [c if c != "" else f"Unnamed: {i}" for i, c in enumerate(r)]
                data = norm_rows[r_idx + 1:]
                return pd.DataFrame(data, columns=headers), r_idx, True
        return pd.DataFrame(), None, True

    if first_article_row == 0:
        # Erste Zeile enthaelt bereits Artikelnummern -> Tabelle hat KEINE Header-Zeile!
        headers = [f"Column {i + 1}" for i in range(max_cols)]
        df = pd.DataFrame(norm_rows, columns=headers)
        return df, None, False

    # Pruefe die Zeile direkt vor first_article_row als potenzielle Header-Zeile
    prev_row = norm_rows[first_article_row - 1]
    prev_non_empty = [c for c in prev_row if c != ""]

    if prev_non_empty and not any(is_valid_wago_article(c) for c in prev_row):
        header_row_idx = first_article_row - 1
        raw_headers = norm_rows[header_row_idx]
        headers: List[str] = []
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
        # Zeile vor first_article_row ist leer -> Tabelle ab first_article_row ohne Header
        headers = [f"Column {i + 1}" for i in range(max_cols)]
        data = norm_rows[first_article_row:]
        df = pd.DataFrame(data, columns=headers)
        return df, None, False


def load_excel_flexible(
    path_or_df: Union[str, Path, pd.DataFrame],
    sheet_name: Optional[str] = None,
    article_column: Optional[str] = None,
) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Laedt eine Excel-Datei oder ein DataFrame struktur-agnostisch und erkennt
    automatisch das beste Worksheet, die Tabellen-Kopfzeile und die Artikelnummern-Spalte.

    Rueckgabe:
        (df, detection_info)
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
        return df, info

    path_str = str(path_or_df)
    wb = openpyxl.load_workbook(path_str, data_only=True)
    sheet_names = list(wb.sheetnames)

    target_sheets = [sheet_name] if (sheet_name and sheet_name in sheet_names) else sheet_names

    best_df: Optional[pd.DataFrame] = None
    best_info: Optional[Dict[str, Any]] = None

    for s_name in target_sheets:
        ws = wb[s_name]
        rows = _extract_sheet_rows_with_formats(ws)
        df_cand, header_row_idx, has_header = _parse_sheet_table(rows)
        if not df_cand.empty:
            # Entferne komplett leere angehaengte Spalten ohne Header
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

        if best_info is None or (
            info_cand["confidence"],
            info_cand["candidates"][0]["valid_count"] if info_cand["candidates"] else 0,
        ) > (
            best_info["confidence"],
            best_info["candidates"][0]["valid_count"] if best_info["candidates"] else 0,
        ):
            best_df = df_cand
            best_info = info_cand

    wb.close()

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
    Erkennt die Artikelnummern-Spalte in einer Excel-Datei oder einem DataFrame
    und gibt ein Diagnose-Dict mit 'column', 'confidence', 'is_ambiguous',
    'candidates', 'short_column', 'long_column' etc. zurueck.
    """
    _, info = load_excel_flexible(path_or_df, sheet_name=sheet_name)
    return info


# ── HTTP-Setup ───────────────────────────────────────────────────────────────

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

_thread_local = threading.local()


def _get_session() -> requests.Session:
    """
    Gibt eine thread-lokale requests.Session mit Retry-Logik zurueck.
    """
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
            "User-Agent": "WAGO-AAS-PCF-Query/2.0",
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

    Unterscheidung:
        404              -> _NOT_FOUND  (Ressource im AAS nicht vorhanden)
        401 / 403        -> _ERROR      (Authentifizierungsfehler)
        Timeout          -> _ERROR      (Zeitueberschreitung)
        5xx / sonst.     -> _ERROR      (Server-/Verbindungsfehler oder ungueltiges JSON)
        Erfolg (200)     -> dict        (JSON-Antwort)
    """
    _set_last_error(INTERNAL_STATUS_QUERY_ERROR)
    try:
        r = _get_session().get(url, timeout=timeout)
        if r.status_code == 404:
            return _NOT_FOUND
        if r.status_code in (401, 403):
            _set_last_error(INTERNAL_STATUS_AUTH_ERROR)
            print(f"  Warning: HTTP {r.status_code} (Authentication error) fuer {url}")
            return _ERROR
        if r.status_code >= 500:
            _set_last_error(INTERNAL_STATUS_SERVER_ERROR)
            print(f"  Warning: HTTP {r.status_code} (Server error) fuer {url}")
            return _ERROR
        r.raise_for_status()
        try:
            data = r.json()
        except (ValueError, TypeError) as e:
            _set_last_error(INTERNAL_STATUS_INVALID_RESP)
            print(f"  Warning: Ungueltige JSON-Antwort fuer {url}: {str(e)[:100]}")
            return _ERROR
        if not isinstance(data, dict):
            _set_last_error(INTERNAL_STATUS_INVALID_RESP)
            return _ERROR
        return data
    except requests.exceptions.Timeout as e:
        _set_last_error(INTERNAL_STATUS_TIMEOUT)
        print(f"  Warning: Timeout ({type(e).__name__}) fuer {url}")
        return _ERROR
    except requests.exceptions.HTTPError as e:
        code = getattr(e.response, "status_code", None)
        if code in (401, 403):
            _set_last_error(INTERNAL_STATUS_AUTH_ERROR)
        elif code and code >= 500:
            _set_last_error(INTERNAL_STATUS_SERVER_ERROR)
        else:
            _set_last_error(INTERNAL_STATUS_QUERY_ERROR)
        print(f"  Warning: HTTP {code} fuer {url}")
        return _ERROR
    except requests.RequestException as e:
        _set_last_error(INTERNAL_STATUS_QUERY_ERROR)
        print(f"  Warning: {type(e).__name__}: {str(e)[:100]}")
        return _ERROR
    except (ValueError, TypeError) as e:
        _set_last_error(INTERNAL_STATUS_INVALID_RESP)
        print(f"  Warning: JSON-Fehler: {str(e)[:100]}")
        return _ERROR


def _safe_jget(url: str, timeout: Tuple[int, int] = REQUEST_TIMEOUT) -> Any:
    """
    Wrapper um _jget(), der auch direkt geworfene Exceptions (z. B. bei gemocktem _jget
    in Unit-Tests) sauber in _ERROR + Diagnosegrund uebersetzt.
    """
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


# ── AAS-Hilfsfunktionen ──────────────────────────────────────────────────────

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
    """
    Konstruiert die direkte AAS-Submodel-URL fuer das CarbonFootprint-Submodel
    der gegebenen (normalisierten) Artikelnummer.

    URL-Muster (SmartService AAS API v3.2):
        ID      : https://www.wago.com/submodel/{article_short}/CarbonFootprint
        Endpoint: {AAS_ENV_BASE}/submodels/{base64(ID)}
    """
    canonical = normalize_article_number(article_number)
    submodel_id = SUBMODEL_ID_TEMPLATE.format(article=canonical)
    return _submodel_url_from_id(submodel_id)


def _registry_lookup_url(article_number: str) -> str:
    """
    Konstruiert die Shell-Registry-Lookup-URL ueber globalAssetId (AAS v3.2 Objektformat).
    """
    canonical = normalize_article_number(article_number)
    asset_id = ASSET_ID_TEMPLATE.format(article=canonical)
    asset_id_obj = {"name": "globalAssetId", "value": asset_id}
    asset_id_encoded = _b64url_encode(json.dumps(asset_id_obj))
    return f"{AAS_REGISTRY_BASE}/shell-descriptors?assetIds={urllib.parse.quote(asset_id_encoded)}"


def _extract_all_values(elements: Any, path: str = "") -> Dict[str, Any]:
    """
    Traversiert rekursiv alle submodelElements (inkl. SubmodelElementList und
    SubmodelElementCollection) und gibt ein Dict {Pfad: Wert} zurueck.

    Kompatibel mit AAS v3.0 / v3.2 (SubmodelElementList ohne idShort bei Listenelementen).
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


# ── Kernfunktionen (PCF-Evaluation & API-Abfrage) ────────────────────────────

def _evaluate_pcf_result_detailed(result: Any) -> Tuple[Optional[float], str, str]:
    """
    Wertet eine Submodel-API-Antwort aus und gibt
    (pcf_value, customer_status, internal_status) zurueck.
    """
    if result is _NOT_FOUND:
        return None, STATUS_NOT_YET_AVAILABLE, INTERNAL_STATUS_NOT_FOUND

    if result is _ERROR:
        return None, STATUS_QUERY_ERROR, _get_last_error()

    if not isinstance(result, dict):
        return None, STATUS_QUERY_ERROR, INTERNAL_STATUS_INVALID_RESP

    # Pruefe, ob die Antwort eine valide AAS-Submodel-Struktur besitzt
    has_sm_elements = "submodelElements" in result and isinstance(result.get("submodelElements"), list)
    is_submodel_type = result.get("modelType") == "Submodel" and (
        result.get("submodelElements") is None or isinstance(result.get("submodelElements"), list)
    )
    if not has_sm_elements and not is_submodel_type:
        return None, STATUS_QUERY_ERROR, INTERNAL_STATUS_INVALID_RESP

    all_values = _extract_all_values(result.get("submodelElements") or [])
    for path, value in all_values.items():
        if "pcfco2eq" in path.lower():
            parsed = _parse_float(value)
            if parsed is not None:
                return parsed, STATUS_PCF_AVAILABLE, INTERNAL_STATUS_PCF_AVAILABLE

    return None, STATUS_NOT_YET_AVAILABLE, INTERNAL_STATUS_AAS_NO_PCF


def _evaluate_pcf_result(result: Any) -> Tuple[Optional[float], str]:
    """
    Gibt (pcf_value, customer_status) fuer eine Submodel-Antwort zurueck.
    """
    pcf_val, cust_status, _ = _evaluate_pcf_result_detailed(result)
    return pcf_val, cust_status


def get_pcf_result(article_number: Any) -> Dict[str, Any]:
    """
    Gibt ein Dict mit PCF-Wert, kundenfreundlichem Status, internem Diagnosestatus
    sowie Kurz-/Langformat der Artikelnummer zurueck.

    Nutzt:
    1. Automatische Normalisierung auf das WAGO-Kurzformat (damit '221-412' und
       '0221-0412' identisch aufgeloest werden).
    2. 'Fast Path' (direkter Abruf des CarbonFootprint-Submodels).
    3. 'Safe Fallback' (Shell-Descriptor-Lookup ueber globalAssetId).
    """
    parsed_art = parse_article_number(article_number)
    if not parsed_art["original"]:
        return {
            "pcf": None,
            "status": STATUS_NOT_YET_AVAILABLE,
            "internal_status": INTERNAL_STATUS_NOT_FOUND,
            "article_short": "",
            "article_long": "",
        }

    query_article = parsed_art["canonical"]
    short_art     = parsed_art["short"]
    long_art      = parsed_art["long"]

    # 1. Fast Path: Direkter Abruf des CarbonFootprint-Submodels
    url = _submodel_url(query_article)
    fast_result = _safe_jget(url)

    pcf_val, pcf_status, internal_status = _evaluate_pcf_result_detailed(fast_result)
    if pcf_status == STATUS_PCF_AVAILABLE:
        return {
            "pcf": pcf_val,
            "status": STATUS_PCF_AVAILABLE,
            "internal_status": INTERNAL_STATUS_PCF_AVAILABLE,
            "article_short": short_art,
            "article_long": long_art,
        }

    # Technische Fehler (Timeout, Auth-Fehler, Server-Fehler, ungueltige Response)
    # duerfen NICHT als "Not yet available" fehlinterpretiert werden.
    if pcf_status == STATUS_QUERY_ERROR:
        return {
            "pcf": None,
            "status": STATUS_QUERY_ERROR,
            "internal_status": internal_status,
            "article_short": short_art,
            "article_long": long_art,
        }

    # 2. Safe Fallback: Shell-Descriptor-Lookup via globalAssetId
    reg_url = _registry_lookup_url(query_article)
    reg_res = _safe_jget(reg_url, timeout=(10, 30))

    if reg_res is _ERROR:
        return {
            "pcf": None,
            "status": STATUS_QUERY_ERROR,
            "internal_status": _get_last_error(),
            "article_short": short_art,
            "article_long": long_art,
        }

    if reg_res is _NOT_FOUND:
        return {
            "pcf": None,
            "status": STATUS_NOT_YET_AVAILABLE,
            "internal_status": INTERNAL_STATUS_NOT_FOUND,
            "article_short": short_art,
            "article_long": long_art,
        }

    if not isinstance(reg_res, dict):
        return {
            "pcf": None,
            "status": STATUS_QUERY_ERROR,
            "internal_status": INTERNAL_STATUS_INVALID_RESP,
            "article_short": short_art,
            "article_long": long_art,
        }

    if isinstance(reg_res.get("result"), list):
        shells = reg_res["result"]
    elif "submodelDescriptors" in reg_res or "submodels" in reg_res:
        shells = [reg_res]
    else:
        return {
            "pcf": None,
            "status": STATUS_QUERY_ERROR,
            "internal_status": INTERNAL_STATUS_INVALID_RESP,
            "article_short": short_art,
            "article_long": long_art,
        }

    if not shells:
        return {
            "pcf": None,
            "status": STATUS_NOT_YET_AVAILABLE,
            "internal_status": internal_status,
            "article_short": short_art,
            "article_long": long_art,
        }

    # Shell existiert im AAS -> Submodel-Descriptoren nach PCF / CarbonFootprint durchsuchen
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
                    fb_val, fb_status, fb_internal = _evaluate_pcf_result_detailed(sm_res)
                    if fb_status == STATUS_PCF_AVAILABLE:
                        return {
                            "pcf": fb_val,
                            "status": STATUS_PCF_AVAILABLE,
                            "internal_status": INTERNAL_STATUS_PCF_AVAILABLE,
                            "article_short": short_art,
                            "article_long": long_art,
                        }
                    if fb_status == STATUS_QUERY_ERROR:
                        return {
                            "pcf": None,
                            "status": STATUS_QUERY_ERROR,
                            "internal_status": fb_internal,
                            "article_short": short_art,
                            "article_long": long_art,
                        }

    # Shell ist vorhanden, aber kein PCF-Submodel / kein PCF-Wert hinterlegt
    return {
        "pcf": None,
        "status": STATUS_NOT_YET_AVAILABLE,
        "internal_status": INTERNAL_STATUS_AAS_NO_PCF,
        "article_short": short_art,
        "article_long": long_art,
    }


def get_pcf(article_number: Any) -> Optional[float]:
    """
    Rueckwaerts-kompatibler Wrapper: gibt nur den PCF-Wert (float | None) zurueck.
    """
    return get_pcf_result(article_number)["pcf"]


def _query_with_timing(article_number: str) -> Dict[str, Any]:
    """
    Fuehrt get_pcf_result() aus und loggt langsame Requests.
    Geeignet fuer Worker-Threads.
    """
    t0 = time.perf_counter()
    try:
        res = get_pcf_result(article_number)
    except Exception as e:
        print(f"  Unerwarteter Fehler fuer {article_number}: {e}")
        parsed = parse_article_number(article_number)
        res = {
            "pcf": None,
            "status": STATUS_QUERY_ERROR,
            "internal_status": INTERNAL_STATUS_QUERY_ERROR,
            "article_short": parsed["short"],
            "article_long": parsed["long"],
        }
    dur = time.perf_counter() - t0
    if dur >= SLOW_REQUEST_THRESHOLD:
        diag = res.get("internal_status") or res.get("status")
        print(f"  Slow query: {article_number} - {dur:.1f}s [{diag}]")
    return res


# ── Excel-Aufbereitung & Export ──────────────────────────────────────────────

def _extract_row_articles(
    df: pd.DataFrame,
    detection_info: Dict[str, Any],
    article_column: Optional[str] = None,
) -> Tuple[List[str], List[str], List[str]]:
    """
    Extrahiert fuer jede Zeile des DataFrames:
        (canonical_articles, short_articles, long_articles)

    Wenn im Input bereits zwei zusammengehoerige Spalten (Short & Long) existieren,
    werden beide wiederverwendet, gegeneinander validiert und fehlende Eintraege ergaenzt.
    """
    primary_col = article_column or detection_info.get("column")
    short_col   = detection_info.get("short_column")
    long_col    = detection_info.get("long_column")
    has_both    = bool(detection_info.get("has_both_formats")) and not article_column

    if not has_both:
        # Pruefe auch explizit benannte Spalten im DataFrame
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
                    # Bei Widerspruch: primaere Spalte bevorzugen
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
    Erzeugt das Export-DataFrame:
    - Bewahrt alle bestehenden Spalten und Originalwerte.
    - Stellt sicher, dass sowohl 'Article Number (Short)' als auch 'Article Number (Long)'
      vorhanden sind (wiederverwendet bestehende Spalten ohne Duplikate oder ergaenzt fehlende).
    - Schreibt kundenfreundliche PCF-Werte und 'PCF Status' ('PCF available',
      'Not yet available', 'Query error') ohne interne Diagnosegruende.
    """
    out_df = df.copy()
    primary_col = article_column or detection_info.get("column")
    canonical_list, short_list, long_list = _extract_row_articles(
        out_df, detection_info, article_column=article_column
    )

    col_lookup = {str(c).strip().lower(): c for c in out_df.columns}
    existing_short_col = col_lookup.get(COL_ARTICLE_SHORT.lower()) or detection_info.get("short_column")
    existing_long_col  = col_lookup.get(COL_ARTICLE_LONG.lower()) or detection_info.get("long_column")

    # Wenn im Input bereits ein zusammengehoeriges (Short, Long)-Paar unter anderen Namen existiert,
    # validiere und vervollstaendige diese Spalten ebenfalls in-place.
    if detection_info.get("has_both_formats") and not article_column:
        s_col = detection_info.get("short_column")
        l_col = detection_info.get("long_column")
        if s_col in out_df.columns:
            out_df[s_col] = short_list
        if l_col in out_df.columns:
            out_df[l_col] = long_list

    # Bestimme Einfuegeposition fuer fehlende Short/Long-Spalten (direkt nach der Artikelspalte)
    anchor_col = existing_short_col or primary_col
    insert_pos = (
        list(out_df.columns).index(anchor_col) + 1
        if (anchor_col and anchor_col in out_df.columns)
        else 0
    )

    # 1. Article Number (Short)
    if COL_ARTICLE_SHORT in out_df.columns:
        out_df[COL_ARTICLE_SHORT] = short_list
    elif existing_short_col and str(existing_short_col).strip().lower() == COL_ARTICLE_SHORT.lower():
        out_df.rename(columns={existing_short_col: COL_ARTICLE_SHORT}, inplace=True)
        out_df[COL_ARTICLE_SHORT] = short_list
    else:
        out_df.insert(insert_pos, COL_ARTICLE_SHORT, short_list)
        insert_pos += 1

    # 2. Article Number (Long)
    if COL_ARTICLE_LONG in out_df.columns:
        out_df[COL_ARTICLE_LONG] = long_list
    elif existing_long_col and str(existing_long_col).strip().lower() == COL_ARTICLE_LONG.lower():
        out_df.rename(columns={existing_long_col: COL_ARTICLE_LONG}, inplace=True)
        out_df[COL_ARTICLE_LONG] = long_list
    else:
        if COL_ARTICLE_SHORT in out_df.columns:
            long_pos = list(out_df.columns).index(COL_ARTICLE_SHORT) + 1
        else:
            long_pos = insert_pos
        out_df.insert(long_pos, COL_ARTICLE_LONG, long_list)

    # 3. PCF & PCF Status berechnen
    pcf_values: List[Any]   = []
    pcf_statuses: List[str] = []

    for canonical in canonical_list:
        if not canonical:
            pcf_values.append(STATUS_NOT_YET_AVAILABLE)
            pcf_statuses.append("")
        else:
            res = cache.get(canonical, {"pcf": None, "status": STATUS_QUERY_ERROR})
            status = res.get("status", STATUS_QUERY_ERROR)
            if res.get("pcf") is not None:
                pcf_values.append(res["pcf"])
                pcf_statuses.append(STATUS_PCF_AVAILABLE)
            elif status == STATUS_QUERY_ERROR:
                # Technischer Fehler darf nicht als "Not yet available" ausgegeben werden
                pcf_values.append("")
                pcf_statuses.append(STATUS_QUERY_ERROR)
            else:
                pcf_values.append(STATUS_NOT_YET_AVAILABLE)
                pcf_statuses.append(STATUS_NOT_YET_AVAILABLE)

    out_df[COL_PCF]        = pcf_values
    out_df[COL_PCF_STATUS] = pcf_statuses

    return out_df


# ── Hauptprogramm (Batch, ohne GUI) ──────────────────────────────────────────

def main(
    input_file: str,
    output_file: str,
    progress_callback: Optional[Callable[[int, int, str, Dict[str, Any]], None]] = None,
    article_column: Optional[str] = None,
    sheet_name: Optional[str] = None,
) -> Optional[Dict[str, Any]]:
    """
    Verarbeitet eine Excel-Datei mit WAGO-Artikelnummern.

    Parameter:
        input_file        : Pfad zur Eingabedatei (str)
        output_file       : Pfad zur Ausgabedatei (str)
        progress_callback : Optionale Funktion(current, total, article, result_dict)
        article_column    : Optionale explizite Spaltenauswahl (ueberschreibt Auto-Detection)
        sheet_name        : Optionaler Worksheet-Name

    Rueckgabe:
        dict mit Zusammenfassung oder None bei Fehler
    """
    print(f"Lese {input_file} ...")
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
    unique_articles = list(dict.fromkeys(a for a in canonical_all if a))
    unique_total    = len(unique_articles)
    total           = len(canonical_all)

    print(
        f"Erkannte Spalte: '{detected_col}' "
        f"(Sheet: '{detection_info.get('sheet_name')}', Confidence: {detection_info.get('confidence', 0):.0%})"
    )
    print(f"Gefunden: {total} Zeilen, {unique_total} eindeutige Produkte (normalisiert)")
    print(f"Starte Abfrage mit {WORKERS} parallelen Threads ...\n")

    cache: Dict[str, Dict[str, Any]] = {}
    cache_lock = threading.Lock()
    done_count = [0]

    def _worker(article: str) -> Tuple[str, Dict[str, Any]]:
        res = _query_with_timing(article)
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
                    pcf_str = f"{res['pcf']} kg CO2e"
                elif res["status"] == STATUS_QUERY_ERROR:
                    pcf_str = STATUS_QUERY_ERROR
                else:
                    pcf_str = STATUS_NOT_YET_AVAILABLE
                print(f"  {article:<25} {pcf_str:<22} [{res['status']}]")

    out_df = prepare_output_dataframe(
        df, detection_info, cache, article_column=article_column
    )

    try:
        out_df.to_excel(output_file, index=False)
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

    for canonical in canonical_all:
        if not canonical:
            continue
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
        "unique_queried":    unique_total,
        "pcf_available":     cust_counts[STATUS_PCF_AVAILABLE],
        "not_yet_available": cust_counts[STATUS_NOT_YET_AVAILABLE],
        "query_errors":      cust_counts[STATUS_QUERY_ERROR],
        # Interne Diagnosezahlen fuer Troubleshooting beibehalten:
        "aas_no_pcf":        internal_counts[INTERNAL_STATUS_AAS_NO_PCF],
        "not_found":         internal_counts[INTERNAL_STATUS_NOT_FOUND],
        "detected_column":   detected_col,
        "sheet_name":        detection_info.get("sheet_name"),
        "output_file":       output_file,
    }

    print(f"""
Finished.

Rows total        : {total}
Unique queried    : {unique_total}
PCF available     : {summary['pcf_available']}
Not yet available : {summary['not_yet_available']}
Query errors      : {summary['query_errors']}

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

    C_BG           = "#1e2228"
    C_PANEL        = "#262c35"
    C_BORDER       = "#353d47"
    C_ACCENT       = "#8cc63f"
    C_ACCENT_HOVER = "#9dd94f"
    C_SUCCESS      = "#4caf73"
    C_WARN         = "#f0a500"
    C_ERROR        = "#e05252"
    C_TEXT         = "#e8eaed"
    C_MUTED        = "#8a9bb0"
    C_BTN_FG       = "#ffffff"

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

        self._input_path: Optional[str]              = None
        self._output_path: Optional[str]             = None
        self._df: Optional[pd.DataFrame]             = None
        self._detection_info: Dict[str, Any]         = {}
        self._selected_sheet: Optional[str]          = None
        self._selected_column: Optional[str]         = None
        self._canonical_articles: List[str]          = []
        self._query_thread: Optional[threading.Thread] = None
        self._stop_event                             = threading.Event()

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

        def get_resource_path(relative_path: str) -> str:
            if hasattr(sys, "_MEIPASS"):
                return os.path.join(sys._MEIPASS, relative_path)
            return os.path.join(os.path.abspath(os.path.dirname(__file__)), relative_path)

        try:
            from PIL import Image, ImageTk
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

        # Optionale Spalten-/Sheet-Auswahl (automatisch vorbelegt, manuell anpassbar)
        self._frm_selectors = tk.Frame(self, bg=self.C_BG)
        self._frm_selectors.pack(fill="x", padx=16, pady=(0, 4))

        tk.Label(
            self._frm_selectors, text="Sheet:", font=self.FONT_SMALL,
            fg=self.C_MUTED, bg=self.C_BG,
        ).pack(side="left", padx=(0, 4))
        self._sheet_var = tk.StringVar(value="")
        self._cmb_sheet = ttk.Combobox(
            self._frm_selectors, textvariable=self._sheet_var,
            state="disabled", width=16, font=self.FONT_SMALL,
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
            state="disabled", width=22, font=self.FONT_SMALL,
        )
        self._cmb_col.pack(side="left")
        self._cmb_col.bind("<<ComboboxSelected>>", self._on_column_changed)

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
        self._section("Preview  (first 10 unique products: Short / Long)")
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

        self._result_rows: Dict[str, tk.Label] = {}
        result_defs = [
            ("pcf_available",     "PCF available:",     self.C_SUCCESS),
            ("not_yet_available", "Not yet available:", self.C_WARN),
            ("query_errors",      "Query errors:",      self.C_ERROR),
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
            font=self.FONT_SMALL, fg=self.C_MUTED, bg=self.C_BG,
        )
        lbl_footer.pack(side="left")

        btn_about = tk.Button(
            frm_footer, text="About", command=self._on_about,
            font=self.FONT_SMALL, fg=self.C_MUTED, bg=self.C_BG,
            relief="flat", cursor="hand2", bd=0,
            activebackground=self.C_BG, activeforeground=self.C_TEXT,
        )
        btn_about.pack(side="right")
        btn_about.bind("<Enter>", lambda e: btn_about.config(fg=self.C_TEXT))
        btn_about.bind("<Leave>", lambda e: btn_about.config(fg=self.C_MUTED))

    # ── Widget-Helfer ─────────────────────────────────────────────────────────

    def _divider(self):
        tk.Frame(self, bg=self.C_BORDER, height=1).pack(fill="x", padx=8, pady=4)

    def _section(self, text: str):
        tk.Label(
            self, text=text.upper(),
            font=("Segoe UI", 7, "bold"), fg=self.C_MUTED, bg=self.C_BG,
        ).pack(anchor="w", padx=16, pady=(4, 0))

    def _make_panel(self) -> tk.Frame:
        return tk.Frame(
            self, bg=self.C_PANEL, relief="flat",
            highlightbackground=self.C_BORDER, highlightthickness=1,
        )

    def _make_button(self, parent: tk.Widget, text: str, command: Callable, accent: bool = False) -> tk.Button:
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
        messagebox.showinfo(
            "About",
            "WAGO AAS PCF Query\nDeveloped by Alexander Flekler\nCorporate Sustainability\nVersion 2.0",
        )

    def _on_select_file(self):
        path = filedialog.askopenfilename(
            title="WAGO PCF Query – Excel-Datei auswaehlen",
            filetypes=[("Excel-Dateien", "*.xlsx"), ("Alle Dateien", "*.*")],
        )
        if path:
            self._load_file(path)

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
        self._lbl_path.config(text=path, fg=self.C_TEXT)
        self._df = None

        try:
            df, info = load_excel_flexible(
                path, sheet_name=sheet_name, article_column=article_column
            )
        except Exception as e:
            self._lbl_summary.config(
                text=f"Could not read file:\n{e}", fg=self.C_ERROR,
            )
            self._lbl_preview.config(text="—", fg=self.C_MUTED)
            self._cmb_sheet.config(state="disabled")
            self._cmb_col.config(state="disabled")
            return

        self._df = df
        self._detection_info = info
        self._selected_sheet = info.get("sheet_name")
        self._selected_column = article_column or info.get("column")

        # Update Comboboxen fuer Sheet- und Spaltenauswahl
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

        # Vorschau (erste 10 eindeutige Produkte mit Short & Long)
        seen_prev = set()
        preview_lines: List[str] = []
        for c, s, l in valid_pairs:
            if c not in seen_prev:
                seen_prev.add(c)
                if len(preview_lines) < 10:
                    preview_lines.append(f"  {s:<22} ->  {l}")

        if len(unique_canonical) > 10:
            preview_text = "\n".join(preview_lines) + f"\n  ... ({len(unique_canonical) - 10} more)"
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
        """
        Orchestrierungs-Thread: verteilt eindeutige normalisierte Artikel auf Worker via
        ThreadPoolExecutor, sammelt Ergebnisse und sendet GUI-Updates thread-safe via after().
        """
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

        cache: Dict[str, Dict[str, Any]] = {}
        cache_lock = threading.Lock()
        done_count = [0]

        def _worker(article: str) -> Tuple[str, Dict[str, Any]]:
            if self._stop_event.is_set():
                return article, {"pcf": None, "status": STATUS_QUERY_ERROR}

            res = _query_with_timing(article)

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
        out_df = prepare_output_dataframe(
            self._df, self._detection_info, cache, article_column=override_col
        )

        summary = None
        try:
            out_df.to_excel(self._output_path, index=False)
            summary = {
                "unique_queried":    unique_total,
                "pcf_available":     counts[STATUS_PCF_AVAILABLE],
                "not_yet_available": counts[STATUS_NOT_YET_AVAILABLE],
                "query_errors":      counts[STATUS_QUERY_ERROR],
                "output_file":       self._output_path,
            }
        except Exception as e:
            print(f"Fehler beim Schreiben der Ausgabedatei: {e}")
            err = str(e)
            self.after(0, lambda msg=err: messagebox.showerror(
                "Write Error", f"Could not write output file:\n{msg}",
            ))

        self.after(0, lambda s=summary: self._on_query_done(s))

    # ── GUI-Update-Callbacks (Main-Thread via after()) ─────────────────────────

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
            text=f"Current article: {article}", fg=self.C_ACCENT,
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
        self._btn_start.config(state=state)
        if running:
            self._cmb_sheet.config(state="disabled")
            self._cmb_col.config(state="disabled")
        else:
            sheets = self._detection_info.get("available_sheets") or []
            self._cmb_sheet.config(state="readonly" if len(sheets) > 1 else "disabled")
            cols = list(self._df.columns) if self._df is not None else []
            self._cmb_col.config(state="readonly" if len(cols) > 1 else "disabled")

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


# ── Einstiegspunkt ────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = PCFQueryApp()
    app.mainloop()
