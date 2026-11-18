"""
portfolio_coverage.py
=====================
Internal WAGO PCF Portfolio Coverage Analysis & Excel Report Generator.

Enumerates accessible Asset Administration Shells (AAS) from the live WAGO
SmartService AAS API, deduplicates products across article number formats,
verifies published PCF values and units, compares missing AAS PCFs against the
embedded Q4 2026 snapshot, and exports a multi-worksheet internal Excel report.
"""

from __future__ import annotations

import base64
import datetime
import io
import os
import re
import shutil
import tempfile
import threading
import time
import zipfile
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Set, Tuple, Union
from xml.sax.saxutils import escape as xml_escape

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import requests

from offline_pcf_data import (
    OFFLINE_PORTAL_URL,
    OFFLINE_SNAPSHOT_VERSION,
    get_offline_pcf_result,
)

# ── WAGO Corporate Design Palette ────────────────────────────────────────────

WAGO_GREEN         = "6EC800"
WAGO_WHITE         = "FFFFFF"
WAGO_LIGHT_GREY    = "EFF0F1"
WAGO_ANTHRACITE_15 = "DEDFE1"
WAGO_ANTHRACITE_40 = "A5A8AF"
WAGO_ANTHRACITE_70 = "616A73"
WAGO_ANTHRACITE    = "1F2837"
WAGO_BLACK         = "000000"

# Classification & Action Constants (Section 7.3)
CLASS_PCF_AVAILABLE_AAS   = "PCF available in AAS"
CLASS_PCF_OFFLINE_ONLY    = "PCF missing in AAS but available in embedded Q4 2026 dataset"
CLASS_PCF_MISSING_BOTH    = "PCF not found in either data source"
CLASS_TECHNICAL_ERROR     = "Product could not be evaluated due to technical errors"

ACTION_PUBLISHED          = "PCF published"
ACTION_REVIEW_PUBLICATION = "Review publication or data integration"
ACTION_REVIEW_CALCULATION = "Review potential calculation requirement"
ACTION_TECH_INVESTIGATION = "Technical investigation"


def extract_product_series(short_article: str) -> str:
    """
    Extrahiert die WAGO-Produktserie aus einer normalisierten Kurz-Artikelnummer
    (z. B. '221-412' -> '221', '2000-115' -> '2000', '231-004/046-000' -> '231').
    """
    if not short_article:
        return "Unknown"
    first_seg = str(short_article).split("/")[0]
    if "-" in first_seg:
        series = first_seg.split("-")[0].strip()
        if series:
            return series
    return "Other"


def _extract_article_candidate_from_shell(shell: Dict[str, Any]) -> str:
    """
    Extrahiert den Artikelnummern-Kandidaten aus einem AAS-Shell-Objekt oder Descriptor.
    Prueft der Reihe nach:
    1. idShort
    2. globalAssetId ('https://wago.com/ids/asset/<article>')
    3. id ('https://wago.com/ids/aas/<article>')
    """
    id_short = str(shell.get("idShort") or "").strip()
    if id_short:
        return id_short

    asset_info = shell.get("assetInformation")
    global_asset_id = ""
    if isinstance(asset_info, dict):
        global_asset_id = str(asset_info.get("globalAssetId") or "").strip()
    if not global_asset_id:
        global_asset_id = str(shell.get("globalAssetId") or "").strip()

    for prefix in ("https://wago.com/ids/asset/", "http://wago.com/ids/asset/"):
        if global_asset_id.startswith(prefix):
            return global_asset_id[len(prefix):].strip()

    aas_id = str(shell.get("id") or "").strip()
    for prefix in ("https://wago.com/ids/aas/", "http://wago.com/ids/aas/"):
        if aas_id.startswith(prefix):
            return aas_id[len(prefix):].strip()

    return ""


def _extract_pcf_submodel_refs_from_shell(shell: Dict[str, Any]) -> List[str]:
    """
    Sucht innerhalb eines Shell-Objekts (`submodels`) oder Shell-Descriptors
    (`submodelDescriptors`) nach Referenzen auf ein CarbonFootprint-/PCF-Submodel.
    """
    pcf_refs: List[str] = []

    # 1. AAS Shell format: submodels = [{"keys": [{"type": "Submodel", "value": "..."}]}]
    submodels = shell.get("submodels")
    if isinstance(submodels, list):
        for sm_ref in submodels:
            if not isinstance(sm_ref, dict):
                continue
            keys = sm_ref.get("keys")
            if isinstance(keys, list):
                for k in keys:
                    if isinstance(k, dict):
                        val = str(k.get("value") or "").strip()
                        val_l = val.lower()
                        if val and ("carbonfootprint" in val_l or "/pcf" in val_l):
                            if val not in pcf_refs:
                                pcf_refs.append(val)

    # 2. Shell Descriptor format: submodelDescriptors = [{"id": "...", "idShort": "..."}]
    descriptors = shell.get("submodelDescriptors")
    if isinstance(descriptors, list):
        for desc in descriptors:
            if not isinstance(desc, dict):
                continue
            sm_id = str(desc.get("id") or "").strip()
            id_short = str(desc.get("idShort") or "").strip().lower()
            sm_id_l = sm_id.lower()
            if sm_id and (
                "carbonfootprint" in id_short
                or "pcf" in id_short
                or "carbonfootprint" in sm_id_l
                or "/pcf" in sm_id_l
            ):
                if sm_id not in pcf_refs:
                    pcf_refs.append(sm_id)

    return pcf_refs


def _request_json_with_backoff(
    session_getter: Callable[[], requests.Session],
    url: str,
    timeout: Tuple[int, int] = (10, 25),
    max_retries: int = 3,
    base_backoff: float = 0.5,
    stop_event: Optional[threading.Event] = None,
) -> Tuple[Optional[Dict[str, Any]], Optional[int], str]:
    """
    Fuehrt einen HTTP GET Request mit Rate-Limit-Behandlung (HTTP 429 + Retry-After)
    und exponentiellem Backoff durch.

    Rueckgabe:
        (json_dict_or_none, status_code, error_reason)
    """
    last_reason = "Query error"
    last_code: Optional[int] = None

    for attempt in range(max_retries):
        if stop_event and stop_event.is_set():
            return None, None, "Cancelled"

        try:
            resp = session_getter().get(url, timeout=timeout)
            last_code = resp.status_code

            if resp.status_code == 404:
                return None, 404, "Not found"

            if resp.status_code in (401, 403):
                return None, resp.status_code, "Authentication error"

            if resp.status_code == 429:
                last_reason = "Rate limit exceeded (HTTP 429)"
                retry_after = resp.headers.get("Retry-After")
                wait_sec = base_backoff * (2 ** attempt)
                if retry_after:
                    try:
                        wait_sec = max(wait_sec, float(retry_after))
                    except ValueError:
                        pass
                if attempt < max_retries - 1:
                    time.sleep(min(wait_sec, 10.0))
                    continue
                return None, 429, last_reason

            if resp.status_code >= 500:
                last_reason = f"Server error (HTTP {resp.status_code})"
                if attempt < max_retries - 1:
                    time.sleep(base_backoff * (2 ** attempt))
                    continue
                return None, resp.status_code, last_reason

            resp.raise_for_status()
            try:
                data = resp.json()
            except (ValueError, TypeError):
                return None, resp.status_code, "Unexpected API response"

            if not isinstance(data, dict):
                return None, resp.status_code, "Unexpected API response"

            return data, 200, ""

        except requests.exceptions.Timeout:
            last_reason = "Request timeout"
            if attempt < max_retries - 1:
                time.sleep(base_backoff * (2 ** attempt))
                continue
        except requests.exceptions.HTTPError as e:
            code = getattr(getattr(e, "response", None), "status_code", None)
            last_code = code
            if code in (401, 403):
                return None, code, "Authentication error"
            last_reason = f"HTTP error ({code})"
            if attempt < max_retries - 1:
                time.sleep(base_backoff * (2 ** attempt))
                continue
        except requests.RequestException as e:
            last_reason = f"Network error ({type(e).__name__})"
            if attempt < max_retries - 1:
                time.sleep(base_backoff * (2 ** attempt))
                continue
        except Exception:
            return None, last_code, "Unexpected API response"

    return None, last_code, last_reason


def run_portfolio_coverage_analysis(
    base_url: Optional[str] = None,
    page_size: int = 500,
    max_shells: Optional[int] = None,
    workers: int = 5,
    use_bulk_submodels: bool = True,
    progress_callback: Optional[Callable[[Dict[str, Any]], None]] = None,
    stop_event: Optional[threading.Event] = None,
    session_getter: Optional[Callable[[], requests.Session]] = None,
) -> Dict[str, Any]:
    """
    Fuehrt die WAGO AAS PCF Portfolio-Coverage-Analyse durch:
    1. Paginierter Abruf aller erreichbaren AAS Shells via Cursor-Pagination.
    2. Filterung von Nicht-Produkt-Shells und Deduplizierung auf eindeutige Produkte
       (anhand der normalisierten WAGO-Artikelnummer).
    3. Pruefung der PCF-Submodel-Referenzen und Verifikation der tatsaechlichen
       PCF-Werte und Einheiten (inkl. Bulk-Vorabruf bei grossen Mengen).
    4. Read-Only-Abgleich fehlender AAS-PCFs mit dem eingebetteten Q4-2026-Snapshot.
    5. Berechnung aller Portfolio-Kennzahlen und Aggregation nach Produktserien.
    """
    import aas_pcf_query as core

    api_base = (base_url or core.AAS_BASE_URL).rstrip("/")
    get_sess = session_getter or core._get_session

    start_time = time.time()
    total_shells_retrieved = 0
    non_product_shells: List[Dict[str, Any]] = []
    enumeration_errors: List[Dict[str, Any]] = []
    enumeration_complete = True
    was_cancelled = False

    # canonical_short -> product info dict
    products_map: Dict[str, Dict[str, Any]] = {}

    cursor: Optional[str] = None
    seen_cursors: Set[str] = set()
    page_num = 0

    def _emit_progress(phase: str, processed_products: int = 0, total_to_verify: int = 0):
        if not progress_callback:
            return
        elapsed = time.time() - start_time
        eta_sec: Optional[float] = None
        if phase == "verifying" and processed_products >= 3 and total_to_verify > processed_products:
            rate = elapsed / max(processed_products, 1)
            eta_sec = rate * (total_to_verify - processed_products)
        progress_callback({
            "phase": phase,
            "shells_retrieved": total_shells_retrieved,
            "unique_products": len(products_map),
            "processed_products": processed_products,
            "total_to_verify": total_to_verify,
            "elapsed_seconds": elapsed,
            "eta_seconds": eta_sec,
        })

    # ── Phase 1: Paginierter Abruf der AAS Shells ────────────────────────────
    while True:
        if stop_event and stop_event.is_set():
            was_cancelled = True
            enumeration_complete = False
            break

        effective_limit = page_size
        if max_shells is not None:
            remaining = max_shells - total_shells_retrieved
            if remaining <= 0:
                break
            effective_limit = min(page_size, remaining)

        url = f"{api_base}/shells?limit={effective_limit}"
        if cursor:
            url += f"&cursor={cursor}"

        data, status_code, err_reason = _request_json_with_backoff(
            get_sess, url, stop_event=stop_event
        )

        if data is None:
            if err_reason == "Cancelled":
                was_cancelled = True
            else:
                enumeration_errors.append({
                    "endpoint": url,
                    "status_code": status_code,
                    "diagnostic": err_reason,
                })
            enumeration_complete = False
            break

        shells_page = data.get("result")
        if not isinstance(shells_page, list):
            enumeration_errors.append({
                "endpoint": url,
                "status_code": status_code,
                "diagnostic": "Unexpected API response (missing result list)",
            })
            enumeration_complete = False
            break

        page_num += 1
        for sh in shells_page:
            if not isinstance(sh, dict):
                continue
            total_shells_retrieved += 1

            candidate = _extract_article_candidate_from_shell(sh)
            parsed = core.parse_article_number(candidate)
            aas_id = str(sh.get("id") or "").strip()

            if not parsed["is_valid"]:
                non_product_shells.append({
                    "aas_id": aas_id,
                    "id_short": str(sh.get("idShort") or ""),
                    "reason": "Non-product or unrecognized article identifier",
                })
                if max_shells is not None and total_shells_retrieved >= max_shells:
                    break
                continue

            canonical = parsed["canonical"]
            pcf_refs = _extract_pcf_submodel_refs_from_shell(sh)

            if canonical not in products_map:
                products_map[canonical] = {
                    "canonical": canonical,
                    "article_short": parsed["short"],
                    "article_long": parsed["long"],
                    "series": extract_product_series(parsed["short"]),
                    "aas_ids": [aas_id] if aas_id else [],
                    "pcf_submodel_refs": list(pcf_refs),
                    "shell_count": 1,
                    "aas_pcf": None,
                    "aas_unit": "",
                    "aas_status": None,
                    "internal_status": "",
                }
            else:
                entry = products_map[canonical]
                entry["shell_count"] += 1
                if aas_id and aas_id not in entry["aas_ids"]:
                    entry["aas_ids"].append(aas_id)
                for ref in pcf_refs:
                    if ref not in entry["pcf_submodel_refs"]:
                        entry["pcf_submodel_refs"].append(ref)

            if max_shells is not None and total_shells_retrieved >= max_shells:
                break

        _emit_progress("enumerating")

        if max_shells is not None and total_shells_retrieved >= max_shells:
            break

        paging_meta = data.get("paging_metadata")
        next_cursor = paging_meta.get("cursor") if isinstance(paging_meta, dict) else None
        if not next_cursor or next_cursor in seen_cursors or len(shells_page) == 0:
            break
        seen_cursors.add(next_cursor)
        cursor = str(next_cursor)

    # ── Phase 2: PCF-Submodel-Verifikation ───────────────────────────────────
    # Produkte ohne PCF-Submodel-Referenz haben bestaetigt keinen veroeffentlichten PCF in der AAS.
    products_needing_verification: List[Dict[str, Any]] = []
    for p_entry in products_map.values():
        if not p_entry["pcf_submodel_refs"]:
            p_entry["aas_pcf"] = None
            p_entry["aas_unit"] = ""
            p_entry["aas_status"] = core.STATUS_NOT_YET_AVAILABLE
            p_entry["internal_status"] = core.INTERNAL_STATUS_AAS_NO_PCF
        else:
            products_needing_verification.append(p_entry)

    total_to_verify = len(products_needing_verification)
    verified_count = 0

    # 2a. Optionaler Bulk-Vorabruf der CarbonFootprint-Submodels bei groesseren Analysen
    bulk_submodel_map: Dict[str, Dict[str, Any]] = {}
    if (
        use_bulk_submodels
        and not was_cancelled
        and total_to_verify >= 25
    ):
        sm_cursor: Optional[str] = None
        seen_sm_cursors: Set[str] = set()
        needed_sm_ids: Set[str] = {
            ref for p in products_needing_verification for ref in p["pcf_submodel_refs"]
        }

        while needed_sm_ids - set(bulk_submodel_map.keys()):
            if stop_event and stop_event.is_set():
                was_cancelled = True
                break
            sm_url = f"{api_base}/submodels?idShort=CarbonFootprint&limit={min(page_size, 500)}"
            if sm_cursor:
                sm_url += f"&cursor={sm_cursor}"
            sm_data, _, _ = _request_json_with_backoff(get_sess, sm_url, stop_event=stop_event)
            if not isinstance(sm_data, dict) or not isinstance(sm_data.get("result"), list):
                break
            sm_list = sm_data["result"]
            if not sm_list:
                break
            for sm_obj in sm_list:
                if isinstance(sm_obj, dict) and sm_obj.get("id"):
                    bulk_submodel_map[str(sm_obj["id"]).strip()] = sm_obj
            sm_meta = sm_data.get("paging_metadata")
            next_sm_cur = sm_meta.get("cursor") if isinstance(sm_meta, dict) else None
            if not next_sm_cur or next_sm_cur in seen_sm_cursors:
                break
            seen_sm_cursors.add(next_sm_cur)
            sm_cursor = str(next_sm_cur)

    # 2b. Verifikation jedes Produkts mit PCF-Submodel-Referenz
    remaining_for_individual_fetch: List[Dict[str, Any]] = []
    for p_entry in products_needing_verification:
        resolved = False
        for ref_id in p_entry["pcf_submodel_refs"]:
            if ref_id in bulk_submodel_map:
                sm_obj = bulk_submodel_map[ref_id]
                val, unit, cust_st, int_st = core._evaluate_pcf_and_unit_detailed(sm_obj)
                if cust_st == core.STATUS_PCF_AVAILABLE:
                    p_entry["aas_pcf"] = val
                    p_entry["aas_unit"] = unit
                    p_entry["aas_status"] = core.STATUS_PCF_AVAILABLE
                    p_entry["internal_status"] = int_st
                    resolved = True
                    break
                elif cust_st == core.STATUS_NOT_YET_AVAILABLE:
                    p_entry["aas_pcf"] = None
                    p_entry["aas_unit"] = ""
                    p_entry["aas_status"] = core.STATUS_NOT_YET_AVAILABLE
                    p_entry["internal_status"] = int_st
                    resolved = True
        if resolved:
            verified_count += 1
            _emit_progress("verifying", verified_count, total_to_verify)
        else:
            remaining_for_individual_fetch.append(p_entry)

    def _verify_single_product(p_entry: Dict[str, Any]) -> None:
        nonlocal verified_count, was_cancelled
        if stop_event and stop_event.is_set():
            was_cancelled = True
            p_entry["aas_pcf"] = None
            p_entry["aas_unit"] = ""
            p_entry["aas_status"] = core.STATUS_QUERY_ERROR
            p_entry["internal_status"] = "Cancelled / Not evaluated"
            return

        last_cust_st = core.STATUS_NOT_YET_AVAILABLE
        last_int_st = core.INTERNAL_STATUS_AAS_NO_PCF

        for ref_id in p_entry["pcf_submodel_refs"]:
            encoded = base64.b64encode(ref_id.encode("utf-8")).decode("ascii")
            sm_url = f"{api_base}/submodels/{encoded}"
            sm_dict, code, err_reason = _request_json_with_backoff(
                get_sess, sm_url, stop_event=stop_event
            )
            if err_reason == "Cancelled":
                was_cancelled = True
                p_entry["aas_pcf"] = None
                p_entry["aas_unit"] = ""
                p_entry["aas_status"] = core.STATUS_QUERY_ERROR
                p_entry["internal_status"] = "Cancelled / Not evaluated"
                return
            if code == 404:
                last_cust_st = core.STATUS_NOT_YET_AVAILABLE
                last_int_st = core.INTERNAL_STATUS_AAS_NO_PCF
                continue
            if sm_dict is None:
                last_cust_st = core.STATUS_QUERY_ERROR
                last_int_st = err_reason or core.INTERNAL_STATUS_QUERY_ERROR
                continue

            val, unit, cust_st, int_st = core._evaluate_pcf_and_unit_detailed(sm_dict)
            if cust_st == core.STATUS_PCF_AVAILABLE:
                p_entry["aas_pcf"] = val
                p_entry["aas_unit"] = unit
                p_entry["aas_status"] = core.STATUS_PCF_AVAILABLE
                p_entry["internal_status"] = int_st
                return
            last_cust_st = cust_st
            last_int_st = int_st

        p_entry["aas_pcf"] = None
        p_entry["aas_unit"] = ""
        p_entry["aas_status"] = last_cust_st
        p_entry["internal_status"] = last_int_st

    if remaining_for_individual_fetch:
        lock = threading.Lock()

        def _worker(p_entry: Dict[str, Any]):
            _verify_single_product(p_entry)
            with lock:
                nonlocal verified_count
                verified_count += 1
                vc = verified_count
            _emit_progress("verifying", vc, total_to_verify)

        with ThreadPoolExecutor(max_workers=max(1, workers)) as executor:
            futures = [executor.submit(_worker, p) for p in remaining_for_individual_fetch]
            for f in as_completed(futures):
                f.result()

    # ── Phase 3: Abgleich mit eingebettetem Q4-2026-Snapshot & Kennzahlen ────
    all_products: List[Dict[str, Any]] = []
    missing_pcf_products: List[Dict[str, Any]] = []
    publication_candidates: List[Dict[str, Any]] = []
    calculation_candidates: List[Dict[str, Any]] = []
    query_error_records: List[Dict[str, Any]] = []

    series_stats: Dict[str, Dict[str, int]] = defaultdict(
        lambda: {
            "total": 0,
            "available": 0,
            "missing": 0,
            "pub_candidates": 0,
            "calc_candidates": 0,
            "errors": 0,
        }
    )

    available_count = 0
    missing_count = 0
    pub_cand_count = 0
    calc_cand_count = 0
    tech_error_count = 0

    for canonical in sorted(products_map.keys()):
        p = products_map[canonical]
        offline_res = get_offline_pcf_result(canonical)
        offline_pcf = offline_res.get("pcf")
        offline_unit = offline_res.get("unit") or ""

        aas_st = p["aas_status"]
        series = p["series"]
        series_stats[series]["total"] += 1

        if aas_st == core.STATUS_PCF_AVAILABLE and p["aas_pcf"] is not None:
            classification = CLASS_PCF_AVAILABLE_AAS
            action = ACTION_PUBLISHED
            available_count += 1
            series_stats[series]["available"] += 1
        elif aas_st == core.STATUS_NOT_YET_AVAILABLE:
            missing_count += 1
            series_stats[series]["missing"] += 1
            if offline_pcf is not None:
                classification = CLASS_PCF_OFFLINE_ONLY
                action = ACTION_REVIEW_PUBLICATION
                pub_cand_count += 1
                series_stats[series]["pub_candidates"] += 1
            else:
                classification = CLASS_PCF_MISSING_BOTH
                action = ACTION_REVIEW_CALCULATION
                calc_cand_count += 1
                series_stats[series]["calc_candidates"] += 1
        else:
            classification = CLASS_TECHNICAL_ERROR
            action = ACTION_TECH_INVESTIGATION
            tech_error_count += 1
            series_stats[series]["errors"] += 1

        record = {
            "article_short": p["article_short"],
            "article_long": p["article_long"],
            "series": series,
            "aas_id": ", ".join(p["aas_ids"]),
            "shell_count": p["shell_count"],
            "pcf_aas": p["aas_pcf"],
            "pcf_unit_aas": p["aas_unit"] if p["aas_pcf"] is not None else "",
            "pcf_offline": offline_pcf,
            "pcf_unit_offline": offline_unit if offline_pcf is not None else "",
            "classification": classification,
            "potential_action": action,
            "internal_status": p["internal_status"],
        }
        all_products.append(record)

        if aas_st == core.STATUS_NOT_YET_AVAILABLE:
            missing_pcf_products.append(record)
            if offline_pcf is not None:
                publication_candidates.append(record)
            else:
                calculation_candidates.append(record)
        elif aas_st == core.STATUS_QUERY_ERROR:
            query_error_records.append(record)

    # Fuege etwaige Enumerations-Fehler ebenfalls den Query-Errors hinzu
    for enum_err in enumeration_errors:
        query_error_records.append({
            "article_short": "",
            "article_long": "",
            "series": "",
            "aas_id": enum_err.get("endpoint", ""),
            "shell_count": 0,
            "pcf_aas": None,
            "pcf_unit_aas": "",
            "pcf_offline": None,
            "pcf_unit_offline": "",
            "classification": CLASS_TECHNICAL_ERROR,
            "potential_action": ACTION_TECH_INVESTIGATION,
            "internal_status": enum_err.get("diagnostic", "Enumeration error"),
        })

    total_unique = len(all_products)
    successfully_evaluated = available_count + missing_count
    total_tech_errors = tech_error_count + len(enumeration_errors)

    pcf_coverage_pct = (
        (available_count / successfully_evaluated * 100.0)
        if successfully_evaluated > 0
        else 0.0
    )
    analysis_completeness_pct = (
        (successfully_evaluated / total_unique * 100.0)
        if total_unique > 0
        else (0.0 if not enumeration_complete else 100.0)
    )

    series_breakdown: List[Dict[str, Any]] = []
    for s_name in sorted(series_stats.keys(), key=lambda x: (len(x), x)):
        st = series_stats[s_name]
        s_eval = st["available"] + st["missing"]
        s_cov = (st["available"] / s_eval * 100.0) if s_eval > 0 else 0.0
        s_comp = (s_eval / st["total"] * 100.0) if st["total"] > 0 else 0.0
        series_breakdown.append({
            "series": s_name,
            "total_products": st["total"],
            "available_pcfs": st["available"],
            "missing_pcfs": st["missing"],
            "publication_candidates": st["pub_candidates"],
            "calculation_candidates": st["calc_candidates"],
            "technical_errors": st["errors"],
            "pcf_coverage_pct": round(s_cov, 2),
            "analysis_completeness_pct": round(s_comp, 2),
        })

    is_partial = was_cancelled or (not enumeration_complete) or (max_shells is not None)

    return {
        "analysis_timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "data_source": f"WAGO SmartService AAS API ({api_base})",
        "offline_snapshot_version": OFFLINE_SNAPSHOT_VERSION,
        "is_partial": is_partial,
        "was_cancelled": was_cancelled,
        "enumeration_complete": enumeration_complete and (max_shells is None),
        "max_shells_limit": max_shells,
        "total_shells_retrieved": total_shells_retrieved,
        "non_product_shells_skipped": len(non_product_shells),
        "total_unique_products": total_unique,
        "successfully_evaluated_products": successfully_evaluated,
        "available_pcfs": available_count,
        "missing_pcfs": missing_count,
        "publication_candidates_count": pub_cand_count,
        "calculation_candidates_count": calc_cand_count,
        "technical_errors": total_tech_errors,
        "pcf_coverage_pct": round(pcf_coverage_pct, 2),
        "analysis_completeness_pct": round(analysis_completeness_pct, 2),
        "elapsed_seconds": round(time.time() - start_time, 2),
        "all_products": all_products,
        "missing_pcf_products": missing_pcf_products,
        "publication_candidates": publication_candidates,
        "calculation_candidates": calculation_candidates,
        "query_errors": query_error_records,
        "series_breakdown": series_breakdown,
    }


# ── Portfolio Excel Export (7 Worksheets, WAGO Corporate Design) ─────────────

_CTRL_CHARS_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _clean_xml_text(val: Any) -> str:
    """Escapes text for OpenXML inline strings while stripping invalid XML control chars."""
    if val is None:
        return ""
    s = str(val)
    if not s:
        return ""
    return xml_escape(_CTRL_CHARS_RE.sub("", s))


def _extract_template_style_ids(sheet_xml: str) -> Tuple[str, str, str, str]:
    """
    Extracts (s_odd_txt, s_odd_gen, s_even_txt, s_even_gen) from the 2 template
    rows (<row r="5"> and <row r="6">) generated by openpyxl in the skeleton sheet.
    """
    m5_txt = re.search(r'<c r="A5" s="(\d+)"', sheet_xml)
    m5_gen = re.search(r'<c r="B5" s="(\d+)"', sheet_xml)
    m6_txt = re.search(r'<c r="A6" s="(\d+)"', sheet_xml)
    m6_gen = re.search(r'<c r="B6" s="(\d+)"', sheet_xml)
    return (
        m5_txt.group(1) if m5_txt else "0",
        m5_gen.group(1) if m5_gen else "0",
        m6_txt.group(1) if m6_txt else "0",
        m6_gen.group(1) if m6_gen else "0",
    )


def _configure_frozen_sheet_view(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    first_data_cell: str = "A5",
    tab_selected: bool = False,
) -> None:
    """
    Stellt sicher, dass das Worksheet-View nach dem Setzen von freeze_panes eine
    konsistente aktive Zelle im unteren Datenbereich (bottomLeft) besitzt und
    nicht versehentlich mehrere Worksheets gleichzeitig gruppiert (tabSelected) sind.
    """
    ws.freeze_panes = first_data_cell
    if ws.views and ws.views.sheetView:
        sv = ws.views.sheetView[0]
        sv.tabSelected = tab_selected
        if sv.selection:
            sv.selection[0].activeCell = first_data_cell
            sv.selection[0].sqref = first_data_cell


def _setup_sheet_skeleton(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    headers: List[str],
    col_widths: List[int],
    title: str,
    subtitle: str,
    s_fonts_fills: Dict[str, Any],
    num_rows: int = 0,
) -> None:
    """
    Populates the header rows (1..4), column widths, freeze panes, auto-filter
    (with the exact final row count so xl/workbook.xml _FilterDatabase is accurate),
    and 2 template rows (5..6) in openpyxl so that styles.xml and worksheet XML
    contain the exact WAGO Corporate Design formatting.
    """
    max_col = max(len(headers), 1)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    c_title = ws.cell(1, 1, title)
    c_title.font = s_fonts_fills["font_title"]
    c_title.fill = s_fonts_fills["fill_banner"]
    c_title.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 26

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)
    c_sub = ws.cell(2, 1, subtitle)
    c_sub.font = s_fonts_fills["font_sub"]
    ws.row_dimensions[2].height = 18

    header_row = 4
    ws.row_dimensions[header_row].height = 22
    for c_idx, h_text in enumerate(headers, start=1):
        cell = ws.cell(header_row, c_idx, h_text)
        cell.font = s_fonts_fills["font_hdr"]
        cell.fill = s_fonts_fills["fill_hdr"]
        cell.alignment = Alignment(vertical="center", horizontal="left")
        cell.border = s_fonts_fills["thin_border"]

    # Template Row 5 (odd row: text '@' and general/numeric)
    c_a5 = ws.cell(5, 1, "T")
    c_a5.data_type = "s"
    c_a5.number_format = "@"
    c_a5.font = s_fonts_fills["font_cell"]
    c_a5.border = s_fonts_fills["thin_border"]

    c_b5 = ws.cell(5, 2, 1.0)
    c_b5.font = s_fonts_fills["font_cell"]
    c_b5.border = s_fonts_fills["thin_border"]

    # Template Row 6 (even zebra row: text '@' and general/numeric)
    c_a6 = ws.cell(6, 1, "T")
    c_a6.data_type = "s"
    c_a6.number_format = "@"
    c_a6.font = s_fonts_fills["font_cell"]
    c_a6.fill = s_fonts_fills["fill_zebra"]
    c_a6.border = s_fonts_fills["thin_border"]

    c_b6 = ws.cell(6, 2, 1.0)
    c_b6.font = s_fonts_fills["font_cell"]
    c_b6.fill = s_fonts_fills["fill_zebra"]
    c_b6.border = s_fonts_fills["thin_border"]

    # Set exact final last_row so openpyxl writes the complete _FilterDatabase range into xl/workbook.xml!
    last_row = max(header_row + num_rows, header_row + 1)
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(max_col)}{last_row}"
    _configure_frozen_sheet_view(ws, first_data_cell="A5", tab_selected=False)

    for c_idx, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(c_idx)].width = width


def _stream_product_sheet_xml(
    zout: zipfile.ZipFile,
    arcname: str,
    skeleton_xml: str,
    records: List[Dict[str, Any]],
    styles: Tuple[str, str, str, str],
    on_chunk: Optional[Callable[[int], None]] = None,
    stop_event: Optional[threading.Event] = None,
) -> None:
    """
    Streams product rows directly into the OpenXML worksheet entry inside the .xlsx
    archive in chunks without allocating millions of openpyxl.cell.Cell objects.
    """
    s_odd_txt, s_odd_gen, s_even_txt, s_even_gen = styles
    total_rows = len(records)
    last_row = max(4 + total_rows, 5)

    idx_r5 = skeleton_xml.find('<row r="5"')
    idx_sd_end = skeleton_xml.find("</sheetData>")
    prefix = skeleton_xml[:idx_r5]
    suffix = skeleton_xml[idx_sd_end:]

    prefix = re.sub(r'<dimension ref="[^"]*"', f'<dimension ref="A1:J{last_row}"', prefix, count=1)
    suffix = re.sub(r'<autoFilter ref="[^"]*"', f'<autoFilter ref="A4:J{last_row}"', suffix, count=1)

    with zout.open(arcname, "w") as fp:
        fp.write(prefix.encode("utf-8"))
        if total_rows == 0:
            fp.write(b'<row r="5"/>')
        buf: List[str] = []
        chunk_size = 2500

        for idx, p in enumerate(records):
            if stop_event is not None and (idx & 0x3FF) == 0 and stop_event.is_set():
                raise RuntimeError("Export cancelled by user.")

            r_num = 5 + idx
            if (idx & 1) == 0:
                st_t = s_odd_txt
                st_g = s_odd_gen
            else:
                st_t = s_even_txt
                st_g = s_even_gen

            a_short = _clean_xml_text(p.get("article_short", ""))
            a_long  = _clean_xml_text(p.get("article_long", ""))
            series  = _clean_xml_text(p.get("series", ""))
            aas_id  = _clean_xml_text(p.get("aas_id", ""))

            pcf_aas = p.get("pcf_aas")
            if isinstance(pcf_aas, (int, float)) and not isinstance(pcf_aas, bool):
                c_e = f'<c r="E{r_num}" s="{st_g}" t="n"><v>{pcf_aas}</v></c>'
            else:
                c_e = f'<c r="E{r_num}" s="{st_g}" t="inlineStr"><is><t>Not yet available</t></is></c>'

            u_aas = _clean_xml_text(p.get("pcf_unit_aas", ""))

            pcf_off = p.get("pcf_offline")
            if isinstance(pcf_off, (int, float)) and not isinstance(pcf_off, bool):
                c_g = f'<c r="G{r_num}" s="{st_g}" t="n"><v>{pcf_off}</v></c>'
            else:
                c_g = f'<c r="G{r_num}" s="{st_g}" t="inlineStr"><is><t>Not yet available</t></is></c>'

            u_off  = _clean_xml_text(p.get("pcf_unit_offline", ""))
            classf = _clean_xml_text(p.get("classification", ""))
            action = _clean_xml_text(p.get("potential_action", ""))

            row_xml = (
                f'<row r="{r_num}">'
                f'<c r="A{r_num}" s="{st_t}" t="inlineStr"><is><t>{a_short}</t></is></c>'
                f'<c r="B{r_num}" s="{st_t}" t="inlineStr"><is><t>{a_long}</t></is></c>'
                f'<c r="C{r_num}" s="{st_t}" t="inlineStr"><is><t>{series}</t></is></c>'
                f'<c r="D{r_num}" s="{st_t}" t="inlineStr"><is><t>{aas_id}</t></is></c>'
                f'{c_e}'
                f'<c r="F{r_num}" s="{st_g}" t="inlineStr"><is><t>{u_aas}</t></is></c>'
                f'{c_g}'
                f'<c r="H{r_num}" s="{st_g}" t="inlineStr"><is><t>{u_off}</t></is></c>'
                f'<c r="I{r_num}" s="{st_g}" t="inlineStr"><is><t>{classf}</t></is></c>'
                f'<c r="J{r_num}" s="{st_g}" t="inlineStr"><is><t>{action}</t></is></c>'
                f'</row>'
            )
            buf.append(row_xml)

            if len(buf) >= chunk_size:
                fp.write("".join(buf).encode("utf-8"))
                n_written = len(buf)
                buf.clear()
                if on_chunk is not None:
                    on_chunk(n_written)

        if buf:
            fp.write("".join(buf).encode("utf-8"))
            n_written = len(buf)
            buf.clear()
            if on_chunk is not None:
                on_chunk(n_written)

        fp.write(suffix.encode("utf-8"))


def _style_table_worksheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    headers: List[str],
    rows: List[List[Any]],
    col_widths: List[int],
    text_col_indices: Set[int],
    numeric_col_indices: Set[int],
    title: str,
    subtitle: str,
) -> None:
    """Hilfsfunktion zum Formatieren eines kleineren Tabellen-Worksheets im WAGO Corporate Design."""
    font_title = Font(name="Segoe UI", size=13, bold=True, color=WAGO_WHITE)
    font_sub   = Font(name="Segoe UI", size=9, italic=True, color=WAGO_ANTHRACITE_70)
    font_hdr   = Font(name="Segoe UI", size=10, bold=True, color=WAGO_WHITE)
    font_cell  = Font(name="Segoe UI", size=10, color=WAGO_ANTHRACITE)

    fill_banner = PatternFill(start_color=WAGO_ANTHRACITE, end_color=WAGO_ANTHRACITE, fill_type="solid")
    fill_hdr    = PatternFill(start_color=WAGO_GREEN, end_color=WAGO_GREEN, fill_type="solid")
    fill_zebra  = PatternFill(start_color=WAGO_LIGHT_GREY, end_color=WAGO_LIGHT_GREY, fill_type="solid")
    thin_border = Border(
        left=Side(style="thin", color=WAGO_ANTHRACITE_15),
        right=Side(style="thin", color=WAGO_ANTHRACITE_15),
        top=Side(style="thin", color=WAGO_ANTHRACITE_15),
        bottom=Side(style="thin", color=WAGO_ANTHRACITE_15),
    )

    max_col = max(len(headers), 1)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    c_title = ws.cell(1, 1, title)
    c_title.font = font_title
    c_title.fill = fill_banner
    c_title.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 26

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)
    c_sub = ws.cell(2, 1, subtitle)
    c_sub.font = font_sub
    ws.row_dimensions[2].height = 18

    header_row = 4
    ws.row_dimensions[header_row].height = 22
    for c_idx, h_text in enumerate(headers, start=1):
        cell = ws.cell(header_row, c_idx, h_text)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.alignment = Alignment(vertical="center", horizontal="left")
        cell.border = thin_border

    for r_offset, row_vals in enumerate(rows):
        r_idx = header_row + 1 + r_offset
        is_even = r_offset % 2 == 1
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws.cell(r_idx, c_idx)
            if c_idx in text_col_indices:
                cell.value = "" if val is None else str(val)
                cell.data_type = "s"
                cell.number_format = "@"
            elif c_idx in numeric_col_indices and isinstance(val, (int, float)):
                cell.value = val
            else:
                cell.value = "" if val is None else val
            cell.font = font_cell
            if is_even:
                cell.fill = fill_zebra
            cell.border = thin_border

    if not rows:
        ws.cell(header_row + 1, 1, None)
    last_row = max(header_row + len(rows), header_row + 1)
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(max_col)}{last_row}"
    _configure_frozen_sheet_view(ws, first_data_cell=f"A{header_row + 1}", tab_selected=False)

    for c_idx, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(c_idx)].width = width


def export_portfolio_coverage_excel(
    analysis_result: Dict[str, Any],
    output_file: Union[str, Path],
    progress_callback: Optional[Callable[[Dict[str, Any]], None]] = None,
    stop_event: Optional[threading.Event] = None,
) -> str:
    """
    Erzeugt den internen WAGO PCF Portfolio Coverage Excel-Bericht mit 7 Worksheets:
    1. Portfolio Overview
    2. All Products
    3. Missing PCFs
    4. Publication Candidates
    5. Calculation Candidates
    6. Query Errors
    7. Coverage by Product Series

    Nutzt eine hochperformante Streaming-OpenXML-Generierung fuer grosse Produktlisten
    (>100.000 Produkte in wenigen Sekunden statt Minuten), schreibt zunaechst in eine
    temporaere Datei und verschiebt diese erst nach erfolgreichem Abschluss atomar an
    den Zielpfad.
    """
    t_start = time.perf_counter()

    all_prods  = analysis_result.get("all_products") or []
    miss_prods = analysis_result.get("missing_pcf_products") or []
    pub_cands  = analysis_result.get("publication_candidates") or []
    calc_cands = analysis_result.get("calculation_candidates") or []
    err_recs   = analysis_result.get("query_errors") or []
    ser_recs   = analysis_result.get("series_breakdown") or []

    total_records = (
        len(all_prods)
        + len(miss_prods)
        + len(pub_cands)
        + len(calc_cands)
        + len(err_recs)
        + len(ser_recs)
    )
    processed_records = 0

    def _emit(stage: str, sheet_name: str, force_pct: Optional[float] = None) -> None:
        if progress_callback is None:
            return
        elapsed = time.perf_counter() - t_start
        if force_pct is not None:
            pct = force_pct
        elif total_records > 0:
            pct = min(96.0, 4.0 + (processed_records / total_records) * 90.0)
        else:
            pct = 50.0
        eta: Optional[float] = None
        if 2500 <= processed_records < total_records and elapsed > 0.05:
            rate = processed_records / elapsed
            eta = max(0.0, (total_records - processed_records) / rate)
        progress_callback({
            "stage": stage,
            "worksheet": sheet_name,
            "processed_records": processed_records,
            "total_records": total_records,
            "percent": round(pct, 1),
            "elapsed_seconds": round(elapsed, 2),
            "eta_seconds": round(eta, 1) if eta is not None else None,
        })

    _emit("Preparing portfolio data...", "Portfolio Overview", force_pct=2.0)
    if stop_event is not None and stop_event.is_set():
        raise RuntimeError("Export cancelled by user.")

    wb = openpyxl.Workbook()

    # ── Worksheet 1: Portfolio Overview ──────────────────────────────────────
    _emit("Writing Portfolio Overview...", "Portfolio Overview", force_pct=4.0)
    ws_ov = wb.active
    ws_ov.title = "Portfolio Overview"

    font_title_ov = Font(name="Segoe UI", size=14, bold=True, color=WAGO_WHITE)
    font_title    = Font(name="Segoe UI", size=13, bold=True, color=WAGO_WHITE)
    font_sub      = Font(name="Segoe UI", size=9, italic=True, color=WAGO_ANTHRACITE_70)
    font_sec      = Font(name="Segoe UI", size=11, bold=True, color=WAGO_ANTHRACITE)
    font_lbl      = Font(name="Segoe UI", size=10, bold=True, color=WAGO_ANTHRACITE)
    font_val      = Font(name="Segoe UI", size=10, color=WAGO_ANTHRACITE)
    font_desc     = Font(name="Segoe UI", size=9, color=WAGO_ANTHRACITE_70)
    font_link     = Font(name="Segoe UI", size=10, underline="single", color="0563C1")
    font_warn     = Font(name="Segoe UI", size=10, bold=True, color="C05000")
    font_hdr      = Font(name="Segoe UI", size=10, bold=True, color=WAGO_WHITE)
    font_cell     = Font(name="Segoe UI", size=10, color=WAGO_ANTHRACITE)

    fill_banner = PatternFill(start_color=WAGO_ANTHRACITE, end_color=WAGO_ANTHRACITE, fill_type="solid")
    fill_sec    = PatternFill(start_color=WAGO_ANTHRACITE_15, end_color=WAGO_ANTHRACITE_15, fill_type="solid")
    fill_hdr    = PatternFill(start_color=WAGO_GREEN, end_color=WAGO_GREEN, fill_type="solid")
    fill_zebra  = PatternFill(start_color=WAGO_LIGHT_GREY, end_color=WAGO_LIGHT_GREY, fill_type="solid")
    thin_border = Border(
        left=Side(style="thin", color=WAGO_ANTHRACITE_15),
        right=Side(style="thin", color=WAGO_ANTHRACITE_15),
        top=Side(style="thin", color=WAGO_ANTHRACITE_15),
        bottom=Side(style="thin", color=WAGO_ANTHRACITE_15),
    )

    s_fonts_fills = {
        "font_title": font_title,
        "font_sub": font_sub,
        "font_hdr": font_hdr,
        "font_cell": font_cell,
        "fill_banner": fill_banner,
        "fill_hdr": fill_hdr,
        "fill_zebra": fill_zebra,
        "thin_border": thin_border,
    }

    ws_ov.merge_cells("A1:C1")
    c_t = ws_ov["A1"]
    c_t.value = "WAGO PCF Portfolio Coverage"
    c_t.font = font_title_ov
    c_t.fill = fill_banner
    c_t.alignment = Alignment(vertical="center", indent=1)
    ws_ov.row_dimensions[1].height = 30

    is_partial = bool(analysis_result.get("is_partial"))
    scope_text = (
        "INCOMPLETE / PARTIAL ANALYSIS (Interrupted, cancelled, or bounded sample)"
        if is_partial
        else "Complete accessible WAGO AAS product portfolio"
    )

    overview_rows = [
        ("Analysis Date & Time", analysis_result.get("analysis_timestamp", "")),
        ("Primary Data Source", analysis_result.get("data_source", "")),
        ("Comparison Reference", f"Embedded Offline Dataset ({analysis_result.get('offline_snapshot_version', 'Q4 2026')})"),
        ("Scope of Analysis", scope_text),
        ("WAGO AAS Portal", OFFLINE_PORTAL_URL),
    ]

    r_curr = 3
    for lbl, val in overview_rows:
        c1 = ws_ov.cell(r_curr, 1, lbl)
        c2 = ws_ov.cell(r_curr, 2, val)
        c1.font = font_lbl
        c2.font = font_warn if (lbl == "Scope of Analysis" and is_partial) else font_val
        if lbl == "WAGO AAS Portal":
            c2.hyperlink = OFFLINE_PORTAL_URL
            c2.font = font_link
        c1.border = thin_border
        c2.border = thin_border
        r_curr += 1

    r_curr += 1
    ws_ov.merge_cells(start_row=r_curr, start_column=1, end_row=r_curr, end_column=3)
    c_m_hdr = ws_ov.cell(r_curr, 1, "Portfolio Coverage Metrics")
    c_m_hdr.font = font_sec
    c_m_hdr.fill = fill_sec
    r_curr += 1

    metrics_rows = [
        ("Total AAS Shells Retrieved", analysis_result.get("total_shells_retrieved", 0), "Total Asset Administration Shells inspected"),
        ("Non-Product Shells Excluded", analysis_result.get("non_product_shells_skipped", 0), "Shells without a valid WAGO product article number"),
        ("Total Unique Products Identified", analysis_result.get("total_unique_products", 0), "Deduplicated across shells and short/long formats"),
        ("Successfully Evaluated Products", analysis_result.get("successfully_evaluated_products", 0), "Products with confirmed PCF availability or absence"),
        ("Available PCFs (in AAS)", analysis_result.get("available_pcfs", 0), "Products with verified numeric PcfCO2eq in AAS"),
        ("Confirmed Missing PCFs (in AAS)", analysis_result.get("missing_pcfs", 0), "Products confirmed without a valid PCF in AAS"),
        ("Technical Errors", analysis_result.get("technical_errors", 0), "Products/requests that failed due to technical errors (excluded from missing PCFs)"),
        ("PCF Coverage (%)", f"{analysis_result.get('pcf_coverage_pct', 0.0):.2f}%", "Available PCFs / Successfully evaluated products × 100"),
        ("Analysis Completeness (%)", f"{analysis_result.get('analysis_completeness_pct', 0.0):.2f}%", "Successfully evaluated products / Total identified unique products × 100"),
    ]

    for lbl, val, desc in metrics_rows:
        c1 = ws_ov.cell(r_curr, 1, lbl)
        c2 = ws_ov.cell(r_curr, 2, val)
        c3 = ws_ov.cell(r_curr, 3, desc)
        c1.font = font_lbl
        c2.font = font_val
        c3.font = font_desc
        c1.border = thin_border
        c2.border = thin_border
        c3.border = thin_border
        r_curr += 1

    r_curr += 1
    ws_ov.merge_cells(start_row=r_curr, start_column=1, end_row=r_curr, end_column=3)
    c_c_hdr = ws_ov.cell(r_curr, 1, "Comparison with Embedded Q4 2026 PCF Dataset")
    c_c_hdr.font = font_sec
    c_c_hdr.fill = fill_sec
    r_curr += 1

    comp_rows = [
        ("PCF Available in AAS", analysis_result.get("available_pcfs", 0), ACTION_PUBLISHED),
        ("PCF Missing in AAS / Available in Q4 2026 Snapshot (Publication Candidates)", analysis_result.get("publication_candidates_count", 0), ACTION_REVIEW_PUBLICATION),
        ("PCF Missing in Both Sources (Calculation Candidates)", analysis_result.get("calculation_candidates_count", 0), ACTION_REVIEW_CALCULATION),
        ("Technical Errors", analysis_result.get("technical_errors", 0), ACTION_TECH_INVESTIGATION),
    ]
    for lbl, val, act in comp_rows:
        c1 = ws_ov.cell(r_curr, 1, lbl)
        c2 = ws_ov.cell(r_curr, 2, val)
        c3 = ws_ov.cell(r_curr, 3, act)
        c1.font = font_lbl
        c2.font = font_val
        c3.font = font_val
        c1.border = thin_border
        c2.border = thin_border
        c3.border = thin_border
        r_curr += 1

    r_curr += 1
    ws_ov.merge_cells(start_row=r_curr, start_column=1, end_row=r_curr, end_column=3)
    c_meth_hdr = ws_ov.cell(r_curr, 1, "Methodology & Limitations")
    c_meth_hdr.font = font_sec
    c_meth_hdr.fill = fill_sec
    r_curr += 1

    notes = [
        ("Methodology", "AAS shells are retrieved via cursor-based pagination from the live WAGO SmartService AAS API. Article numbers are normalized to deduplicate products across representations. Products without a CarbonFootprint submodel reference are confirmed as missing a published PCF; products with a CarbonFootprint reference are inspected to verify a valid numeric PcfCO2eq value and unit."),
        ("Q4 2026 Comparison", "Products without a published AAS PCF are compared read-only against the embedded Q4 2026 dataset (9,635 records). Offline PCF values never replace live AAS values."),
        ("Limitations", "1) Product-count-based PCF coverage is not a proxy for revenue- or sales-volume-weighted coverage. 2) Absence of a PCF in both the AAS and the fixed Q4 2026 snapshot does not prove that a PCF has not been calculated internally since Q4 2026. 3) If analysis completeness is below 100% or scope is marked partial, total counts may not represent the complete accessible portfolio."),
    ]
    for lbl, txt in notes:
        c1 = ws_ov.cell(r_curr, 1, lbl)
        ws_ov.merge_cells(start_row=r_curr, start_column=2, end_row=r_curr, end_column=3)
        c2 = ws_ov.cell(r_curr, 2, txt)
        c1.font = font_lbl
        c2.font = font_val
        c2.alignment = Alignment(wrap_text=True, vertical="top")
        c1.border = thin_border
        c2.border = thin_border
        ws_ov.row_dimensions[r_curr].height = 36
        r_curr += 1

    ws_ov.column_dimensions["A"].width = 48
    ws_ov.column_dimensions["B"].width = 36
    ws_ov.column_dimensions["C"].width = 58

    # Common product table columns for Worksheets 2, 3, 4, 5
    prod_headers = [
        "Article Number (Short)",
        "Article Number (Long)",
        "Product Series",
        "AAS Identifier",
        "PCF (AAS)",
        "PCF Unit (AAS)",
        "PCF (Offline Q4 2026)",
        "PCF Unit (Offline Q4 2026)",
        "Internal Availability Classification",
        "Potential Action",
    ]
    prod_widths = [24, 26, 16, 44, 18, 16, 22, 22, 48, 38]

    # ── Worksheets 2..5 Skeletons (populated via fast XML streaming below) ───
    ws_all = wb.create_sheet("All Products")
    _setup_sheet_skeleton(
        ws_all, prod_headers, prod_widths,
        "WAGO PCF Portfolio Coverage – All Identified Products",
        "Complete list of unique products identified in the WAGO AAS portfolio.",
        s_fonts_fills,
        num_rows=len(all_prods),
    )

    ws_miss = wb.create_sheet("Missing PCFs")
    _setup_sheet_skeleton(
        ws_miss, prod_headers, prod_widths,
        "WAGO PCF Portfolio Coverage – Confirmed Missing PCFs in AAS",
        "Worklist of products confirmed without a valid published PCF in the WAGO AAS.",
        s_fonts_fills,
        num_rows=len(miss_prods),
    )

    ws_pub = wb.create_sheet("Publication Candidates")
    _setup_sheet_skeleton(
        ws_pub, prod_headers, prod_widths,
        "WAGO PCF Portfolio Coverage – Publication Candidates",
        "Products without a PCF in the AAS that already have a PCF value in the embedded Q4 2026 dataset.",
        s_fonts_fills,
        num_rows=len(pub_cands),
    )

    ws_calc = wb.create_sheet("Calculation Candidates")
    _setup_sheet_skeleton(
        ws_calc, prod_headers, prod_widths,
        "WAGO PCF Portfolio Coverage – Calculation Candidates",
        "Products without a PCF in either the WAGO AAS or the embedded Q4 2026 dataset.",
        s_fonts_fills,
        num_rows=len(calc_cands),
    )

    # ── Worksheet 6: Query Errors ────────────────────────────────────────────
    ws_err = wb.create_sheet("Query Errors")
    err_headers = [
        "Article Number (Short)",
        "Article Number (Long)",
        "Product Series",
        "AAS Identifier / Endpoint",
        "Internal Diagnostic Reason",
        "Potential Action",
    ]
    err_rows = [
        [
            e.get("article_short", ""),
            e.get("article_long", ""),
            e.get("series", ""),
            e.get("aas_id", ""),
            e.get("internal_status", "Query error"),
            e.get("potential_action", ACTION_TECH_INVESTIGATION),
        ]
        for e in err_recs
    ]
    _style_table_worksheet(
        ws_err,
        err_headers,
        err_rows,
        [24, 26, 16, 48, 36, 32],
        text_col_indices={1, 2, 3, 4, 5, 6},
        numeric_col_indices=set(),
        title="WAGO PCF Portfolio Coverage – Technical Query Errors",
        subtitle="Products or endpoints that could not be evaluated reliably due to technical errors.",
    )
    processed_records += len(err_recs)

    # ── Worksheet 7: Coverage by Product Series ──────────────────────────────
    ws_ser = wb.create_sheet("Coverage by Product Series")
    ser_headers = [
        "Product Series",
        "Total Products",
        "Available PCFs",
        "Missing PCFs",
        "Publication Candidates (Offline Available)",
        "Calculation Candidates",
        "Technical Errors",
        "PCF Coverage (%)",
        "Analysis Completeness (%)",
    ]
    ser_rows = [
        [
            s["series"],
            s["total_products"],
            s["available_pcfs"],
            s["missing_pcfs"],
            s["publication_candidates"],
            s["calculation_candidates"],
            s["technical_errors"],
            s["pcf_coverage_pct"],
            s["analysis_completeness_pct"],
        ]
        for s in ser_recs
    ]
    _style_table_worksheet(
        ws_ser,
        ser_headers,
        ser_rows,
        [18, 16, 16, 16, 36, 24, 18, 20, 24],
        text_col_indices={1},
        numeric_col_indices={2, 3, 4, 5, 6, 7, 8, 9},
        title="WAGO PCF Portfolio Coverage – Aggregated Coverage by Product Series",
        subtitle="Product-count-based PCF coverage grouped by WAGO product series.",
    )
    processed_records += len(ser_recs)

    # Save skeleton workbook to an in-memory buffer
    skeleton_buf = io.BytesIO()
    wb.save(skeleton_buf)
    wb.close()
    skeleton_buf.seek(0)

    out_path = Path(output_file).resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_path_str = tempfile.mkstemp(prefix="wago_portfolio_", suffix=".xlsx")
    os.close(fd)

    stream_targets = {
        "xl/worksheets/sheet2.xml": ("Exporting All Products...", "All Products", all_prods),
        "xl/worksheets/sheet3.xml": ("Exporting Missing PCFs...", "Missing PCFs", miss_prods),
        "xl/worksheets/sheet4.xml": ("Exporting Publication Candidates...", "Publication Candidates", pub_cands),
        "xl/worksheets/sheet5.xml": ("Exporting Calculation Candidates...", "Calculation Candidates", calc_cands),
    }

    try:
        with zipfile.ZipFile(skeleton_buf, "r") as zin, zipfile.ZipFile(
            tmp_path_str, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=1
        ) as zout:
            sheet2_skel = zin.read("xl/worksheets/sheet2.xml").decode("utf-8")
            styles = _extract_template_style_ids(sheet2_skel)

            for item in zin.infolist():
                arc = item.filename
                if arc in stream_targets:
                    stage_msg, sh_name, recs = stream_targets[arc]
                    _emit(stage_msg, sh_name)
                    skel_xml = zin.read(arc).decode("utf-8")

                    def _on_chunk(n: int, st_msg: str = stage_msg, s_nm: str = sh_name) -> None:
                        nonlocal processed_records
                        processed_records += n
                        _emit(st_msg, s_nm)

                    _stream_product_sheet_xml(
                        zout,
                        arc,
                        skel_xml,
                        recs,
                        styles,
                        on_chunk=_on_chunk,
                        stop_event=stop_event,
                    )
                else:
                    zout.writestr(item, zin.read(arc))

        if stop_event is not None and stop_event.is_set():
            raise RuntimeError("Export cancelled by user.")

        _emit("Finalizing Excel workbook...", "Saving File", force_pct=98.0)
        shutil.move(tmp_path_str, str(out_path))
    except Exception:
        if os.path.exists(tmp_path_str):
            try:
                os.remove(tmp_path_str)
            except OSError:
                pass
        raise

    return str(out_path)

