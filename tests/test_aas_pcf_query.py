"""
Automated test suite for WAGO AAS PCF Query Tool (aas_pcf_query.py,
offline_pcf_data.py, and portfolio_coverage.py).

Covers:
1. Article number normalization & bidirectional conversion
2. Excel import (.xlsx and legacy .xls) & structure-agnostic detection
3. Standardized Excel export ('PCF Results' first sheet, PCF Unit, original workbook preservation)
4. API behavior & PCF unit / reference quantity handling
5. Single Query workflow & GUI (Drag-and-drop, 3 workflow tabs, clipboard, Q4 2026 notice)
6. PCF Portfolio Coverage (pagination, deduplication, non-product filtering, Q4 2026 comparison,
   rate-limit retry, partial/cancelled analysis, 7-worksheet Excel export)
7. Compiled executable verification without external reference Excel files
"""

from __future__ import annotations

import struct
import threading
from pathlib import Path
from typing import Any, List
from unittest.mock import MagicMock, patch

import openpyxl
import pandas as pd
import pytest
import requests

import aas_pcf_query as pcf_tool
from aas_pcf_query import (
    COL_ARTICLE_LONG,
    COL_ARTICLE_SHORT,
    COL_PCF,
    COL_PCF_UNIT,
    DEFAULT_PCF_UNIT,
    INTERNAL_STATUS_AAS_NO_PCF,
    INTERNAL_STATUS_AUTH_ERROR,
    INTERNAL_STATUS_INVALID_RESP,
    INTERNAL_STATUS_NOT_FOUND,
    INTERNAL_STATUS_PCF_AVAILABLE,
    INTERNAL_STATUS_SERVER_ERROR,
    INTERNAL_STATUS_TIMEOUT,
    MODE_OFFLINE,
    MODE_ONLINE,
    PCF_RESULTS_SHEET_NAME,
    STATUS_NOT_YET_AVAILABLE,
    STATUS_PCF_AVAILABLE,
    STATUS_QUERY_ERROR,
    WORKFLOW_BATCH,
    WORKFLOW_PORTFOLIO,
    WORKFLOW_SINGLE,
    convert_article_number,
    detect_article_column,
    detect_excel_file_format,
    export_batch_pcf_excel,
    export_portfolio_coverage_excel,
    format_single_result_for_clipboard,
    get_pcf,
    get_pcf_result,
    is_valid_wago_article,
    load_excel_flexible,
    main,
    normalize_article_number,
    parse_article_number,
    prepare_output_dataframe,
    query_single_article,
    read_pcf_results_table,
    run_portfolio_coverage_analysis,
    to_long_article_number,
    to_short_article_number,
)
from offline_pcf_data import (
    EMBEDDED_PCF_DATA,
    OFFLINE_EXCEL_HEADER_NOTE,
    OFFLINE_NOTICE_TEXT,
    OFFLINE_NOTICE_TITLE,
    OFFLINE_PORTAL_URL,
    get_offline_pcf,
    get_offline_pcf_result,
    validate_and_index_pcf_records,
)


def _write_test_xls_file(path: Path, rows: List[List[Any]]) -> None:
    """
    Writes a valid binary BIFF .xls worksheet file that xlrd parses natively.
    """
    def _rec(opcode: int, data: bytes) -> bytes:
        return struct.pack("<HH", opcode, len(data)) + data

    buf = bytearray()
    # BOF BIFF2 (0x0009): version=2, type=0x0010 (worksheet)
    buf += _rec(0x0009, struct.pack("<HH", 2, 0x0010))
    # CODEPAGE (0x0042): Windows-1252 (1252)
    buf += _rec(0x0042, struct.pack("<H", 1252))

    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row):
            if val is None or val == "":
                continue
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                # NUMBER (0x0003)
                buf += _rec(0x0003, struct.pack("<HH3sd", r_idx, c_idx, b"\x00\x00\x00", float(val)))
            else:
                s_bytes = str(val).encode("cp1252", errors="replace")[:255]
                # LABEL (0x0004)
                buf += _rec(0x0004, struct.pack("<HH3sB", r_idx, c_idx, b"\x00\x00\x00", len(s_bytes)) + s_bytes)

    # EOF (0x000A)
    buf += _rec(0x000A, b"")
    path.write_bytes(bytes(buf))


def _make_submodel_response(
    pcf_value: Any = "0.0116",
    article: str = "221-412",
    ref_unit: str = "piece",
    ref_qty: str = "1",
):
    return {
        "id": f"https://www.wago.com/submodel/{article}/CarbonFootprint",
        "idShort": "CarbonFootprint",
        "modelType": "Submodel",
        "submodelElements": [
            {
                "idShort": "ProductCarbonFootprints",
                "modelType": "SubmodelElementList",
                "value": [
                    {
                        "modelType": "SubmodelElementCollection",
                        "value": [
                            {
                                "idShort": "PcfCO2eq",
                                "modelType": "Property",
                                "value": str(pcf_value) if pcf_value is not None else "",
                            },
                            {
                                "idShort": "ReferenceImpactUnitForCalculation",
                                "modelType": "Property",
                                "value": ref_unit,
                            },
                            {
                                "idShort": "QuantityOfMeasureForCalculation",
                                "modelType": "Property",
                                "value": ref_qty,
                            },
                        ],
                    }
                ],
            }
        ],
    }


# ═══════════════════════════════════════════════════════════════════════════
# 1. Article Number Normalization & Bidirectional Conversion Tests
# ═══════════════════════════════════════════════════════════════════════════

class TestArticleNumberNormalization:

    def test_standard_short_long_conversion_221_412(self):
        assert to_long_article_number("221-412") == "0221-0412"
        assert to_short_article_number("0221-0412") == "221-412"
        assert to_short_article_number("221-412") == "221-412"
        assert to_long_article_number("0221-0412") == "0221-0412"
        assert convert_article_number("221-412") == ("221-412", "0221-0412")
        assert convert_article_number("0221-0412") == ("221-412", "0221-0412")

    def test_four_digit_item_conversion_221_2411(self):
        assert to_long_article_number("221-2411") == "0221-2411"
        assert to_short_article_number("0221-2411") == "221-2411"
        assert convert_article_number("221-2411") == ("221-2411", "0221-2411")
        assert convert_article_number("0221-2411") == ("221-2411", "0221-2411")

    @pytest.mark.parametrize(
        "short_fmt, long_fmt",
        [
            ("221-412", "0221-0412"),
            ("221-2411", "0221-2411"),
            ("2000-115", "2000-0115"),
            ("2000-1201", "2000-1201"),
            ("2002-6307", "2002-6307"),
            ("280-101", "0280-0101"),
            ("285-1175", "0285-1175"),
            ("750-103", "0750-0103"),
            ("231-004/046-000", "0231-0004/0046-0000"),
            ("734-001/K010-2937/000-100", "0734-0001/K010-2937/0000-0100"),
            ("8001-001/K010-2937/000-2800", "8001-0001/K010-2937/0000-2800"),
            ("8001-099/Z000-1809", "8001-0099/Z000-1809"),
            ("2091-1122/000-604/8000-050", "2091-1122/0000-0604/8000-0050"),
            ("280-561/281-413_HI", "0280-0561/0281-0413_HI"),
            ("770-103-JPE", "0770-0103-JPE"),
            ("770-213/KN", "0770-0213/KN"),
            ("835-0FAR/215-411", "0835-0FAR/0215-0411"),
        ],
    )
    def test_additional_wago_formats_bidirectional(self, short_fmt: str, long_fmt: str):
        assert is_valid_wago_article(short_fmt)
        assert is_valid_wago_article(long_fmt)
        assert to_long_article_number(short_fmt) == long_fmt
        assert to_short_article_number(long_fmt) == short_fmt

    def test_duplicate_products_in_different_representations_normalize_identically(self):
        variants_221_412 = ["221-412", "0221-0412", "221-0412", "0221-412", "2210412", "02210412", " 0221-0412 "]
        assert {normalize_article_number(v) for v in variants_221_412} == {"221-412"}

    def test_leading_zeros_lost_in_numeric_or_mixed_values(self):
        p1 = parse_article_number(2210412)
        assert p1["is_valid"] is True
        assert p1["short"] == "221-412"
        assert p1["long"] == "0221-0412"

        p2 = parse_article_number("2212411.0")
        assert p2["is_valid"] is True
        assert p2["short"] == "221-2411"
        assert p2["long"] == "0221-2411"

    @pytest.mark.parametrize(
        "invalid_val",
        [
            "INVALID",
            "12-34",
            "1-2",
            "221-41",
            "221-4",
            "12345",
            "123456",
            "123456789",
            "51170648",
            "2026-11-16",
            "99999-111",
            "ABC-DEF",
            "221-412/",
            "/221-412",
        ],
    )
    def test_invalid_or_ambiguous_article_numbers_preserve_original(self, invalid_val: str):
        parsed = parse_article_number(invalid_val)
        assert parsed["is_valid"] is False
        assert parsed["short"] == invalid_val
        assert parsed["long"] == invalid_val


# ═══════════════════════════════════════════════════════════════════════════
# 2. Structure-Agnostic Excel Detection & .xls / .xlsx Import Tests
# ═══════════════════════════════════════════════════════════════════════════

class TestExcelImportAndDetection:

    def test_xlsx_and_legacy_xls_import_and_detection(self, tmp_path: Path):
        # 1. Legacy .xls file with arbitrary column position & numeric leading-zero loss
        xls_path = tmp_path / "legacy_input.xls"
        _write_test_xls_file(
            xls_path,
            [
                ["Pos", "Description", "Materialnummer", "Qty"],
                [1, "Connector", "221-412", 10],
                [2, "Inline Connector", "0221-2411", 20],
                [3, "Numeric 0280-0101", 2800101, 5],
            ],
        )

        df_xls, info_xls = load_excel_flexible(xls_path)
        assert info_xls["file_format"] == "xls"
        assert info_xls["column"] == "Materialnummer"
        canonical, short_list, long_list = pcf_tool._extract_row_articles(df_xls, info_xls)
        assert short_list == ["221-412", "221-2411", "280-101"]
        assert long_list == ["0221-0412", "0221-2411", "0280-0101"]

        # 2. Standard .xlsx file
        xlsx_path = tmp_path / "standard_input.xlsx"
        pd.DataFrame({
            "Article Number": ["221-412", "280-101"],
            "Description": ["A", "B"],
        }).to_excel(xlsx_path, index=False)

        df_xlsx, info_xlsx = load_excel_flexible(xlsx_path)
        assert info_xlsx["file_format"] == "xlsx"
        assert info_xlsx["column"] == "Article Number"
        assert len(df_xlsx) == 2

    def test_missing_or_unusual_headers_and_introductory_rows(self, tmp_path: Path):
        # Headerless Excel file
        headerless_path = tmp_path / "headerless.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Connector 2-conductor", "221-412", 50])
        ws.append(["Connector 3-conductor", "221-413", 100])
        wb.save(headerless_path)

        df_hl, info_hl = load_excel_flexible(headerless_path)
        assert info_hl["has_header"] is False
        assert info_hl["column"] == "Column 2"
        assert len(df_hl) == 2

        # Introductory rows before actual table
        intro_path = tmp_path / "intro_rows.xlsx"
        wb2 = openpyxl.Workbook()
        ws2 = wb2.active
        ws2.append(["Customer Sustainability Inquiry"])
        ws2.append(["Date: 2026-11-18"])
        ws2.append([])
        ws2.append(["Pos", "Materialnummer", "Bezeichnung"])
        ws2.append(["1", "0221-0412", "Verbindungsklemme"])
        ws2.append(["2", "0221-2411", "Durchgangsverbinder"])
        wb2.save(intro_path)

        df_intro, info_intro = load_excel_flexible(intro_path)
        assert info_intro["column"] == "Materialnummer"
        assert len(df_intro) == 2

    def test_multiple_candidate_columns_and_multi_sheet(self, tmp_path: Path):
        df_same = pd.DataFrame({
            "Material lang ber.": ["0221-0412", "0221-2411"],
            "Material lang kompr.": ["221-412", "221-2411"],
        })
        info_same = detect_article_column(df_same)
        assert info_same["has_both_formats"] is True
        assert info_same["is_ambiguous"] is False

        df_competing = pd.DataFrame({
            "Col_1": ["221-412", "280-101"],
            "Col_2": ["750-103", "787-1212"],
        })
        info_comp = detect_article_column(df_competing)
        assert info_comp["is_ambiguous"] is True


# ═══════════════════════════════════════════════════════════════════════════
# 3. API Behavior & PCF Unit / Reference Quantity Handling Tests
# ═══════════════════════════════════════════════════════════════════════════

class TestAPIBehaviorAndUnits:

    def test_successful_pcf_and_unit_retrieval_online(self):
        sample_sm = _make_submodel_response("0.0116", "221-412", ref_unit="piece", ref_qty="1")
        with patch.object(pcf_tool, "_jget", return_value=sample_sm):
            res_short = get_pcf_result("221-412")
            res_long  = get_pcf_result("0221-0412")

            assert res_short["pcf"] == pytest.approx(0.0116)
            assert res_short["unit"] == "kg CO2e"
            assert res_short["status"] == STATUS_PCF_AVAILABLE
            assert res_long["pcf"] == pytest.approx(0.0116)
            assert res_long["unit"] == "kg CO2e"
            assert get_pcf("0221-0412") == pytest.approx(0.0116)

    def test_non_piece_reference_unit_and_quantity_preserved_without_silent_conversion(self):
        # e.g. kg CO2e per kg of product
        sm_kg = _make_submodel_response("2.5", "285-135", ref_unit="kg", ref_qty="1")
        with patch.object(pcf_tool, "_jget", return_value=sm_kg):
            res_kg = get_pcf_result("285-135")
            assert res_kg["pcf"] == pytest.approx(2.5)
            assert res_kg["unit"] == "kg CO2e / kg"

        # e.g. kg CO2e per 100 piece
        sm_100 = _make_submodel_response("1.16", "221-412", ref_unit="piece", ref_qty="100")
        with patch.object(pcf_tool, "_jget", return_value=sm_100):
            res_100 = get_pcf_result("221-412")
            assert res_100["pcf"] == pytest.approx(1.16)
            assert res_100["unit"] == "kg CO2e / 100 piece"

    def test_product_without_available_pcf_has_empty_unit_and_not_yet_available_status(self):
        with patch.object(pcf_tool, "_jget", side_effect=[pcf_tool._NOT_FOUND, {"result": []}]):
            res = get_pcf_result("999-999")
            assert res["pcf"] is None
            assert res["unit"] == ""
            assert res["status"] == STATUS_NOT_YET_AVAILABLE
            assert res["internal_status"] == INTERNAL_STATUS_NOT_FOUND

    @pytest.mark.parametrize("status_code", [401, 403])
    def test_authentication_failure_is_query_error(self, status_code: int):
        mock_resp = MagicMock()
        mock_resp.status_code = status_code
        mock_session = MagicMock()
        mock_session.get.return_value = mock_resp

        with patch.object(pcf_tool, "_get_session", return_value=mock_session):
            res = get_pcf_result("221-412")
            assert res["pcf"] is None
            assert res["unit"] == ""
            assert res["status"] == STATUS_QUERY_ERROR
            assert res["internal_status"] == INTERNAL_STATUS_AUTH_ERROR

    def test_api_timeout_and_unexpected_response_are_query_errors(self):
        mock_session = MagicMock()
        mock_session.get.side_effect = requests.exceptions.ReadTimeout("Timeout")
        with patch.object(pcf_tool, "_get_session", return_value=mock_session):
            res = get_pcf_result("221-412")
            assert res["status"] == STATUS_QUERY_ERROR
            assert res["internal_status"] == INTERNAL_STATUS_TIMEOUT


# ═══════════════════════════════════════════════════════════════════════════
# 4. Redesigned Excel Export ('PCF Results' + Original Sheets Preservation)
# ═══════════════════════════════════════════════════════════════════════════

class TestRedesignedExcelExport:

    def test_pcf_results_sheet_first_and_original_sheets_formulas_preserved(self, tmp_path: Path):
        input_path  = tmp_path / "customer_workbook.xlsx"
        output_path = tmp_path / "customer_workbook_PCF.xlsx"

        # Create multi-sheet workbook with a formula and custom columns
        wb_in = openpyxl.Workbook()
        ws1 = wb_in.active
        ws1.title = "BOM_Sheet"
        ws1.append(["Pos", "Part_No", "Qty", "Line_Total"])
        ws1.append([1, "221-412", 10, "=C2*2"])
        ws1.append([2, "0221-2411", 5, "=C3*2"])
        ws1.append([3, "999-999", 1, "=C4*2"])

        ws2 = wb_in.create_sheet("Notes_Sheet")
        ws2.append(["Project", "Customer X"])
        ws2.append(["Total Qty", "=SUM(BOM_Sheet!C2:C4)"])
        wb_in.save(input_path)

        orig_bytes = input_path.read_bytes()

        # Run batch query in Offline Mode
        summary = main(str(input_path), str(output_path), mode=MODE_OFFLINE)
        assert summary is not None
        assert output_path.exists()

        # 1. Original input file on disk is 100% untouched
        assert input_path.read_bytes() == orig_bytes

        # 2. Inspect exported workbook structure & formulas
        wb_out = openpyxl.load_workbook(output_path, data_only=False)
        assert wb_out.sheetnames == ["PCF Results", "BOM_Sheet", "Notes_Sheet"]

        # Original sheets are completely unchanged (no columns added, formulas intact)
        ws_bom = wb_out["BOM_Sheet"]
        assert [ws_bom.cell(1, c).value for c in range(1, ws_bom.max_column + 1)] == [
            "Pos", "Part_No", "Qty", "Line_Total"
        ]
        assert ws_bom["D2"].value == "=C2*2"
        assert ws_bom["B2"].value == "221-412"
        assert ws_bom["B3"].value == "0221-2411"

        ws_notes = wb_out["Notes_Sheet"]
        assert ws_notes["B2"].value == "=SUM(BOM_Sheet!C2:C4)"

        # 3. Inspect 'PCF Results' Information Header & Table
        ws_pcf = wb_out["PCF Results"]
        assert ws_pcf["A1"].value == "WAGO Product Carbon Footprint Data"
        assert "Data snapshot: Q4 2026" in str(ws_pcf["B4"].value)
        assert ws_pcf["B8"].hyperlink is not None
        assert ws_pcf["B8"].hyperlink.target == "https://aas.wago.com/"
        assert ws_pcf["A9"].value == OFFLINE_EXCEL_HEADER_NOTE
        assert ws_pcf.freeze_panes == "A12"
        assert ws_pcf.auto_filter.ref == "A11:D14"

        # Table headers on Row 11: exact 4 columns, NO PCF Status column
        table_headers = [ws_pcf.cell(11, c).value for c in range(1, ws_pcf.max_column + 1)]
        assert table_headers == [COL_ARTICLE_SHORT, COL_ARTICLE_LONG, COL_PCF, COL_PCF_UNIT]
        assert "PCF Status" not in table_headers

        # Row 12 (221-412): text article numbers, numeric PCF cell, 'kg CO2e' unit
        assert ws_pcf.cell(12, 1).value == "221-412"
        assert ws_pcf.cell(12, 1).data_type == "s"
        assert ws_pcf.cell(12, 2).value == "0221-0412"
        assert ws_pcf.cell(12, 2).data_type == "s"
        assert isinstance(ws_pcf.cell(12, 3).value, float)
        assert ws_pcf.cell(12, 3).value == pytest.approx(0.0116)
        assert ws_pcf.cell(12, 4).value == "kg CO2e"

        # Row 14 (999-999): missing PCF -> 'Not yet available' in PCF cell, empty unit
        assert ws_pcf.cell(14, 1).value == "999-999"
        assert ws_pcf.cell(14, 2).value == "0999-0999"
        assert ws_pcf.cell(14, 3).value == "Not yet available"
        assert ws_pcf.cell(14, 4).value in ("", None)
        wb_out.close()

    def test_legacy_xls_to_xlsx_batch_export_preserves_original_sheet(self, tmp_path: Path):
        xls_in   = tmp_path / "legacy_order.xls"
        xlsx_out = tmp_path / "legacy_order_PCF.xlsx"

        _write_test_xls_file(
            xls_in,
            [
                ["Item", "WAGO_No", "Note"],
                [1, "280-101", "Terminal"],
                [2, "855-9910", "Transformer"],
            ],
        )

        summary = main(str(xls_in), str(xlsx_out), mode=MODE_OFFLINE)
        assert summary is not None
        assert xlsx_out.exists()

        wb = openpyxl.load_workbook(xlsx_out)
        assert wb.sheetnames[0] == "PCF Results"
        assert len(wb.sheetnames) == 2  # 'PCF Results' + preserved original sheet from .xls
        wb.close()

        df_res = read_pcf_results_table(xlsx_out)
        assert list(df_res.columns) == [COL_ARTICLE_SHORT, COL_ARTICLE_LONG, COL_PCF, COL_PCF_UNIT]
        assert df_res[COL_ARTICLE_SHORT].tolist() == ["280-101", "855-9910"]
        assert df_res[COL_ARTICLE_LONG].tolist() == ["0280-0101", "0855-9910"]
        assert float(df_res.loc[0, COL_PCF]) == pytest.approx(0.0411)
        assert df_res.loc[0, COL_PCF_UNIT] == "kg CO2e"
        assert df_res.loc[1, COL_PCF] == "Not yet available"
        assert df_res.loc[1, COL_PCF_UNIT] == ""


# ═══════════════════════════════════════════════════════════════════════════
# 5. Single Query, Drag-and-Drop & GUI Workflow Tests
# ═══════════════════════════════════════════════════════════════════════════

class TestSingleQueryAndGUI:

    def test_single_query_online_and_offline_and_clipboard_format(self):
        # Offline mode single query
        res_off = query_single_article("0221-0412", mode=MODE_OFFLINE)
        assert res_off["article_short"] == "221-412"
        assert res_off["article_long"] == "0221-0412"
        assert res_off["pcf"] == pytest.approx(0.0116)
        assert res_off["unit"] == "kg CO2e"
        assert "Q4 2026" in res_off["data_source"]

        assert format_single_result_for_clipboard(res_off, value_only=True) == "0.0116"
        full_clip = format_single_result_for_clipboard(res_off, value_only=False)
        assert "Article Number (Short): 221-412" in full_clip
        assert "Article Number (Long):  0221-0412" in full_clip
        assert "PCF Unit:               kg CO2e" in full_clip

        # Missing article in single query
        res_miss = query_single_article("999-999", mode=MODE_OFFLINE)
        assert res_miss["pcf"] is None
        assert res_miss["pcf_display"] == "Not yet available"
        assert res_miss["unit"] == ""

    def test_gui_workflows_drag_and_drop_single_query_and_offline_notice(self, tmp_path: Path):
        xlsx_file = tmp_path / "drop_test.xlsx"
        pd.DataFrame({"Article Number": ["221-412", "280-101"]}).to_excel(xlsx_file, index=False)
        invalid_file = tmp_path / "unsupported.txt"
        invalid_file.write_text("not an excel file", encoding="utf-8")

        app = pcf_tool.PCFQueryApp()
        try:
            # 1. Default mode is Online, notice hidden
            assert app.get_mode() == MODE_ONLINE
            assert app._offline_notice_visible is False

            # 2. Switch to Offline Mode -> German Q4 2026 notice appears
            app.set_mode(MODE_OFFLINE)
            assert app._offline_notice_visible is True
            assert app._lbl_offline_title.cget("text") == OFFLINE_NOTICE_TITLE
            assert app._lbl_offline_text.cget("text") == OFFLINE_NOTICE_TEXT
            assert app._lbl_offline_link.cget("text") == OFFLINE_PORTAL_URL

            # 3. Switch to Single Query tab & execute query via Enter-key handler
            app.select_workflow(WORKFLOW_SINGLE)
            assert app.get_active_workflow() == WORKFLOW_SINGLE
            app._single_article_var.set("0221-0412")
            app._on_single_query()
            assert app._single_fields["article_short"].cget("text") == "221-412"
            assert app._single_fields["article_long"].cget("text") == "0221-0412"
            assert app._single_fields["pcf_display"].cget("text") == "0.0116"
            assert app._single_fields["unit"].cget("text") == "kg CO2e"

            # Copy to clipboard
            copied_val = app._copy_single_result(value_only=True)
            assert copied_val == "0.0116"

            # 4. Drag-and-drop validation: invalid file rejected without crash, valid file loaded
            assert app.handle_dropped_file(str(invalid_file)) is False
            assert app.handle_dropped_file(str(xlsx_file)) is True
            assert app.get_active_workflow() == WORKFLOW_BATCH
            assert app._input_path == str(xlsx_file)

            # Clear / remove file
            app.clear_selected_file()
            assert app._input_path is None

            # 5. Switch to Portfolio Coverage -> Offline notice hidden, live API required
            app.select_workflow(WORKFLOW_PORTFOLIO)
            assert app.get_active_workflow() == WORKFLOW_PORTFOLIO
            assert app._offline_notice_visible is False
        finally:
            app.destroy()


# ═══════════════════════════════════════════════════════════════════════════
# 6. PCF Portfolio Coverage & 7-Worksheet Excel Report Tests
# ═══════════════════════════════════════════════════════════════════════════

class TestPortfolioCoverage:

    def test_portfolio_enumeration_pagination_deduplication_and_classifications(self, tmp_path: Path):
        # Page 1 of /shells:
        # - 221-412 (has CarbonFootprint ref -> valid PCF in AAS)
        # - 0221-0412 (duplicate of 221-412 in long format -> deduplicated!)
        # - Non-product shell ('SystemTemplateShell') -> excluded!
        page1 = {
            "paging_metadata": {"cursor": "cur_page_2"},
            "result": [
                {
                    "id": "https://wago.com/ids/aas/221-412",
                    "idShort": "221-412",
                    "submodels": [
                        {"keys": [{"type": "Submodel", "value": "https://www.wago.com/submodel/221-412/CarbonFootprint"}]}
                    ],
                },
                {
                    "id": "https://wago.com/ids/aas/0221-0412",
                    "idShort": "0221-0412",
                    "submodels": [],
                },
                {
                    "id": "https://wago.com/ids/aas/SystemTemplate",
                    "idShort": "SystemTemplateShell",
                    "submodels": [],
                },
            ],
        }
        # Page 2 of /shells:
        # - 734-001/K010-2937/000-100 (no CarbonFootprint ref in AAS, but exists in Q4 2026 offline -> Publication Candidate!)
        # - 855-9910 (no CarbonFootprint ref in AAS, and not in Q4 2026 offline -> Calculation Candidate!)
        # - 280-101 (has CarbonFootprint ref, but submodel fetch fails with 500 -> Technical Error!)
        page2 = {
            "paging_metadata": {},
            "result": [
                {
                    "id": "https://wago.com/ids/aas/734-001/K010-2937/000-100",
                    "idShort": "734-001/K010-2937/000-100",
                    "submodels": [
                        {"keys": [{"type": "Submodel", "value": "https://www.wago.com/submodel/734-001/K010-2937/000-100/Nameplate"}]}
                    ],
                },
                {
                    "id": "https://wago.com/ids/aas/855-9910",
                    "idShort": "855-9910",
                    "submodels": [],
                },
                {
                    "id": "https://wago.com/ids/aas/280-101",
                    "idShort": "280-101",
                    "submodels": [
                        {"keys": [{"type": "Submodel", "value": "https://www.wago.com/submodel/280-101/CarbonFootprint"}]}
                    ],
                },
            ],
        }

        sm_221_412 = _make_submodel_response("0.0116", "221-412")

        mock_session = MagicMock()

        def _fake_get(url: str, timeout=None):
            resp = MagicMock()
            if "/shells?" in url and "cursor=cur_page_2" not in url:
                resp.status_code = 200
                resp.json.return_value = page1
                return resp
            if "/shells?" in url and "cursor=cur_page_2" in url:
                resp.status_code = 200
                resp.json.return_value = page2
                return resp
            # Submodel fetch for 221-412 vs 280-101
            b64_221 = pcf_tool._b64_encode("https://www.wago.com/submodel/221-412/CarbonFootprint")
            if b64_221 in url:
                resp.status_code = 200
                resp.json.return_value = sm_221_412
                return resp
            # 280-101 submodel fails with HTTP 500
            resp.status_code = 500
            return resp

        mock_session.get.side_effect = _fake_get

        result = run_portfolio_coverage_analysis(
            use_bulk_submodels=False,
            session_getter=lambda: mock_session,
        )

        # Verify portfolio counts & formulas
        assert result["total_shells_retrieved"] == 6
        assert result["non_product_shells_skipped"] == 1
        assert result["total_unique_products"] == 4  # 221-412, 734-001/..., 855-9910, 280-101
        assert result["available_pcfs"] == 1         # 221-412
        assert result["missing_pcfs"] == 2           # 734-001/... and 855-9910
        assert result["publication_candidates_count"] == 1  # 734-001/K010-2937/000-100
        assert result["calculation_candidates_count"] == 1  # 855-9910
        assert result["technical_errors"] == 1       # 280-101
        assert result["successfully_evaluated_products"] == 3  # 1 available + 2 missing

        # PCF Coverage = 1 / 3 * 100 = 33.33%
        assert result["pcf_coverage_pct"] == pytest.approx(33.33, abs=0.01)
        # Analysis Completeness = 3 / 4 * 100 = 75.00%
        assert result["analysis_completeness_pct"] == pytest.approx(75.00, abs=0.01)

        # Export 7-worksheet Excel report and verify all sheets exist
        report_path = tmp_path / "portfolio_report.xlsx"
        export_portfolio_coverage_excel(result, report_path)
        assert report_path.exists()

        wb = openpyxl.load_workbook(report_path)
        assert wb.sheetnames == [
            "Portfolio Overview",
            "All Products",
            "Missing PCFs",
            "Publication Candidates",
            "Calculation Candidates",
            "Query Errors",
            "Coverage by Product Series",
        ]
        wb.close()

    def test_portfolio_rate_limiting_429_and_cancellation(self):
        # Verify HTTP 429 Retry-After is respected and recovers on retry
        resp_429 = MagicMock()
        resp_429.status_code = 429
        resp_429.headers = {"Retry-After": "0.05"}

        resp_200 = MagicMock()
        resp_200.status_code = 200
        resp_200.json.return_value = {
            "paging_metadata": {},
            "result": [
                {
                    "id": "https://wago.com/ids/aas/855-9910",
                    "idShort": "855-9910",
                    "submodels": [],
                }
            ],
        }

        mock_session = MagicMock()
        mock_session.get.side_effect = [resp_429, resp_200]

        res = run_portfolio_coverage_analysis(
            use_bulk_submodels=False,
            session_getter=lambda: mock_session,
        )
        assert res["total_unique_products"] == 1
        assert res["missing_pcfs"] == 1

        # Verify cancellation produces partial result cleanly
        stop_ev = threading.Event()
        stop_ev.set()
        res_cancel = run_portfolio_coverage_analysis(
            stop_event=stop_ev,
            session_getter=lambda: mock_session,
        )
        assert res_cancel["was_cancelled"] is True
        assert res_cancel["is_partial"] is True

    def test_large_portfolio_excel_export_100k_performance_and_completeness(self, tmp_path: Path):
        """
        Verifies that exporting >100,000 products across all 7 worksheets completes
        quickly, emits accurate stage/progress callbacks (never 100% before save),
        and preserves all worksheets, metrics, PCF values/units, and text article formats.
        """
        import time
        import zipfile
        from portfolio_coverage import (
            CLASS_PCF_AVAILABLE_AAS,
            CLASS_PCF_OFFLINE_ONLY,
            CLASS_PCF_MISSING_BOTH,
            ACTION_PUBLISHED,
            ACTION_REVIEW_PUBLICATION,
            ACTION_REVIEW_CALCULATION,
        )

        n_products = 105000
        all_products = []
        missing_pcf_products = []
        publication_candidates = []
        calculation_candidates = []

        for i in range(n_products):
            series = f"{200 + (i % 50)}"
            short_art = f"{series}-{100 + (i % 900)}"
            long_art = f"0{series}-0{100 + (i % 900)}"
            mod = i % 10
            if mod < 4:
                rec = {
                    "article_short": short_art,
                    "article_long": long_art,
                    "series": series,
                    "aas_id": f"https://wago.com/ids/aas/{short_art}",
                    "shell_count": 1,
                    "pcf_aas": 0.0116,
                    "pcf_unit_aas": "kg CO2e",
                    "pcf_offline": 0.0116,
                    "pcf_unit_offline": "kg CO2e",
                    "classification": CLASS_PCF_AVAILABLE_AAS,
                    "potential_action": ACTION_PUBLISHED,
                    "internal_status": "PCF available",
                }
                all_products.append(rec)
            elif mod == 4:
                rec = {
                    "article_short": short_art,
                    "article_long": long_art,
                    "series": series,
                    "aas_id": f"https://wago.com/ids/aas/{short_art}",
                    "shell_count": 1,
                    "pcf_aas": None,
                    "pcf_unit_aas": "",
                    "pcf_offline": 0.025,
                    "pcf_unit_offline": "kg CO2e",
                    "classification": CLASS_PCF_OFFLINE_ONLY,
                    "potential_action": ACTION_REVIEW_PUBLICATION,
                    "internal_status": "AAS shell exists, no PCF",
                }
                all_products.append(rec)
                missing_pcf_products.append(rec)
                publication_candidates.append(rec)
            else:
                rec = {
                    "article_short": short_art,
                    "article_long": long_art,
                    "series": series,
                    "aas_id": f"https://wago.com/ids/aas/{short_art}",
                    "shell_count": 1,
                    "pcf_aas": None,
                    "pcf_unit_aas": "",
                    "pcf_offline": None,
                    "pcf_unit_offline": "",
                    "classification": CLASS_PCF_MISSING_BOTH,
                    "potential_action": ACTION_REVIEW_CALCULATION,
                    "internal_status": "AAS shell exists, no PCF",
                }
                all_products.append(rec)
                missing_pcf_products.append(rec)
                calculation_candidates.append(rec)

        large_result = {
            "analysis_timestamp": "2026-11-18 11:30:00",
            "data_source": "WAGO SmartService AAS API",
            "offline_snapshot_version": "Q4 2026",
            "is_partial": False,
            "was_cancelled": False,
            "enumeration_complete": True,
            "max_shells_limit": None,
            "total_shells_retrieved": n_products,
            "non_product_shells_skipped": 0,
            "total_unique_products": n_products,
            "successfully_evaluated_products": n_products,
            "available_pcfs": len(all_products) - len(missing_pcf_products),
            "missing_pcfs": len(missing_pcf_products),
            "publication_candidates_count": len(publication_candidates),
            "calculation_candidates_count": len(calculation_candidates),
            "technical_errors": 0,
            "pcf_coverage_pct": 40.0,
            "analysis_completeness_pct": 100.0,
            "elapsed_seconds": 60.0,
            "all_products": all_products,
            "missing_pcf_products": missing_pcf_products,
            "publication_candidates": publication_candidates,
            "calculation_candidates": calculation_candidates,
            "query_errors": [],
            "series_breakdown": [
                {
                    "series": "200",
                    "total_products": n_products,
                    "available_pcfs": len(all_products) - len(missing_pcf_products),
                    "missing_pcfs": len(missing_pcf_products),
                    "publication_candidates": len(publication_candidates),
                    "calculation_candidates": len(calculation_candidates),
                    "technical_errors": 0,
                    "pcf_coverage_pct": 40.0,
                    "analysis_completeness_pct": 100.0,
                }
            ],
        }

        progress_events = []
        out_xlsx = tmp_path / "portfolio_105k.xlsx"

        t0 = time.perf_counter()
        export_portfolio_coverage_excel(
            large_result,
            out_xlsx,
            progress_callback=lambda info: progress_events.append(dict(info)),
        )
        duration = time.perf_counter() - t0

        assert out_xlsx.exists()
        assert duration < 20.0, f"105k export took too long: {duration:.2f}s"

        stages_seen = {ev["stage"] for ev in progress_events}
        assert "Preparing portfolio data..." in stages_seen
        assert "Writing Portfolio Overview..." in stages_seen
        assert "Exporting All Products..." in stages_seen
        assert "Exporting Missing PCFs..." in stages_seen
        assert "Exporting Publication Candidates..." in stages_seen
        assert "Exporting Calculation Candidates..." in stages_seen
        assert "Finalizing Excel workbook..." in stages_seen
        # Progress callback must never report 100% before the file is saved
        assert all(ev["percent"] < 100.0 for ev in progress_events)

        # Verify dimensions and sheet integrity inside the generated .xlsx
        with zipfile.ZipFile(out_xlsx, "r") as zf:
            s2_head = zf.read("xl/worksheets/sheet2.xml")[:1024].decode("utf-8")
            assert f'<dimension ref="A1:J{4 + n_products}"' in s2_head

    def test_gui_background_portfolio_export_responsiveness_cancellation_and_error_recovery(
        self, tmp_path: Path
    ):
        """
        Verifies background Portfolio Coverage export in PCFQueryApp:
        - GUI stays responsive while background thread exports
        - Duplicate simultaneous exports are prevented
        - Cancellation deletes temp file and does not produce incomplete output
        - Failed export preserves analysis results for retry
        """
        import time

        sample_result = {
            "analysis_timestamp": "2026-11-18 11:35:00",
            "data_source": "WAGO SmartService AAS API",
            "offline_snapshot_version": "Q4 2026",
            "is_partial": False,
            "was_cancelled": False,
            "enumeration_complete": True,
            "total_shells_retrieved": 500,
            "non_product_shells_skipped": 0,
            "total_unique_products": 500,
            "successfully_evaluated_products": 500,
            "available_pcfs": 250,
            "missing_pcfs": 250,
            "publication_candidates_count": 100,
            "calculation_candidates_count": 150,
            "technical_errors": 0,
            "pcf_coverage_pct": 50.0,
            "analysis_completeness_pct": 100.0,
            "elapsed_seconds": 2.0,
            "all_products": [
                {
                    "article_short": f"221-{100 + i}",
                    "article_long": f"0221-0{100 + i}",
                    "series": "221",
                    "aas_id": f"https://wago.com/ids/aas/221-{100 + i}",
                    "shell_count": 1,
                    "pcf_aas": 0.0116 if i % 2 == 0 else None,
                    "pcf_unit_aas": "kg CO2e" if i % 2 == 0 else "",
                    "pcf_offline": 0.0116,
                    "pcf_unit_offline": "kg CO2e",
                    "classification": "PCF available in AAS",
                    "potential_action": "PCF published",
                    "internal_status": "PCF available",
                }
                for i in range(500)
            ],
            "missing_pcf_products": [],
            "publication_candidates": [],
            "calculation_candidates": [],
            "query_errors": [],
            "series_breakdown": [],
        }

        app = pcf_tool.PCFQueryApp()
        try:
            app.select_workflow(WORKFLOW_PORTFOLIO)
            app._apply_portfolio_result(sample_result)
            assert app._portfolio_result is sample_result

            # 1. Start background export and verify duplicate export prevention
            out_ok = tmp_path / "bg_export_ok.xlsx"
            started_first = app.start_portfolio_export(str(out_ok), show_error_dialog=False)
            assert started_first is True
            assert app._portfolio_export_running is True

            # Second simultaneous export attempt must be rejected
            started_second = app.start_portfolio_export(
                str(tmp_path / "bg_export_dup.xlsx"), show_error_dialog=False
            )
            assert started_second is False

            # Pump GUI events while background thread finishes (proving GUI responsiveness)
            deadline = time.time() + 15.0
            while app._portfolio_export_running and time.time() < deadline:
                app.update()
                time.sleep(0.02)
            app.update()

            assert app._portfolio_export_running is False
            assert out_ok.exists()
            assert app._port_export_progress_var.get() == pytest.approx(100.0)
            assert str(app._btn_port_open_excel.cget("state")) == "normal"

            # 2. Test cancellation: no incomplete file may be left at destination
            out_cancel = tmp_path / "bg_export_cancelled.xlsx"
            app._portfolio_export_stop_event.set()
            with patch("aas_pcf_query.export_portfolio_coverage_excel", side_effect=RuntimeError("Export cancelled by user.")):
                app.start_portfolio_export(str(out_cancel), show_error_dialog=False)
                deadline = time.time() + 5.0
                while app._portfolio_export_running and time.time() < deadline:
                    app.update()
                    time.sleep(0.01)
                app.update()

            assert not out_cancel.exists()
            assert app._portfolio_result is sample_result
            assert app._port_export_progress_var.get() == pytest.approx(0.0)
            assert str(app._btn_port_export.cget("state")) == "normal"

            # 3. Test export error recovery: analysis results remain available after failure
            out_fail = tmp_path / "bg_export_failed.xlsx"
            with patch("aas_pcf_query.export_portfolio_coverage_excel", side_effect=OSError("Disk full")):
                app.start_portfolio_export(str(out_fail), show_error_dialog=False)
                deadline = time.time() + 5.0
                while app._portfolio_export_running and time.time() < deadline:
                    app.update()
                    time.sleep(0.01)
                app.update()

            assert not out_fail.exists()
            assert app._portfolio_result is sample_result
            assert "Export failed" in app._lbl_port_export_stage.cget("text")
            assert str(app._btn_port_export.cget("state")) == "normal"
        finally:
            app.destroy()


# ═══════════════════════════════════════════════════════════════════════════
# 6b. Application Icon, Window Size & Responsive Layout Tests
# ═══════════════════════════════════════════════════════════════════════════

class TestIconAndResponsiveLayout:

    def test_application_icon_multi_resolution_and_runtime_fallback(self):
        from PIL import Image

        repo_dir = Path(__file__).resolve().parent.parent
        ico_file = repo_dir / "wago_pcf_icon.ico"
        assert ico_file.exists(), "wago_pcf_icon.ico must exist in repository root"

        with Image.open(ico_file) as im:
            sizes = set(im.info.get("sizes", set()))
            expected = {(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)}
            assert expected.issubset(sizes), f"Missing ICO resolutions: {expected - sizes}"

        # Verify fallback icon generation works even if external file is absent
        fallback_bytes = pcf_tool._generate_fallback_ico_bytes()
        assert len(fallback_bytes) > 500

    def test_window_sizing_responsive_layout_and_dpi_display_compatibility(self):
        # 1. Verify responsive geometry calculation across corporate resolutions & DPI scales
        resolutions = [
            (1366, 768),
            (1920, 1080),
            (2560, 1440),
            (3840, 2160),
        ]
        dpi_scales = [1.0, 1.25, 1.5, 2.0]

        for phys_w, phys_h in resolutions:
            for scale in dpi_scales:
                eff_w = int(phys_w / scale)
                eff_h = max(420, int((phys_h - 48) / scale))
                w, h, px, py, min_w, min_h = pcf_tool.compute_responsive_window_geometry(
                    720, 820, eff_w, eff_h, work_x=0, work_y=0
                )
                assert w <= eff_w, f"Width {w} exceeds work area {eff_w} at {phys_w}x{phys_h} @{scale*100}%"
                assert h <= eff_h, f"Height {h} exceeds work area {eff_h} at {phys_w}x{phys_h} @{scale*100}%"
                assert min_w <= w
                assert min_h <= h
                assert px >= 0 and py >= 0

        # 2. Verify live GUI layout, workflow switching, offline notice, and scrollbar activation
        app = pcf_tool.PCFQueryApp()
        try:
            app.update_idletasks()
            app.update()

            # Verify icon loaded
            assert app._icon_path is not None and Path(app._icon_path).is_file()
            assert len(app._app_icon_photos) > 0

            # Verify window dimensions fit within available desktop work area
            _, _, work_w, work_h = pcf_tool.get_available_desktop_work_area(app)
            win_w = app.winfo_width()
            win_h = app.winfo_height()
            assert win_w >= 580
            assert win_h >= 520
            assert win_w <= work_w
            assert win_h <= work_h

            # Switch across all 3 workflows and Offline Mode; verify footer stays visible
            for wf in (WORKFLOW_SINGLE, WORKFLOW_BATCH, WORKFLOW_PORTFOLIO):
                app.select_workflow(wf)
                app.set_mode(MODE_OFFLINE)
                app.update_idletasks()
                footer_y = app._frm_footer_container.winfo_y()
                footer_h = app._frm_footer_container.winfo_height()
                assert footer_h > 0
                assert footer_y + footer_h <= app.winfo_height() + 2

            # Simulate a compact 1366x768 @ 150% viewport (height 560) and verify scrollbar activates
            app.select_workflow(WORKFLOW_PORTFOLIO)
            app.geometry("640x560")
            app.update_idletasks()
            app.update()
            app._update_scrollbar_visibility()
            assert app._vscroll_visible is True
        finally:
            app.destroy()


# ═══════════════════════════════════════════════════════════════════════════
# 7. Offline Dataset Integrity & Compiled Executable Verification
# ═══════════════════════════════════════════════════════════════════════════

class TestOfflineIntegrityAndExecutable:

    def test_conflicting_source_records_detected(self):
        with pytest.raises(ValueError, match="Conflicting PCF values"):
            validate_and_index_pcf_records([
                ("221-412", 0.0116),
                ("0221-0412", 0.0999),
            ])

    def test_complex_configurable_articles_in_offline_mode_and_dataset_completeness(self, tmp_path: Path):
        """
        Verifies both explicitly reported complex article numbers in Offline Mode
        (direct lookup, Single Query, Batch Query, short & long formats) and validates
        round-trip normalization across all 9,635 embedded records.
        """
        from offline_pcf_data import EMBEDDED_PCF_DATA

        art1_long  = "0734-0025/K050-1745/0000-00200PW25-0000"
        art1_short = "734-025/K050-1745/000-00200PW25-000"
        pcf1_exp   = 0.040994997

        art2_long  = "0734-0025/K050-1745/0000-0100PW25-0000"
        art2_short = "734-025/K050-1745/000-0100PW25-000"
        pcf2_exp   = 0.069879606

        # 1. Normalization & bidirectional conversion preserves complex middle suffix
        for s_form, l_form in [(art1_short, art1_long), (art2_short, art2_long)]:
            p_l = parse_article_number(l_form)
            p_s = parse_article_number(s_form)
            assert p_l["is_valid"] is True
            assert p_s["is_valid"] is True
            assert p_l["short"] == s_form
            assert p_l["long"] == l_form
            assert p_s["short"] == s_form
            assert p_s["long"] == l_form
            assert p_l["canonical"] == s_form
            assert p_s["canonical"] == s_form

        # 2. Direct embedded dataset lookup & Single Query in Offline Mode
        for query_in, exp_val, exp_s, exp_l in [
            (art1_long,  pcf1_exp, art1_short, art1_long),
            (art1_short, pcf1_exp, art1_short, art1_long),
            (art2_long,  pcf2_exp, art2_short, art2_long),
            (art2_short, pcf2_exp, art2_short, art2_long),
        ]:
            direct_res = get_offline_pcf_result(query_in)
            assert direct_res["pcf"] == pytest.approx(exp_val, rel=1e-12)
            assert direct_res["unit"] == "kg CO2e"
            assert direct_res["article_short"] == exp_s
            assert direct_res["article_long"] == exp_l

            single_res = query_single_article(query_in, mode=MODE_OFFLINE)
            assert single_res["pcf"] == pytest.approx(exp_val, rel=1e-12)
            assert single_res["unit"] == "kg CO2e"
            assert single_res["article_short"] == exp_s
            assert single_res["article_long"] == exp_l

        # 3. Batch Query in Offline Mode
        batch_in  = tmp_path / "complex_batch_in.xlsx"
        batch_out = tmp_path / "complex_batch_out.xlsx"
        pd.DataFrame({"Material lang ber.": [art1_long, art2_long]}).to_excel(batch_in, index=False)
        summary = main(str(batch_in), str(batch_out), mode=MODE_OFFLINE)
        assert summary is not None
        assert summary["pcf_available"] == 2
        assert summary["not_yet_available"] == 0

        df_out = read_pcf_results_table(batch_out)
        assert df_out[COL_ARTICLE_SHORT].tolist() == [art1_short, art2_short]
        assert df_out[COL_ARTICLE_LONG].tolist() == [art1_long, art2_long]
        assert float(df_out.loc[0, COL_PCF]) == pytest.approx(pcf1_exp, rel=1e-12)
        assert float(df_out.loc[1, COL_PCF]) == pytest.approx(pcf2_exp, rel=1e-12)

        # 4. Complete 9,635-record embedded dataset round-trip check
        assert len(EMBEDDED_PCF_DATA) == 9635
        for key_short, val_expected in EMBEDDED_PCF_DATA.items():
            p_info = parse_article_number(key_short)
            assert p_info["is_valid"], f"Embedded key invalid: {key_short}"
            assert p_info["canonical"] == key_short
            # Query via Long format must return the exact same PCF value
            res_long = get_offline_pcf_result(p_info["long"])
            assert res_long["pcf"] == val_expected, f"Long lookup failed for {key_short} -> {p_info['long']}"

    def test_drag_and_drop_paths_spaces_umlauts_nested_multiple_and_mainloop_stability(
        self, tmp_path: Path
    ):
        """
        Verifies that drag-and-drop handles .xlsx, .xls, spaces, German umlauts,
        deeply nested directories, unsupported files, multiple dropped files, and
        file replacement without crashing app.mainloop() or auto-starting the query.
        """
        # Prepare test files
        f_std = tmp_path / "standard.xlsx"
        pd.DataFrame({"Article Number": ["221-412"]}).to_excel(f_std, index=False)

        f_xls = tmp_path / "legacy_format.xls"
        _write_test_xls_file(f_xls, [["Article Number"], ["280-101"]])

        f_space = tmp_path / "WAGO File With Spaces 2026.xlsx"
        pd.DataFrame({"Article Number": ["221-2411"]}).to_excel(f_space, index=False)

        f_umlaut = tmp_path / "Artikelübersicht_ÄÖÜß_Werte.xlsx"
        pd.DataFrame({"Article Number": ["750-103"]}).to_excel(f_umlaut, index=False)

        nested_dir = tmp_path / "lvl1" / "lvl2" / "lvl3" / "lvl4" / "lvl5" / "lvl6"
        nested_dir.mkdir(parents=True, exist_ok=True)
        f_nested = nested_dir / "deeply_nested_workbook.xlsx"
        pd.DataFrame({"Article Number": ["2000-115"]}).to_excel(f_nested, index=False)

        f_unsupp = tmp_path / "not_an_excel_file.txt"
        f_unsupp.write_text("invalid content", encoding="utf-8")

        app = pcf_tool.PCFQueryApp()
        try:
            # 1. Standard .xlsx
            assert app.handle_dropped_file(str(f_std)) is True
            assert app._input_path == str(f_std)
            assert app._query_thread is None  # Must not auto-start query

            # 2. Legacy .xls (replaces already selected file)
            assert app.handle_dropped_file(str(f_xls)) is True
            assert app._input_path == str(f_xls)

            # 3. File name with spaces (also test Tcl brace-wrapped event format)
            assert app.handle_dropped_file(f"{{{f_space}}}") is True
            assert app._input_path == str(f_space)

            # 4. File name with German umlauts
            assert app.handle_dropped_file(str(f_umlaut)) is True
            assert app._input_path == str(f_umlaut)

            # 5. Deeply nested directory
            assert app.handle_dropped_file(str(f_nested)) is True
            assert app._input_path == str(f_nested)

            # 6. Unsupported file format (rejected safely without crash)
            assert app.handle_dropped_file(str(f_unsupp)) is False
            assert "Unsupported file" in app._lbl_drop_hint.cget("text")

            # 7. Multiple dropped files (selects first valid Excel file gracefully)
            assert app.handle_dropped_files([str(f_space), str(f_std)]) is True
            assert app._input_path == str(f_space)
            assert "1 of 2 dropped files" in app._lbl_drop_hint.cget("text")

            # 8. Real Win32 WM_DROPFILES posted during app.mainloop() (verifies GIL safety)
            def _send_win32_drop():
                pcf_tool._post_win32_hdrop_message(app.winfo_id(), [str(f_umlaut)])
                app.after(150, app.quit)

            app.after(30, _send_win32_drop)
            app.mainloop()
            assert app._input_path == str(f_umlaut)
        finally:
            app.destroy()

    def test_excel_table_filters_functional_in_batch_and_portfolio_exports(self, tmp_path: Path):
        """
        Verifies that exported Excel workbooks (Batch Query 'PCF Results' and
        Portfolio Coverage 7-sheet report) have:
        1. Exactly 1 worksheet with tabSelected='1' (preventing Excel [Group] mode
           which disables AutoFilter dropdowns in the Excel UI).
        2. Matching <autoFilter ref="..."/> and <definedName name="_xlnm._FilterDatabase">
           covering all data rows (never just the header row).
        """
        import re
        import zipfile

        # Part A: Batch Query export from an input workbook where original sheet has tabSelected=True
        in_xlsx = tmp_path / "input_grouped.xlsx"
        out_xlsx = tmp_path / "output_grouped_PCF.xlsx"
        wb_in = openpyxl.Workbook()
        ws_in = wb_in.active
        ws_in.title = "Original_Data"
        ws_in.views.sheetView[0].tabSelected = True
        ws_in.append(["Material", "Qty"])
        ws_in.append(["0734-0025/K050-1745/0000-00200PW25-0000", 10])
        ws_in.append(["0734-0025/K050-1745/0000-0100PW25-0000", 20])
        ws_in.append(["999-999", 5])
        wb_in.save(in_xlsx)
        wb_in.close()

        summary = main(str(in_xlsx), str(out_xlsx), mode=MODE_OFFLINE)
        assert summary is not None and out_xlsx.exists()

        with zipfile.ZipFile(out_xlsx, "r") as zf:
            wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
            s1_xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
            s2_xml = zf.read("xl/worksheets/sheet2.xml").decode("utf-8")

            # _FilterDatabase in workbook.xml must cover A11:D14
            assert "'PCF Results'!$A$11:$D$14" in wb_xml
            assert '<autoFilter ref="A11:D14"/>' in s1_xml

            # Only sheet1 (PCF Results) may have tabSelected="1"; sheet2 must NOT have tabSelected="1"
            assert 'tabSelected="1"' in s1_xml
            assert 'tabSelected="1"' not in s2_xml

        # Part B: Portfolio Coverage export across all filterable worksheets
        port_res = {
            "analysis_timestamp": "2026-11-18 13:00:00",
            "data_source": "WAGO SmartService AAS API",
            "offline_snapshot_version": "Q4 2026",
            "is_partial": False,
            "was_cancelled": False,
            "enumeration_complete": True,
            "total_shells_retrieved": 3,
            "non_product_shells_skipped": 0,
            "total_unique_products": 3,
            "successfully_evaluated_products": 3,
            "available_pcfs": 1,
            "missing_pcfs": 2,
            "publication_candidates_count": 1,
            "calculation_candidates_count": 1,
            "technical_errors": 0,
            "pcf_coverage_pct": 33.33,
            "analysis_completeness_pct": 100.0,
            "elapsed_seconds": 1.0,
            "all_products": [
                {
                    "article_short": "221-412",
                    "article_long": "0221-0412",
                    "series": "221",
                    "aas_id": "id1",
                    "shell_count": 1,
                    "pcf_aas": 0.0116,
                    "pcf_unit_aas": "kg CO2e",
                    "pcf_offline": 0.0116,
                    "pcf_unit_offline": "kg CO2e",
                    "classification": "PCF available in AAS",
                    "potential_action": "PCF published",
                    "internal_status": "PCF available",
                },
                {
                    "article_short": "734-025/K050-1745/000-00200PW25-000",
                    "article_long": "0734-0025/K050-1745/0000-00200PW25-0000",
                    "series": "734",
                    "aas_id": "id2",
                    "shell_count": 1,
                    "pcf_aas": None,
                    "pcf_unit_aas": "",
                    "pcf_offline": 0.040994997,
                    "pcf_unit_offline": "kg CO2e",
                    "classification": "PCF missing in AAS, available in Q4 2026 snapshot",
                    "potential_action": "Review publication or data integration",
                    "internal_status": "AAS shell exists, no PCF",
                },
                {
                    "article_short": "855-9910",
                    "article_long": "0855-9910",
                    "series": "855",
                    "aas_id": "id3",
                    "shell_count": 1,
                    "pcf_aas": None,
                    "pcf_unit_aas": "",
                    "pcf_offline": None,
                    "pcf_unit_offline": "",
                    "classification": "PCF missing in both sources",
                    "potential_action": "Review potential calculation requirement",
                    "internal_status": "AAS shell exists, no PCF",
                },
            ],
            "missing_pcf_products": [
                {
                    "article_short": "734-025/K050-1745/000-00200PW25-000",
                    "article_long": "0734-0025/K050-1745/0000-00200PW25-0000",
                    "series": "734",
                    "aas_id": "id2",
                    "shell_count": 1,
                    "pcf_aas": None,
                    "pcf_unit_aas": "",
                    "pcf_offline": 0.040994997,
                    "pcf_unit_offline": "kg CO2e",
                    "classification": "PCF missing in AAS, available in Q4 2026 snapshot",
                    "potential_action": "Review publication or data integration",
                    "internal_status": "AAS shell exists, no PCF",
                },
                {
                    "article_short": "855-9910",
                    "article_long": "0855-9910",
                    "series": "855",
                    "aas_id": "id3",
                    "shell_count": 1,
                    "pcf_aas": None,
                    "pcf_unit_aas": "",
                    "pcf_offline": None,
                    "pcf_unit_offline": "",
                    "classification": "PCF missing in both sources",
                    "potential_action": "Review potential calculation requirement",
                    "internal_status": "AAS shell exists, no PCF",
                },
            ],
            "publication_candidates": [
                {
                    "article_short": "734-025/K050-1745/000-00200PW25-000",
                    "article_long": "0734-0025/K050-1745/0000-00200PW25-0000",
                    "series": "734",
                    "aas_id": "id2",
                    "shell_count": 1,
                    "pcf_aas": None,
                    "pcf_unit_aas": "",
                    "pcf_offline": 0.040994997,
                    "pcf_unit_offline": "kg CO2e",
                    "classification": "PCF missing in AAS, available in Q4 2026 snapshot",
                    "potential_action": "Review publication or data integration",
                    "internal_status": "AAS shell exists, no PCF",
                },
            ],
            "calculation_candidates": [
                {
                    "article_short": "855-9910",
                    "article_long": "0855-9910",
                    "series": "855",
                    "aas_id": "id3",
                    "shell_count": 1,
                    "pcf_aas": None,
                    "pcf_unit_aas": "",
                    "pcf_offline": None,
                    "pcf_unit_offline": "",
                    "classification": "PCF missing in both sources",
                    "potential_action": "Review potential calculation requirement",
                    "internal_status": "AAS shell exists, no PCF",
                },
            ],
            "query_errors": [],
            "series_breakdown": [
                {
                    "series": "221",
                    "total_products": 1,
                    "available_pcfs": 1,
                    "missing_pcfs": 0,
                    "publication_candidates": 0,
                    "calculation_candidates": 0,
                    "technical_errors": 0,
                    "pcf_coverage_pct": 100.0,
                    "analysis_completeness_pct": 100.0,
                }
            ],
        }

        port_xlsx = tmp_path / "portfolio_filter_check.xlsx"
        export_portfolio_coverage_excel(port_res, port_xlsx)

        with zipfile.ZipFile(port_xlsx, "r") as zf:
            wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
            assert "'All Products'!$A$4:$J$7" in wb_xml
            assert "'Missing PCFs'!$A$4:$J$6" in wb_xml
            assert "'Publication Candidates'!$A$4:$J$5" in wb_xml
            assert "'Calculation Candidates'!$A$4:$J$5" in wb_xml
            assert "'Query Errors'!$A$4:$F$5" in wb_xml
            assert "'Coverage by Product Series'!$A$4:$I$5" in wb_xml
            # Never allow header-only $A$4:$J$4 _FilterDatabase
            assert "$A$4:$J$4" not in wb_xml

    def test_mhtml_html_xml2003_and_corrupt_xls_import_and_batch_export(self, tmp_path: Path):
        """
        Verifies content-based file format detection and end-to-end import + batch export for:
        1. Native .xlsx files
        2. Native binary BIFF .xls files
        3. MHTML-based .xls exports (SAP/enterprise MIME-Version 1.0, quoted-printable & base64,
           multi-sheet <x:ExcelWorksheet>, German umlauts, leading zeros, complex WAGO suffixes)
        4. HTML-based .xls exports (nested layout table + colspan/rowspan)
        5. Excel 2003 XML SpreadsheetML .xls exports (with ss:Index)
        6. Genuinely corrupted/unsupported .xls files (clear error message, no 'Expected BOF record',
           and GUI remains responsive for selecting another file).
        """
        import base64
        import quopri

        # 1. Native .xlsx & 2. Native BIFF .xls format detection
        xlsx_file = tmp_path / "native.xlsx"
        pd.DataFrame({"Article": ["221-412"]}).to_excel(xlsx_file, index=False)
        assert detect_excel_file_format(xlsx_file) == "xlsx"

        biff_xls = tmp_path / "native_biff.xls"
        _write_test_xls_file(biff_xls, [["Article"], ["0221-0412"]])
        assert detect_excel_file_format(biff_xls) == "xls_biff"

        # 3. MHTML-based .xls export (simulating SAP multi-sheet MHTML with quoted-printable + base64)
        sheet1_html = """<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">
<head><meta http-equiv=Content-Type content="text/html; charset=utf-8"></head>
<body>
<table>
  <tr><td>SAP Export Report</td><td>Stückliste ÄÖÜß</td><td></td></tr>
  <tr><th>Pos</th><th>Materialnummer</th><th>Bezeichnung</th></tr>
  <tr><td>10</td><td>0221-0412</td><td>COMPACT-Verbindungsklemme 2-Leiter</td></tr>
  <tr><td>20</td><td>0734-0025/K050-1745/0000-00200PW25-0000</td><td>Sondervariante Prüfgerät</td></tr>
  <tr><td>30</td><td>0734-0025/K050-1745/0000-0100PW25-0000</td><td>Sondervariante Zubehör</td></tr>
</table>
</body></html>"""

        sheet2_html = """<html><head><meta charset="utf-8"></head><body>
<table>
  <tr><th>Notiz</th><th>Wert</th></tr>
  <tr><td>Prüfung</td><td>Bestanden (100%)</td></tr>
</table>
</body></html>"""

        qp_sheet1 = quopri.encodestring(sheet1_html.encode("utf-8")).decode("ascii")
        b64_sheet2 = base64.b64encode(sheet2_html.encode("utf-8")).decode("ascii")

        mhtml_content = (
            "MIME-Version: 1.0\r\n"
            'Content-Type: multipart/related; boundary="----=_NextPart_SAP_MHTML"\r\n'
            "\r\n"
            "------=_NextPart_SAP_MHTML\r\n"
            "Content-Location: file:///C:/SAP/Workbook.htm\r\n"
            'Content-Type: text/html; charset="utf-8"\r\n'
            "\r\n"
            "<html xmlns:x=\"urn:schemas-microsoft-com:office:excel\"><head>\r\n"
            "<xml><x:ExcelWorkbook><x:ExcelWorksheets>\r\n"
            "<x:ExcelWorksheet><x:Name>SAP Stückliste</x:Name>"
            "<x:WorksheetSource HRef=\"sheet001.htm\"/></x:ExcelWorksheet>\r\n"
            "<x:ExcelWorksheet><x:Name>Metadaten</x:Name>"
            "<x:WorksheetSource HRef=\"sheet002.htm\"/></x:ExcelWorksheet>\r\n"
            "</x:ExcelWorksheets></x:ExcelWorkbook></xml></head><body></body></html>\r\n"
            "------=_NextPart_SAP_MHTML\r\n"
            "Content-Location: file:///C:/SAP/sheet001.htm\r\n"
            "Content-Transfer-Encoding: quoted-printable\r\n"
            'Content-Type: text/html; charset="utf-8"\r\n'
            "\r\n"
            f"{qp_sheet1}\r\n"
            "------=_NextPart_SAP_MHTML\r\n"
            "Content-Location: file:///C:/SAP/sheet002.htm\r\n"
            "Content-Transfer-Encoding: base64\r\n"
            'Content-Type: text/html; charset="utf-8"\r\n'
            "\r\n"
            f"{b64_sheet2}\r\n"
            "------=_NextPart_SAP_MHTML--\r\n"
        )
        mhtml_xls = tmp_path / "sap_export_mhtml.xls"
        mhtml_xls.write_bytes(mhtml_content.encode("utf-8"))

        assert detect_excel_file_format(mhtml_xls) == "mhtml"
        df_mhtml, det_mhtml = load_excel_flexible(mhtml_xls)
        assert det_mhtml["sheet_name"] == "SAP Stückliste"
        assert det_mhtml["column"] == "Materialnummer"
        assert det_mhtml["header_row"] == 1
        assert det_mhtml["file_format"] == "mhtml"
        assert list(df_mhtml.columns) == ["Pos", "Materialnummer", "Bezeichnung"]
        assert df_mhtml["Materialnummer"].tolist() == [
            "0221-0412",
            "0734-0025/K050-1745/0000-00200PW25-0000",
            "0734-0025/K050-1745/0000-0100PW25-0000",
        ]
        assert "Prüfgerät" in df_mhtml["Bezeichnung"].iloc[1]

        # Run full batch query & export on MHTML .xls
        mhtml_out = tmp_path / "sap_export_mhtml_out.xlsx"
        main(mhtml_xls, mhtml_out, mode=MODE_OFFLINE)
        assert mhtml_out.exists()
        wb_mhtml_out = openpyxl.load_workbook(mhtml_out)
        assert wb_mhtml_out.sheetnames == ["PCF Results", "SAP Stückliste", "Metadaten"]
        # Check preserved sheet content in exported .xlsx
        ws_sap = wb_mhtml_out["SAP Stückliste"]
        assert ws_sap.cell(row=3, column=2).value == "0221-0412"
        wb_mhtml_out.close()

        df_mhtml_res = read_pcf_results_table(mhtml_out)
        assert df_mhtml_res[COL_ARTICLE_LONG].tolist() == [
            "0221-0412",
            "0734-0025/K050-1745/0000-00200PW25-0000",
            "0734-0025/K050-1745/0000-0100PW25-0000",
        ]
        assert float(df_mhtml_res.loc[0, COL_PCF]) == pytest.approx(0.0116)
        assert float(df_mhtml_res.loc[1, COL_PCF]) == pytest.approx(0.040994997, rel=1e-12)
        assert float(df_mhtml_res.loc[2, COL_PCF]) == pytest.approx(0.069879606, rel=1e-12)

        # 4. HTML-based .xls export (with nested layout wrapper table, cp1252 umlauts, colspan)
        html_content = """<html><head><title>HTML Stückliste</title></head><body>
<table><tr><td>
  <table border="1">
    <tr><th colspan="2">Kopfzeile Übersicht</th></tr>
    <tr><th>WAGO Artikel-Nr.</th><th>Beschreibung</th></tr>
    <tr><td>0221-0412</td><td>Klemme für Gehäusegröße Standard</td></tr>
    <tr><td>221-2411</td><td>Durchgangsverbinder</td></tr>
  </table>
</td></tr></table>
</body></html>"""
        html_xls = tmp_path / "html_export.xls"
        html_xls.write_bytes(html_content.encode("cp1252"))

        assert detect_excel_file_format(html_xls) == "html"
        df_html, det_html = load_excel_flexible(html_xls)
        assert det_html["sheet_name"] == "HTML Stückliste"
        assert det_html["column"] == "WAGO Artikel-Nr."
        assert det_html["header_row"] == 1
        assert det_html["file_format"] == "html"
        assert df_html["WAGO Artikel-Nr."].tolist() == ["0221-0412", "221-2411"]
        assert "Gehäusegröße" in df_html["Beschreibung"].iloc[0]

        html_out = tmp_path / "html_export_out.xlsx"
        main(html_xls, html_out, mode=MODE_OFFLINE)
        wb_html_out = openpyxl.load_workbook(html_out)
        assert wb_html_out.sheetnames == ["PCF Results", "HTML Stückliste"]
        wb_html_out.close()

        # 5. Excel 2003 XML SpreadsheetML (.xls)
        xml2003_content = """<?xml version="1.0" encoding="UTF-8"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
 <Worksheet ss:Name="XML2003 Daten">
  <Table>
   <Row>
    <Cell><Data ss:Type="String">Position</Data></Cell>
    <Cell><Data ss:Type="String">Artikelnummer</Data></Cell>
    <Cell><Data ss:Type="String">Kommentar</Data></Cell>
   </Row>
   <Row>
    <Cell><Data ss:Type="Number">1</Data></Cell>
    <Cell><Data ss:Type="String">0221-0412</Data></Cell>
    <Cell><Data ss:Type="String">Leading Zero Preserved</Data></Cell>
   </Row>
   <Row>
    <Cell ss:Index="2"><Data ss:Type="String">0734-0025/K050-1745/0000-00200PW25-0000</Data></Cell>
    <Cell><Data ss:Type="String">Complex Variant</Data></Cell>
   </Row>
  </Table>
 </Worksheet>
</Workbook>"""
        xml_xls = tmp_path / "xml2003_export.xls"
        xml_xls.write_text(xml2003_content, encoding="utf-8")

        assert detect_excel_file_format(xml_xls) == "xml_2003"
        df_xml, det_xml = load_excel_flexible(xml_xls)
        assert det_xml["sheet_name"] == "XML2003 Daten"
        assert det_xml["column"] == "Artikelnummer"
        assert det_xml["header_row"] == 0
        assert det_xml["file_format"] == "xml_2003"
        assert df_xml["Artikelnummer"].tolist() == [
            "0221-0412",
            "0734-0025/K050-1745/0000-00200PW25-0000",
        ]

        xml_out = tmp_path / "xml2003_export_out.xlsx"
        main(xml_xls, xml_out, mode=MODE_OFFLINE)
        wb_xml_out = openpyxl.load_workbook(xml_out)
        assert wb_xml_out.sheetnames == ["PCF Results", "XML2003 Daten"]
        wb_xml_out.close()

        # 6. Genuinely corrupted/unsupported .xls file & GUI recovery
        corrupt_xls = tmp_path / "corrupt_random.xls"
        corrupt_xls.write_bytes(b"\x01\x02\x03NOT_AN_EXCEL_OR_MHTML_FILE\xff\xfe\x00\x00")
        assert detect_excel_file_format(corrupt_xls) == "unknown"

        with pytest.raises(ValueError) as exc_info:
            load_excel_flexible(corrupt_xls)
        err_msg = str(exc_info.value)
        assert "Expected BOF record" not in err_msg
        assert "not a valid or supported Excel workbook" in err_msg

        # Verify GUI stays responsive and can immediately load a valid MHTML .xls afterward
        app = pcf_tool.PCFQueryApp()
        try:
            app.withdraw()
            app._load_file(str(corrupt_xls))
            assert app._df is None
            assert "Could not read file" in app._lbl_summary.cget("text")
            assert "Expected BOF record" not in app._lbl_summary.cget("text")
            assert str(app._btn_start.cget("state")) == "disabled"

            # Now select the valid MHTML .xls file without restarting the application
            app._load_file(str(mhtml_xls))
            assert app._df is not None
            assert app._detection_info["column"] == "Materialnummer"
            assert app._sheet_var.get() == "SAP Stückliste"
            assert "Materialnummer" in app._lbl_summary.cget("text")
            assert str(app._btn_start.cget("state")) == "normal"
        finally:
            app.destroy()

    def test_compiled_executable_works_without_reference_excel_files(self, tmp_path: Path):
        """
        Verifies that the compiled dist/WAGO_PCF_Query.exe:
        1. Runs in Offline Mode when the original reference Excel files in dist/ are unavailable,
           including the corrected complex articles (0734-0025/K050-1745/0000-00200PW25-0000 and
           0734-0025/K050-1745/0000-0100PW25-0000).
        2. Processes Win32 WM_DROPFILES drag-and-drop events in the compiled GUI across all 8
           required scenarios (.xlsx, .xls, spaces, umlauts, nested folder, unsupported file,
           multiple files, replacement) without crashing.
        """
        import json
        import subprocess

        repo_dir = Path(__file__).resolve().parent.parent
        exe_path = repo_dir / "dist" / "WAGO_PCF_Query.exe"
        if not exe_path.exists():
            pytest.skip("Compiled executable dist/WAGO_PCF_Query.exe not present")

        ref_files = [
            repo_dir / "dist" / "2026 Q4_Consolidated PCF list.xlsx",
            repo_dir / "dist" / "Artikelübersicht PCF Werte_PCF.xlsx",
            repo_dir / "dist" / "Artikelübersicht PCF Werte.xlsx",
        ]
        moved_pairs = []

        test_in  = tmp_path / "exe_offline_test_in.xlsx"
        test_out = tmp_path / "exe_offline_test_out.xlsx"

        df_in = pd.DataFrame({
            "WAGO Artikel": [
                "221-412",
                "0221-0412",
                "221-2411",
                "0734-0025/K050-1745/0000-00200PW25-0000",
                "0734-0025/K050-1745/0000-0100PW25-0000",
                "999-999",
            ],
            "Comment": ["short 412", "long 412", "short 2411", "PW25 00200", "PW25 0100", "unknown"],
        })
        df_in.to_excel(test_in, index=False)

        # Prepare Drag-and-Drop test files for the compiled executable
        f_xls = tmp_path / "legacy_drop.xls"
        _write_test_xls_file(f_xls, [["Article Number"], ["280-101"]])
        f_mhtml = tmp_path / "sap_mhtml_drop.xls"
        f_mhtml.write_text(
            "MIME-Version: 1.0\r\n"
            'Content-Type: text/html; charset="utf-8"\r\n\r\n'
            "<html><body><table><tr><th>Artikel</th></tr>"
            "<tr><td>0221-0412</td></tr></table></body></html>\r\n",
            encoding="utf-8",
        )
        f_space = tmp_path / "File With Spaces 2026.xlsx"
        df_in.to_excel(f_space, index=False)
        f_umlaut = tmp_path / "Übersicht_ÄÖÜß_Test.xlsx"
        df_in.to_excel(f_umlaut, index=False)
        nested_dir = tmp_path / "d1" / "d2" / "d3" / "d4" / "d5"
        nested_dir.mkdir(parents=True, exist_ok=True)
        f_nested = nested_dir / "nested_drop.xlsx"
        df_in.to_excel(f_nested, index=False)
        f_bad = tmp_path / "invalid_drop.txt"
        f_bad.write_text("unsupported", encoding="utf-8")

        scenarios_json = tmp_path / "dnd_scenarios.json"
        report_json    = tmp_path / "dnd_report.json"
        scenarios = [
            {"name": "standard_xlsx",   "paths": [str(test_in)]},
            {"name": "legacy_xls",      "paths": [str(f_xls)]},
            {"name": "mhtml_xls",       "paths": [str(f_mhtml)]},
            {"name": "spaces_in_name",  "paths": [str(f_space)]},
            {"name": "german_umlauts",  "paths": [str(f_umlaut)]},
            {"name": "deeply_nested",   "paths": [str(f_nested)]},
            {"name": "unsupported_txt", "paths": [str(f_bad)]},
            {"name": "multiple_files",  "paths": [str(f_space), str(test_in)]},
            {"name": "replace_file",    "paths": [str(test_in)]},
        ]
        scenarios_json.write_text(json.dumps(scenarios), encoding="utf-8")

        try:
            for rf in ref_files:
                if rf.exists():
                    hidden_path = tmp_path / f"_hidden_{rf.name}"
                    rf.rename(hidden_path)
                    moved_pairs.append((hidden_path, rf))

            # 1. Batch offline test on compiled exe
            proc = subprocess.run(
                [str(exe_path), "--batch", str(test_in), str(test_out), "--mode", "offline"],
                cwd=str(tmp_path),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=60,
            )
            assert proc.returncode == 0, f"Executable batch failed: {proc.stderr}\n{proc.stdout}"
            assert test_out.exists()

            # 2. Drag-and-Drop GUI test on compiled exe across all 8 scenarios
            proc_dnd = subprocess.run(
                [str(exe_path), "--test-dnd-gui", str(scenarios_json), str(report_json)],
                cwd=str(tmp_path),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=60,
            )
            assert proc_dnd.returncode == 0, f"Executable DnD crashed: {proc_dnd.stderr}\n{proc_dnd.stdout}"
            assert report_json.exists()
        finally:
            for hidden_path, orig_path in moved_pairs:
                if hidden_path.exists():
                    hidden_path.rename(orig_path)

        wb_out = openpyxl.load_workbook(test_out)
        assert wb_out.sheetnames == ["PCF Results", "Sheet1"]
        wb_out.close()

        df_res = read_pcf_results_table(test_out)
        assert list(df_res.columns) == [COL_ARTICLE_SHORT, COL_ARTICLE_LONG, COL_PCF, COL_PCF_UNIT]
        assert df_res[COL_ARTICLE_SHORT].tolist() == [
            "221-412",
            "221-412",
            "221-2411",
            "734-025/K050-1745/000-00200PW25-000",
            "734-025/K050-1745/000-0100PW25-000",
            "999-999",
        ]
        assert float(df_res.loc[0, COL_PCF]) == pytest.approx(0.0116)
        assert float(df_res.loc[3, COL_PCF]) == pytest.approx(0.040994997, rel=1e-12)
        assert float(df_res.loc[4, COL_PCF]) == pytest.approx(0.069879606, rel=1e-12)
        assert df_res.loc[5, COL_PCF] == "Not yet available"
        assert df_res.loc[5, COL_PCF_UNIT] == ""

        dnd_rep = json.loads(report_json.read_text(encoding="utf-8"))
        assert dnd_rep["status"] == "ok"
        res_by_name = {r["name"]: r for r in dnd_rep["results"]}
        assert res_by_name["standard_xlsx"]["input_path"] == str(test_in)
        assert res_by_name["legacy_xls"]["input_path"] == str(f_xls)
        assert res_by_name["mhtml_xls"]["input_path"] == str(f_mhtml)
        assert res_by_name["spaces_in_name"]["input_path"] == str(f_space)
        assert res_by_name["german_umlauts"]["input_path"] == str(f_umlaut)
        assert res_by_name["deeply_nested"]["input_path"] == str(f_nested)
        assert "Unsupported file" in res_by_name["unsupported_txt"]["hint_text"]
        assert res_by_name["multiple_files"]["input_path"] == str(f_space)
        assert res_by_name["replace_file"]["input_path"] == str(test_in)
        assert all(r["query_running"] is False for r in dnd_rep["results"])
