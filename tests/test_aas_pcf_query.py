"""
Automated test suite for WAGO AAS PCF Query Tool (aas_pcf_query.py).

Covers:
1. Excel structure & article column detection
2. Article number normalization & bidirectional conversion
3. WAGO AAS API behavior (Fast Path, Fallback, Auth errors, Timeouts, Unexpected responses, Batch deduplication)
4. Customer-facing Excel export & status simplification
"""

import json
from pathlib import Path
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
    COL_PCF_STATUS,
    INTERNAL_STATUS_AAS_NO_PCF,
    INTERNAL_STATUS_AUTH_ERROR,
    INTERNAL_STATUS_INVALID_RESP,
    INTERNAL_STATUS_NOT_FOUND,
    INTERNAL_STATUS_PCF_AVAILABLE,
    INTERNAL_STATUS_SERVER_ERROR,
    INTERNAL_STATUS_TIMEOUT,
    STATUS_NOT_YET_AVAILABLE,
    STATUS_PCF_AVAILABLE,
    STATUS_QUERY_ERROR,
    convert_article_number,
    detect_article_column,
    get_pcf,
    get_pcf_result,
    is_valid_wago_article,
    load_excel_flexible,
    main,
    normalize_article_number,
    parse_article_number,
    prepare_output_dataframe,
    to_long_article_number,
    to_short_article_number,
)


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
        assert to_short_article_number("221-2411") == "221-2411"
        assert to_long_article_number("0221-2411") == "0221-2411"
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
        assert to_short_article_number(short_fmt) == short_fmt
        assert to_long_article_number(long_fmt) == long_fmt

    def test_duplicate_products_in_different_representations_normalize_identically(self):
        variants_221_412 = ["221-412", "0221-0412", "221-0412", "0221-412", "2210412", "02210412", " 0221-0412 "]
        normalized = {normalize_article_number(v) for v in variants_221_412}
        assert normalized == {"221-412"}

        variants_221_2411 = ["221-2411", "0221-2411", "2212411", "02212411"]
        normalized_2411 = {normalize_article_number(v) for v in variants_221_2411}
        assert normalized_2411 == {"221-2411"}

    def test_leading_zeros_lost_in_numeric_or_mixed_values(self):
        # 7-digit integer where Excel stripped leading zero of 02210412
        p1 = parse_article_number(2210412)
        assert p1["is_valid"] is True
        assert p1["short"] == "221-412"
        assert p1["long"] == "0221-0412"

        # Float artifact 2212411.0 from pandas/Excel
        p2 = parse_article_number("2212411.0")
        assert p2["is_valid"] is True
        assert p2["short"] == "221-2411"
        assert p2["long"] == "0221-2411"

        # Partial leading zero loss in hyphenated strings
        assert convert_article_number("221-0412") == ("221-412", "0221-0412")
        assert convert_article_number("0221-412") == ("221-412", "0221-0412")
        assert convert_article_number("85-101") == ("085-101", "0085-0101")

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
            "51170648",   # Internal SAP ID (51xxxxxx), not a WAGO article number
            "2026-11-16", # ISO Date
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
        assert to_short_article_number(invalid_val) == invalid_val
        assert to_long_article_number(invalid_val) == invalid_val

    def test_empty_and_none_inputs(self):
        for val in (None, "", "   ", float("nan")):
            parsed = parse_article_number(val)
            assert parsed["is_valid"] is False
            assert parsed["short"] == ""
            assert parsed["long"] == ""
            assert parsed["canonical"] == ""


# ═══════════════════════════════════════════════════════════════════════════
# 2. Structure-Agnostic Excel Detection Tests
# ═══════════════════════════════════════════════════════════════════════════

class TestExcelDetection:

    def test_article_numbers_in_first_column(self, tmp_path: Path):
        df = pd.DataFrame({
            "Article Number": ["221-412", "280-101", "2000-115"],
            "Description": ["Connector", "Terminal block", "TOPJOB S"],
            "Quantity": [10, 20, 30],
        })
        excel_path = tmp_path / "first_col.xlsx"
        df.to_excel(excel_path, index=False)

        info = detect_article_column(excel_path)
        assert info["column"] == "Article Number"
        assert info["is_ambiguous"] is False
        assert info["confidence"] >= 0.85

    def test_article_numbers_in_arbitrary_column(self, tmp_path: Path):
        df = pd.DataFrame({
            "Index": ["1", "2", "3", "4"],
            "Category": ["Splicing", "Rail", "Rail", "WINSTA"],
            "Supplier": ["WAGO", "WAGO", "WAGO", "WAGO"],
            "Part_Code": ["221-412", "0280-0101", "2000-1201", "770-103"],
            "Notes": ["urgent", "stock", "new", "ok"],
        })
        excel_path = tmp_path / "arbitrary_col.xlsx"
        df.to_excel(excel_path, index=False)

        info = detect_article_column(excel_path)
        assert info["column"] == "Part_Code"
        assert info["is_ambiguous"] is False

    def test_missing_or_unusual_headers_and_introductory_rows(self, tmp_path: Path):
        # 1. Unusual / unnamed headers
        df_unusual = pd.DataFrame({
            "Col_X": ["Alpha", "Beta", "Gamma"],
            "Unnamed: 1": ["221-412", "221-2411", "280-101"],
            "Col_Z": ["100", "200", "300"],
        })
        info_unusual = detect_article_column(df_unusual)
        assert info_unusual["column"] == "Unnamed: 1"

        # 2. Headerless Excel file (Row 1 starts directly with data)
        headerless_path = tmp_path / "headerless.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Connector 2-conductor", "221-412", 50])
        ws.append(["Connector 3-conductor", "221-413", 100])
        ws.append(["Inline splicing connector", "221-2411", 25])
        wb.save(headerless_path)

        df_loaded, info_hl = load_excel_flexible(headerless_path)
        assert info_hl["has_header"] is False
        assert len(df_loaded) == 3
        assert info_hl["column"] == "Column 2"
        assert df_loaded["Column 2"].tolist() == ["221-412", "221-413", "221-2411"]

        # 3. Introductory rows before the actual table
        intro_path = tmp_path / "intro_rows.xlsx"
        wb2 = openpyxl.Workbook()
        ws2 = wb2.active
        ws2.append(["Customer Sustainability Inquiry"])
        ws2.append(["Created on: 2026-11-16"])
        ws2.append([])
        ws2.append(["Pos", "Materialnummer", "Bezeichnung"])
        ws2.append(["1", "0221-0412", "Verbindungsklemme"])
        ws2.append(["2", "0221-2411", "Durchgangsverbinder"])
        wb2.save(intro_path)

        df_intro, info_intro = load_excel_flexible(intro_path)
        assert info_intro["column"] == "Materialnummer"
        assert len(df_intro) == 2
        assert df_intro["Materialnummer"].tolist() == ["0221-0412", "0221-2411"]

    def test_multiple_candidate_columns_matched_short_long_vs_ambiguous(self):
        # Case A: Two columns representing Short and Long of the SAME products -> Not ambiguous
        df_same = pd.DataFrame({
            "Material lang ber.": ["0221-0412", "0221-2411", "2000-0115"],
            "Material lang kompr.": ["221-412", "221-2411", "2000-115"],
            "Description": ["A", "B", "C"],
        })
        info_same = detect_article_column(df_same)
        assert info_same["has_both_formats"] is True
        assert info_same["is_ambiguous"] is False
        assert info_same["short_column"] == "Material lang kompr."
        assert info_same["long_column"] == "Material lang ber."

        # Case B: Two columns with DIFFERENT valid article numbers and generic headers -> Ambiguous
        df_competing = pd.DataFrame({
            "Col_1": ["221-412", "280-101", "2000-115"],
            "Col_2": ["750-103", "787-1212", "855-9910"],
            "Description": ["Row 1", "Row 2", "Row 3"],
        })
        info_competing = detect_article_column(df_competing)
        assert info_competing["is_ambiguous"] is True
        assert len(info_competing["candidates"]) == 2

        # Manual override resolves ambiguity cleanly
        _, info_override = load_excel_flexible(df_competing, article_column="Col_2")
        assert info_override["column"] == "Col_2"
        assert info_override["is_ambiguous"] is False

    def test_leading_zeros_lost_through_excel_formatting(self, tmp_path: Path):
        excel_path = tmp_path / "formatted_numbers.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Item No", "Remark"])
        # Cell with custom number format 0000-0000 and integer value 2210412 (leading 0 dropped)
        ws.append([2210412, "Custom formatted 0221-0412"])
        ws["A2"].number_format = "0000-0000"
        # Cell with General number format and integer value 2212411 (leading 0 dropped)
        ws.append([2212411, "General integer 0221-2411"])
        wb.save(excel_path)

        df_loaded, info = load_excel_flexible(excel_path)
        assert info["column"] == "Item No"
        canonical, short_list, long_list = pcf_tool._extract_row_articles(df_loaded, info)
        assert canonical == ["221-412", "221-2411"]
        assert short_list == ["221-412", "221-2411"]
        assert long_list == ["0221-0412", "0221-2411"]

    def test_mixed_short_and_long_formats_and_empty_rows_and_multi_sheet(self, tmp_path: Path):
        excel_path = tmp_path / "multi_sheet.xlsx"
        wb = openpyxl.Workbook()
        ws_cover = wb.active
        ws_cover.title = "Cover"
        ws_cover.append(["Project Overview", "No articles on this sheet"])

        ws_data = wb.create_sheet("Article_List")
        ws_data.append(["Customer Ref", "WAGO ID", "Unrelated Column"])
        ws_data.append(["REF-1", "221-412", "Keep 1"])
        ws_data.append(["", "", ""])  # Empty row in the middle
        ws_data.append(["REF-2", "0221-2411", "Keep 2"])
        ws_data.append(["REF-3", "2000-0115", "Keep 3"])
        wb.save(excel_path)

        df_loaded, info = load_excel_flexible(excel_path)
        assert info["sheet_name"] == "Article_List"
        assert info["column"] == "WAGO ID"
        assert len(df_loaded) == 4
        assert "Unrelated Column" in df_loaded.columns


# ═══════════════════════════════════════════════════════════════════════════
# 3. API Behavior Tests
# ═══════════════════════════════════════════════════════════════════════════

def _make_submodel_response(pcf_value="0.0116", article="221-412"):
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
                                "value": str(pcf_value),
                            }
                        ],
                    }
                ],
            }
        ],
    }


class TestAPIBehavior:

    def test_successful_pcf_retrieval_short_and_long_resolve_identically(self):
        sample_sm = _make_submodel_response("0.0116", "221-412")
        with patch.object(pcf_tool, "_jget", return_value=sample_sm) as mock_jget:
            res_short = get_pcf_result("221-412")
            res_long  = get_pcf_result("0221-0412")

            assert res_short["pcf"] == pytest.approx(0.0116)
            assert res_short["status"] == STATUS_PCF_AVAILABLE
            assert res_short["article_short"] == "221-412"
            assert res_short["article_long"] == "0221-0412"

            assert res_long["pcf"] == pytest.approx(0.0116)
            assert res_long["status"] == STATUS_PCF_AVAILABLE
            assert res_long["article_short"] == "221-412"
            assert res_long["article_long"] == "0221-0412"

            # Both queried the exact same normalized Fast Path URL
            assert mock_jget.call_count == 2
            url_short = mock_jget.call_args_list[0][0][0]
            url_long  = mock_jget.call_args_list[1][0][0]
            assert url_short == url_long
            assert "https://c1.api.wago.com/smartservice-aas/submodels/" in url_short

            # Wrapper get_pcf() returns float directly
            assert get_pcf("0221-0412") == pytest.approx(0.0116)

    def test_successful_pcf_retrieval_via_registry_fallback(self):
        shell_with_custom_pcf = {
            "result": [
                {
                    "id": "https://wago.com/ids/aas/280-101",
                    "idShort": "280-101",
                    "submodelDescriptors": [
                        {
                            "id": "https://www.wago.com/submodel/280-101/CarbonFootprint",
                        }
                    ],
                }
            ]
        }
        submodel_payload = _make_submodel_response("0.0411", "280-101")
        with patch.object(
            pcf_tool,
            "_jget",
            side_effect=[pcf_tool._NOT_FOUND, shell_with_custom_pcf, submodel_payload],
        ) as mock_jget:
            res = get_pcf_result("0280-0101")
            assert res["pcf"] == pytest.approx(0.0411)
            assert res["status"] == STATUS_PCF_AVAILABLE
            assert res["internal_status"] == INTERNAL_STATUS_PCF_AVAILABLE
            assert mock_jget.call_count == 3

    def test_product_without_available_pcf_returns_not_yet_available(self):
        # Case 1: Product exists in AAS Registry, but has no CarbonFootprint submodel
        shell_without_pcf = {
            "result": [
                {
                    "id": "https://wago.com/ids/aas/855-9910",
                    "idShort": "855-9910",
                    "submodelDescriptors": [
                        {"id": "https://www.wago.com/submodel/855-9910/Nameplate", "idShort": "Nameplate"}
                    ],
                }
            ]
        }
        with patch.object(pcf_tool, "_jget", side_effect=[pcf_tool._NOT_FOUND, shell_without_pcf]):
            res = get_pcf_result("855-9910")
            assert res["pcf"] is None
            assert res["status"] == STATUS_NOT_YET_AVAILABLE
            assert res["internal_status"] == INTERNAL_STATUS_AAS_NO_PCF

        # Case 2: Product does not exist in AAS at all (404 on Fast Path, empty registry result)
        with patch.object(pcf_tool, "_jget", side_effect=[pcf_tool._NOT_FOUND, {"result": []}]):
            res_nf = get_pcf_result("999-999")
            assert res_nf["pcf"] is None
            assert res_nf["status"] == STATUS_NOT_YET_AVAILABLE
            assert res_nf["internal_status"] == INTERNAL_STATUS_NOT_FOUND

    @pytest.mark.parametrize("status_code", [401, 403])
    def test_authentication_failure_is_query_error_not_unavailable(self, status_code: int):
        mock_resp = MagicMock()
        mock_resp.status_code = status_code
        mock_session = MagicMock()
        mock_session.get.return_value = mock_resp

        with patch.object(pcf_tool, "_get_session", return_value=mock_session):
            res = get_pcf_result("221-412")
            assert res["pcf"] is None
            assert res["status"] == STATUS_QUERY_ERROR
            assert res["status"] != STATUS_NOT_YET_AVAILABLE
            assert res["internal_status"] == INTERNAL_STATUS_AUTH_ERROR
            # Fast path auth error must not trigger redundant fallback requests
            assert mock_session.get.call_count == 1

    def test_api_timeout_is_query_error_not_unavailable(self):
        mock_session = MagicMock()
        mock_session.get.side_effect = requests.exceptions.ReadTimeout("Read timed out")

        with patch.object(pcf_tool, "_get_session", return_value=mock_session):
            res = get_pcf_result("221-412")
            assert res["pcf"] is None
            assert res["status"] == STATUS_QUERY_ERROR
            assert res["status"] != STATUS_NOT_YET_AVAILABLE
            assert res["internal_status"] == INTERNAL_STATUS_TIMEOUT

    @pytest.mark.parametrize(
        "unexpected_payload",
        [
            {"unexpected": "schema_changed"},
            {"modelType": "Submodel", "submodelElements": "not_a_list"},
            ["a", "json", "list"],
            "plain text non-json",
        ],
    )
    def test_unexpected_or_changed_api_response_handled_as_query_error(self, unexpected_payload):
        mock_resp = MagicMock()
        mock_resp.status_code = 200
        if isinstance(unexpected_payload, str):
            mock_resp.json.side_effect = ValueError("No JSON object could be decoded")
        else:
            mock_resp.json.return_value = unexpected_payload

        mock_session = MagicMock()
        mock_session.get.return_value = mock_resp

        with patch.object(pcf_tool, "_get_session", return_value=mock_session):
            res = get_pcf_result("221-412")
            assert res["pcf"] is None
            assert res["status"] == STATUS_QUERY_ERROR
            assert res["internal_status"] == INTERNAL_STATUS_INVALID_RESP

    def test_batch_processing_deduplicates_across_article_formats(self, tmp_path: Path):
        input_path  = tmp_path / "batch_dups.xlsx"
        output_path = tmp_path / "batch_dups_out.xlsx"

        df = pd.DataFrame({
            "WAGO No": [
                "221-412",
                "0221-0412",
                "2210412",
                "221-2411",
                "0221-2411",
                "",
            ],
            "Note": ["short 412", "long 412", "numeric 412", "short 2411", "long 2411", "empty"],
        })
        df.to_excel(input_path, index=False)

        queried_articles = []

        def fake_get_pcf_result(art):
            queried_articles.append(art)
            if normalize_article_number(art) == "221-412":
                return {
                    "pcf": 0.0116,
                    "status": STATUS_PCF_AVAILABLE,
                    "internal_status": INTERNAL_STATUS_PCF_AVAILABLE,
                    "article_short": "221-412",
                    "article_long": "0221-0412",
                }
            return {
                "pcf": None,
                "status": STATUS_NOT_YET_AVAILABLE,
                "internal_status": INTERNAL_STATUS_AAS_NO_PCF,
                "article_short": "221-2411",
                "article_long": "0221-2411",
            }

        with patch.object(pcf_tool, "get_pcf_result", side_effect=fake_get_pcf_result):
            summary = main(str(input_path), str(output_path))

        assert summary is not None
        assert summary["total_rows"] == 6
        assert summary["unique_queried"] == 2
        # Only 2 unique products were queried despite 5 rows in different representations
        assert sorted(queried_articles) == ["221-2411", "221-412"]
        assert summary["pcf_available"] == 3
        assert summary["not_yet_available"] == 2
        assert summary["query_errors"] == 0


# ═══════════════════════════════════════════════════════════════════════════
# 4. Excel Export & Customer-Facing Status Simplification Tests
# ═══════════════════════════════════════════════════════════════════════════

class TestExcelExport:

    def test_export_generates_both_article_representations_and_preserves_columns(self, tmp_path: Path):
        input_path  = tmp_path / "export_single_col.xlsx"
        output_path = tmp_path / "export_single_col_out.xlsx"

        df = pd.DataFrame({
            "Customer_ID": ["C1", "C2", "C3"],
            "Artikelnummer": ["221-412", "0221-2411", "855-9910"],
            "Qty": ["5", "10", "15"],
        })
        df.to_excel(input_path, index=False)

        def fake_query(art):
            if art in ("221-412", "221-2411"):
                return {
                    "pcf": 0.0116 if art == "221-412" else 0.025,
                    "status": STATUS_PCF_AVAILABLE,
                    "internal_status": INTERNAL_STATUS_PCF_AVAILABLE,
                }
            return {
                "pcf": None,
                "status": STATUS_NOT_YET_AVAILABLE,
                "internal_status": INTERNAL_STATUS_AAS_NO_PCF,
            }

        with patch.object(pcf_tool, "get_pcf_result", side_effect=fake_query):
            main(str(input_path), str(output_path))

        out_df = pd.read_excel(output_path, dtype=str)
        assert COL_ARTICLE_SHORT in out_df.columns
        assert COL_ARTICLE_LONG in out_df.columns
        assert "Customer_ID" in out_df.columns
        assert "Artikelnummer" in out_df.columns
        assert "Qty" in out_df.columns

        # Original values preserved
        assert out_df["Artikelnummer"].tolist() == ["221-412", "0221-2411", "855-9910"]
        assert out_df["Customer_ID"].tolist() == ["C1", "C2", "C3"]
        assert out_df["Qty"].tolist() == ["5", "10", "15"]

        # Both Short and Long representations generated accurately
        assert out_df[COL_ARTICLE_SHORT].tolist() == ["221-412", "221-2411", "855-9910"]
        assert out_df[COL_ARTICLE_LONG].tolist() == ["0221-0412", "0221-2411", "0855-9910"]

    def test_export_reuses_and_validates_existing_short_and_long_columns_without_duplicates(self):
        df = pd.DataFrame({
            COL_ARTICLE_SHORT: ["221-412", "", "0280-0101"],
            COL_ARTICLE_LONG: ["", "0221-2411", "280-101"],
            "Description": ["Item 1", "Item 2", "Item 3"],
        })
        _, info = load_excel_flexible(df)
        cache = {
            "221-412": {"pcf": 0.0116, "status": STATUS_PCF_AVAILABLE},
            "221-2411": {"pcf": 0.025, "status": STATUS_PCF_AVAILABLE},
            "280-101": {"pcf": None, "status": STATUS_NOT_YET_AVAILABLE},
        }
        out_df = prepare_output_dataframe(df, info, cache)

        # No duplicate columns created
        assert list(out_df.columns).count(COL_ARTICLE_SHORT) == 1
        assert list(out_df.columns).count(COL_ARTICLE_LONG) == 1
        assert list(out_df.columns) == [
            COL_ARTICLE_SHORT,
            COL_ARTICLE_LONG,
            "Description",
            COL_PCF,
            COL_PCF_STATUS,
        ]
        # Missing / swapped entries validated and filled
        assert out_df[COL_ARTICLE_SHORT].tolist() == ["221-412", "221-2411", "280-101"]
        assert out_df[COL_ARTICLE_LONG].tolist() == ["0221-0412", "0221-2411", "0280-0101"]

    def test_missing_pcfs_display_not_yet_available_and_exclude_internal_reasons(self, tmp_path: Path):
        input_path  = tmp_path / "status_check.xlsx"
        output_path = tmp_path / "status_check_out.xlsx"

        df = pd.DataFrame({
            "Article Number": ["221-412", "855-9910", "999-999", "280-101"],
        })
        df.to_excel(input_path, index=False)

        mock_results = {
            "221-412": {
                "pcf": 0.0116,
                "status": STATUS_PCF_AVAILABLE,
                "internal_status": INTERNAL_STATUS_PCF_AVAILABLE,
            },
            "855-9910": {
                "pcf": None,
                "status": STATUS_NOT_YET_AVAILABLE,
                "internal_status": INTERNAL_STATUS_AAS_NO_PCF,
            },
            "999-999": {
                "pcf": None,
                "status": STATUS_NOT_YET_AVAILABLE,
                "internal_status": INTERNAL_STATUS_NOT_FOUND,
            },
            "280-101": {
                "pcf": None,
                "status": STATUS_QUERY_ERROR,
                "internal_status": INTERNAL_STATUS_TIMEOUT,
            },
        }

        with patch.object(pcf_tool, "get_pcf_result", side_effect=lambda a: mock_results[a]):
            main(str(input_path), str(output_path))

        out_df = pd.read_excel(output_path, dtype=str).fillna("")

        # Available PCF
        assert float(out_df.loc[0, COL_PCF]) == pytest.approx(0.0116)
        assert out_df.loc[0, COL_PCF_STATUS] == STATUS_PCF_AVAILABLE

        # Missing PCFs (both AAS-no-PCF and Not-found) display exact wording "Not yet available"
        assert out_df.loc[1, COL_PCF] == "Not yet available"
        assert out_df.loc[1, COL_PCF_STATUS] == "Not yet available"
        assert out_df.loc[2, COL_PCF] == "Not yet available"
        assert out_df.loc[2, COL_PCF_STATUS] == "Not yet available"

        # Query error is NOT confused with "Not yet available"
        assert out_df.loc[3, COL_PCF] == ""
        assert out_df.loc[3, COL_PCF_STATUS] == STATUS_QUERY_ERROR

        # Internal diagnostic reasons and legacy "Not available" never appear anywhere in exported Excel
        all_exported_cells = " | ".join(
            str(v) for col in out_df.columns for v in out_df[col].tolist()
        )
        assert "AAS available - PCF not available" not in all_exported_cells
        assert "Article not found" not in all_exported_cells
        assert "Request timeout" not in all_exported_cells
        assert "not available" not in all_exported_cells.lower().replace("not yet available", "")


# ═══════════════════════════════════════════════════════════════════════════
# 5. Offline Mode (Embedded Q4 2026 Snapshot) Tests
# ═══════════════════════════════════════════════════════════════════════════

import offline_pcf_data
from offline_pcf_data import (
    EMBEDDED_PCF_DATA,
    OFFLINE_NOTICE_TEXT,
    OFFLINE_NOTICE_TITLE,
    OFFLINE_PORTAL_URL,
    get_offline_pcf,
    get_offline_pcf_result,
    validate_and_index_pcf_records,
)


class TestOfflineMode:

    def test_embedded_dataset_completeness_and_correct_values(self):
        # Dataset contains all 9,635 consolidated Q4 2026 records embedded in Python
        assert len(EMBEDDED_PCF_DATA) == 9635
        assert get_offline_pcf("221-412") == pytest.approx(0.0116)
        assert get_offline_pcf("280-101") == pytest.approx(0.0411)
        assert get_offline_pcf("2000-115") == pytest.approx(0.001)
        assert get_offline_pcf("221-2411") == pytest.approx(0.01847692085402356)
        assert get_offline_pcf("734-001/K010-2937/000-100") == pytest.approx(0.08532127818568026)

    @pytest.mark.parametrize(
        "short_art, long_art",
        [
            ("221-412", "0221-0412"),
            ("221-2411", "0221-2411"),
            ("280-101", "0280-0101"),
            ("2000-115", "2000-0115"),
            ("734-001/K010-2937/000-100", "0734-0001/K010-2937/0000-0100"),
            ("734-001/K000-002", "0734-0001/K000-002"),
        ],
    )
    def test_both_article_formats_resolve_to_identical_offline_records(self, short_art: str, long_art: str):
        res_s = get_offline_pcf_result(short_art)
        res_l = get_offline_pcf_result(long_art)
        assert res_s["status"] == STATUS_PCF_AVAILABLE
        assert res_l["status"] == STATUS_PCF_AVAILABLE
        assert res_s["pcf"] == pytest.approx(res_l["pcf"])
        assert res_s["article_short"] == res_l["article_short"]
        assert res_s["article_long"] == res_l["article_long"]

    def test_offline_missing_pcf_displays_not_yet_available(self):
        res = get_offline_pcf_result("999-999")
        assert res["pcf"] is None
        assert res["status"] == STATUS_NOT_YET_AVAILABLE
        assert res["internal_status"] == INTERNAL_STATUS_NOT_FOUND

    def test_offline_mode_never_makes_api_calls(self, tmp_path: Path):
        input_path  = tmp_path / "offline_input.xlsx"
        output_path = tmp_path / "offline_output.xlsx"

        df = pd.DataFrame({
            "Article Number": ["221-412", "0221-0412", "221-2411", "0221-2411", "999-999"],
            "Description": ["Short 412", "Long 412", "Short 2411", "Long 2411", "Missing"],
        })
        df.to_excel(input_path, index=False)

        with patch.object(pcf_tool, "_get_session", side_effect=AssertionError("Network session must not be used in Offline Mode!")), \
             patch.object(pcf_tool, "_jget", side_effect=AssertionError("HTTP GET must not be called in Offline Mode!")), \
             patch.object(pcf_tool, "get_pcf_result", side_effect=AssertionError("Online get_pcf_result must not be called in Offline Mode!")):
            summary = main(str(input_path), str(output_path), mode=pcf_tool.MODE_OFFLINE)

        assert summary is not None
        assert summary["mode"] == pcf_tool.MODE_OFFLINE
        assert summary["total_rows"] == 5
        assert summary["unique_queried"] == 3  # 221-412, 221-2411, 999-999 deduplicated
        assert summary["pcf_available"] == 4
        assert summary["not_yet_available"] == 1
        assert summary["query_errors"] == 0

    def test_online_and_offline_generate_identical_excel_output_structure(self, tmp_path: Path):
        input_path   = tmp_path / "struct_in.xlsx"
        out_online   = tmp_path / "struct_online.xlsx"
        out_offline  = tmp_path / "struct_offline.xlsx"

        df = pd.DataFrame({
            "Pos": ["1", "2", "3"],
            "WAGO-Nr": ["280-101", "0280-0312", "999-999"],
            "Note": ["A", "B", "C"],
        })
        df.to_excel(input_path, index=False)

        # Offline run
        main(str(input_path), str(out_offline), mode=pcf_tool.MODE_OFFLINE)

        # Online run (mocked with matching values)
        online_mock_map = {
            "280-101": {"pcf": 0.0411, "status": STATUS_PCF_AVAILABLE, "internal_status": INTERNAL_STATUS_PCF_AVAILABLE},
            "280-312": {"pcf": 0.0241, "status": STATUS_PCF_AVAILABLE, "internal_status": INTERNAL_STATUS_PCF_AVAILABLE},
            "999-999": {"pcf": None, "status": STATUS_NOT_YET_AVAILABLE, "internal_status": INTERNAL_STATUS_NOT_FOUND},
        }
        with patch.object(pcf_tool, "get_pcf_result", side_effect=lambda a: online_mock_map[a]):
            main(str(input_path), str(out_online), mode=pcf_tool.MODE_ONLINE)

        df_on  = pd.read_excel(out_online, dtype=str).fillna("")
        df_off = pd.read_excel(out_offline, dtype=str).fillna("")

        assert list(df_on.columns) == list(df_off.columns)
        assert list(df_off.columns) == [
            "Pos",
            "WAGO-Nr",
            COL_ARTICLE_SHORT,
            COL_ARTICLE_LONG,
            "Note",
            COL_PCF,
            COL_PCF_STATUS,
        ]
        assert df_off[COL_ARTICLE_SHORT].tolist() == ["280-101", "280-312", "999-999"]
        assert df_off[COL_ARTICLE_LONG].tolist() == ["0280-0101", "0280-0312", "0999-0999"]
        assert float(df_off.loc[0, COL_PCF]) == pytest.approx(0.0411)
        assert float(df_off.loc[1, COL_PCF]) == pytest.approx(0.0241)
        assert df_off.loc[2, COL_PCF] == "Not yet available"
        assert df_off.loc[2, COL_PCF_STATUS] == "Not yet available"

    def test_conflicting_source_records_are_detected_not_silently_resolved(self):
        # Identical values across short/long formats succeed without conflict
        clean_idx = validate_and_index_pcf_records([
            ("221-412", 0.0116),
            ("0221-0412", 0.0116),
            ("280-101", 0.0411),
        ])
        assert clean_idx == {"221-412": 0.0116, "280-101": 0.0411}

        # Conflicting PCF values for the same product raise ValueError
        with pytest.raises(ValueError, match="Conflicting PCF values"):
            validate_and_index_pcf_records([
                ("221-412", 0.0116),
                ("0221-0412", 0.0999),
            ])

    def test_ui_q4_2026_notice_displayed_only_in_offline_mode(self):
        app = pcf_tool.PCFQueryApp()
        try:
            # Default is Online Mode, notice is hidden
            assert app.get_mode() == pcf_tool.MODE_ONLINE
            assert app._offline_notice_visible is False
            assert app._frm_offline_notice.winfo_manager() == ""

            # Switch to Offline Mode -> notice is immediately shown with required German text & link
            app.set_mode(pcf_tool.MODE_OFFLINE)
            assert app.get_mode() == pcf_tool.MODE_OFFLINE
            assert app._offline_notice_visible is True
            assert app._frm_offline_notice.winfo_manager() == "pack"
            assert app._lbl_offline_title.cget("text") == OFFLINE_NOTICE_TITLE
            assert app._lbl_offline_text.cget("text") == OFFLINE_NOTICE_TEXT
            assert app._lbl_offline_link.cget("text") == OFFLINE_PORTAL_URL
            assert "Offline-Modus – Datenstand Q4 2026" in app._lbl_offline_title.cget("text")
            assert "https://aas.wago.com/" == app._lbl_offline_link.cget("text")

            # Switch back to Online Mode -> notice is hidden immediately
            app.set_mode(pcf_tool.MODE_ONLINE)
            assert app.get_mode() == pcf_tool.MODE_ONLINE
            assert app._offline_notice_visible is False
            assert app._frm_offline_notice.winfo_manager() == ""
        finally:
            app.destroy()

    def test_compiled_executable_works_without_reference_excel_files(self, tmp_path: Path):
        """
        Verifies that the compiled dist/WAGO_PCF_Query.exe runs in Offline Mode
        from an isolated directory when the original reference Excel files in dist/
        are temporarily moved/renamed so they cannot be accessed at runtime.
        """
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
                "0221-2411",
                "734-001/K010-2937/000-100",
                "999-999",
            ],
            "Comment": ["short 412", "long 412", "short 2411", "long 2411", "Q4 2026 item", "unknown"],
        })
        df_in.to_excel(test_in, index=False)

        try:
            # Make all reference Excel files in dist/ unavailable
            for rf in ref_files:
                if rf.exists():
                    hidden_path = tmp_path / f"_hidden_{rf.name}"
                    rf.rename(hidden_path)
                    moved_pairs.append((hidden_path, rf))
                    assert not rf.exists()

            proc = subprocess.run(
                [str(exe_path), "--batch", str(test_in), str(test_out), "--mode", "offline"],
                cwd=str(tmp_path),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=60,
            )
            assert proc.returncode == 0, f"Executable failed: {proc.stderr}\n{proc.stdout}"
            assert test_out.exists(), "Executable did not create output Excel file"
        finally:
            # Always restore reference Excel files to their original location
            for hidden_path, orig_path in moved_pairs:
                if hidden_path.exists():
                    hidden_path.rename(orig_path)

        df_out = pd.read_excel(test_out, dtype=str).fillna("")
        assert df_out[COL_ARTICLE_SHORT].tolist() == [
            "221-412",
            "221-412",
            "221-2411",
            "221-2411",
            "734-001/K010-2937/000-100",
            "999-999",
        ]
        assert df_out[COL_ARTICLE_LONG].tolist() == [
            "0221-0412",
            "0221-0412",
            "0221-2411",
            "0221-2411",
            "0734-0001/K010-2937/0000-0100",
            "0999-0999",
        ]
        assert float(df_out.loc[0, COL_PCF]) == pytest.approx(0.0116)
        assert float(df_out.loc[1, COL_PCF]) == pytest.approx(0.0116)
        assert float(df_out.loc[2, COL_PCF]) == pytest.approx(0.01847692085402356)
        assert float(df_out.loc[3, COL_PCF]) == pytest.approx(0.01847692085402356)
        assert float(df_out.loc[4, COL_PCF]) == pytest.approx(0.08532127818568026)
        assert df_out.loc[5, COL_PCF] == "Not yet available"
        assert df_out.loc[5, COL_PCF_STATUS] == "Not yet available"


