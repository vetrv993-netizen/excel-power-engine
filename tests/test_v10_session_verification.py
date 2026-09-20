from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl import Workbook

from excel_power_engine.orchestrator import run_workspace
from excel_power_engine.session import WorkbookSession
from excel_power_engine.verification import verify_workbook_change


def make_book(path: Path) -> None:
    wb = Workbook(); ws = wb.active; ws.title = "Data"
    ws.append(["name", "code", "record", "id", "value"])
    ws.append(["Ahmed", "001", "10", "100", 1])
    ws.append(["Mona", "002", "11", "101", 2])
    wb.save(path)


def test_workspace_search_regression_and_targets(tmp_path):
    source, output = tmp_path / "source.xlsx", tmp_path / "out.xlsx"; make_book(source)
    session = WorkbookSession.open(source)
    result = run_workspace(session, [{"kind": "search", "sheet": "Data", "payload": {"term": "Ahmed", "result_id": "people"}}], output, open_after=False)
    assert result["applied_operations"][0]["target_set"]["targets"] == {"Data": ["A2"]}
    assert output.exists() and Path(result["manifest"]).exists()


def test_search_target_to_bulk_and_verification(tmp_path):
    source, output = tmp_path / "source.xlsx", tmp_path / "out.xlsx"; make_book(source)
    result = run_workspace(source, [
        {"kind": "search", "sheet": "Data", "payload": {"term": "Ahmed", "result_id": "people"}},
        {"kind": "bulk", "sheet": "Data", "payload": {"data": "Changed", "start_cell": "A1"}, "target_set": "people"},
    ], output, open_after=False)
    assert result["applied_operations"][-1]["verification"]["status"] == "PASS"
    assert result["applied_operations"][-1]["verification"]["target_values_verified"]


def test_keyed_bulk_rejects_duplicates_and_missing(tmp_path):
    source, output = tmp_path / "source.xlsx", tmp_path / "out.xlsx"; make_book(source)
    operation = {"kind": "keyed-bulk", "sheet": "Data", "payload": {
        "target_key_column": "B", "source_key_field": "code", "mapping": {"E": "value"},
        "source_rows": [{"code": "001", "value": "9"}, {"code": "999", "value": "8"}], "missing_key_policy": "reject"}}
    try:
        run_workspace(source, [operation], output, open_after=False)
    except ValueError as exc:
        assert "rejected keys" in str(exc)
    else:
        raise AssertionError("missing keyed records must reject the operation")


def test_verification_gate_rejects_unexpected_part_change(tmp_path):
    source, output = tmp_path / "source.xlsx", tmp_path / "out.xlsx"; make_book(source)
    with ZipFile(source, "r") as source_zip, ZipFile(output, "w", ZIP_DEFLATED) as output_zip:
        for item in source_zip.infolist():
            content = source_zip.read(item.filename)
            if item.filename == "docProps/core.xml":
                content = content.replace(b"</cp:coreProperties>", b"<dc:subject>changed</dc:subject></cp:coreProperties>")
            output_zip.writestr(item, content)
    result = verify_workbook_change(source, output, allowed_changed_parts={"xl/worksheets/sheet1.xml"})
    assert not result.accepted
    assert result.unexpected_changed_parts == ["docProps/core.xml"]
