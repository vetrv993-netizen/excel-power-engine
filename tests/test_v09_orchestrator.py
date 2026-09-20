from pathlib import Path
import json
from openpyxl import Workbook
from excel_power_engine.orchestrator import preview_workspace, validate_workspace_operations, rollback_workspace, run_workspace


def make_book(path: Path):
    wb=Workbook(); ws=wb.active; ws.title="بيانات"
    ws.append(["الاسم","القيمة"])
    ws.append(["أحمد",1]); ws.append(["محمد",2]); ws.append(["أحمدية",3])
    wb.save(path)


def test_validate_workspace_bulk_and_search():
    ops=validate_workspace_operations([
        {"kind":"search","sheet":"بيانات","payload":{"term":"أحمد"}},
        {"kind":"bulk","sheet":"بيانات","payload":{"start_cell":"B2","data":"9\n8"}},
        {"kind":"format","sheet":"بيانات","ranges":["A2:B3"],"payload":{"fill":"#FFF2CC"}},
    ])
    assert len(ops)==3


def test_preview_search_and_bulk(tmp_path):
    src=tmp_path/"book.xlsx"; make_book(src)
    result=preview_workspace(src,[
        {"kind":"search","sheet":"بيانات","payload":{"term":"أحمد"}},
        {"kind":"bulk","sheet":"بيانات","payload":{"start_cell":"B2","data":"9\n8"}},
    ])
    assert result["operation_count"]==2
    assert result["search_hits"]
    assert result["operations"][1]["change_count"]==2


def test_validate_rejects_unknown_operation():
    try:
        validate_workspace_operations([{"kind":"unknown"}])
    except ValueError:
        return
    raise AssertionError("unknown operation was accepted")


def test_run_workspace_bulk_and_rollback(tmp_path):
    src=tmp_path/"source.xlsx"; out=tmp_path/"out.xlsx"; make_book(src)
    result=run_workspace(src,[{"kind":"bulk","sheet":"بيانات","payload":{"start_cell":"B2","data":"9\n8"}}],out,backup=True,open_after=False)
    assert result["rollback_available"]
    assert out.exists()
    rb=rollback_workspace(out)
    assert rb["rolled_back"]
