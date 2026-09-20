from pathlib import Path

from excel_power_engine.native_excel import NativeExcelBridge


def test_v043_staging_path_uses_xlsm_suffix(tmp_path):
    output = tmp_path / "result.xlsm"
    staging = NativeExcelBridge._prepare_staging_path(output)
    assert staging.suffix.lower() == ".xlsm"
    assert "excel_native_tmp" in staging.name


def test_v043_commit_staging(tmp_path):
    staging = tmp_path / "result.excel_native_tmp.xlsm"
    output = tmp_path / "result.xlsm"
    staging.write_bytes(b"ok")
    NativeExcelBridge._commit_staging(staging, output)
    assert output.read_bytes() == b"ok"
    assert not staging.exists()
