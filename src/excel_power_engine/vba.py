from __future__ import annotations


def extract_vba(path: str):
    try:
        from oletools.olevba import VBA_Parser
    except ImportError as exc:
        raise RuntimeError("Install oletools with: pip install -e \".[security]\"") from exc

    parser = VBA_Parser(path)
    modules = []
    try:
        if not parser.detect_vba_macros():
            return modules
        for filename, stream_path, vba_filename, code in parser.extract_macros():
            modules.append({
                "filename": filename,
                "stream_path": stream_path,
                "module": vba_filename,
                "code": code,
            })
        return modules
    finally:
        parser.close()


def has_vba(path: str) -> bool:
    from .inspect import inspect_workbook
    return inspect_workbook(path).has_vba
