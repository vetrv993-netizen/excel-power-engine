from __future__ import annotations

from typing import Any


def describe_dataframe(df: Any) -> dict[str, Any]:
    return {
        "rows": int(len(df)),
        "columns": int(len(df.columns)),
        "columns_names": [str(c) for c in df.columns],
        "nulls": {str(k): int(v) for k, v in df.isna().sum().items()},
        "dtypes": {str(k): str(v) for k, v in df.dtypes.items()},
    }


def to_polars(df: Any):
    try:
        import polars as pl
    except ImportError as exc:
        raise RuntimeError("polars is not installed. Install with: pip install -e \".[fast]\"") from exc
    return pl.from_pandas(df)
