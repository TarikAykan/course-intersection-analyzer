from __future__ import annotations

from pathlib import Path

import pandas as pd


def export_results_to_excel(df: pd.DataFrame, output_path: str) -> None:
    path = Path(output_path)
    if path.suffix.lower() not in {".xlsx", ".xls"}:
        raise ValueError("Cikti dosya uzantisi .xlsx veya .xls olmali.")

    try:
        if path.suffix.lower() == ".xlsx":
            df.to_excel(path, index=False, engine="openpyxl")
        else:
            df.to_excel(path, index=False)
    except Exception as exc:
        raise ValueError(f"Sonuc Excel'e aktarilamadi: {exc}") from exc
