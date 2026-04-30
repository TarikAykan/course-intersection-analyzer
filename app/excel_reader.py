from __future__ import annotations

from pathlib import Path

import pandas as pd
import requests


SUPPORTED_EXTENSIONS = {".xlsx", ".xls"}


def read_excel_file(file_path: str) -> pd.DataFrame:
    """Excel dosyasini okuyup DataFrame olarak dondurur."""
    path = Path(file_path)
    if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise ValueError("Lutfen .xlsx veya .xls uzantili bir dosya secin.")

    try:
        if path.suffix.lower() == ".xlsx":
            df = pd.read_excel(path, engine="openpyxl")
        else:
            df = pd.read_excel(path)
    except ImportError as exc:
        raise ValueError(
            ".xls dosyasi okumak icin ek olarak 'xlrd' paketini kurmaniz gerekebilir."
        ) from exc
    except Exception as exc:
        raise ValueError(f"Excel dosyasi okunamadi: {exc}") from exc

    if df.empty:
        raise ValueError("Excel dosyasi bos gorunuyor.")

    return df


def get_columns(df: pd.DataFrame) -> list[str]:
    return [str(col) for col in df.columns]


def read_api_data(api_url: str, timeout: int = 20) -> pd.DataFrame:
    """API'den JSON veri cekip DataFrame olarak dondurur."""
    if not api_url.strip():
        raise ValueError("API URL bos birakilamaz.")

    try:
        response = requests.get(api_url.strip(), timeout=timeout)
        response.raise_for_status()
    except requests.RequestException as exc:
        raise ValueError(f"API verisi alinamadi: {exc}") from exc

    try:
        payload = response.json()
    except ValueError as exc:
        raise ValueError("API yaniti gecerli JSON formatinda degil.") from exc

    if isinstance(payload, dict):
        if "data" in payload and isinstance(payload["data"], list):
            records = payload["data"]
        else:
            raise ValueError(
                "API JSON yapisi uygun degil. Beklenen yapi: liste veya {'data': [...]}."
            )
    elif isinstance(payload, list):
        records = payload
    else:
        raise ValueError("API JSON yapisi uygun degil. Liste tipinde veri bekleniyor.")

    df = pd.DataFrame(records)
    if df.empty:
        raise ValueError("API'den gelen veri bos.")

    return df
