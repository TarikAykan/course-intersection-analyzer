from pathlib import Path

import pandas as pd


def main() -> None:
    base_dir = Path(__file__).resolve().parent
    csv_path = base_dir / "ornek_veri.csv"
    xlsx_path = base_dir / "ornek_veri.xlsx"

    df = pd.read_csv(csv_path)
    df.to_excel(xlsx_path, index=False, engine="openpyxl")
    print(f"Ornek Excel dosyasi olusturuldu: {xlsx_path}")


if __name__ == "__main__":
    main()
