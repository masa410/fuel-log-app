"""
既存の「ソリオ燃費早見表.xlsm」へ渡す「最新データのCSV」を書き出すモジュール。

【2026-10 の仕様変更】
- Excelファイル本体へは、このモジュールから書き込まない。Excel側のマクロ（modFuelSync）が、
  このCSVを読んで自分のシートを更新する（Excelが開いたままでも反映できる。xlsmを外部から
  書き換えるとボタンなどが壊れる心配もなくなる）。
- スプレッドシート（クラウド）を唯一の正とする。CSVは全記録を日付・ID順に書き出す。
- CSVの置き場所は「ダウンロード」フォルダ。Excel側が読み終えると自動で削除する。
- このPCに local_settings.py があるときだけ動作する（クラウド上では何もしない）。

CSVの列: id, record_date(YYYY-MM-DD), odometer_km, fuel_liters, fuel_unit_price, note
文字コード: UTF-8（BOM付き）
"""

from __future__ import annotations

import csv
import os
from datetime import datetime

import pandas as pd

try:
    from local_settings import EXCEL_PATH
except ImportError:
    EXCEL_PATH = None

try:
    from local_settings import EXPORT_DIR
except ImportError:
    EXPORT_DIR = os.path.join(os.path.expanduser("~"), "Downloads")

EXPORT_FILENAME = "燃費_最新データ.csv"
CSV_HEADER = ["id", "record_date", "odometer_km", "fuel_liters", "fuel_unit_price", "note"]


def is_available() -> bool:
    """このPC（local_settings.pyがある環境）でだけTrue。クラウド上ではFalse。"""
    return bool(EXCEL_PATH)


def export_path() -> str:
    return os.path.join(EXPORT_DIR, EXPORT_FILENAME)


def _num(v) -> str:
    if v is None or pd.isna(v):
        return ""
    f = float(v)
    return str(int(f)) if f == int(f) else repr(f)


def _status_file_path():
    if not EXCEL_PATH:
        return None
    return os.path.join(os.path.dirname(EXCEL_PATH), "sync_status.txt")


def _write_status_file(status: str, count: int, last_date: str):
    path = _status_file_path()
    if not path:
        return
    lines = [
        f"checked_at={datetime.now().isoformat(timespec='seconds')}",
        f"status={status}",
        f"exported={count}",
        f"cloud_last_date={last_date}",
    ]
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


def export_records_for_excel(records_df: pd.DataFrame) -> dict:
    """全記録をCSVに書き出す。戻り値: {"status": "not_configured"|"exported"|"error", "count": 件数, "path": パス}"""
    if not is_available() or records_df is None:
        return {"status": "not_configured", "count": 0, "path": ""}

    df = records_df.sort_values(["record_date", "id"]).reset_index(drop=True)
    path = export_path()
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(CSV_HEADER)
            for _, r in df.iterrows():
                note = r.get("note")
                w.writerow([
                    int(r["id"]),
                    r["record_date"].strftime("%Y-%m-%d"),
                    _num(r["odometer_km"]),
                    _num(r["fuel_liters"]),
                    _num(r.get("fuel_unit_price")),
                    note if isinstance(note, str) else "",
                ])
        os.replace(tmp, path)
    except OSError as e:
        _write_status_file("error", 0, "")
        return {"status": "error", "count": 0, "path": path, "error": str(e)}

    last = df["record_date"].max().strftime("%Y-%m-%d") if len(df) else ""
    _write_status_file("exported", len(df), last)
    return {"status": "exported", "count": len(df), "path": path}
