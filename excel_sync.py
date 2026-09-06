"""
既存の「ソリオ燃費早見表.xlsm」へ、自宅PC使用時のみ自動で記録を反映するモジュール。

- このファイルがこのPC上に見つかる場合だけ動作する（クラウド上では何もしない＝安全）。
- Excelが開いていて書き込めない場合は、エラーにせず黙ってスキップする。
- 書き込み前に必ずバックアップを取る（誤って壊した場合に復元できるように）。
- B/C/D/F列（日付・給油量・メーター値・単価）だけを値として書き込み、
  E/G/I列（燃費・金額・区間距離）とH列（オイル交換時期の警告）は、
  直前の行の数式をそのままコピーして再現する（シートの既存の作り方に合わせる）。
"""

from __future__ import annotations

import glob
import os
import re
import shutil
from datetime import datetime

import openpyxl
import pandas as pd

EXCEL_PATH = os.path.join(
    os.path.expanduser("~"), "OneDrive", "デスクトップ", "真紀フォルダー",
    "家計管理関連", "ソリオ燃費早見表.xlsm",
)
SHEET_NAME = "燃費確認表"
BACKUP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "excel_backups")
MAX_BACKUPS = 10

COL_DATE = 2   # B
COL_FUEL = 3   # C
COL_ODO = 4    # D
COL_EFF = 5    # E (数式)
COL_PRICE = 6  # F
COL_COST = 7   # G (数式)
COL_NOTE = 8   # H (通常はオイル交換警告の数式。メモがあれば手入力で上書き)
COL_DIST = 9   # I (数式)

CELL_REF_RE = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")


def is_available() -> bool:
    return os.path.exists(EXCEL_PATH)


def _backup():
    os.makedirs(BACKUP_DIR, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = os.path.join(BACKUP_DIR, f"ソリオ燃費早見表_backup_{stamp}.xlsm")
    shutil.copy2(EXCEL_PATH, dest)
    backups = sorted(glob.glob(os.path.join(BACKUP_DIR, "ソリオ燃費早見表_backup_*.xlsm")))
    for old in backups[:-MAX_BACKUPS]:
        os.remove(old)


def _last_used_row(ws) -> int:
    for r in range(ws.max_row, 1, -1):
        if ws.cell(row=r, column=COL_DATE).value is not None:
            return r
    return 1


def _existing_dates(ws, last_row) -> set:
    dates = set()
    for r in range(2, last_row + 1):
        v = ws.cell(row=r, column=COL_DATE).value
        if hasattr(v, "strftime"):
            dates.add(v.strftime("%Y-%m-%d"))
    return dates


def _shift_formula(formula: str, from_row: int, to_row: int) -> str:
    delta = to_row - from_row

    def repl(m):
        col_abs, col, row_abs, row = m.groups()
        if row_abs == "$":
            return m.group(0)
        return f"{col_abs}{col}{row_abs}{int(row) + delta}"

    return CELL_REF_RE.sub(repl, formula)


def _find_formula_template(ws, last_row, col) -> str | None:
    for r in range(last_row, 1, -1):
        v = ws.cell(row=r, column=col).value
        if isinstance(v, str) and v.startswith("="):
            return v
    return None


def sync_records_to_excel(records_df: pd.DataFrame):
    """Googleスプレッドシート上の記録のうち、Excelにまだ無い日付の分を追記する。
    戻り値: 追加した件数。ファイルが無い/開けない等で実行できなかった場合は None。
    """
    if not is_available() or records_df is None or records_df.empty:
        return None

    try:
        _backup()
        wb = openpyxl.load_workbook(EXCEL_PATH, data_only=False, keep_vba=True)
    except PermissionError:
        return None

    ws = wb[SHEET_NAME]
    last_row = _last_used_row(ws)
    existing = _existing_dates(ws, last_row)

    to_add = records_df[~records_df["record_date"].dt.strftime("%Y-%m-%d").isin(existing)]
    to_add = to_add.sort_values("record_date")
    if to_add.empty:
        return 0

    e_template = _find_formula_template(ws, last_row, COL_EFF) if last_row > 1 else None
    g_template = _find_formula_template(ws, last_row, COL_COST) if last_row > 1 else None
    i_template = _find_formula_template(ws, last_row, COL_DIST) if last_row > 1 else None
    h_template = _find_formula_template(ws, last_row, COL_NOTE) if last_row > 1 else None
    template_row = last_row

    row = last_row
    added = 0
    for _, rec in to_add.iterrows():
        row += 1
        ws.cell(row=row, column=COL_DATE, value=rec["record_date"].to_pydatetime())
        ws.cell(row=row, column=COL_FUEL, value=float(rec["fuel_liters"]))
        ws.cell(row=row, column=COL_ODO, value=float(rec["odometer_km"]))
        if pd.notna(rec.get("fuel_unit_price")):
            ws.cell(row=row, column=COL_PRICE, value=float(rec["fuel_unit_price"]))

        if e_template:
            ws.cell(row=row, column=COL_EFF, value=_shift_formula(e_template, template_row, row))
        if g_template:
            ws.cell(row=row, column=COL_COST, value=_shift_formula(g_template, template_row, row))
        if i_template:
            ws.cell(row=row, column=COL_DIST, value=_shift_formula(i_template, template_row, row))

        note = rec.get("note")
        if isinstance(note, str) and note.strip():
            ws.cell(row=row, column=COL_NOTE, value=note)
        elif h_template:
            ws.cell(row=row, column=COL_NOTE, value=_shift_formula(h_template, template_row, row))

        added += 1

    wb.save(EXCEL_PATH)
    return added
