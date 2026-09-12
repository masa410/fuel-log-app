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
from datetime import datetime, timedelta

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment

try:
    from local_settings import EXCEL_PATH
except ImportError:
    EXCEL_PATH = None
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
    return bool(EXCEL_PATH) and os.path.exists(EXCEL_PATH)


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


def _cell_date_str(v) -> str:
    """日付セルの値を"YYYY-MM-DD"文字列にする。
    大半の行はdatetimeで入っているが、シートの一部の古い行(2〜4行目)は
    日付書式が適用されておらずExcelのシリアル値(数値)のまま保存されているため、
    その場合も変換する（import_from_excel.pyのto_date()と同じ換算式）。
    """
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, (int, float)):
        return (datetime(1899, 12, 30) + timedelta(days=v)).strftime("%Y-%m-%d")
    return ""


def _existing_dates(ws, last_row) -> set:
    dates = set()
    for r in range(2, last_row + 1):
        v = ws.cell(row=r, column=COL_DATE).value
        s = _cell_date_str(v)
        if s:
            dates.add(s)
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


def _status_file_path():
    if not EXCEL_PATH:
        return None
    return os.path.join(os.path.dirname(EXCEL_PATH), "sync_status.txt")


def _write_status_file(status: str, added: int, cloud_last_date: str = "", excel_last_date: str = ""):
    path = _status_file_path()
    if not path:
        return
    up_to_date = "true" if (status == "synced" and cloud_last_date == excel_last_date) else "false"
    lines = [
        f"checked_at={datetime.now().isoformat(timespec='seconds')}",
        f"status={status}",
        f"added={added}",
        f"cloud_last_date={cloud_last_date}",
        f"excel_last_date={excel_last_date}",
        f"up_to_date={up_to_date}",
    ]
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


def sync_records_to_excel(records_df: pd.DataFrame):
    """Googleスプレッドシート上の記録のうち、Excelにまだ無い日付の分を追記する。
    戻り値: {"status": "not_configured"|"locked"|"synced", "added": 追加件数} の dict。
    """
    if not is_available() or records_df is None or records_df.empty:
        return {"status": "not_configured", "added": 0}

    cloud_last_date = records_df["record_date"].max().strftime("%Y-%m-%d") if not records_df.empty else ""

    try:
        _backup()
        wb = openpyxl.load_workbook(EXCEL_PATH, data_only=False, keep_vba=True)
    except PermissionError:
        _write_status_file("locked", 0, cloud_last_date)
        return {"status": "locked", "added": 0}

    ws = wb[SHEET_NAME]
    last_row = _last_used_row(ws)
    existing = _existing_dates(ws, last_row)
    excel_last_date = _cell_date_str(ws.cell(row=last_row, column=COL_DATE).value)

    to_add = records_df[~records_df["record_date"].dt.strftime("%Y-%m-%d").isin(existing)]
    to_add = to_add.sort_values("record_date")
    if to_add.empty:
        _write_status_file("synced", 0, cloud_last_date, excel_last_date)
        return {"status": "synced", "added": 0}

    e_template = _find_formula_template(ws, last_row, COL_EFF) if last_row > 1 else None
    g_template = _find_formula_template(ws, last_row, COL_COST) if last_row > 1 else None
    i_template = _find_formula_template(ws, last_row, COL_DIST) if last_row > 1 else None
    h_template = _find_formula_template(ws, last_row, COL_NOTE) if last_row > 1 else None
    template_row = last_row

    # B/C/D/F列(日付・給油量・メーター値・単価)は値だけを書くセルのため数式は無いが、
    # 単位付きの表示形式("m月d日"、"9.26ℓ"等)は前の行からコピーしないとデフォルト書式に
    # 戻ってしまう(2026-09-12に実際に発生: 日付がyyyy-mm-dd表記になり単位も消えた)。
    date_fmt = ws.cell(row=last_row, column=COL_DATE).number_format if last_row > 1 else None
    fuel_fmt = ws.cell(row=last_row, column=COL_FUEL).number_format if last_row > 1 else None
    odo_fmt = ws.cell(row=last_row, column=COL_ODO).number_format if last_row > 1 else None
    price_fmt = ws.cell(row=last_row, column=COL_PRICE).number_format if last_row > 1 else None

    row = last_row
    added = 0
    for _, rec in to_add.iterrows():
        row += 1
        date_cell = ws.cell(row=row, column=COL_DATE, value=rec["record_date"].to_pydatetime())
        fuel_cell = ws.cell(row=row, column=COL_FUEL, value=float(rec["fuel_liters"]))
        odo_cell = ws.cell(row=row, column=COL_ODO, value=float(rec["odometer_km"]))
        # 表示形式に加えて中央揃えも明示する。openpyxlは新規に値を書いたセルへ
        # 既存の見た目（列全体の中央揃え等）を自動で引き継がないため。
        date_cell.alignment = Alignment(horizontal="center")
        fuel_cell.alignment = Alignment(horizontal="center")
        odo_cell.alignment = Alignment(horizontal="center")
        if date_fmt:
            date_cell.number_format = date_fmt
        if fuel_fmt:
            fuel_cell.number_format = fuel_fmt
        if odo_fmt:
            odo_cell.number_format = odo_fmt
        if pd.notna(rec.get("fuel_unit_price")):
            price_cell = ws.cell(row=row, column=COL_PRICE, value=float(rec["fuel_unit_price"]))
            price_cell.alignment = Alignment(horizontal="center")
            if price_fmt:
                price_cell.number_format = price_fmt

        if e_template:
            cell = ws.cell(row=row, column=COL_EFF, value=_shift_formula(e_template, template_row, row))
            cell.alignment = Alignment(horizontal="center")
        if g_template:
            cell = ws.cell(row=row, column=COL_COST, value=_shift_formula(g_template, template_row, row))
            cell.alignment = Alignment(horizontal="center")
        if i_template:
            cell = ws.cell(row=row, column=COL_DIST, value=_shift_formula(i_template, template_row, row))
            cell.alignment = Alignment(horizontal="center")

        note = rec.get("note")
        if isinstance(note, str) and note.strip():
            note_cell = ws.cell(row=row, column=COL_NOTE, value=note)
            note_cell.alignment = Alignment(horizontal="center")
        elif h_template:
            note_cell = ws.cell(row=row, column=COL_NOTE, value=_shift_formula(h_template, template_row, row))
            note_cell.alignment = Alignment(horizontal="center")

        added += 1

    wb.save(EXCEL_PATH)
    _write_status_file("synced", added, cloud_last_date, cloud_last_date)
    return {"status": "synced", "added": added}
