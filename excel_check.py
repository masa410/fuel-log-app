"""
Excel(ソリオ燃費早見表.xlsm)とスプレッドシートの内容を照合する（読み取り専用）。
------------------------------------------------------------------
- Excelへは一切書き込まない。差分の一覧を返すだけ。
- 比較するのは A列(ID)・B列(日付)・C列(給油量)・D列(メーター値)・F列(単価)。
  数式の列（燃費・金額・区間距離）とメモ(H列)は比べない。
- IDが入っていない行（初回のID付与前）は「IDなし」として数える。
- このPCに local_settings.py があるときだけ使える（アプリ側で表示を切り替える）。
"""

from __future__ import annotations

from datetime import datetime, timedelta

import openpyxl

try:
    from local_settings import EXCEL_PATH
except ImportError:
    EXCEL_PATH = None

SHEET_NAME = "燃費確認表"
COL_ID, COL_DATE, COL_FUEL, COL_ODO, COL_PRICE = 1, 2, 3, 4, 6


def _date_str(v) -> str:
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return (datetime(1899, 12, 30) + timedelta(days=v)).strftime("%Y-%m-%d")
    return ""


def _num(v):
    if v is None or v == "" or isinstance(v, bool):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def read_excel_rows(path: str) -> list:
    """Excelの記録行を [{row, id, date, fuel, odo, price}, ...] で返す（値のみ。読み取り専用で開く）。"""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True, keep_vba=False)
    ws = wb[SHEET_NAME]
    out = []
    for r, row in enumerate(ws.iter_rows(min_row=2, max_col=COL_PRICE, values_only=True), start=2):
        d = _date_str(row[COL_DATE - 1])
        if not d:
            continue
        rid = _num(row[COL_ID - 1])
        out.append({
            "row": r,
            "id": int(rid) if rid is not None else None,
            "date": d,
            "fuel": _num(row[COL_FUEL - 1]),
            "odo": _num(row[COL_ODO - 1]),
            "price": _num(row[COL_PRICE - 1]),
        })
    wb.close()
    return out


def _same(a, b, tol=0.05) -> bool:
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    return abs(a - b) <= tol


def compare(df, path: str | None = None) -> dict:
    """スプレッドシートの記録 df と Excel を比べる。
    戻り値: {"summary": 文, "rows": [{区分, ID, 日付, 内容}, ...]}"""
    path = path or EXCEL_PATH
    try:
        ex = read_excel_rows(path)
    except PermissionError:
        return {"summary": "Excelファイルを読めませんでした（開いている・同期中など）。閉じてからやり直してください。", "rows": []}
    except Exception as e:  # ファイル無し・シート名違いなど
        return {"summary": f"Excelファイルを読めませんでした：{e}", "rows": []}

    no_id = [e for e in ex if e["id"] is None]
    if no_id:
        return {
            "summary": f"Excelの{len(no_id)}行にIDがありません。初回のID付与（Excelの設定シートのボタン）を先に行ってください。",
            "rows": [],
        }

    cloud = {int(r["id"]): r for _, r in df.iterrows()}
    ex_by_id = {e["id"]: e for e in ex}
    rows, same = [], 0
    for rid, e in ex_by_id.items():
        c = cloud.get(rid)
        if c is None:
            rows.append({"区分": "Excelだけにある", "ID": rid, "日付": e["date"], "内容": f"{e['row']}行目"})
            continue
        diffs = []
        if e["date"] != c["record_date"].strftime("%Y-%m-%d"):
            diffs.append(f"日付 Excel {e['date']} / クラウド {c['record_date']:%Y-%m-%d}")
        c_price = None if c["fuel_unit_price"] != c["fuel_unit_price"] else float(c["fuel_unit_price"])
        for label, ev, cv in (("給油量", e["fuel"], float(c["fuel_liters"])),
                              ("メーター", e["odo"], float(c["odometer_km"])),
                              ("単価", e["price"], c_price)):
            if not _same(ev, cv):
                diffs.append(f"{label} Excel {ev} / クラウド {cv}")
        if diffs:
            rows.append({"区分": "値が違う", "ID": rid, "日付": e["date"], "内容": "、".join(diffs)})
        else:
            same += 1
    for rid, c in cloud.items():
        if rid not in ex_by_id:
            rows.append({"区分": "クラウドだけにある", "ID": rid, "日付": c["record_date"].strftime("%Y-%m-%d"),
                         "内容": "Excelへ未反映"})
    summary = (f"一致 {same} 件 ／ 値が違う {sum(1 for r in rows if r['区分']=='値が違う')} 件 ／ "
               f"Excelだけ {sum(1 for r in rows if r['区分']=='Excelだけにある')} 件 ／ "
               f"クラウドだけ {sum(1 for r in rows if r['区分']=='クラウドだけにある')} 件")
    return {"summary": summary, "rows": rows}
