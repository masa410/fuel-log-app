"""
Googleスプレッドシートをデータストアとして扱うモジュール
------------------------------------------------------
app.py と migrate_to_sheets.py の両方から利用する。
認証情報は st.secrets["gcp_service_account"] / st.secrets["spreadsheet_id"] から読む。
"""

import json
from datetime import date, datetime

import pandas as pd
import streamlit as st

import record_edit
from gsheets_client import SheetsClient

SHEET_NAME = "records"
HEADERS = [
    "id", "record_date", "odometer_km", "distance_km", "fuel_liters",
    "fuel_unit_price", "fuel_cost", "efficiency_km_per_l", "note", "created_at",
]
HISTORY_SHEET = "history"
HISTORY_HEADERS = ["changed_at", "action", "record_id", "before", "after"]
NUMERIC_COLS = [
    "odometer_km", "distance_km", "fuel_liters",
    "fuel_unit_price", "fuel_cost", "efficiency_km_per_l",
]


@st.cache_resource
def _get_client() -> SheetsClient:
    info = dict(st.secrets["gcp_service_account"])
    client = SheetsClient(info, st.secrets["spreadsheet_id"])
    client.ensure_sheet(SHEET_NAME, HEADERS)
    return client


def load_records() -> pd.DataFrame:
    client = _get_client()
    values = client.get_all_values(SHEET_NAME)
    if len(values) < 2:
        return pd.DataFrame(columns=HEADERS)

    header, rows = values[0], values[1:]
    rows = [r + [""] * (len(header) - len(r)) for r in rows]
    df = pd.DataFrame(rows, columns=header)

    df["record_date"] = pd.to_datetime(df["record_date"])
    for col in NUMERIC_COLS:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df["id"] = pd.to_numeric(df["id"], errors="coerce").astype("Int64")
    df = df.sort_values(["record_date", "id"]).reset_index(drop=True)
    return df


def get_last_odometer():
    df = load_records()
    if df.empty:
        return None
    return df.iloc[-1]["odometer_km"]


def _next_id(df: pd.DataFrame) -> int:
    if df.empty:
        return 1
    return int(df["id"].max()) + 1


def insert_record(record_date, odometer_km, distance_km, fuel_liters, fuel_unit_price,
                   fuel_cost, efficiency, note, created_at=None):
    client = _get_client()
    df = load_records()
    new_id = _next_id(df)
    row = [
        new_id,
        record_date.isoformat() if hasattr(record_date, "isoformat") else record_date,
        odometer_km,
        distance_km if distance_km is not None else "",
        fuel_liters,
        fuel_unit_price if fuel_unit_price is not None else "",
        fuel_cost if fuel_cost is not None else "",
        efficiency if efficiency is not None else "",
        note or "",
        created_at or datetime.now().isoformat(timespec="seconds"),
    ]
    client.append_row(SHEET_NAME, [str(v) for v in row])


def insert_records_bulk(records: list):
    """一括移行用。records: insert_record と同じキーを持つdictのリスト。
    load_records()を1回だけ呼び、まとめて1回のAPI呼び出しで追加する。
    """
    if not records:
        return
    client = _get_client()
    df = load_records()
    next_id = _next_id(df)
    rows = []
    for rec in records:
        row = [
            next_id,
            rec["record_date"],
            rec["odometer_km"],
            rec.get("distance_km") if rec.get("distance_km") is not None else "",
            rec["fuel_liters"],
            rec.get("fuel_unit_price") if rec.get("fuel_unit_price") is not None else "",
            rec.get("fuel_cost") if rec.get("fuel_cost") is not None else "",
            rec.get("efficiency_km_per_l") if rec.get("efficiency_km_per_l") is not None else "",
            rec.get("note") or "",
            rec.get("created_at") or datetime.now().isoformat(timespec="seconds"),
        ]
        rows.append([str(v) for v in row])
        next_id += 1
    client.append_rows(SHEET_NAME, rows)


def _cell(v) -> str:
    """スプレッドシートへ書く文字列にする（欠損は空、日付はYYYY-MM-DD）。"""
    if v is None:
        return ""
    if isinstance(v, (pd.Timestamp, datetime, date)):
        return v.strftime("%Y-%m-%d")
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    return str(v)


def _append_history(client, action, record_id, before, after):
    client.ensure_sheet(HISTORY_SHEET, HISTORY_HEADERS)
    client.append_row(HISTORY_SHEET, [
        datetime.now().isoformat(timespec="seconds"), action, str(record_id),
        json.dumps(before, ensure_ascii=False) if before is not None else "",
        json.dumps(after, ensure_ascii=False) if after is not None else "",
    ])


def _apply_updates(client, updates, target_id):
    """updates（record_edit の計画）をシートへ反映する。 (対象の修正前, 修正後) を dict で返す。"""
    values = client.get_all_values(SHEET_NAME)
    header = values[0]
    id_col = header.index("id")
    by_id = {}
    for i, row in enumerate(values[1:], start=1):
        row = row + [""] * (len(header) - len(row))
        by_id[str(row[id_col])] = (i, row)

    writes, before, after = [], None, None
    for rid, changed in updates:
        i, row = by_id[str(rid)]
        old = dict(zip(header, row))
        new_row = list(row)
        for col, v in changed.items():
            new_row[header.index(col)] = _cell(v)
        writes.append((i, new_row))
        if str(rid) == str(target_id):
            before, after = old, dict(zip(header, new_row))
    client.batch_update_rows(SHEET_NAME, writes)
    if before is None and str(target_id) in by_id:
        i, row = by_id[str(target_id)]
        before = dict(zip(header, row))
    return before, after


def update_record(record_id, new: dict, today=None):
    """記録を修正する。new: record_date・odometer_km・fuel_liters・fuel_unit_price・fuel_cost・note。
    戻り値: (成功したか, 画面に出す文)。範囲外などは書き込まずに理由を返す。"""
    client = _get_client()
    df = load_records()
    updates, err = record_edit.plan_edit(df, record_id, new, today)
    if err:
        return False, err
    if not updates:
        return True, "変更はありませんでした。"
    before, after = _apply_updates(client, updates, record_id)
    _append_history(client, "update", record_id, before, after)
    return True, "修正しました。"


def delete_record(record_id):
    """記録を削除する。次の記録の距離・燃費を再計算し、変更履歴に残す。"""
    client = _get_client()
    df = load_records()
    updates, err = record_edit.plan_delete(df, record_id)
    if err:
        return False, err
    if updates:
        _apply_updates(client, updates, record_id)
    values = client.get_all_values(SHEET_NAME)
    if len(values) < 2:
        return False, "対象の記録が見つかりません。"
    header = values[0]
    id_col = header.index("id")
    for i, row in enumerate(values[1:], start=1):
        if len(row) > id_col and str(row[id_col]) == str(record_id):
            before = dict(zip(header, row + [""] * (len(header) - len(row))))
            client.delete_row(SHEET_NAME, i)
            _append_history(client, "delete", record_id, before, None)
            return True, "削除しました。"
    return False, "対象の記録が見つかりません。"
