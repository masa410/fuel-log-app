"""
Googleスプレッドシートをデータストアとして扱うモジュール
------------------------------------------------------
app.py と migrate_to_sheets.py の両方から利用する。
認証情報は st.secrets["gcp_service_account"] / st.secrets["spreadsheet_id"] から読む。
"""

from datetime import datetime

import pandas as pd
import streamlit as st

from gsheets_client import SheetsClient

SHEET_NAME = "records"
HEADERS = [
    "id", "record_date", "odometer_km", "distance_km", "fuel_liters",
    "fuel_unit_price", "fuel_cost", "efficiency_km_per_l", "note", "created_at",
]
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


def delete_record(record_id):
    client = _get_client()
    values = client.get_all_values(SHEET_NAME)
    if len(values) < 2:
        return
    id_col = values[0].index("id")
    for i, row in enumerate(values[1:], start=1):
        if len(row) > id_col and str(row[id_col]) == str(record_id):
            client.delete_row(SHEET_NAME, i)
            return
