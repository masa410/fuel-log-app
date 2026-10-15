"""
記録の修正・削除の計画を立てる純粋な関数群（Streamlit・ネットワークに依存しない）。
------------------------------------------------------------------
sheets_db.py から呼ばれる。ここでは「何を書き換えるか」だけを決め、書き込みはしない。
テストは test_record_edit.py。

並び順は load_records() と同じ (record_date, id) の昇順。
距離・燃費の計算は app.py の記録追加と同じ式：
    距離 = メーター値 - 直前の記録のメーター値（負なら保存しない）
    燃費 = 距離 / 給油量（小数第2位。距離が0または無しなら保存しない）
"""

from __future__ import annotations

from datetime import date

import pandas as pd


def calc_distance_efficiency(prev_odo, odo, liters):
    """直前メーター値・今回メーター値・給油量から (距離, 燃費) を返す。算出できなければ (None, None)。"""
    if prev_odo is None or pd.isna(prev_odo) or odo is None or pd.isna(odo):
        return None, None
    distance = round(float(odo) - float(prev_odo), 1)
    if distance < 0:
        return None, None
    efficiency = round(distance / float(liters), 2) if (distance and liters and float(liters) > 0) else None
    return distance, efficiency


def _sorted(df: pd.DataFrame) -> pd.DataFrame:
    return df.sort_values(["record_date", "id"]).reset_index(drop=True)


def _same(a, b) -> bool:
    a_na = a is None or (not isinstance(a, str) and pd.isna(a))
    b_na = b is None or (not isinstance(b, str) and pd.isna(b))
    if a_na or b_na:
        return a_na and b_na
    if isinstance(a, (str, pd.Timestamp)) or isinstance(b, (str, pd.Timestamp)):
        return str(a) == str(b)
    return abs(float(a) - float(b)) < 1e-9


def edit_bounds(df: pd.DataFrame, record_id, today: date | None = None) -> dict:
    """修正フォームに出す入力可能範囲を返す。
    {"min_date": date|None, "max_date": date, "min_odo": float|None, "max_odo": float|None}
    """
    today = today or date.today()
    s = _sorted(df)
    idx = s.index[s["id"] == record_id]
    if len(idx) == 0:
        raise KeyError(record_id)
    pos = int(idx[0])
    prev = s.iloc[pos - 1] if pos > 0 else None
    nxt = s.iloc[pos + 1] if pos + 1 < len(s) else None
    return {
        "min_date": prev["record_date"].date() if prev is not None else None,
        "max_date": nxt["record_date"].date() if nxt is not None else today,
        "min_odo": float(prev["odometer_km"]) if prev is not None and pd.notna(prev["odometer_km"]) else None,
        "max_odo": float(nxt["odometer_km"]) if nxt is not None and pd.notna(nxt["odometer_km"]) else None,
    }


def plan_edit(df: pd.DataFrame, record_id, new: dict, today: date | None = None):
    """修正の計画を返す。 (updates, error)

    new: record_date(date)・odometer_km・fuel_liters・fuel_unit_price(None可)・fuel_cost(None可)・note
    updates: [(id, {列名: 新しい値, ...}), ...]  変わる列だけ。対象の記録と、距離・燃費が変わる前後の記録。
    error: 受け付けられないときの理由（日本語）。updates は空。
    """
    today = today or date.today()
    s = _sorted(df)
    if not (s["id"] == record_id).any():
        return [], "対象の記録が見つかりません。"

    new_date = pd.Timestamp(new["record_date"])
    odo = float(new["odometer_km"])
    liters = float(new["fuel_liters"])
    if liters <= 0:
        return [], "給油量は0より大きい値を入力してください。"
    price = new.get("fuel_unit_price")
    if price is not None and float(price) <= 0:
        return [], "単価は0より大きい値を入力してください（入力しない場合は空にしてください）。"

    b = edit_bounds(df, record_id, today)
    if b["min_date"] is not None and new_date.date() < b["min_date"]:
        return [], f"日付は {b['min_date']:%Y/%m/%d}（前の記録）以降にしてください。"
    if new_date.date() > b["max_date"]:
        if b["max_date"] == today:
            return [], f"日付は今日（{today:%Y/%m/%d}）以前にしてください。"
        return [], f"日付は {b['max_date']:%Y/%m/%d}（次の記録）以前にしてください。"
    if b["min_odo"] is not None and odo < b["min_odo"]:
        return [], f"メーター値は {b['min_odo']:,.0f} km（前の記録）以上にしてください。"
    if b["max_odo"] is not None and odo > b["max_odo"]:
        return [], f"メーター値は {b['max_odo']:,.0f} km（次の記録）以下にしてください。"

    # 修正後の並び順で再計算する（同じ日付の記録どうしは id 順のため、順序が入れ替わる場合がある）
    cand = s.copy()
    m = cand["id"] == record_id
    cand.loc[m, "record_date"] = new_date
    cand.loc[m, "odometer_km"] = odo
    cand.loc[m, "fuel_liters"] = liters
    cand.loc[m, "fuel_unit_price"] = float(price) if price is not None else float("nan")
    cand.loc[m, "fuel_cost"] = float(new["fuel_cost"]) if new.get("fuel_cost") is not None else float("nan")
    cand.loc[m, "note"] = new.get("note") or ""
    cand = _sorted(cand)

    j = int(cand.index[cand["id"] == record_id][0])
    if j > 0 and cand.iloc[j - 1]["odometer_km"] > odo:
        return [], "この日付だと、同じ日付の別の記録とメーター値の順序が合いません。日付かメーター値を確認してください。"
    if j + 1 < len(cand) and cand.iloc[j + 1]["odometer_km"] < odo:
        return [], "この日付だと、同じ日付の別の記録とメーター値の順序が合いません。日付かメーター値を確認してください。"

    # 距離・燃費を再計算する記録：対象、修正前の次の記録、修正後の次の記録
    old_pos = int(s.index[s["id"] == record_id][0])
    affected = {record_id}
    if old_pos + 1 < len(s):
        affected.add(s.iloc[old_pos + 1]["id"])
    if j + 1 < len(cand):
        affected.add(cand.iloc[j + 1]["id"])

    old_by_id = s.set_index("id")
    updates = []
    for pos in range(len(cand)):
        rid = cand.iloc[pos]["id"]
        if rid not in affected:
            continue
        row = cand.iloc[pos]
        prev_odo = cand.iloc[pos - 1]["odometer_km"] if pos > 0 else None
        dist, eff = calc_distance_efficiency(prev_odo, row["odometer_km"], row["fuel_liters"])
        want = {"distance_km": dist, "efficiency_km_per_l": eff}
        if rid == record_id:
            want.update({
                "record_date": row["record_date"],
                "odometer_km": row["odometer_km"],
                "fuel_liters": row["fuel_liters"],
                "fuel_unit_price": row["fuel_unit_price"],
                "fuel_cost": row["fuel_cost"],
                "note": row["note"],
            })
        old = old_by_id.loc[rid]
        changed = {k: v for k, v in want.items() if not _same(old[k], v)}
        if changed:
            updates.append((rid, changed))
    return updates, None


def plan_delete(df: pd.DataFrame, record_id):
    """削除の計画を返す。 (updates, error)。updates は削除後に距離・燃費が変わる次の記録の分。"""
    s = _sorted(df)
    idx = s.index[s["id"] == record_id]
    if len(idx) == 0:
        return [], "対象の記録が見つかりません。"
    pos = int(idx[0])
    if pos + 1 >= len(s):
        return [], None
    nxt = s.iloc[pos + 1]
    prev_odo = s.iloc[pos - 1]["odometer_km"] if pos > 0 else None
    dist, eff = calc_distance_efficiency(prev_odo, nxt["odometer_km"], nxt["fuel_liters"])
    want = {"distance_km": dist, "efficiency_km_per_l": eff}
    changed = {k: v for k, v in want.items() if not _same(nxt[k], v)}
    return ([(nxt["id"], changed)] if changed else []), None
