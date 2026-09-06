"""
既存の fuel_log.db (sqlite) の中身を、Googleスプレッドシートへ一括移行するスクリプト。

事前準備:
  1. .streamlit/secrets.toml を用意する（README.md参照。json_to_secrets.py で自動生成可能）
  2. サービスアカウントのメールアドレスを、移行先スプレッドシートに「編集者」として共有しておく

実行方法:
  python migrate_to_sheets.py
"""

import os
import sqlite3

from sheets_db import insert_records_bulk, load_records

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "fuel_log.db")


def main():
    if not os.path.exists(DB_PATH):
        print(f"fuel_log.db が見つかりません: {DB_PATH}")
        return

    existing = load_records()
    already_migrated_dates = set(existing["record_date"].dt.strftime("%Y-%m-%d")) if not existing.empty else set()

    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute(
        """
        SELECT record_date, odometer_km, distance_km, fuel_liters, fuel_unit_price,
               fuel_cost, efficiency_km_per_l, note, created_at
        FROM records ORDER BY record_date ASC, id ASC
        """
    ).fetchall()
    conn.close()

    to_migrate, skipped = [], 0
    for row in rows:
        (record_date, odometer_km, distance_km, fuel_liters, fuel_unit_price,
         fuel_cost, efficiency, note, created_at) = row
        if record_date in already_migrated_dates:
            skipped += 1
            continue
        to_migrate.append({
            "record_date": record_date, "odometer_km": odometer_km,
            "distance_km": distance_km, "fuel_liters": fuel_liters,
            "fuel_unit_price": fuel_unit_price, "fuel_cost": fuel_cost,
            "efficiency_km_per_l": efficiency, "note": note, "created_at": created_at,
        })

    insert_records_bulk(to_migrate)
    print(f"移行完了: {len(to_migrate)}件をスプレッドシートに追加、{skipped}件は既に存在したためスキップしました。")


if __name__ == "__main__":
    main()
