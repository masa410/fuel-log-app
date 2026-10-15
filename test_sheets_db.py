"""sheets_db.py の修正・削除の書き込みを、メモリ上の偽クライアントで確認する（ネットワーク不要・ダミーデータのみ）。
python -m unittest test_sheets_db -v"""
import json
import unittest
from datetime import date

import sheets_db


class FakeClient:
    def __init__(self, rows):
        self.sheets = {"records": [list(sheets_db.HEADERS)] + [list(r) for r in rows]}
        self.calls = []

    def ensure_sheet(self, title, headers):
        if title not in self.sheets:
            self.sheets[title] = [list(headers)]

    def get_all_values(self, title):
        return [list(r) for r in self.sheets.get(title, [])]

    def batch_update_rows(self, title, updates):
        self.calls.append(("batch", len(updates)))
        for idx, row in updates:
            self.sheets[title][idx] = list(row)

    def append_row(self, title, row):
        self.sheets[title].append(list(row))

    def delete_row(self, title, idx):
        self.calls.append(("delete", idx))
        del self.sheets[title][idx]


def rec(i, d, odo, dist, liters, price, cost, eff):
    return [str(i), d, str(odo), str(dist), str(liters), str(price), str(cost), str(eff), "", "2026-10-01T00:00:00"]


ROWS = [
    rec(1, "2026-09-01", 10000.0, "", 30.0, 170, 5100, ""),
    rec(2, "2026-09-10", 10400.0, 400.0, 25.0, 172, 4300, 16.0),
    rec(3, "2026-09-20", 10800.0, 400.0, 20.0, 175, 3500, 20.0),
    rec(4, "2026-10-01", 11300.0, 500.0, 28.0, 168, 4704, 17.86),
]


class DbTests(unittest.TestCase):
    def setUp(self):
        self.fake = FakeClient(ROWS)
        self._orig = sheets_db._get_client
        sheets_db._get_client = lambda: self.fake
        self.today = date(2026, 10, 15)

    def tearDown(self):
        sheets_db._get_client = self._orig

    def test_update_writes_row_neighbor_and_history(self):
        new = {"record_date": date(2026, 9, 12), "odometer_km": 10500.0, "fuel_liters": 25.0,
               "fuel_unit_price": 172.0, "fuel_cost": 4300.0, "note": "修正テスト"}
        ok, msg = sheets_db.update_record(2, new, self.today)
        self.assertTrue(ok, msg)
        recs = {r[0]: r for r in self.fake.sheets["records"][1:]}
        self.assertEqual(recs["2"][1], "2026-09-12")
        self.assertEqual(float(recs["2"][2]), 10500.0)
        self.assertEqual(float(recs["2"][3]), 500.0)
        self.assertEqual(float(recs["2"][7]), 20.0)
        self.assertEqual(recs["2"][8], "修正テスト")
        self.assertEqual(float(recs["3"][3]), 300.0)       # 次の記録の距離が再計算される
        self.assertEqual(float(recs["3"][7]), 15.0)
        self.assertEqual(recs["4"][3], "500.0")            # その先は変わらない
        self.assertEqual(recs["2"][9], "2026-10-01T00:00:00")   # created_at は保持
        hist = self.fake.sheets["history"]
        self.assertEqual(len(hist), 2)
        self.assertEqual(hist[1][1], "update")
        self.assertEqual(json.loads(hist[1][3])["odometer_km"], "10400.0")
        self.assertEqual(json.loads(hist[1][4])["odometer_km"], "10500.0")

    def test_out_of_range_writes_nothing(self):
        new = {"record_date": date(2026, 8, 1), "odometer_km": 10400.0, "fuel_liters": 25.0,
               "fuel_unit_price": 172.0, "fuel_cost": 4300.0, "note": ""}
        before = [list(r) for r in self.fake.sheets["records"]]
        ok, msg = sheets_db.update_record(2, new, self.today)
        self.assertFalse(ok)
        self.assertIn("以降", msg)
        self.assertEqual(self.fake.sheets["records"], before)
        self.assertNotIn("history", self.fake.sheets)

    def test_delete_recalculates_next_and_logs(self):
        ok, msg = sheets_db.delete_record(2)
        self.assertTrue(ok, msg)
        ids = [r[0] for r in self.fake.sheets["records"][1:]]
        self.assertEqual(ids, ["1", "3", "4"])
        r3 = self.fake.sheets["records"][2]
        self.assertEqual(float(r3[3]), 800.0)
        self.assertEqual(float(r3[7]), 40.0)
        self.assertEqual(self.fake.sheets["history"][1][1], "delete")

    def test_delete_last(self):
        ok, _ = sheets_db.delete_record(4)
        self.assertTrue(ok)
        self.assertEqual([r[0] for r in self.fake.sheets["records"][1:]], ["1", "2", "3"])

    def test_clearing_price_leaves_cell_empty(self):
        new = {"record_date": date(2026, 9, 10), "odometer_km": 10400.0, "fuel_liters": 25.0,
               "fuel_unit_price": None, "fuel_cost": None, "note": ""}
        ok, _ = sheets_db.update_record(2, new, self.today)
        self.assertTrue(ok)
        row = [r for r in self.fake.sheets["records"] if r[0] == "2"][0]
        self.assertEqual((row[5], row[6]), ("", ""))


if __name__ == "__main__":
    unittest.main()
