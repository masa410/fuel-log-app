"""record_edit.py のテスト（ダミーデータのみ。ネットワーク不要）。 python -m unittest test_record_edit -v"""
import unittest
from datetime import date

import pandas as pd

import record_edit as re_


def make_df(rows):
    # rows: (id, 日付, メーター, 給油量, 単価, 金額, メモ)  距離・燃費は計算で入れる
    recs, prev = [], None
    for rid, d, odo, liters, price, cost, note in rows:
        dist, eff = re_.calc_distance_efficiency(prev, odo, liters)
        recs.append({
            "id": rid, "record_date": pd.Timestamp(d), "odometer_km": float(odo),
            "distance_km": dist if dist is not None else float("nan"),
            "fuel_liters": float(liters),
            "fuel_unit_price": float(price) if price else float("nan"),
            "fuel_cost": float(cost) if cost else float("nan"),
            "efficiency_km_per_l": eff if eff is not None else float("nan"),
            "note": note, "created_at": "",
        })
        prev = odo
    return pd.DataFrame(recs)


BASE = [
    (1, "2026-09-01", 10000, 30.0, 170, 5100, ""),
    (2, "2026-09-10", 10400, 25.0, 172, 4300, ""),
    (3, "2026-09-20", 10800, 20.0, 175, 3500, ""),
    (4, "2026-10-01", 11300, 28.0, 168, 4704, ""),
]


def new_of(rid, df, **over):
    r = df[df["id"] == rid].iloc[0]
    d = {
        "record_date": r["record_date"].date(), "odometer_km": r["odometer_km"],
        "fuel_liters": r["fuel_liters"],
        "fuel_unit_price": None if pd.isna(r["fuel_unit_price"]) else r["fuel_unit_price"],
        "fuel_cost": None if pd.isna(r["fuel_cost"]) else r["fuel_cost"],
        "note": r["note"],
    }
    d.update(over)
    return d


class EditTests(unittest.TestCase):
    def setUp(self):
        self.df = make_df(BASE)
        self.today = date(2026, 10, 15)

    def test_bounds_middle(self):
        b = re_.edit_bounds(self.df, 2, self.today)
        self.assertEqual(b["min_date"], date(2026, 9, 1))
        self.assertEqual(b["max_date"], date(2026, 9, 20))
        self.assertEqual((b["min_odo"], b["max_odo"]), (10000.0, 10800.0))

    def test_bounds_last_uses_today(self):
        b = re_.edit_bounds(self.df, 4, self.today)
        self.assertEqual(b["max_date"], self.today)
        self.assertIsNone(b["max_odo"])

    def test_bounds_first_has_no_lower(self):
        b = re_.edit_bounds(self.df, 1, self.today)
        self.assertIsNone(b["min_date"])
        self.assertIsNone(b["min_odo"])

    def test_date_inside_ok_and_same_day_ok(self):
        ups, err = re_.plan_edit(self.df, 2, new_of(2, self.df, record_date=date(2026, 9, 20)), self.today)
        self.assertIsNone(err)

    def test_date_out_of_range(self):
        _, err = re_.plan_edit(self.df, 2, new_of(2, self.df, record_date=date(2026, 8, 31)), self.today)
        self.assertIn("以降", err)
        _, err = re_.plan_edit(self.df, 2, new_of(2, self.df, record_date=date(2026, 9, 21)), self.today)
        self.assertIn("以前", err)

    def test_last_record_cannot_be_future(self):
        _, err = re_.plan_edit(self.df, 4, new_of(4, self.df, record_date=date(2026, 10, 16)), self.today)
        self.assertIn("今日", err)

    def test_odometer_out_of_range(self):
        _, err = re_.plan_edit(self.df, 2, new_of(2, self.df, odometer_km=9999), self.today)
        self.assertIn("以上", err)
        _, err = re_.plan_edit(self.df, 2, new_of(2, self.df, odometer_km=10801), self.today)
        self.assertIn("以下", err)

    def test_liters_and_price_must_be_positive(self):
        _, err = re_.plan_edit(self.df, 2, new_of(2, self.df, fuel_liters=0), self.today)
        self.assertIn("給油量", err)
        _, err = re_.plan_edit(self.df, 2, new_of(2, self.df, fuel_unit_price=0), self.today)
        self.assertIn("単価", err)

    def test_edit_odometer_recalculates_self_and_next(self):
        ups, err = re_.plan_edit(self.df, 2, new_of(2, self.df, odometer_km=10500), self.today)
        self.assertIsNone(err)
        by = dict(ups)
        self.assertEqual(by[2]["distance_km"], 500.0)
        self.assertEqual(by[2]["efficiency_km_per_l"], 20.0)      # 500/25
        self.assertEqual(by[3]["distance_km"], 300.0)             # 10800-10500
        self.assertEqual(by[3]["efficiency_km_per_l"], 15.0)      # 300/20
        self.assertNotIn(4, by)                                   # その先は変わらない

    def test_edit_liters_changes_only_efficiency_of_self(self):
        ups, _ = re_.plan_edit(self.df, 2, new_of(2, self.df, fuel_liters=20.0), self.today)
        by = dict(ups)
        self.assertEqual(by[2]["efficiency_km_per_l"], 20.0)      # 400/20
        self.assertNotIn("distance_km", by[2])
        self.assertNotIn(3, by)

    def test_no_change_gives_no_updates(self):
        ups, err = re_.plan_edit(self.df, 2, new_of(2, self.df), self.today)
        self.assertEqual((ups, err), ([], None))

    def test_same_day_order_flip_is_rejected(self):
        # id=3 を id=4 と同じ日付にしても id 順（3→4）は崩れないので許可される
        ups, err = re_.plan_edit(self.df, 3, new_of(3, self.df, record_date=date(2026, 10, 1)), self.today)
        self.assertIsNone(err)

    def test_delete_recalculates_next(self):
        ups, err = re_.plan_delete(self.df, 2)
        self.assertIsNone(err)
        by = dict(ups)
        self.assertEqual(by[3]["distance_km"], 800.0)             # 10800-10000
        self.assertEqual(by[3]["efficiency_km_per_l"], 40.0)      # 800/20

    def test_delete_last_has_no_updates(self):
        self.assertEqual(re_.plan_delete(self.df, 4), ([], None))

    def test_delete_first_makes_next_have_no_distance(self):
        ups, _ = re_.plan_delete(self.df, 1)
        by = dict(ups)
        self.assertTrue(pd.isna(by[2]["distance_km"]) or by[2]["distance_km"] is None)

    def test_negative_distance_is_not_saved(self):
        self.assertEqual(re_.calc_distance_efficiency(10000, 9990, 10), (None, None))


if __name__ == "__main__":
    unittest.main()
