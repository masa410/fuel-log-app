"""app.py の画面（修正・削除）を、Streamlitの AppTest で確認する。ダミーデータ・偽のDB関数のみ。
python -m unittest test_app_ui -v"""
import unittest
from datetime import date

import pandas as pd
from streamlit.testing.v1 import AppTest

import excel_sync
import sheets_db
from test_record_edit import BASE, make_df


class AppUiTests(unittest.TestCase):
    def setUp(self):
        self.df = make_df(BASE)
        self.calls = []
        self._saved = {
            "load_records": sheets_db.load_records,
            "get_last_odometer": sheets_db.get_last_odometer,
            "update_record": sheets_db.update_record,
            "delete_record": sheets_db.delete_record,
            "is_available": excel_sync.is_available,
        }
        sheets_db.load_records = lambda: self.df
        sheets_db.get_last_odometer = lambda: float(self.df.iloc[-1]["odometer_km"])
        sheets_db.update_record = lambda rid, new, today=None: (self.calls.append(("update", rid, new)) or (True, "修正しました。"))
        sheets_db.delete_record = lambda rid: (self.calls.append(("delete", rid)) or (True, "削除しました。"))
        excel_sync.is_available = lambda: False

    def tearDown(self):
        for k, v in self._saved.items():
            if k == "is_available":
                excel_sync.is_available = v
            else:
                setattr(sheets_db, k, v)

    def make_app(self):
        at = AppTest.from_file("app.py", default_timeout=30)
        at.session_state["authenticated"] = True
        return at

    def test_history_tab_renders_without_error(self):
        at = self.make_app().run()
        self.assertFalse(at.exception, at.exception)
        labels = [b.label for b in at.button]
        self.assertIn("✏️ 選択した記録を修正する", labels)

    def test_edit_form_shows_current_values_and_saves(self):
        at = self.make_app()
        at.session_state["edit_id"] = 2
        at.run()
        self.assertFalse(at.exception, at.exception)
        odo = [n for n in at.number_input if n.label.startswith("メーター表示") and n.key == "e_odo_2"][0]
        self.assertEqual(odo.value, 10400.0)
        odo.set_value(10500.0)
        [b for b in at.button if b.key == "e_save_2"][0].click()
        at.run()
        self.assertFalse(at.exception, at.exception)
        kind, rid, new = self.calls[0]
        self.assertEqual((kind, rid), ("update", 2))
        self.assertEqual(new["odometer_km"], 10500.0)
        self.assertEqual(new["record_date"], date(2026, 9, 10))
        self.assertNotIn("edit_id", at.session_state)

    def test_validation_error_is_shown_and_form_stays(self):
        sheets_db.update_record = lambda rid, new, today=None: (False, "日付は 2026/09/01（前の記録）以降にしてください。")
        at = self.make_app()
        at.session_state["edit_id"] = 2
        at.run()
        [b for b in at.button if b.key == "e_save_2"][0].click()
        at.run()
        # エラーはポップアップで出す。修正フォームは残る
        self.assertEqual(at.session_state["edit_id"], 2)
        self.assertTrue(any("以降にしてください" in m.value for m in at.markdown))

    def test_delete_requires_confirmation(self):
        at = self.make_app()
        at.session_state["edit_id"] = 3
        at.run()
        delete_btn = [b for b in at.button if b.key == "e_delete_3"][0]
        self.assertTrue(delete_btn.disabled)
        [c for c in at.checkbox if c.key == "e_confirm_3"][0].check()
        at.run()
        [b for b in at.button if b.key == "e_delete_3"][0].click()
        at.run()
        self.assertEqual(self.calls[0], ("delete", 3))


class AddFlowTests(unittest.TestCase):
    """記録追加の流れ（二重記録の防止・初期化・ポップアップ）。偽のDB関数のみ。"""

    def setUp(self):
        self.df = make_df(BASE)
        self.inserted = []
        self._saved = {k: getattr(sheets_db, k) for k in
                       ("load_records", "get_last_odometer", "insert_record")}
        self._saved_avail = excel_sync.is_available
        sheets_db.load_records = lambda: self.df
        sheets_db.get_last_odometer = lambda: float(self.df.iloc[-1]["odometer_km"])

        def fake_insert(record_date, odo, dist, liters, price, cost, eff, note):
            self.inserted.append((record_date, odo, dist, liters))
            new_id = int(self.df["id"].max()) + 1
            self.df = pd.concat([self.df, pd.DataFrame([{
                "id": new_id, "record_date": pd.Timestamp(record_date), "odometer_km": float(odo),
                "distance_km": dist if dist is not None else float("nan"), "fuel_liters": float(liters),
                "fuel_unit_price": float("nan"), "fuel_cost": float("nan"),
                "efficiency_km_per_l": eff if eff is not None else float("nan"), "note": note, "created_at": "",
            }])], ignore_index=True)
            return new_id

        sheets_db.insert_record = fake_insert
        excel_sync.is_available = lambda: False

    def tearDown(self):
        for k, v in self._saved.items():
            setattr(sheets_db, k, v)
        excel_sync.is_available = self._saved_avail

    def make_app(self):
        at = AppTest.from_file("app.py", default_timeout=30)
        at.session_state["authenticated"] = True
        return at.run()

    @staticmethod
    def record_button(at):
        return [b for b in at.button if b.label == "この内容で記録する"][0]

    def fill(self, at, odo, liters, day=None):
        at.number_input(key="f_odo").set_value(odo)
        at.number_input(key="f_liters").set_value(liters)
        if day is not None:
            at.date_input(key="f_date").set_value(day)
        at.run()

    def test_initial_form_values(self):
        at = self.make_app()
        self.assertFalse(at.exception, at.exception)
        self.assertEqual(at.number_input(key="f_odo").value, 11300.0)
        self.assertEqual(at.number_input(key="f_liters").value, 0.0)
        self.assertFalse(self.record_button(at).disabled)

    def test_unit_price_is_auto_calculated_from_cost_and_liters(self):
        at = self.make_app()
        at.number_input(key="f_liters").set_value(30.0)
        at.number_input(key="f_cost").set_value(5100.0)
        at.run()
        price = [n for n in at.number_input if n.label.startswith("給油単価")][0]
        self.assertEqual(price.value, 170.0)
        # 金額を変えたら、自動計算も追従する
        at.number_input(key="f_cost").set_value(5400.0)
        at.run()
        price = [n for n in at.number_input if n.label.startswith("給油単価")][0]
        self.assertEqual(price.value, 180.0)

    def test_record_then_form_resets_and_button_locks(self):
        at = self.make_app()
        self.fill(at, 11700.0, 30.0)
        self.record_button(at).click()
        at.run()
        self.assertFalse(at.exception, at.exception)
        self.assertEqual(len(self.inserted), 1)
        self.assertEqual(self.inserted[0][1:], (11700.0, 400.0, 30.0))
        # 入力欄：メーターだけ新しい値、他は0
        self.assertEqual(at.number_input(key="f_odo").value, 11700.0)
        self.assertEqual(at.number_input(key="f_liters").value, 0.0)
        self.assertEqual(at.number_input(key="f_cost").value, 0.0)
        self.assertEqual(at.text_input(key="f_note").value, "")
        # 記録ボタンは使えない（ポップアップを閉じた後も）
        self.assertTrue(self.record_button(at).disabled)
        self.assertTrue(at.session_state["locked"])
        self.assertTrue(any("今回の入力が完了しました" in m.value for m in at.markdown))
        # 再実行（ポップアップを閉じた状態）でも鍵は残り、二重記録できない
        at.run()
        self.assertTrue(self.record_button(at).disabled)
        self.assertEqual(len(self.inserted), 1)

    def test_lock_released_when_liters_entered(self):
        at = self.make_app()
        self.fill(at, 11700.0, 30.0)
        self.record_button(at).click()
        at.run()
        at.number_input(key="f_liters").set_value(25.0)
        at.run()
        self.assertFalse(at.session_state["locked"])
        self.assertFalse(self.record_button(at).disabled)

    def test_zero_liters_is_refused_with_popup(self):
        at = self.make_app()
        self.record_button(at).click()
        at.run()
        self.assertEqual(self.inserted, [])
        self.assertFalse(at.session_state["locked"])
        self.assertTrue(any("給油量を入力してください" in m.value for m in at.markdown))

    def test_smaller_odometer_is_not_recorded(self):
        at = self.make_app()
        self.fill(at, 11000.0, 20.0)
        self.record_button(at).click()
        at.run()
        self.assertEqual(self.inserted, [])
        self.assertTrue(any("前回より小さく" in m.value for m in at.markdown))

    def test_same_as_last_record_is_refused(self):
        at = self.make_app()
        self.fill(at, 11300.0, 28.0, day=date(2026, 10, 1))
        self.record_button(at).click()
        at.run()
        self.assertEqual(self.inserted, [])
        self.assertFalse(at.session_state["locked"])
        self.assertTrue(any("直前の記録と同じ内容" in m.value for m in at.markdown))

    def test_insert_failure_unlocks_and_shows_popup(self):
        def boom(*a, **k):
            raise RuntimeError("network")
        sheets_db.insert_record = boom
        at = self.make_app()
        self.fill(at, 11700.0, 30.0)
        self.record_button(at).click()
        at.run()
        self.assertFalse(at.exception, at.exception)
        self.assertFalse(at.session_state["locked"])
        self.assertTrue(any("記録できませんでした" in m.value for m in at.markdown))

    def test_history_with_highlight_and_tab_switch_renders(self):
        at = self.make_app()
        at.session_state["hl_id"] = 4
        at.session_state["go_history"] = True
        at.run()
        self.assertFalse(at.exception, at.exception)
        self.assertFalse(at.session_state["go_history"])  # 1回だけ使って消える

    def test_highlight_cleared_when_input_changes(self):
        at = self.make_app()
        at.session_state["hl_id"] = 4
        at.run()
        at.number_input(key="f_odo").set_value(11500.0)
        at.run()
        self.assertIsNone(at.session_state["hl_id"])


if __name__ == "__main__":
    unittest.main()
