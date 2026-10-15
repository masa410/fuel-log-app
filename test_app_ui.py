"""app.py の画面（修正・削除）を、Streamlitの AppTest で確認する。ダミーデータ・偽のDB関数のみ。
python -m unittest test_app_ui -v"""
import unittest
from datetime import date

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
        self.assertTrue(any("以降にしてください" in e.value for e in at.error))
        self.assertEqual(at.session_state["edit_id"], 2)

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


if __name__ == "__main__":
    unittest.main()
