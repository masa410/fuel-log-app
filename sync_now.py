"""
Excel(ソリオ燃費早見表.xlsm)を開く前に、デスクトップのランチャーから呼び出される単発スクリプト。
また、Excelシート上の「修正内容反映」ボタンからも呼び出される。

Googleスプレッドシートの最新データを「ダウンロード」フォルダのCSVに書き出して終了する
（Streamlitのブラウザ画面は開かない）。Excelへの反映は、Excel側のマクロがCSVを読んで行う。
"""

import os
import sys

os.chdir(os.path.dirname(os.path.abspath(__file__)))

import excel_sync
from sheets_db import load_records


def main():
    if not excel_sync.is_available():
        print("SYNC_SKIPPED: local_settings.py not found on this PC")
        return

    try:
        result = excel_sync.export_records_for_excel(load_records())
    except Exception as e:
        print(f"SYNC_ERROR: {e}")
        sys.exit(1)

    print(f"SYNC_RESULT: {result}")
    if result.get("status") != "exported":
        sys.exit(1)


if __name__ == "__main__":
    main()
