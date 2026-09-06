"""
Excel(ソリオ燃費早見表.xlsm)を開く前に、デスクトップのランチャーから
呼び出される単発同期スクリプト。

Googleスプレッドシートの最新データをExcelファイルへ反映するだけで終了する
(Streamlitのブラウザ画面は開かない)。Excelがまだ開いていないタイミングで
実行することが前提(ランチャー側でこのスクリプト→Excel起動、の順に呼ぶこと)。
"""

import os
import sys

os.chdir(os.path.dirname(os.path.abspath(__file__)))

import excel_sync
from sheets_db import load_records


def main():
    if not excel_sync.is_available():
        print("SYNC_SKIPPED: Excel file not found on this PC")
        return

    try:
        result = excel_sync.sync_records_to_excel(load_records())
    except Exception as e:
        print(f"SYNC_ERROR: {e}")
        sys.exit(1)

    print(f"SYNC_RESULT: {result}")


if __name__ == "__main__":
    main()
