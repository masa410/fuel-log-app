"""
Google Cloud Consoleでダウンロードしたサービスアカウントの鍵(JSONファイル)を、
Streamlitの .streamlit/secrets.toml に変換するヘルパースクリプト。

使い方:
  python json_to_secrets.py "ダウンロードしたjsonファイルのパス" "スプレッドシートのID"

スプレッドシートのIDは、URLの
  https://docs.google.com/spreadsheets/d/【ここがID】/edit
の【ここがID】の部分です。
"""

import json
import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def toml_escape(value: str) -> str:
    return value.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n")


def main():
    if len(sys.argv) != 3:
        print("使い方: python json_to_secrets.py <サービスアカウントjsonのパス> <スプレッドシートID>")
        return

    json_path, spreadsheet_id = sys.argv[1], sys.argv[2]
    with open(json_path, encoding="utf-8") as f:
        creds = json.load(f)

    lines = [f'spreadsheet_id = "{toml_escape(spreadsheet_id)}"', "", "[gcp_service_account]"]
    for key, value in creds.items():
        lines.append(f'{key} = "{toml_escape(str(value))}"')

    out_dir = os.path.join(BASE_DIR, ".streamlit")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "secrets.toml")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

    print(f"作成しました: {out_path}")
    print("このファイルは絶対にGitHubにアップロードしないでください（.gitignoreで除外済みです）。")


if __name__ == "__main__":
    main()
