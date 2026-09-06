# 燃費管理アプリ（クラウド版）

メーターと給油レシート/アプリ画面の写真を撮るだけで、走行距離と燃費(km/L)を自動記録・グラフ化するアプリです。
Googleスプレッドシートをデータの保存先にすることで、**外出先からスマホでも同じデータにアクセス**できます。

## できること

- カメラでメーターを撮影 → OCRで数値を自動抽出（手動修正も可能）
- カメラで給油レシート/給油アプリ画面を撮影 → 給油量を自動抽出（手動修正も可能）
- 前回記録との差分から、走行距離・燃費を自動計算
- 履歴を一覧・グラフ（燃費の推移、週間走行距離）で確認
- データはGoogleスプレッドシートに保存 → クラウド上にアプリを公開すれば、自宅の外・スマホからもアクセス可能
- 写真そのものは保存しません（その場でOCR読み取りに使うだけ）

## 全体の仕組み

```
スマホ / PC のブラウザ
      │  (インターネット経由)
      ▼
Streamlit Community Cloud 上のこのアプリ
      │  (Googleスプレッドシート API)
      ▼
Googleスプレッドシート（データ保存先）
```

## セットアップ手順（初回のみ）

### 1. Googleスプレッドシートとサービスアカウントを準備する

1. Googleアカウントで [Google スプレッドシート](https://sheets.google.com) を開き、新規スプレッドシートを作成する（名前は自由。例:「燃費管理データ」）。
   - URLの `https://docs.google.com/spreadsheets/d/【ここがID】/edit` の **【ここがID】** の部分を控えておく（スプレッドシートID）。
2. [Google Cloud Console](https://console.cloud.google.com/) で新しいプロジェクトを作成する。
3. 「APIとサービス」→「ライブラリ」から **Google Sheets API** を有効化する。
4. 「APIとサービス」→「認証情報」→「認証情報を作成」→「サービスアカウント」を作成する。
5. 作成したサービスアカウントの「鍵」タブから、JSON形式の鍵を作成・ダウンロードする。
6. サービスアカウントのメールアドレス（`xxxx@xxxx.iam.gserviceaccount.com` の形式）を、手順1で作ったスプレッドシートに**編集者として共有**する。

### 2. ローカルでの動作確認用に secrets.toml を作る

ダウンロードしたJSON鍵ファイルを使って、以下を実行します。

```
python json_to_secrets.py "ダウンロードしたjsonファイルのパス" "スプレッドシートID"
```

`.streamlit/secrets.toml` が自動生成されます（このファイルは`.gitignore`で除外されており、GitHubにはアップロードされません）。

### 3. 依存ライブラリをインストール

```
pip install -r requirements.txt
```

**Tesseract OCR本体**もインストールしてください（pipのライブラリだけでは動きません）。
- Windows: https://github.com/UB-Mannheim/tesseract/wiki からインストーラーを取得してインストール
- インストール先が自動でPATHに入らない場合は、`app.py`の先頭付近に以下を追記してください
  ```python
  import pytesseract
  pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
  ```

### 4. 既存データ（fuel_log.db）をスプレッドシートに移行する

同じフォルダに過去のsqliteデータ(`fuel_log.db`)がある場合、以下で一括移行できます（何度実行しても、同じ日付のデータは重複登録されません）。

```
python migrate_to_sheets.py
```

### 5. ローカルで起動確認

```
streamlit run app.py
```

`http://localhost:8501` で正常に表示され、スプレッドシート側にもデータが追加されることを確認してください。

### 6. クラウドに公開する（外出先からのアクセスに必須）

1. GitHubにこのフォルダの中身をリポジトリとしてアップロードする（`fuel_log.db`・`photos/`・`.streamlit/secrets.toml` は`.gitignore`で自動的に除外されます）。
2. [Streamlit Community Cloud](https://share.streamlit.io/) にGitHubアカウントでログインし、「New app」から上記リポジトリ・`app.py` を指定してデプロイする。
3. デプロイ設定の「Secrets」に、ローカルで作った `.streamlit/secrets.toml` の中身をそのまま貼り付ける。
4. 発行されたURL（例: `https://xxxxx.streamlit.app`）に、スマホ・PCからどこでもアクセスできます。

### 7. 既存Excel(ソリオ燃費早見表.xlsm)への自動同期

`run_app.bat`でアプリを開いたときに、Googleスプレッドシート側の未反映分をExcelへ自動反映しますが、
それとは別に、**Excelファイルを開く前に**同期しておきたい場合は`sync_now.py`を使います。

```
python sync_now.py
```

このスクリプトはブラウザを開かず、同期だけを行って終了します。デスクトップのランチャー(`.bat`)から
Excelを開く前にこれを呼ぶようにしておくと、Excelを開くたびに自動で最新化されます。

**重要**: Excelがそのファイルを開いている間は、そのファイル自体への書き込みはOSレベルでブロックされます。
そのため`sync_now.py`は**必ずExcelを開く前**に実行してください(開いた後・`Workbook_Open`から呼んでも
ロックされていて反映されません)。ExcelのVBA側(`Workbook_Open`)は、同期を実行するのではなく、この
`sync_now.py`が直前に作った同期結果(`sync_status.txt`)を読んでポップアップ表示するだけの役割です。

## 現状の制限

- **OCRの精度**: メーターのデジタル表示や光の反射があると読み取りを間違えることがあります。必ず表示された数値を確認・修正してから保存してください。
- **給油アプリとの自動連携はしていません**: 給油アプリ自体にAPIが無いことが多いため、アプリの画面を「写真として撮る」ことで代用しています。
- **写真は保存されません**: OCRでの読み取りにその場で使うだけで、記録には数値のみが残ります。
