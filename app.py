"""
燃費管理アプリ
--------------------------------
- メーター写真 / 給油レシート写真 を撮影・アップロード
- OCRで数値を自動抽出（うまく読めない場合は手動で修正可能）
- 前回記録との差分から走行距離・燃費(km/L)を自動計算
- 履歴をグラフで確認
- データはGoogleスプレッドシートに保存 → PC・スマホ・外出先どこからでも同じデータにアクセス可能

起動方法・セットアップは同じフォルダの README.md を参照してください。
"""

import re
from datetime import date

import plotly.express as px
import pytesseract
import streamlit as st
from PIL import Image

import excel_sync
from sheets_db import get_last_odometer, insert_record, load_records, delete_record

st.set_page_config(page_title="燃費管理", page_icon="🚗", layout="centered")


def show_excel_sync_result(result):
    """Excel同期の結果をユーザーに表示する（クラウド環境では何も表示しない）。"""
    if not result or result["status"] == "not_configured":
        return
    if result["status"] == "locked":
        st.warning("⚠️ Excelファイルが開いているため、Excelへの反映をスキップしました。閉じてから開き直すと反映されます。")
    elif result["status"] == "synced":
        if result["added"] > 0:
            st.success(f"📗 既存Excelにも{result['added']}件反映しました。")
        else:
            st.caption("📗 Excelは最新の状態です。")


# 自宅PCでこのアプリを開いたときだけ、既存Excel(ソリオ燃費早見表.xlsm)に
# 前回までの未反映分をまとめて反映する（クラウド上やExcelが開いている場合は何もしない）。
if "excel_catchup_done" not in st.session_state:
    try:
        catchup_result = excel_sync.sync_records_to_excel(load_records())
    except Exception:
        catchup_result = None
    show_excel_sync_result(catchup_result)
    st.session_state["excel_catchup_done"] = True

if st.session_state.pop("record_added", False):
    st.success("記録しました！「履歴・グラフ」タブで確認できます。")
    show_excel_sync_result(st.session_state.pop("last_excel_sync_result", None))


# ---------------------------------------------------------------------------
# OCR ヘルパー
# ---------------------------------------------------------------------------
def ocr_best_number(image: Image.Image, decimals=False):
    """画像から一番それらしい数値を推定して返す（読めなければ None）。"""
    config = "--psm 6 -c tessedit_char_whitelist=0123456789.,"
    try:
        text = pytesseract.image_to_string(image, config=config)
    except Exception:
        return None, ""
    candidates = re.findall(r"\d[\d,]*\.?\d*", text)
    candidates = [c.replace(",", "") for c in candidates if c.replace(",", "").strip(".")]
    if not candidates:
        return None, text
    # 一番桁数の多い（＝メーター値やレシート数量らしい）ものを採用
    best = max(candidates, key=len)
    try:
        value = float(best) if (decimals or "." in best) else int(best)
        return value, text
    except ValueError:
        return None, text


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.title("🚗 燃費管理")
st.caption("メーターと給油の写真を撮るだけで、走行距離・燃費を自動記録します")

tab_add, tab_history = st.tabs(["📷 新しい記録を追加", "📊 履歴・グラフ"])

# --- 記録追加タブ -----------------------------------------------------------
with tab_add:
    last_odo = get_last_odometer()
    if last_odo is not None:
        st.info(f"前回のメーター値: **{last_odo:,.0f} km**")
    else:
        st.info("まだ記録がありません。最初の記録を入力してください。")

    record_date = st.date_input("日付", value=date.today())

    st.subheader("① 走行距離メーター")
    odo_input_mode = st.radio("入力方法", ["カメラで撮影", "写真をアップロード"], horizontal=True, key="odo_mode")
    if odo_input_mode == "カメラで撮影":
        odo_file = st.camera_input("メーターを撮影", key="odo_camera")
    else:
        odo_file = st.file_uploader("メーター写真を選択", type=["jpg", "jpeg", "png"], key="odo_upload")

    odo_guess = None
    if odo_file is not None:
        odo_img = Image.open(odo_file)
        odo_guess, odo_raw_text = ocr_best_number(odo_img)
        with st.expander("OCR結果を確認（読み違えていたら下の欄で修正してください）"):
            st.text(odo_raw_text or "(文字を検出できませんでした)")

    odometer_km = st.number_input(
        "メーター表示 (km) ※OCRの自動読み取り結果。違っていたら修正してください",
        min_value=0.0, value=float(odo_guess) if odo_guess else float(last_odo or 0), step=1.0,
    )

    st.subheader("② 給油レシート / 給油アプリ画面")
    fuel_input_mode = st.radio("入力方法", ["カメラで撮影", "写真をアップロード"], horizontal=True, key="fuel_mode")
    if fuel_input_mode == "カメラで撮影":
        fuel_file = st.camera_input("レシート・アプリ画面を撮影", key="fuel_camera")
    else:
        fuel_file = st.file_uploader("レシート・アプリ画面の写真を選択", type=["jpg", "jpeg", "png"], key="fuel_upload")

    fuel_guess = None
    if fuel_file is not None:
        fuel_img = Image.open(fuel_file)
        fuel_guess, fuel_raw_text = ocr_best_number(fuel_img, decimals=True)
        with st.expander("OCR結果を確認（読み違えていたら下の欄で修正してください）"):
            st.text(fuel_raw_text or "(文字を検出できませんでした)")

    col1, col2 = st.columns(2)
    with col1:
        fuel_liters = st.number_input(
            "給油量 (L) ※OCRの自動読み取り結果。違っていたら修正してください",
            min_value=0.0, value=float(fuel_guess) if fuel_guess else 0.0, step=0.1,
        )
    with col2:
        fuel_cost = st.number_input("給油金額 (円・任意)", min_value=0.0, value=0.0, step=100.0)

    fuel_unit_price = st.number_input(
        "給油単価 (円/L・任意。金額と給油量から自動計算されます)",
        min_value=0.0,
        value=round(fuel_cost / fuel_liters, 1) if (fuel_cost > 0 and fuel_liters > 0) else 0.0,
        step=1.0,
    )

    note = st.text_input("メモ（任意）")

    if st.button("この内容で記録する", type="primary"):
        if fuel_liters <= 0:
            st.error("給油量を入力してください。")
        else:
            distance_km = None
            if last_odo is not None:
                distance_km = round(odometer_km - last_odo, 1)
                if distance_km < 0:
                    st.warning("メーター値が前回より小さくなっています。値を確認してください。距離は保存しません。")
                    distance_km = None
            efficiency = round(distance_km / fuel_liters, 2) if distance_km else None

            insert_record(
                record_date, odometer_km, distance_km, fuel_liters,
                fuel_unit_price if fuel_unit_price > 0 else None,
                fuel_cost if fuel_cost > 0 else None, efficiency, note,
            )
            try:
                st.session_state["last_excel_sync_result"] = excel_sync.sync_records_to_excel(load_records())
            except Exception:
                st.session_state["last_excel_sync_result"] = None
            st.session_state["record_added"] = True
            st.cache_data.clear()
            st.rerun()

# --- 履歴タブ ---------------------------------------------------------------
with tab_history:
    df = load_records()
    if df.empty:
        st.write("まだ記録がありません。")
    else:
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("累計走行距離", f"{df['distance_km'].sum():,.0f} km")
        col2.metric("累計給油量", f"{df['fuel_liters'].sum():,.1f} L")
        avg_eff = df["efficiency_km_per_l"].dropna()
        col3.metric("平均燃費", f"{avg_eff.mean():.2f} km/L" if len(avg_eff) else "―")
        col4.metric("累計給油金額", f"¥{df['fuel_cost'].sum():,.0f}")

        st.subheader("燃費の推移")
        eff_df = df.dropna(subset=["efficiency_km_per_l"])
        if not eff_df.empty:
            fig = px.line(eff_df, x="record_date", y="efficiency_km_per_l", markers=True,
                           labels={"record_date": "日付", "efficiency_km_per_l": "燃費 (km/L)"})
            st.plotly_chart(fig, use_container_width=True)

        st.subheader("週間走行距離")
        dist_df = df.dropna(subset=["distance_km"])
        if not dist_df.empty:
            fig2 = px.bar(dist_df, x="record_date", y="distance_km",
                           labels={"record_date": "日付", "distance_km": "走行距離 (km)"})
            st.plotly_chart(fig2, use_container_width=True)

        st.subheader("記録一覧")
        show_df = df[["record_date", "odometer_km", "distance_km", "fuel_liters",
                       "fuel_unit_price", "fuel_cost", "efficiency_km_per_l", "note"]].copy()
        show_df.columns = ["日付", "メーター(km)", "走行距離(km)", "給油量(L)", "単価(円/L)", "金額(円)", "燃費(km/L)", "メモ"]
        st.dataframe(show_df.sort_values("日付", ascending=False), use_container_width=True, hide_index=True)

        with st.expander("記録を削除する"):
            del_id = st.selectbox("削除する記録のID", df["id"].tolist())
            if st.button("削除する", type="secondary"):
                delete_record(del_id)
                st.success("削除しました。")
                st.rerun()
