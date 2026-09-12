"""
燃費管理アプリ
--------------------------------
- メーター表示・給油量・単価をシンプルなフォームで入力
- 前回記録との差分から走行距離・燃費(km/L)を自動計算
- 履歴をグラフで確認
- データはGoogleスプレッドシートに保存 → PC・スマホ・外出先どこからでも同じデータにアクセス可能

起動方法・セットアップは同じフォルダの README.md を参照してください。
"""

from datetime import date

import plotly.express as px
import streamlit as st

import excel_sync
from sheets_db import get_last_odometer, insert_record, load_records, delete_record

st.set_page_config(page_title="燃費管理", page_icon="🚗", layout="centered")

# 入力欄・ボタンを大きく見やすくするための共通スタイル
st.markdown(
    """
    <style>
    div[data-testid="stNumberInput"] input,
    div[data-testid="stTextInput"] input,
    div[data-testid="stDateInput"] input {
        font-size: 1.4rem !important;
        padding: 0.6rem 0.75rem !important;
        height: auto !important;
    }
    div[data-testid="stNumberInput"] label,
    div[data-testid="stTextInput"] label,
    div[data-testid="stDateInput"] label {
        font-size: 1.05rem !important;
        font-weight: 600 !important;
    }
    div[data-testid="stButton"] button {
        font-size: 1.2rem !important;
        padding: 0.6rem 1.2rem !important;
        height: auto !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def check_password() -> bool:
    """パスワードを確認する。一度正しく入力すれば、同じブラウザセッション内は再入力不要。"""
    if st.session_state.get("authenticated", False):
        return True

    st.title("🚗 燃費管理")
    with st.form("password_form"):
        col1, col2 = st.columns([5, 1])
        with col1:
            pwd = st.text_input("パスワード", type="password")
        with col2:
            st.write("")
            submitted = st.form_submit_button("→", use_container_width=True)

    if submitted:
        if pwd == st.secrets.get("app_password", ""):
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.session_state["authenticated"] = False

    if st.session_state.get("authenticated") is False:
        st.error("パスワードが違います。")
    return False


if not check_password():
    st.stop()


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
# UI
# ---------------------------------------------------------------------------
st.title("🚗 燃費管理")
st.caption("メーターと給油の数値を入力するだけで、走行距離・燃費を自動記録します")

tab_add, tab_history = st.tabs(["📝 新しい記録を追加", "📊 履歴・グラフ"])

# --- 記録追加タブ -----------------------------------------------------------
with tab_add:
    last_odo = get_last_odometer()
    if last_odo is not None:
        st.info(f"前回のメーター値: **{last_odo:,.0f} km**")
    else:
        st.info("まだ記録がありません。最初の記録を入力してください。")

    record_date = st.date_input("日付", value=date.today())

    odometer_km = st.number_input(
        "メーター表示 (km)",
        min_value=0.0, value=float(last_odo or 0), step=1.0,
    )

    fuel_liters = st.number_input(
        "給油量 (L)",
        min_value=0.0, value=0.0, step=0.1,
    )

    fuel_cost = st.number_input("給油金額 (円・任意)", min_value=0.0, value=0.0, step=100.0)

    fuel_unit_price = st.number_input(
        "給油単価 (円/L・任意。金額と給油量から自動計算されます)",
        min_value=0.0,
        value=round(fuel_cost / fuel_liters, 1) if (fuel_cost > 0 and fuel_liters > 0) else 0.0,
        step=1.0,
    )

    note = st.text_input("メモ（任意）")

    if st.button("この内容で記録する", type="primary", use_container_width=True):
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
