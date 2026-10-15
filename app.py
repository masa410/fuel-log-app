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

import streamlit as st

import excel_check
import excel_sync
import record_edit
from sheets_db import (
    delete_record, get_last_odometer, insert_record, load_records, update_record,
)

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
    div[data-testid="stButton"] button,
    div[data-testid="stFormSubmitButton"] button {
        font-size: 1.2rem !important;
        padding: 0.6rem 1.2rem !important;
        height: auto !important;
    }
    div[data-testid="stButton"] button[kind="primary"],
    div[data-testid="stFormSubmitButton"] button[kind="primary"] {
        background-color: #e02424 !important;
        color: #ffffff !important;
        border-color: #e02424 !important;
    }
    div[data-testid="stDataFrame"] {
        font-size: 1.1rem !important;
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
        col1, col2 = st.columns([3, 2])
        with col1:
            pwd = st.text_input("パスワード", type="password")
        with col2:
            st.write("")
            submitted = st.form_submit_button("ログイン", type="primary", use_container_width=True)

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
    """Excel用データの書き出し結果を表示する（クラウド環境では何も表示しない）。"""
    if not result or result["status"] == "not_configured":
        return
    if result["status"] == "error":
        st.warning("⚠️ Excel用のデータを書き出せませんでした。Excelへは反映されません。")
    elif result["status"] == "exported":
        st.caption(
            f"📗 Excel用の最新データ（{result['count']}件）を書き出しました。"
            "Excelを開く（または「修正内容反映」ボタン）と反映されます。"
        )


# 自宅PCでこのアプリを開いたときだけ、既存Excel(ソリオ燃費早見表.xlsm)へ渡す
# 最新データのCSVを書き出す（クラウド上では何もしない）。Excelへの書き込みはExcel側のマクロが行う。
if "excel_catchup_done" not in st.session_state:
    try:
        catchup_result = excel_sync.export_records_for_excel(load_records())
    except Exception:
        catchup_result = None
    show_excel_sync_result(catchup_result)
    st.session_state["excel_catchup_done"] = True

if st.session_state.pop("record_added", False):
    st.success("記録しました！「履歴」タブで確認できます。")
    show_excel_sync_result(st.session_state.pop("last_excel_sync_result", None))

_done_msg = st.session_state.pop("edit_done_message", None)
if _done_msg:
    st.success(_done_msg)
    show_excel_sync_result(st.session_state.pop("last_excel_sync_result", None))


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.title("🚗 燃費管理")
st.caption("メーターと給油の数値を入力するだけで、走行距離・燃費を自動記録します")

tab_add, tab_history = st.tabs(["📝 新しい記録を追加", "📊 履歴"])

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
                st.session_state["last_excel_sync_result"] = excel_sync.export_records_for_excel(load_records())
            except Exception:
                st.session_state["last_excel_sync_result"] = None
            st.session_state["record_added"] = True
            st.cache_data.clear()
            st.rerun()

# --- 履歴タブ ---------------------------------------------------------------
def _refresh_excel_export():
    try:
        st.session_state["last_excel_sync_result"] = excel_sync.export_records_for_excel(load_records())
    except Exception:
        st.session_state["last_excel_sync_result"] = None


def render_edit_form(df, edit_id):
    """選んだ記録の修正・削除フォーム。範囲外の入力は書き込まずに理由を表示する。"""
    row = df[df["id"] == edit_id]
    if row.empty:
        st.session_state.pop("edit_id", None)
        st.rerun()
    row = row.iloc[0]
    b = record_edit.edit_bounds(df, edit_id)
    min_d = b["min_date"] or date(2000, 1, 1)

    st.subheader("✏️ 記録を修正する")
    st.caption(
        f"日付は {min_d:%Y/%m/%d} ～ {b['max_date']:%Y/%m/%d} の範囲で変更できます。"
        + (f" メーター値は {b['min_odo']:,.0f} km 以上" if b["min_odo"] is not None else "")
        + (f"、{b['max_odo']:,.0f} km 以下" if b["max_odo"] is not None else "")
        + ("。" if (b["min_odo"] is not None or b["max_odo"] is not None) else "")
    )
    k = f"_{edit_id}"
    e_date = st.date_input("日付", value=row["record_date"].date(), min_value=min_d,
                           max_value=b["max_date"], key="e_date" + k)
    e_odo = st.number_input("メーター表示 (km)", min_value=0.0, value=float(row["odometer_km"]),
                            step=1.0, key="e_odo" + k)
    e_liters = st.number_input("給油量 (L)", min_value=0.0, value=float(row["fuel_liters"]),
                               step=0.1, key="e_liters" + k)
    cost0 = 0.0 if row["fuel_cost"] != row["fuel_cost"] else float(row["fuel_cost"])
    price0 = 0.0 if row["fuel_unit_price"] != row["fuel_unit_price"] else float(row["fuel_unit_price"])
    e_cost = st.number_input("給油金額 (円・任意)", min_value=0.0, value=cost0, step=100.0, key="e_cost" + k)
    e_price = st.number_input("給油単価 (円/L・任意)", min_value=0.0, value=price0, step=1.0, key="e_price" + k)
    e_note = st.text_input("メモ（任意）", value=row["note"] if isinstance(row["note"], str) else "",
                           key="e_note" + k)

    c1, c2, c3 = st.columns(3)
    if c1.button("この内容で修正する", type="primary", use_container_width=True, key="e_save" + k):
        ok, msg = update_record(edit_id, {
            "record_date": e_date, "odometer_km": e_odo, "fuel_liters": e_liters,
            "fuel_unit_price": e_price if e_price > 0 else None,
            "fuel_cost": e_cost if e_cost > 0 else None, "note": e_note,
        })
        if ok:
            _refresh_excel_export()
            st.session_state["edit_done_message"] = msg
            st.session_state.pop("edit_id", None)
            st.rerun()
        else:
            st.error(msg)
            st.toast(msg, icon="⚠️")
    if c2.button("キャンセル", use_container_width=True, key="e_cancel" + k):
        st.session_state.pop("edit_id", None)
        st.rerun()
    with c3:
        confirm = st.checkbox("削除してよい", key="e_confirm" + k)
    if st.button("🗑️ この記録を削除する", disabled=not confirm, key="e_delete" + k):
        ok, msg = delete_record(edit_id)
        if ok:
            _refresh_excel_export()
            st.session_state["edit_done_message"] = msg
            st.session_state.pop("edit_id", None)
            st.rerun()
        else:
            st.error(msg)


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

        st.subheader("📋 記録一覧")
        st.caption("修正・削除したい行を選んで、下の「選択した記録を修正する」を押してください。")
        ordered = df.sort_values(["record_date", "id"], ascending=False).reset_index(drop=True)
        show_df = ordered[["record_date", "odometer_km", "distance_km", "fuel_liters",
                           "fuel_unit_price", "fuel_cost", "efficiency_km_per_l", "note"]].copy()
        show_df["record_date"] = show_df["record_date"].dt.strftime("%Y/%m/%d")
        show_df.columns = ["日付", "メーター(km)", "走行距離(km)", "給油量(L)", "単価(円/L)", "金額(円)", "燃費(km/L)", "メモ"]
        event = st.dataframe(
            show_df,
            use_container_width=True,
            hide_index=True,
            height=600,
            on_select="rerun",
            selection_mode="single-row",
            column_config={
                "メモ": st.column_config.TextColumn("メモ", width=200),
            },
        )
        picked = event.selection.rows if event and event.selection else []
        if picked:
            sel = ordered.iloc[picked[0]]
            st.caption(f"選択中：{sel['record_date']:%Y/%m/%d}　{sel['odometer_km']:,.0f} km　{sel['fuel_liters']:.1f} L")
        else:
            st.caption("※ 行の左端のチェック（□）を押して選ぶと、下の「修正する」ボタンが押せます。")
        if st.button("✏️ 選択した記録を修正する", disabled=not picked, type="primary"):
            st.session_state["edit_id"] = int(ordered.iloc[picked[0]]["id"])
            st.rerun()

        # 修正フォームは、ボタンのすぐ下に出す（一覧の下までスクロールしていても画面の外にならない）
        edit_id = st.session_state.get("edit_id")
        if edit_id is not None:
            st.divider()
            render_edit_form(df, edit_id)

        if excel_sync.is_available():
            with st.expander("Excelとの確認（読み取りのみ・何も書き換えません）"):
                if st.button("Excelの内容と照合する"):
                    st.session_state["excel_check_result"] = excel_check.compare(df)
                res = st.session_state.get("excel_check_result")
                if res:
                    st.write(res["summary"])
                    if res["rows"]:
                        st.dataframe(res["rows"], use_container_width=True, hide_index=True)
