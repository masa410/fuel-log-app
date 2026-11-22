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

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

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
    /* 記録済みで押せないときのボタン（押せないことが分かる色） */
    div[data-testid="stButton"] button[kind="primary"]:disabled {
        background-color: #5a2a2a !important;
        color: #999999 !important;
        border-color: #5a2a2a !important;
    }
    /* ポップアップ：画面中央・大きい文字・大きいボタン */
    div[data-testid="stDialog"] > div { align-items: center !important; }
    div[data-testid="stDialog"] div[role="dialog"] { font-size: 1.3rem; }
    div[data-testid="stDialog"] li { font-size: 1.35rem; line-height: 1.8; }
    div[data-testid="stDialog"] p { font-size: 1.25rem; }
    div[data-testid="stDialog"] button { font-size: 1.4rem !important; padding: 0.9rem 1.2rem !important; }
    /* 履歴の装飾を消すための裏ボタン（画面には出さない。スクリプトが押す） */
    .st-key-hlclear { display: none; }
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

# 記録・修正・削除のあとのExcel用データの書き出し結果（完了メッセージ自体はポップアップで出す）
show_excel_sync_result(st.session_state.pop("last_excel_sync_result", None))


# ---------------------------------------------------------------------------
# ポップアップ（画面中央・大きく表示。ボタンでしか閉じない）
# ---------------------------------------------------------------------------
@st.dialog("✅ 記録しました", width="large", dismissible=False)
def dialog_record_done(rec):
    dist = "―" if rec["distance_km"] is None else f"{rec['distance_km']:,.0f} km"
    eff = "―" if rec["efficiency"] is None else f"{rec['efficiency']:.2f} km/L"
    st.markdown(
        "### 今回の入力が完了しました\n"
        f"- 日付：{rec['record_date']:%Y/%m/%d}\n"
        f"- メーター：{rec['odometer_km']:,.0f} km\n"
        f"- 給油量：{rec['fuel_liters']:.1f} L\n"
        f"- 走行距離：{dist}\n"
        f"- 燃費：{eff}"
    )
    st.write("履歴画面へ切り替えて今回の入力内容を確認する場合は、下のボタンをタップしてください。")
    c1, c2 = st.columns(2)
    if c1.button("📊 履歴を見る", type="primary", use_container_width=True, key="dlg_go_history"):
        st.session_state["hl_id"] = rec["id"]
        st.session_state["go_history"] = True
        st.rerun()
    if c2.button("閉じる", use_container_width=True, key="dlg_close"):
        st.rerun()


@st.dialog("⚠️ 確認してください", width="large", dismissible=False)
def dialog_warn(msg):
    st.markdown(f"### {msg}")
    if st.button("OK", type="primary", use_container_width=True, key="dlg_warn_ok"):
        st.rerun()


@st.dialog("✅ 完了", width="large", dismissible=False)
def dialog_edit_done(msg):
    st.markdown(f"### {msg}")
    if st.button("OK", type="primary", use_container_width=True, key="dlg_done_ok"):
        st.rerun()


def clear_highlight():
    st.session_state["hl_id"] = None


def on_liters_change():
    """給油量を入れ直したら、記録ボタンの鍵を外す（装飾は消す）。"""
    if st.session_state.get("f_liters", 0) > 0:
        st.session_state["locked"] = False
    clear_highlight()


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.title("🚗 燃費管理")
st.caption("メーターと給油の数値を入力するだけで、走行距離・燃費を自動記録します")

tab_add, tab_history = st.tabs(["📝 新しい記録を追加", "📊 履歴"])

# --- 記録追加タブ -----------------------------------------------------------
S = st.session_state
S.setdefault("locked", False)       # 記録した直後は True（給油量を入れ直すまで記録ボタンは使えない）
S.setdefault("hl_id", None)         # 履歴で装飾する行のID
S.setdefault("pending", None)       # 画面の最後に出すポップアップ
S.setdefault("go_history", False)   # 「履歴を見る」を押した直後だけ True
S.setdefault("form_nonce", 0)
if "f_date" not in S:
    S["reset_form"] = True

# 裏ボタン（履歴の装飾を消す）。画面には出ない。
st.button("hl", key="hlclear", on_click=clear_highlight)

with tab_add:
    last_odo = get_last_odometer()
    if S.pop("reset_form", False):
        # 記録後（および最初）の入力欄：メーターは最新値、給油量・金額・単価は0、メモは空、日付は今日
        S["f_date"] = date.today()
        S["f_odo"] = float(last_odo or 0)
        S["f_liters"] = 0.0
        S["f_cost"] = 0.0
        S["f_note"] = ""
        S["form_nonce"] += 1
    if last_odo is not None:
        st.info(f"前回のメーター値: **{last_odo:,.0f} km**")
    else:
        st.info("まだ記録がありません。最初の記録を入力してください。")

    record_date = st.date_input("日付", key="f_date", on_change=clear_highlight)

    odometer_km = st.number_input(
        "メーター表示 (km)", min_value=0.0, step=1.0, key="f_odo", on_change=clear_highlight,
    )

    fuel_liters = st.number_input(
        "給油量 (L)", min_value=0.0, step=0.1, key="f_liters", on_change=on_liters_change,
    )

    fuel_cost = st.number_input(
        "給油金額 (円・任意)", min_value=0.0, step=100.0, key="f_cost", on_change=clear_highlight,
    )

    fuel_unit_price = st.number_input(
        "給油単価 (円/L・任意。金額と給油量から自動計算されます)",
        min_value=0.0,
        value=round(fuel_cost / fuel_liters, 1) if (fuel_cost > 0 and fuel_liters > 0) else 0.0,
        step=1.0,
        key=f"f_price_{S['form_nonce']}",
        on_change=clear_highlight,
    )

    note = st.text_input("メモ（任意）", key="f_note", on_change=clear_highlight)

    if S["locked"]:
        st.caption("🔒 記録済みです。次の記録は、給油量を入力すると押せるようになります。")
    if st.button("この内容で記録する", type="primary", use_container_width=True, disabled=S["locked"]):
        S["locked"] = True  # 先に鍵をかける（二重記録の防止）
        df_now = load_records()
        if fuel_liters <= 0:
            S["locked"] = False
            S["pending"] = ("warn", "給油量を入力してください。")
        elif last_odo is not None and odometer_km < last_odo:
            S["locked"] = False
            S["pending"] = ("warn", "メーター値が前回より小さくなっています。値を確認してください。（記録していません）")
        elif record_edit.is_duplicate_of_last(df_now, record_date, odometer_km, fuel_liters):
            S["locked"] = False
            S["pending"] = ("warn", "直前の記録と同じ内容です。二重記録を防ぐため、記録していません。")
        else:
            distance_km = None
            if last_odo is not None:
                distance_km = round(odometer_km - last_odo, 1)
            efficiency = round(distance_km / fuel_liters, 2) if distance_km else None
            try:
                new_id = insert_record(
                    record_date, odometer_km, distance_km, fuel_liters,
                    fuel_unit_price if fuel_unit_price > 0 else None,
                    fuel_cost if fuel_cost > 0 else None, efficiency, note,
                )
            except Exception:
                S["locked"] = False
                S["pending"] = ("warn", "記録できませんでした。通信状態を確認して、もう一度お試しください。")
                st.rerun()
            try:
                S["last_excel_sync_result"] = excel_sync.export_records_for_excel(load_records())
            except Exception:
                S["last_excel_sync_result"] = None
            S["pending"] = ("done", {
                "id": new_id, "record_date": record_date, "odometer_km": odometer_km,
                "fuel_liters": fuel_liters, "distance_km": distance_km, "efficiency": efficiency,
            })
            S["reset_form"] = True
            st.cache_data.clear()
        st.rerun()

# --- 履歴タブ ---------------------------------------------------------------
def _refresh_excel_export():
    try:
        st.session_state["last_excel_sync_result"] = excel_sync.export_records_for_excel(load_records())
    except Exception:
        st.session_state["last_excel_sync_result"] = None


def _fmt_num(v):
    """数値の表示（小数は4桁までで末尾の0を消す。空欄は空）。"""
    if v is None or v != v:
        return ""
    return f"{v:.4f}".rstrip("0").rstrip(".")


def style_highlight(show_df, hl_pos):
    """指定の行だけ、黄色の背景＋黒の太字にした表を返す。"""
    def paint(row):
        if row.name != hl_pos:
            return [""] * len(row)
        return ["background-color:#ffe94d; color:#000; font-weight:800"] * len(row)
    num_cols = [c for c in show_df.columns if c not in ("日付", "メモ")]
    return show_df.style.apply(paint, axis=1).format({c: _fmt_num for c in num_cols}, na_rep="")


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
            st.session_state["pending"] = ("edit", msg)
            st.session_state["hl_id"] = edit_id  # 修正した行を装飾する
            st.session_state.pop("edit_id", None)
        else:
            st.session_state["pending"] = ("warn", msg)
        st.rerun()
    if c2.button("キャンセル", use_container_width=True, key="e_cancel" + k):
        st.session_state.pop("edit_id", None)
        st.rerun()
    with c3:
        confirm = st.checkbox("削除してよい", key="e_confirm" + k)
    if st.button("🗑️ この記録を削除する", disabled=not confirm, key="e_delete" + k):
        ok, msg = delete_record(edit_id)
        if ok:
            _refresh_excel_export()
            st.session_state["pending"] = ("edit", msg)
            st.session_state["hl_id"] = None
            st.session_state.pop("edit_id", None)
        else:
            st.session_state["pending"] = ("warn", msg)
        st.rerun()


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
        hl_id = st.session_state.get("hl_id")
        table = show_df
        if hl_id is not None and (ordered["id"] == hl_id).any():
            hl_pos = int((ordered["id"] == hl_id).to_numpy().nonzero()[0][0])
            table = style_highlight(show_df, hl_pos)
        event = st.dataframe(
            table,
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


# ---------------------------------------------------------------------------
# 画面の最後：タブ切替・入力欄の補助（要確認の仕掛け）とポップアップ
# ---------------------------------------------------------------------------
# 要確認の仕掛け（Streamlitの更新で動かなくなる可能性がある）
#  1. 「履歴を見る」を押した直後だけ、履歴タブを自動で開く
#  2. 「新しい記録を追加」タブを押したら、履歴の装飾を消す
#  3. 値が0の数値欄をタップしたら全選択する（そのまま入力すると置き換わる）
# 2・3 は親ページ側に1回だけ設置する（iframeが消えても働き続けるため）。
if S.get("go_history"):
    S["go_history"] = False
    components.html(
        """<script>
        const tabs = window.parent.document.querySelectorAll('button[role="tab"]');
        for (const t of tabs) { if (t.innerText.includes('履歴')) { t.click(); break; } }
        </script>""",
        height=0,
    )

components.html(
    """<script>
    const P = window.parent;
    if (!P.__fuelInit) {
      const sc = P.document.createElement('script');
      sc.textContent = `
        window.__fuelInit = true;
        document.addEventListener('click', (e) => {
          const t = e.target.closest && e.target.closest('button[role="tab"]');
          if (t && t.innerText.includes('新しい記録')) {
            const b = document.querySelector('.st-key-hlclear button');
            if (b) b.click();
          }
        }, true);
        document.addEventListener('focusin', (e) => {
          const el = e.target;
          if (el && el.tagName === 'INPUT' && el.type === 'number' && parseFloat(el.value) === 0) {
            setTimeout(() => el.select(), 0);
          }
        });`;
      P.document.head.appendChild(sc);
    }
    </script>""",
    height=0,
)

_pending = S.get("pending")
if _pending:
    S["pending"] = None
    _kind, _payload = _pending
    if _kind == "done":
        dialog_record_done(_payload)
    elif _kind == "warn":
        dialog_warn(_payload)
    else:
        dialog_edit_done(_payload)
