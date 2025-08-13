import streamlit as st
from generate_schedule import run
from PIL import Image  # ロゴ表示用


# ① ページ設定＋背景グラデ
st.set_page_config(
    page_title="✈ NAGU 乗務割整形支援ツール",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# --- サイドバー（STAGEバッジ＋差分比較トグル） ---
st.sidebar.markdown("### 🧪 STAGE（テスト環境）")
st.sidebar.subheader("前回Excel比較（任意）")

# 比較ON/OFFのトグル
use_diff = st.sidebar.checkbox("前回Excelと比較して変更セルをハイライト", value=False)

# 比較ONのときだけアップローダを出す（← ここが1か所だけ）
prev_up = None
if use_diff:
    prev_up = st.sidebar.file_uploader(
        "前回の出力Excel (.xlsx)",
        type=["xlsx"],
        key="prev_xlsx",  # ← このkeyはアプリ全体で一意
    )


st.markdown(
    """
    <style>
    .stApp {
      background: linear-gradient(135deg, #e0f7fa 0%, #e1bee7 100%);
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ② ロゴをセンタリング表示
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    try:
        logo = Image.open("image/nagu_logo.png")  # 実際のロゴファイルパスに書き換え
        st.image(logo, use_container_width=True)
    except FileNotFoundError:
        st.write("ロゴ画像が見つかりません。")

# --- ファイルアップロード UI ---
st.sidebar.header("入力ファイル")
# スケジュールCSVアップロード
sched_file = st.sidebar.file_uploader("スケジュールCSVを選択", type=["csv"])
# 設定ファイル（PREF.xlsx）アップロード
pref_file = st.sidebar.file_uploader("設定ファイル（PREF.xlsx）を選択", type=["xlsx"])
# SIM Slot List Excelアップロード（任意）
simslot_file = st.sidebar.file_uploader(
    "SIM Slot List Excelを選択（任意）",
    type=["xlsx"],
    help="SIM訓練実績が入ったExcelファイルをアップロードしてください（未アップロード可）",
)

# --- 実行ボタン ---
# --- 実行ボタン ---
if st.sidebar.button("実行"):
    if not sched_file or not pref_file:
        st.sidebar.error("スケジュールCSVと設定ファイル(PREF.xlsx)は必須です。")
    else:
        try:
            import io

            sched_io = io.BytesIO(sched_file.getvalue())
            pref_io = io.BytesIO(pref_file.getvalue())
            sim_io = io.BytesIO(simslot_file.getvalue()) if simslot_file else ""

            # ✅ 比較ONのときだけ前回Excelを渡す
            prev_arg = (
                prev_up.getvalue() if (use_diff and prev_up is not None) else None
            )

            result = run(
                schedule_file=sched_io,
                pref_file=pref_io,
                sim_file=sim_io,
                prev_excel=prev_arg,  # ← ここが肝！
            )

            if result is None:
                st.sidebar.error(
                    "スケジュールブロックが見つかりませんでした。入力ファイルを確認してください。"
                )
            else:
                csv_out, xlsx_out = result
                st.success("処理が完了しました！")

                # CSV ダウンロード
                with open(csv_out, "rb") as f_csv:
                    st.download_button(
                        label="CSVをダウンロード",
                        data=f_csv,
                        file_name=csv_out,
                        mime="text/csv",
                    )
                # Excel ダウンロード
                with open(xlsx_out, "rb") as f_xlsx:
                    st.download_button(
                        label="Excelをダウンロード",
                        data=f_xlsx,
                        file_name=xlsx_out,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
        except Exception as e:
            st.sidebar.error(f"処理中にエラーが発生しました: {e}")
