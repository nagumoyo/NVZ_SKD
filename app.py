import streamlit as st
from generate_schedule import run

st.title("スケジュール整形ツール")

# --- ファイルアップロード UI ---
st.sidebar.header("入力ファイル")

# スケジュールCSVアップロード
sched_file = st.sidebar.file_uploader("スケジュールCSVを選択", type=["csv"])

# 設定ファイル（PREF.xlsx）アップロード
pref_file = st.sidebar.file_uploader("設定ファイル（PREF.xlsx）を選択", type=["xlsx"])

# SIM Slot List Excelアップロード
simslot_file = st.sidebar.file_uploader(
    "SIM Slot List Excelを選択",
    type=["xlsx"],
    help="SIM訓練実績が入ったExcelファイルをアップロードしてください",
)

# --- 実行ボタン ---
if st.sidebar.button("実行"):
    # 必要ファイルのチェック
    if not sched_file or not pref_file or not simslot_file:
        st.sidebar.error(
            "スケジュールCSV／設定ファイル(PREF.xlsx)／SIM Slot List Excel のすべてをアップロードしてください。"
        )
    else:
        try:
            # run の呼び出し（emp ファイルは不要になりました）
            result = run(
                schedule_file=sched_file,
                pref_file=pref_file,
                sim_file=simslot_file,
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
            st.error(f"エラーが発生しました: {e}")
