import streamlit as st
from generate_schedule import run

st.title("スケジュール整形ツール")

# --- ファイルアップロード UI ---
st.sidebar.header("入力ファイル")
sched_file = st.sidebar.file_uploader("スケジュールCSVを選択", type=["csv"])
emp_file = st.sidebar.file_uploader("職員番号CSVを選択", type=["csv"])
pref_file = st.sidebar.file_uploader("設定ファイル（PREF.xlsx）を選択", type=["xlsx"])
simslot_file = st.sidebar.file_uploader(
    "SIM Slot List Excelを選択",
    type=["xlsx"],
    help="SIM訓練実績が入ったExcelファイルをアップロードしてください",
)

# --- 実行ボタン ---
if st.sidebar.button("実行"):
    # 全3ファイル必須
    if not sched_file or not emp_file or not simslot_file:
        st.sidebar.error(
            "スケジュールCSV／職員番号CSV／SIM Slot List Excel のすべてをアップロードしてください。"
        )
    else:
        try:
            # run 関数に simslot_file を追加で渡す
            csv_out, xlsx_out = run(
                schedule_file=sched_file,
                emp_file=emp_file,
                pref_file=pref_file,
                sim_file=simslot_file,
            )
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
