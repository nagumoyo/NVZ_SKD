import streamlit as st
from generate_schedule import run

# スケジュール整形アプリケーション
st.title("スケジュール整形ツール")

# --- ファイルアップロード UI ---
st.sidebar.header("入力ファイル")
sched_file = st.sidebar.file_uploader("スケジュールCSVを選択", type=["csv"])
emp_file = st.sidebar.file_uploader("職員番号CSVを選択", type=["csv"])

# --- 実行ボタン ---
if st.sidebar.button("実行"):
    if not sched_file or not emp_file:
        st.sidebar.error("両方のファイルをアップロードしてください。")
    else:
        try:
            # 変換処理
            csv_out, xlsx_out = run(sched_file, emp_file)
            st.success("処理が完了しました！")
            # ダウンロードリンク
            with open(csv_out, "rb") as f_csv:
                st.download_button(
                    label="CSVをダウンロード",
                    data=f_csv,
                    file_name=csv_out,
                    mime="text/csv",
                )
            with open(xlsx_out, "rb") as f_xlsx:
                st.download_button(
                    label="Excelをダウンロード",
                    data=f_xlsx,
                    file_name=xlsx_out,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        except Exception as e:
            st.error(f"エラーが発生しました: {e}")
