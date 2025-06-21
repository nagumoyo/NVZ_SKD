import streamlit as st
from generate_schedule import run

st.title("スケジュール整形ツール")
sched_file = st.file_uploader("スケジュールCSVを選択", type="csv")
emp_file = st.file_uploader("職員番号CSVを選択", type="csv")

if st.button("実行"):
    if not sched_file or not emp_file:
        st.error("両方のファイルをアップロードしてください")
    else:
        csv_out, xlsx_out = run(sched_file, emp_file)
        st.success("処理が完了しました！")
        # CSV ダウンロード
        with open(csv_out, "rb") as f:
            st.download_button("CSV をダウンロード", f, file_name=csv_out)
        # Excel ダウンロード
        with open(xlsx_out, "rb") as f:
            st.download_button("Excel をダウンロード", f, file_name=xlsx_out)
