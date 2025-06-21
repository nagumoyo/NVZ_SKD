import streamlit as st
from generate_schedule import run  # モジュール化した関数名に合わせてください

st.title("スケジュール整形ツール")
sched = st.file_uploader("スケジュールCSV", type="csv")
emp   = st.file_uploader("職員番号CSV", type="csv")
if st.button("実行"):
    out_path = run(sched, emp)
    st.success("処理が完了しました！")
    st.download_button("結果をダウンロード", open(out_path, "rb"), file_name="schedule.xlsx")
