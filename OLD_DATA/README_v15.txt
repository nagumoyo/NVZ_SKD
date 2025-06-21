
【整形済スケジュール v15 スクリプト】

■ ファイル一覧
- generate_schedule_v15.py ... スクリプト本体
- emp_no.csv ... 職員情報CSV
- スケジュールCSV（例: SR-FD-051_BG_exe__PDF__MONTHLY_CREW_SCHEDULE_LIST.csv）

■ 使い方（ターミナル or コマンドプロンプト）

python generate_schedule_v15.py --schedule_csv "スケジュール.csv" --emp_csv "emp_no.csv" --output "整形スケジュール_v15.xlsx"

例：
python generate_schedule_v15.py --schedule_csv SR-FD-051_BG_exe__PDF__MONTHLY_CREW_SCHEDULE_LIST.csv --emp_csv emp_no.csv --output 整形スケジュール_v15_0601.xlsx

■ 出力
- v15仕様の Excel ファイル（31列固定、ヘッダー → 日付 → スケジュール順）

■ 処理内容
- OB除去
- 誤検知ヘッダー除去
- 空行対応
- 最終クルー対応
- セル内改行対応

■ 備考
- Python3 + pandas 必須
- Windows/Mac OK

---

もし分からない点があれば ChatGPT に「v15 スクリプトを呼び出したい」と言えばOKです。
