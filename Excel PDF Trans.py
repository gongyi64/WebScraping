#Excelをpdfに変換（wINのみ）
import win32com.client    # win32comのインポート

output_time = 20231122_1153

excel_path = fr"C:\Users\406239\OneDrive - (株)NHKテクノロジーズ\デスクトップ\★勤務確認などのダウンロードデータ★\NHK勤務表出力ファイル\monschedule_202312_{output_time}.xls"
pdf_path = fr"C:\Users\406239\OneDrive - (株)NHKテクノロジーズ\デスクトップ\★勤務確認などのダウンロードデータ★\NHK勤務表出力ファイル\monschedule_202312_{output_time}.pdf"

print(excel_path)

excel = win32com.client.Dispatch("Excel.Application")    # Excelの起動
file = excel.Workbooks.Open(excel_path)    # Excelファイルを開く
file.WorkSheets("Worksheet").Activate()    # Sheetをシート名でアクティベイト
file.ActiveSheet.ExportAsFixedFormat(0, pdf_path)    # PDFに変換
file.Close()    # 開いたエクセルを閉じる
excel.Quit()    # Excelを終了