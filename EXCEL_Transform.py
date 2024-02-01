# import pyexcel as p
#
# p.save_book_as(file_name='   .xls',dest_file_name = '   .xlsx')

import xlwings as xw

path = r'C:\Users\406239\OneDrive - (株)NHKテクノロジーズ\デスクトップ\★勤務確認などのダウンロードデータ★\NHK勤務表出力ファイル\monschedule_202312_20240104_1024.xls'

wb = xw.Book(path)

path = r'C:\Users\406239\OneDrive - (株)NHKテクノロジーズ\デスクトップ\★勤務確認などのダウンロードデータ★\NHK勤務表出力ファイル\monschedule_202312_20240104_1024.xlsx'

wb.save(path)