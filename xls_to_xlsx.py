# #import pyexcel as p
#
# #p.save_book_as(file_name= "C:\\Users\\406239\\OneDrive - (株)NHKテクノロジーズ\デスクトップ\\★勤務確認などのダウンロードデータ★\\NHK勤務表出力ファイル\\monschedule_202312.xls",dest_file_name = "C:\\Users\\406239\\OneDrive - (株)NHKテクノロジーズ\\デスクトップ\\★勤務確認などのダウンロードデータ★\\NHK勤務表出力ファイル\\monschedule_202312_20231122_1153.xlsx")
#
# #p.save_book_as(file_name= "C:\\Users\\406239\\OneDrive - (株)NHKテクノロジーズ\デスクトップ\\monschedule_202312.xls",dest_file_name = "C:\\Users\\406239\\OneDrive - (株)NHKテクノロジーズ\\デスクトップ\\★勤務確認などのダウンロードデータ★\\NHK勤務表出力ファイル\\monschedule_202312_20231122_1153.xlsx")
#
#
# from openpyxl import load_workbook
# import pyexcel
# import getopt
# import sys
#
# print('sys.argv         : ', sys.argv)
# print('type(sys.argv)   : ', type(sys.argv))
# print('len(sys.argv)    : ', len(sys.argv))
#
# print()
#
# print('sys.argv[0]      : ', sys.argv[0])
# print('sys.argv[1]      : ', sys.argv[1])
# print('sys.argv[2]      : ', sys.argv[2])
# print('type(sys.argv[0]): ', type(sys.argv[0]))
# print('type(sys.argv[1]): ', type(sys.argv[1]))
# print('type(sys.argv[2]): ', type(sys.argv[2]))
#
# #
# # xls_path =
# # xlsx_path =
#
# def main():
#     xls_path = sys.argv[1]
#     xlsx_path = xls_path + 'x'
#
#     pyexcel.save_book_as(file_name=xls_path,
#                          dest_file_name=xlsx_path)
#
#     wbook = load_workbook(xlsx_path, data_only=True)
#
#     for sheet_name in wbook.get_sheet_names():
#         print
#         '## SHEET_NAME=' + sheet_name
#         wsheet = wbook.get_sheet_by_name(sheet_name)
#
#         cell = wsheet.cell(row=8, column=11)
#         val = cell.value
#         val_2 = cell.internal_value
#         print
#         "%s,%s" % (val, val_2)
#
#         return
#
#
# if __name__ == '__main__':
#     main()


import pyexcel as p
import glob

def convert_xls_to_xlsx():
    it = glob.glob("*.xls")
    for xls in it:
        xlsx = "{}".format(xls) + "x"
        print(xlsx)
        p.save_book_as(file_name='{}'.format(xls), dest_file_name='{}'.format(xlsx))
