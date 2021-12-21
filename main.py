import openpyxl


main_workbook = openpyxl.load_workbook('main_report.xlsx')
main_worksheet = main_workbook.worksheets[0]
main_worksheet.delete_rows(0, 1)
main_workbook.save('main_report.xlsx')
