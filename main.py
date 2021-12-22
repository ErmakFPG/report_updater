import openpyxl
from openpyxl.styles import PatternFill


main_workbook = openpyxl.load_workbook('ремонт_2021.xlsx')
main_worksheet = main_workbook.worksheets[1]
child_workbook = openpyxl.load_workbook('ГЗ_ТА.xlsx')
child_worksheet = child_workbook.worksheets[3]

index = 1
for i in range(1, 100):
    if child_worksheet["A" + str(i)].value == index:
        index += 1
        current_child_event = "".join(child_worksheet["B" + str(i)].value.split())
        for j in range(7, 54):
            current_main_event = "".join(main_worksheet["B" + str(j)].value.split())
            if current_main_event == current_child_event:
                main_worksheet["E" + str(j)] = child_worksheet["G" + str(i)].value
                print("E" + str(j))
                child_worksheet["A" + str(i)].fill =\
                    PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid')
                main_worksheet["A" + str(j)].fill =\
                    PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid')
                break


main_workbook.save('ремонт_2021.xlsx')
child_workbook.save('ГЗ_ТА.xlsx')
