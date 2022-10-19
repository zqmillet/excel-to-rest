from excel_to_rest import excel_to_rest

def test_excel_to_rest():
    print()
    print(excel_to_rest('./testcases/statics/workbook_1.xlsx', 'Sheet1'))
    print(excel_to_rest('./testcases/statics/workbook_1.xlsx', 'Sheet2'))
