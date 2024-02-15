import openpyxl

# パスの定義
static_path = "static"
search_method_excel_path = static_path + "/input_excel/search_method.xlsx"
output_reins_excel_path = static_path + "/output_excel/output_reins.xlsx"


def excel_to_list(input_excel_path: str = "input.xlsx"):
    workbook = openpyxl.load_workbook(input_excel_path)
    sheet = workbook.active
    row_num = sheet.max_row
    col_num = sheet.max_column
    data_list = []
    for row in range(1, row_num+1):
        row_data = []
        for col in range(1, col_num+1):
            cell_value = sheet.cell(row=row, column=col).value
            row_data.append(cell_value)
        data_list.append(row_data)
    return data_list

def list_to_excel(to_excel_list: list , output_excel_path: str = "output.xlsx"):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    row_num = len(to_excel_list)
    col_num = len(to_excel_list[0])
    for row in range(row_num):
        for col in range(col_num):
            sheet.cell(row=row+1, column=col+1).value = to_excel_list[row][col]
    workbook.save(output_excel_path)
    

def main():
    data_list = excel_to_list(search_method_excel_path)
    print(data_list)
    list_to_excel(data_list , output_reins_excel_path)


if __name__ == "__main__":
    main()