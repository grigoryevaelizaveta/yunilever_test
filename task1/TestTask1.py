import xlwings as xw

def color_rows_by_status(file_path, sheet_name="Sheet1"):
    try:
        wb = xw.Book(file_path)
        sheet = wb.sheets[sheet_name]

        last_row = sheet.range("A" + str(sheet.cells.last_cell.row)).end("up").row 
        data_range = sheet.range(f"A2:C{last_row}") 

        for row in data_range.rows:
            status = row.value[1]  

            if status == "Done":
                row.color = (0, 255, 0) 
            elif status == "In progress":
                row.color = (255, 0, 0) 
            else:
                row.color = None 
        wb.save()
        wb.close() 
        print(f"Файл '{file_path}' успешно обработан и сохранен.")

    except FileNotFoundError:
        print(f"Ошибка: Файл '{file_path}' не найден.")
    except Exception as e:
        print(f"Произошла ошибка: {e}")

file_path = "TestTask1.xlsx" 
color_rows_by_status(file_path)
