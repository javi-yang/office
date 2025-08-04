import xlwings as xw
import datetime

def extract_margin_data(input_path):
    app = xw.App(visible=False)
    wb = app.books.open(input_path)
    results = []

    for sheet in wb.sheets:
        used_range = sheet.used_range
        values = used_range.value
        if not values:
            continue
        start_row = used_range.row
        start_col = used_range.column
        rows = len(values)
        cols = len(values[0]) if rows > 0 else 0
        for row_idx, row in enumerate(values):
            for col_idx, cell in enumerate(row):
                if cell and "Margin" in str(cell):
                    excel_row = start_row + row_idx
                    excel_col = start_col + col_idx
                    # 向下查找负数并标红
                    scan_row = excel_row + 1
                    while True:
                        margin_val = sheet.range((scan_row, excel_col)).value
                        if margin_val is None:
                            break
                        try:
                            margin_num = float(margin_val)
                        except (TypeError, ValueError):
                            scan_row += 1
                            continue
                        if margin_num < 0:
                            sheet.range((scan_row, excel_col)).color = (255, 0, 0)  # 红色
                            # 向左查找第一个Frequency
                            freq_col = None
                            for c in range(excel_col - 1, start_col - 1, -1):
                                freq_val = sheet.range((excel_row, c)).value
                                if freq_val and "Frequency" in str(freq_val):
                                    freq_col = c
                                    break
                            if freq_col:
                                sheet.range((scan_row, freq_col)).color = (255, 0, 0)  # 红色
                                freq_value = sheet.range((scan_row, freq_col)).value
                                results.append([sheet.name, freq_value, margin_num])
                        scan_row += 1

    # 保存结果到新文件
    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M")
    output_path = f"D:\\ALAP\\output_{timestamp}.xlsx"
    out_app = xw.App(visible=False)
    out_wb = out_app.books.add()
    out_sheet = out_wb.sheets[0]
    out_sheet.range("A1").value = ["Sheet", "Frequency", "Margin"]
    if results:
        out_sheet.range("A2").value = results
    out_wb.save(output_path)
    out_wb.close()
    out_app.quit()

    wb.save(input_path)
    wb.close()
    app.quit()

if __name__ == "__main__":
    extract_margin_data(r"D:\ALAP\test.xlsx")