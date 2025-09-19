import xlwings as xw
ex_content=[]

def generate_prog_data(sheet):
    global ex_content
    result = []
    prefix = "PROG:DATA "  # 預設前綴字串
   # app = xw.App(visible=False, add_book=False)
    try:
      #  wb = app.books.open(filename)
        #sht = wb.sheets[sheetname]

        # 找出 B 欄最後一列
        last_row = sheet.cells(1, "B").end("down").row

        # 從 B2:C(最後) 取值
        arr = sheet.range(f"B2:C{last_row}").value

        for row in arr:
            if row[0] is not None and row[1] is not None:
                # 轉換為整數（避免 3.0）
                b_val = int(row[0]) if float(row[0]).is_integer() else row[0]
                c_val = int(row[1]) if float(row[1]).is_integer() else row[1]
                ex_content.append(c_val)
                # 拼接指令字串
                new_str = prefix + f"{b_val},LIST,{c_val},0,{c_val}"
                result.append(new_str)
    finally:
       # wb.close()
       # app.quit()

     return result




def generate_prog_data2(sheet, row_count=None):
    """
    sheet: xlwings 的工作表物件 (例如 wb.sheets[0] 或 wb.sheets[1])
    row_count: 從 B2 開始要讀取的列數。若為 None，會自動偵測 B 欄最後一列。
    """
    prefix = "PROG:DATA:LIST "
    result = []

    # 自動偵測列數（從 B2 向下到最後一筆連續資料）
    if row_count is None:
        last_row = sheet.cells(2, "B").end("down").row  # 回到最後一筆
        row_count = max(0, last_row - 1)  # 扣掉表頭 B1

    if row_count <= 0:
        return result

    # 固定讀 B~F 共 5 欄，從 B2 開始往下 row_count 列
    arr = sheet.range("B2").resize(row_count, 5).value

    for row in arr:
        # row = [B, C, D, E, F]
        if not row or any(v is None for v in row):
            continue

        vals = []
        for v in row:
            if isinstance(v, (int, float)) and float(v).is_integer():
                vals.append(str(int(v)))
            else:
                vals.append(str(v).strip())

        # 組字串
        new_str = (f"{prefix} {vals[0]},{vals[1]},AUTO,CC,2,"
                   f"{vals[2]},{vals[3]},{vals[3]},{vals[4]},-1,-1,-1,-1,-1,-1,1")
        result.append(new_str)

    return result