import csv, openpyxl
import json

UNIT_PRICE = 4000
START_ROW_NUMBER = 15
END_ROW_NUMBER = 27
INSERT_START_ROW_NUMBER = 28
SUB_TOTAL_ROW_NUMBER = 28
TOTAL_ROW_NUMBER = 29
PAYEE_NUMBER = 33

def load_input_csv():
    i = 0
    filename = 'tmp/input.csv'
    fixedRows = []
    with open(filename, encoding='utf8', newline='') as f:
        csvreader = csv.reader(f)
        for row in csvreader:
            i += 1
            if i < 4:
                continue
            if len(row) != 6:
                continue
            fixedRows.append(
                [row[1], row[2], '時間', UNIT_PRICE, float(row[2]) * UNIT_PRICE]
            )
    return fixedRows

def append_rows(ws, insert_count):
    if insert_count < 1:
        return ws
    # もともとのマージされた行を解除
    # 小計
    ws.unmerge_cells("A28:M28")
    ws.unmerge_cells("N28:P28")
    # 合計
    ws.unmerge_cells("A29:M29")
    ws.unmerge_cells("N29:P29")

    side = openpyxl.styles.borders.Side(style='thin', color='000000')
    border_aro1 = openpyxl.styles.borders.Border(top=side, bottom=side, left=side)
    border_aro2 = openpyxl.styles.borders.Border(top=side, bottom=side)
    border_aro3 = openpyxl.styles.borders.Border(top=side, bottom=side, left=side, right=side)

    copy_font = ws.cell(row=START_ROW_NUMBER, column=1).font
    alignment = openpyxl.styles.Alignment(horizontal="centerContinuous")
    rightAlignment = openpyxl.styles.Alignment(horizontal="right")

    # 追加の行を挿入
    end_insert_row_number = END_ROW_NUMBER + insert_count
    for i in range(INSERT_START_ROW_NUMBER, end_insert_row_number): 
        ws.insert_rows(i)

        #A ~ Gを結合
        ws.merge_cells("A%d:G%d" % (i, i))
        ws["A%d" % i].border = border_aro1
        ws["A%d" % i].font = copy_font._StyleProxy__target
        ws["A%d" % i].alignment = alignment
        for key in ["B", "C", "D", "E", "F"]:
            ws["%s%d" % (key, i)].border = border_aro2
        ws["G%d" % i].border = border_aro3

        for keys in [["H", "I"], ["J", "K"], ["L", "M"]]:
            ws.merge_cells("%s%d:%s%d" % (keys[0], i, keys[1], i))
            ws["%s%d" % (keys[0], i)].border = border_aro1
            ws["%s%d" % (keys[1], i)].border = border_aro3
            ws["%s%d" % (keys[0], i)].font = copy_font._StyleProxy__target
            ws["%s%d" % (keys[0], i)].alignment = alignment

        #N ~ Pを結合
        ws.merge_cells("N%d:P%d" % (i, i))
        ws["N%d" % i].border = border_aro1
        ws["O%d" % i].border = border_aro2
        ws["P%d" % i].border = border_aro3
        ws["N%d" % i].font = copy_font._StyleProxy__target
        ws["N%d" % i].alignment = rightAlignment

    
    sum_format = "=SUM(N15:N%d)" % (end_insert_row_number - 1)
    # 小計の更新
    sub_total_insert_row_number = SUB_TOTAL_ROW_NUMBER+insert_count - 1
    ws.cell(row=sub_total_insert_row_number, column=1).value = "小計（内税）"
    ws.cell(row=sub_total_insert_row_number, column=14).value = sum_format
    ws.merge_cells("A%d:M%d" % (sub_total_insert_row_number, sub_total_insert_row_number))
    ws.merge_cells("N%d:P%d" % (sub_total_insert_row_number, sub_total_insert_row_number))

    # 合計の更新
    total_insert_row = TOTAL_ROW_NUMBER+insert_count - 1
    ws.cell(row=total_insert_row, column=1).value = "合計"
    ws.cell(row=total_insert_row, column=14).value = sum_format
    ws.merge_cells("A%d:M%d" % (total_insert_row, total_insert_row))
    ws.merge_cells("N%d:P%d" % (total_insert_row, total_insert_row))

    # 請求金額の参照を更新
    ws.cell(row=11, column=4).value = "=N%d" % total_insert_row

    return ws

def insert_values(ws, fixed_rows):
    insert_count = len(fixed_rows) - (END_ROW_NUMBER - START_ROW_NUMBER)
    end_insert_row_number = END_ROW_NUMBER + insert_count
    i = 0
    for row_number in range(START_ROW_NUMBER, end_insert_row_number):
        insert_values = fixed_rows[i]
        ws.cell(row=row_number, column=1).value = insert_values[0]
        ws.cell(row=row_number, column=8).value = insert_values[1]
        ws.cell(row=row_number, column=10).value = insert_values[2]
        ws.cell(row=row_number, column=12).value = insert_values[3]
        ws.cell(row=row_number, column=12).number_format = "#,##0"
        ws.cell(row=row_number, column=14).value = insert_values[4]
        ws.cell(row=row_number, column=14).number_format = "#,##0"

        i += 1


if __name__ == "__main__":
    fixed_rows = load_input_csv()
    wb = openpyxl.load_workbook('tmp/template.xlsx')
    ws = wb["Sheet1"]

    config = {}
    with open('config/config.json') as f:
        config = json.load(f)

    # 請求情報を書き込んでいく
    ws.cell(row=7, column=4).value = config["件名"]
    ws.cell(row=8, column=4).value = config["お支払い期限"]
    ws.cell(row=1, column=12).value = "請求日:%s" % config["請求日"]
    ws.cell(row=3, column=13).value = config["郵便番号"]
    ws.cell(row=4, column=13).value = config["住所"]
    ws.cell(row=5, column=13).value = config["氏名"]
    ws.cell(row=7, column=13).value = config[ "電話番号"]
    ws.cell(row=9, column=13).value = config["メールアドレス"]

    insert_count = len(fixed_rows) - (END_ROW_NUMBER - START_ROW_NUMBER)
    append_rows(ws, insert_count)
    insert_values(ws, fixed_rows)

    bank_row = PAYEE_NUMBER + insert_count - 1
    bank_number_row = bank_row + 1
    bank_name_row = bank_row + 2
    ws.cell(row=bank_row, column=1).value = config["銀行"]
    ws.cell(row=bank_number_row, column=1).value = config["口座番号"]
    ws.cell(row=bank_name_row, column=1).value = "口座名義 %s" % config["口座名義"]

    wb.save("tmp/output.xlsx")
    wb.close()