import xlrd
from datetime import datetime

def excel_birthdays(file_name):
    wb = xlrd.open_workbook(file_name)
    ws = wb.sheet_by_index(0)
    flag = 0
    list_of_adresses = []
    for rownum in range(ws.nrows)[1:]:
        excel_row = ws.row_values(rownum)
        for num, cell in enumerate(excel_row):
            if cell == "":
                pass
            else:
                if not flag:
                    flag = 1
                    break
                try:
                    birthday = datetime(*xlrd.xldate_as_tuple(excel_row[num + 2], wb.datemode)).strftime("%Y-%m-%d")
                except:
                    print("error date format in {}, {} row, {} cell. Must be dd.mm.yyyy (например 21.06.1990)".format(file_name, rownum, num + 3))
                    break
                today_date = datetime.today().strftime("%Y-%m-%d")
                if birthday == today_date:
                    list_of_adresses.append(excel_row[num:])
                break
    return (list_of_adresses)

if __name__ == "__main__":
    adr = excel_birthdays("ДР_ДСР.xlsx")
    print(adr)

