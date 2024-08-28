import openpyxl
import win32com.client as win32
from docx import Document
input_file_path = 'Book1.xlsx'
wb = openpyxl.load_workbook(input_file_path)
sheet = wb.active
for row in range(2, 20):
     print(sheet[f'B{row}'].value)
     cell_value = float(input())
     sheet[f'A{row}'] = cell_value
type_of_coding = int(input("Турбокод - 1 \nСверотный - 2 \nБез кодирования - 3 \n"))
print1,print2 =0,0
lossalgorithm = float(input(f'Пороговое отношение Еб/N0, дБ, при Р ош={sheet["A19"].value} с учетом потерь на реализацию: '))
if type_of_coding == 1:
    print1 = 35
    print2 = 42
    sheet["B22"] = lossalgorithm
elif type_of_coding == 2:
    print1 = 44
    print2 = 51
    sheet["B23"] = lossalgorithm
else:
    print1 = 53
    print2 = 61
    sheet["B24"] = lossalgorithm
distance = float(input("Дальность Д до КА, км "))
sheet["B26"] = distance
output_file_path = 'Book1_new.xlsx'
wb.save(output_file_path)
print(f"Значения успешно записаны в новый файл: {output_file_path}")
input_file_path = r"C:\Users\user\PycharmProjects\pythonProject\Book1_new.xlsx"
excel = win32.Dispatch('Excel.Application')
wb = excel.Workbooks.Open(input_file_path)
excel.Visible = False
wb.RefreshAll()
excel.CalculateFull()
sheet = wb.Sheets(1)

for i in range(27,31):
    cell_value = sheet.Range(f'B{i}').Value
    cell_value1 = sheet.Range(f'A{i}').Value
    print(f"{cell_value1} {cell_value}")
for i in range(print1,print2):
    cell_value = sheet.Range(f'B{i}').Value
    cell_value1 = sheet.Range(f'A{i}').Value
    print(f"{cell_value1} {cell_value}")



wb.Close(SaveChanges=True)


excel.Quit()

