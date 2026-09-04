from pathlib import Path

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from rich import print
from rich.markdown import Markdown

ROOT_FOLDER = Path(__file__).parent
WORKBOOKS_FOLDER = ROOT_FOLDER / "workbooks"
Path.mkdir(WORKBOOKS_FOLDER, exist_ok=True)

WORKBOOK_PATH = WORKBOOKS_FOLDER / "workbook.xlsx"

workbook = Workbook()

# Creating a custom sheet
sheet_name = "My Sheet"
workbook.create_sheet(sheet_name)
worksheet: Worksheet = workbook[sheet_name]  # setting sheet as active

workbook.remove(workbook["Sheet"])  # Remove standard sheet

# creating headers
worksheet.cell(1, 1, "Name")
worksheet.cell(1, 2, "Age")
worksheet.cell(1, 3, "Grade")

students = [
    ["Caio", 26, 5.5],
    ["Guina", 45, 7.5],
    ["Jailson", 40, 8.5],
    ["Yoruichi", 27, 10],
    ["Claire", 25, 9.5],
    ["Amir", 22, 10],
    ["Mikasa", 20, 10],
]

# Hard way, but G.O.A.T. for other situations

# for i, student_row in enumerate(students, start=2):
#     for j, student_column in enumerate(student_row, start=1):
#         worksheet.cell(i, j, student_column)
#         print(i, j, student_column)

# Easy

for student in students:
    worksheet.append(student)

print(Markdown("---"))

workbook.save(WORKBOOK_PATH)
