import os

from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation

# Create the workbook and worksheet we'll be working with
wb = Workbook()
ws = wb.active

# Create a data-validation object with list validation
dv = DataValidation(
    type="list",
    formula1='"корпусном,металлическом,мягкой мебели,столярном"',
    allow_blank=True,
)

# Optionally set a custom error message
dv.error = "Запись отсутствует в списке"
dv.errorTitle = "Ошибочный ввод"

# Optionally set a custom prompt message
dv.prompt = "Выберите цех"
dv.promptTitle = "Список цехов"

# Add the data-validation object to the worksheet
ws.add_data_validation(dv)
# Create some cells, and add them to the data-validation object
c1 = ws["A1"]
c1.value = "Dog"
dv.add(c1)

wb.save(r"d:\test.xlsx")
os.startfile(r"d:\test.xlsx")
