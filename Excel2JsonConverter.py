import xlrd
from collections import OrderedDict
import simplejson as json

# Open the workbook and select the first worksheet
wb = xlrd.open_workbook('input_excel.xlsx')
sh = wb.sheet_by_index(0)
# List to hold dictionaries
input_checker_list = []
# Iterate through each row in worksheet and fetch values into dict
for rownum in range(1, sh.nrows):
    inputDataChecker = OrderedDict()
    row_values = sh.row_values(rownum)
    inputDataChecker['1st Columns'] = row_values[0]
    inputDataChecker['2nd Columns'] = row_values[1]
    inputDataChecker['3rd Columns'] = row_values[2]
    input_3P_checker_list.append(inputDataChecker)
# Serialize the list of dicts to JSON
j = json.dumps(input_checker_list)
# Write to file
with open('data.json', 'w') as f:
    f.write(j)
