import xlrd
from collections import OrderedDict
import simplejson as json

# Open the workbook and select the first worksheet
wb = xlrd.open_workbook('master_HMS_3P_Tools.xlsx')
sh = wb.sheet_by_index(0)
# List to hold dictionaries
input_3P_checker_list = []
# Iterate through each row in worksheet and fetch values into dict
for rownum in range(1, sh.nrows):
    inputDataChecker = OrderedDict()
    row_values = sh.row_values(rownum)
    inputDataChecker['3rd Party Library'] = row_values[0]
    inputDataChecker['Supported in HMS'] = row_values[1]
    inputDataChecker['Aditional Information'] = row_values[2]
    input_3P_checker_list.append(inputDataChecker)
# Serialize the list of dicts to JSON
j = json.dumps(input_3P_checker_list)
# Write to file
with open('data.json', 'w') as f:
    f.write(j)
