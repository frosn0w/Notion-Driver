from openpyxl import load_workbook
from time import sleep
import sys

# Data Initialisation
path_in = ''
path_out = ''
wb = load_workbook(path_in)
ws = wb.active
bottomline = 30


# Transform to markdown
for row in range(2,bottomline):
    row = str(row)
    body = ws['C'+row].value + ':\n' + ws['D'+row].value + '\n'
    for col in 'GHIJK':
        url =str(ws[col+row].value).split('=',1)[0]
        if len(url)>0:
            body = body + '![](' + url + '=s1000' + ')' + '\n'
            sleep(1)
        else:continue
    body = body + '\n---\n'
    with open(path_out,"a")as f:
        f.writelines(body)
