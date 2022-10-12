import webbrowser
import openpyxl

for ii in range(1,6):
    wb = openpyxl.load_workbook('点名情况课程'+str(ii)+'.xlsx')
    ws = wb['课程'+str(ii)]

    a = []
    for i in range(1,91):
        a.append((ws.cell(i+1,22).value*10,i))
    a.sort()
