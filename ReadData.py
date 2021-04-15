import xlrd

file = 'data/订单.xlsx'

wb = xlrd.open_workbook(filename=file)#打开文件
print(wb.sheet_names())#获取所有表格名字

sheet = wb.sheet_by_name('Sheet1')#通过名字获取表格

order_num = sheet.nrows
feature_num = sheet.ncols

order_list = []

for i in range(order_num):
    if i==0: continue
    currentOrder =  sheet.row_values(i)
    order = {'ID':str(sheet.cell(i, 0).value), \
             'productType': sheet.cell(i, 1).value, \
             'machineType': sheet.cell(i, 2).value, \
             'needleNum':   sheet.cell(i, 3).value, \
             'combNum'    : sheet.cell(i, 4).value  \
             }
    order_list.append(order)




