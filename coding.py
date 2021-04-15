import numpy as np
import matplotlib.pyplot as plt
import xlrd
import time

def readOrderDataFromExcel():
    file = 'data/Huayudata.xlsx'
    wb = xlrd.open_workbook(filename=file)  # 打开文件
    sheet = wb.sheet_by_name('生产订单')  # 通过名字获取表格

    order10_num = sheet.nrows          # 订单数量，表格行数，注意要将第一行表头去除
    order11_num = sheet.ncols        # 多少列

    order_list = [] # 设置一个新的list数组用于保存订单信息
    for i in range(order10_num):
        # 逐行读取表格信息
        if i == 0: continue # 跳过第一行表头
        # 用字典来保存一个订单的生产单号、产品号。
        order = {'produceID': str(sheet.cell(i, 16).value),
                 'productName': sheet.cell(i, 12).value,
                 }

        order_list.append(order)
    return order_list


def readMachineDataFromExcel():
    file = 'data/Huayudata.xlsx'
    wb = xlrd.open_workbook(filename=file)  # 打开文件
    sheet = wb.sheet_by_name('机台信息')  # 通过名字获取表格
    machine10_num = sheet.nrows          # 行数
    machine11_num = sheet.ncols        # 列数

    machine_list = [] #设置一个新的list数组用于保存机器有效信息
    for i in range(machine10_num):
        # 逐行读取表格信息
        if i == 0: continue # 跳过第一行表头
        #用字典来保存一个机器的有效信息，目前有机台编号、梳节数、开档、排机类型、针数、
        Machine = {'ID': str(sheet.cell(i, 5).value), \
                 'machineType': sheet.cell(i, 7).value, \
                 'needleNum': sheet.cell(i, 8).value, \
                 'combNum': sheet.cell(i, 4).value, \
                 'openDown': sheet.cell(i, 3).value, \
                   }
        machine_list.append(Machine)
    return machine_list


def readProductdataFromExcel():
    file = 'data/Huayudata.xlsx'
    wb = xlrd.open_workbook(filename=file)  # 打开文件
    sheet = wb.sheet_by_name('产品信息')  # 通过名字获取表格
    product10_num = sheet.nrows          # 行数
    product11_num = sheet.ncols        # 列数

    product_list = [] #设置一个新的list数组用于保存产品有效信息
    for i in range(product10_num):
        # 逐行读取表格信息
        if i == 0: continue # 跳过第一行表头
        #用字典来保存一个产品的有效信息，包括产品ID、排机类型、针数、梳节数
        product1 = {'ID': str(sheet.cell(i, 7).value), \
                 'proMachineType': sheet.cell(i, 8).value, \
                 'proNeedleNum': sheet.cell(i, 12).value, \
                 'proCombNum': sheet.cell(i, 5).value, \
                   }
        Product_list1.append(Machine)
    return product_list1



def readProductdataFromExcel():
    file = 'data/Huayudata.xlsx'
    wb = xlrd.open_workbook(filename=file)  # 打开文件
    sheet = wb.sheet_by_name('产品信息表')  # 通过名字获取表格
    product20_num = sheet.nrows          # 行数
    product21_num = sheet.ncols        # 列数

    product_list = [] #设置一个新的list数组用于保存产品有效信息
    for i in range(product20_num):
        # 逐行读取表格信息
        if i == 0: continue # 跳过第一行表头
        #用字典来保存一个产品ID对应梳节的原料、条数。
        Product2 = {'ID2': str(sheet.cell(i, 6).value), \
                 'proMachineType2': sheet.cell(i, 8).value, \
                 'proNeedleNum2': sheet.cell(i, 12).value, \
                 'proCombNum2': sheet.cell(i, 5).value, \
                   }
        Product_list2.append(Machine)
    return product_list2


def selectMachine(machine_list, order_list):
    # 为每个订单选择可用设备
    n_order   = len(order_list)     # 获取订单数量
    n_machnie = len(machine_list)   # 获取设备数量
    order_machine = list()          # 空列表用于存可用设备信息
    for i in np.arange(n_order):
        currentOrder = order_list[i]            # 每次取一个订单进行选择
        machineType = currentOrder['machineType']  # 获取订单需要的设备类型
        needleNum  = currentOrder['needleNum']   # 获取订单需要的针型
        combNum     = currentOrder['combNum']      # 获取订单需要的梳栉数

        machine_selected = list()
        #从设备表中选取可用的设备
        for m in range(n_machnie):
            currentMachine = machine_list[m]    # 取一个设备判断是否可以生产
            try:
                (currentMachine['machineType'] == machineType) & (needleNum == currentMachine['needleNum']) & (
                            currentMachine['combNum'] > combNum)
            except:
                print(currentMachine['machineType'] , machineType, needleNum , currentMachine['needleNum'],
                            currentMachine['combNum'] , combNum)
            if (currentMachine['machineType'] == machineType) & (needleNum ==currentMachine['needleNum']) & (currentMachine['combNum'] >=combNum):
                machine_selected.append(currentMachine['ID']);  #将设备编号加入到可用列表
                                           #获取设备订单剩余时间、
                                            # 当前产品类型
                                            # 原料规格、条数、
                                            # 库存大于3吨、
                                            # 3年内生产过什么产品
                                            # 生产次品率

            if (m == n_machnie) & (1>len(machine_selected)): #没有可加工设备
                print('Error: no avaliable machine!!order number:'+i)
                print('Error: no avaliable machine!!')
                print('Error: no avaliable machine!!')
                exit(0)
        currentOder_machine = {"ID":currentOrder['ID'],"machines":machine_selected}
        order_machine.append(currentOder_machine)
    return order_machine

