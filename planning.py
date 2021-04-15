# -*- coding: utf-8 -*-

import numpy as np
import matplotlib.pyplot as plt
import xlrd
import time


def readOrderDataFromExcel():
    file = 'data/订单.xlsx'
    wb = xlrd.open_workbook(filename=file)  # 打开文件
    sheet = wb.sheet_by_name('Sheet1')      # 通过名字获取表格
    order_num = sheet.nrows                 # 订单数量，表格行数，注意要将第一行表头去除
    feature_num = sheet.ncols               # 多少列

    order_list = [] # 设置一个新的list数组用于保存订单信息
    for i in range(order_num):
        # 逐行读取表格信息
        if i == 0: continue # 跳过第一行表头
        # 用字典来保存一个订单的信息，目前只取6个属性，单号，产品类型，设备类型，针数，梳栉数，订单量
        order = {'ID': str(sheet.cell(i, 0).value),
                 'productType': sheet.cell(i, 1).value,
                 'machineType': sheet.cell(i, 2).value,
                 'needleNum': sheet.cell(i, 3).value,
                 'combNum': sheet.cell(i, 4).value,
                 'volume':sheet.cell(i, 5).value,
                 'material': list(),
                 }
        for j in range(int(order['combNum'])):
            material_temp = {'type': sheet.cell(i, 2*j+6).value, 'param':sheet.cell(i, 2*j+7).value}  # 梳栉原料 # 原料参数
            order['material'].append(material_temp)

        order_list.append(order)
    return order_list


def readMachineDataFromExcel():
    file = 'data/设备.xlsx'
    wb = xlrd.open_workbook(filename=file)  # 打开文件
    sheet = wb.sheet_by_name('Sheet1')  # 通过名字获取表格
    order_num = 558          #订单数量，表格行数，注意要将第一行表头去除
    feature_num = sheet.ncols        # 多少列

    machine_list = [] #设置一个新的list数组用于保存订单信息
    for i in range(order_num):
        # 逐行读取表格信息
        if i == 0: continue # 跳过第一行表头
        #用字典来保存一个订单的信息，目前只取5个属性
        Machine = {'ID': str(sheet.cell(i, 0).value), \
                 'machineType': sheet.cell(i, 1).value, \
                 'needleNum': sheet.cell(i, 2).value, \
                 'combNum': sheet.cell(i, 3).value, \
                 }
        machine_list.append(Machine)
    return machine_list

 # 读取设备状态列表
def readMachineStatusFromExcel():
    file = 'data/当前机台原料对照.xlsx'
    wb = xlrd.open_workbook(filename = file)  # 打开文件
    sheet = wb.sheet_by_name('1')  # 通过名字获取表格
    order_num = 3164          #订单数量，表格行数，注意要将第一行表头去除
    feature_num = sheet.ncols        # 多少列

    machine_status = [] #设置一个新的list数组用于保存订单信息
    machine_ID = '000'
    i = 2  # 跳过第一、二行表头
    while i< order_num:
        # 逐行读取表格信息

        # 判断设备编号是否一致，作为是否进入下一个设备的状态提取
        if machine_ID != str(sheet.cell(i, 2).value):
        # 新的设备，提取设备ID,产品名称，本机安排量、计划完成时间
            machine_ID =  sheet.cell(i, 2).value
            status = {'ID': machine_ID,
                      'productType': sheet.cell(i, 5).value,
                      'productVolume':sheet.cell(i, 4).value,
                      'speed':sheet.cell(i, 3).value,
                      'completionTime':xlrd.xldate.xldate_as_datetime(sheet.cell(i,21).value,0),
                      'materials':[]}
        # 同一个设备，提出原料信息
        if  sheet.cell(i, 6).value !='' :
                material_temp = {'type': sheet.cell(i, 6).value,
                             'param': sheet.cell(i, 8).value,
                                 'demandVolume':sheet.cell(i, 13).value,
                                 'remain':sheet.cell(i, 15).value}  # 梳栉原料 # 原料参数
                status['materials'].append(material_temp)
        if i ==(order_num-1):
            machine_status.append(status)
            del status
        else:
            if machine_ID != str(sheet.cell(i+1, 2).value):
                machine_status.append(status)
                del status
        i +=1

    return machine_status


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





def get_between_days(start_date, end_date):
    # 获得两个日期之间的天数
    start_sec = time.mktime(time.strptime(start_date,'%Y-%m-%d'))
    if end_date =='1900-01-01':
        work_days = -1000
    else:
        end_sec = time.mktime(time.strptime(end_date,'%Y-%m-%d'))
        work_days = int((end_sec - start_sec)/(24*60*60))
    return work_days

def planning(order_data, order_machine, machine_status):
    # 根据规则确定最终使用设备
    order4planning = order_machine                 # 待分配订单
    machines_ID  = [m['ID']  for m in machine_status] # 所有可用设备
    result = []#[{'orderID': order['ID'], 'machineID':''} for order in order_machine]                 # 最终分配结果

    # 检查是不是所有设备的状态都可用
    for order in order_machine:
        for m in order['machines']:
            if m in machines_ID:
                continue
            else:
                print(m, 'do not have status')
                order['machines'].remove(m) # 删除没有状态的设备



    for order in order4planning:
        result_temp = {'ID': order['ID'], 'machine': {}}

        # 首先确定订单中只有一个可用设备的
        if len(order['machines']) < 2 :
            if len(order['machines']) ==1:
                 result_temp['machine'] = order['machines'][0] #获取订单编号和设备编号
            else:
                 result_temp['machine'] = ''  # 获取订单编号和设备编号
            # machines_ID.remove(order['machines'])  # 更新设备信息，订单完成时间,现在直接删除,可能两个设备都只能用同一个
            result.append(result_temp)
        else:
            machines = order['machines']
            # 获取每个设备的状态

            # 1 选出小于60小时的设备
                #获取每个设备的完成时间
            completionTime = []
            start_date = time.strftime('%Y-%m-%d', time.localtime(time.time()))
            for m in machines:
                for m_s in machine_status:
                    if m_s['ID'] == m:
                        completionTime.append( get_between_days(start_date,m_s['completionTime'].strftime('%Y-%m-%d')))

            arr = np.array(completionTime)
            ind = np.where(arr<3)[0].tolist()
            if len(ind):
                machines = [machines[ind[i]] for i in range(len(ind))]
                if len(ind)==1:
                    result_temp['machine'] = machines  # 获取订单编号和设备编号
                    #machines_ID.remove(order['machines'])  # 更新设备信息，订单完成时间,现在直接删除,可能两个设备都只能用同一个
                    result.append(result_temp)
                    continue
            # 在生产的产品类型是否一致
            o_productType = [order['productType'] for _, order in enumerate(order_list) if order['ID'] == order["ID"]][0]
            m_productType = [machine['productType'] for _, machine in enumerate(machine_status) for m in machines]
            # 找出相同产品的设备
            ind = [i for i,_ in enumerate(m_productType) if _ ==o_productType]
            if len(ind):
                 machines = [machines[ind[i]] for i in range(len(ind))] #选出符合条件的设备
                 if len(ind) == 1:
                        result_temp['machine'] = machines[ind]  # 获取订单编号和设备编号
                        machines_ID.remove(order['machines'])  # 更新设备信息，订单完成时间,现在直接删除,可能两个设备都只能用同一个
                        result.append(result_temp)
                        continue
            # 2 选择 原料规格一致的设备
                    # 订单原料
            ind = []
            orderMaterial = [order['material'] for _, order in enumerate(order_list) if order['ID'] == order["ID"]][0]
            m_material = [machine['materials'] for _, machine in enumerate(machine_status) if machine['ID'] in machines]
            for material in orderMaterial:
                #规格和名称都一致
                if material['type'] == '75D涤半光' or material['type'] =='30D涤单丝':
                    for j, m_m in enumerate(m_material):
                        for m_m_i in m_m:
                            if m_m_i['type'] == material['type'] and m_m_i['param']== material['param']:
                                ind.append(j)
                                continue
            if len(ind):
                 machines =[machines[ind[i]] for i in range(len(ind))] #选出符合条件的设备
                 if len(ind) == 1:
                        result_temp['machine'] = machines  # 获取订单编号和设备编号
                       # machines_ID.remove(order['machines'])  # 更新设备信息，订单完成时间,现在直接删除,可能两个设备都只能用同一个
                        result.append(result_temp)
                        continue

            # 3 设备库存大于3t

            # 如果都没匹配上就选一个时间最近的安排上
            result_temp['machine'] = machines[0]  # 获取订单编号和设备编号
            # machines_ID.remove(order['machines'])  # 更新设备信息，订单完成时间,现在直接删除,可能两个设备都只能用同一个
            result.append(result_temp)



        # 车速:高低
    return result


   #  # 否则进入下一个环节，根据规则确定最终设备
# 根据规则确定最终设备

def plotResutl(order_machine,machine_list):
    #画出当前设备状态
    machine_order = list()
    for machine in machine_list:
        machine_temp = {'ID':machine['ID'],'order_num':0,'order_IDs':list()}
        machine_order.append(machine_temp)


    for i in range(len(order_machine)):
       machines =  order_machine[i]['machines'] #获取订单对应的设备
       for machine in machines:
        for j, d in enumerate(machine_order):
           if d['ID'] == machine:
               d['order_num'] = d['order_num']+1
               d['order_IDs'].append(order_machine[i]['ID'])


    shop_1 = np.arange(1,80)
    shop_2 = np.arange(186,236)
    shop_3 = np.arange(330,370)
    shop_4 = np.arange(6601,6700)
    shop_5 = np.arange(3001,3045)
    shop_6 = np.arange(1001,1148)
    shop_7 = np.arange(5001,5088)
    shop_8 = np.arange(121,246)
    print(shop_1.shape[0],shop_2.shape[0],shop_3.shape[0],shop_4.shape[0],shop_5.shape[0],shop_6.shape[0],shop_7.shape[0],shop_8.shape[0])
    print([shop_1.shape[0]+shop_2.shape[0]+shop_3.shape[0]+shop_4.shape[0]+shop_5.shape[0]+shop_6.shape[0]+shop_7.shape[0]+shop_8.shape[0]])

   # 画出设备状态,
    names = locals();
    x0 = 0
    y0 = 0
    plt.show()

    for m , d in enumerate(machine_order):
        y = y0 + 8 * int(m % 10) + 5  # 第几列

        x = x0 + 8 * int(m / 10) + 10  # 第几行
        color = (1/(d['order_num']+20),1/(d['order_num']+2),1/(d['order_num']+10))
        plt.scatter(x, y, color= color, marker='o',  s=100,  alpha=0.5)  # 把 corlor 设置为空，通过edgecolors来控制颜色
        plt.text(x-4,y-3,d['ID'], fontsize=6)

    return machine_order
    # for i in range(8): #按车间画
    #     shop = names['shop_' + str(i + 1)]
    #     # for j in range (10): # 按行画，每个车间分为10行
    #     machine_num = shop.shape[0]
    #     width =  int(machine_num/10)#10对应多少列
    #     for m in range(machine_num):
    #         # 确定每个圈的坐标
    #         y = y0+8*int(m%10)+5 #第几列
    #
    #         x = x0+8*int(m/10)+5        #第几行
    #         plt.scatter(x, y, color='g', marker='o', edgecolors='g',s=100)  # 把 corlor 设置为空，通过edgecolors来控制颜色
    #         if m == (machine_num-1):
    #             if i!=7:
    #                 plt.plot([x+4,x+4],[0,80])
    #             x0=x+5
    plt.show()

    #画出订单分配状态


if __name__ == '__main__':

     machine_list = readMachineDataFromExcel() # 获取设备列表， 设备编号、类型、针数、梳栉数、
     order_list = readOrderDataFromExcel()    #获取订单列表   编号 产品类型、需要的设备类型、针数、梳栉数、订单量、材料
     order_machine = selectMachine(machine_list, order_list) #订单可用设备
     achine_order=plotResutl(order_machine,machine_list)    #设备分配到的订单

     machine_status = readMachineStatusFromExcel() #设备当前状态
     result =  planning(order_list, order_machine, machine_status)
     # 打印出订单对应的可用设备
     #   for order in order_machine:
     #       for j, d in enumerate(order_list):
     #          if d['ID'] == order['ID']:
     #              type = d['productType']
     #       print('订单编号：',order['ID'],'产品类型',type)
     #       print('可用设备:',order['machines'])
     #       print()