# -*- coding: utf-8 -*-
import  xdrlib ,sys
import xlrd
import xlwt
import easygui
import os
import pandas as pd

def open_excel(file= 'test.xls'):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception as e:
        print(e)
#根据索引获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的所以  ，by_index：表的索引
def excel_table_byindex(file= 'test.xls',colnameindex=0,by_index=0):
    data = open_excel(file)
    table = data.sheets()[by_index]
    nrows = table.nrows #行数
    ncols = table.ncols #列数
    colnames =  table.row_values(colnameindex) #某一行数据
    list =[]
    for rownum in range(1,nrows):

         row = table.row_values(rownum)
         if row:
             app = {}
             for i in range(len(colnames)):
                app[colnames[i]] = row[i]
             list.append(app)
    return list

#根据名称获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的所以  ，by_name：Sheet1名称
def excel_table_byname(file='test.xls',colnameindex=0,by_name=u'Sheet1'):
    data = open_excel(file)
    table = data.sheet_by_name(by_name)
    nrows = table.nrows #行数
    colnames =  table.row_values(colnameindex) #某一行数据
    list =[]
    for rownum in range(1,nrows):
         row = table.row_values(rownum)
         if row:
             app = {}
             for i in range(len(colnames)):
                app[colnames[i]] = row[i]
             list.append(app)
    return list

def wblist(extension='.xls'):
    '''
    for root, dirs, files in os.walk("./节目单", topdown=False):
        #文件
        for name in files:
            print(os.path.join(root, name))
        #目录
        for name in dirs:
            print(os.path.join(root, name))
    '''
    # 处理文件名
    documnetName = []
    for root, dirs, files in os.walk("./节目单", topdown=False):
        # 文件
        for name in files:
            # print(os.path.join(root, name))
            if extension in name:
                documnetName.append([name[0], name[3:-4], root, name])
    documnetName.sort(key=lambda x: x[1])
    # 删除无用列
    documnetNames = []
    for i in documnetName:
        documnetNames.append(i[2:4])
    return documnetNames

def tableHeader(book, documnetName):
    # 处理表头
    # 读表头
    data = open_excel(documnetName)
    table = data.sheets()[0]
    colnames = table.row_values(0)  # 某一行数据
    # 写表头
    pindao = ['卫视', '经济', '都市', '影视', '少儿', '公共', '农民']
    for i in range(7):
        sheet = book.add_sheet(pindao[i], cell_overwrite_ok=True)
        sheet.row(0).write(0, '日期')
        for j in range(len(colnames)):
            sheet.row(0).write(j + 1, colnames[j])
            if len(colnames[j]) == 0:
                sheet.col(j + 1).width = 256 * 1
            if colnames[j] == '节目名称':
                sheet.col(j + 1).width = 256 * 40


def copydata(book, root, name, myrownum):
    #打印文件名称
    print(name)
    #读取源数据
    data = open_excel(os.path.join(root, name))
    table = data.sheets()[0]
    nrows = table.nrows  # 行数
    ncols = table.ncols  # 列数
    colnames = table.row_values(0)  # 某一行数据
    ndate = name[3:-4]
    sheetName = int(name[0])-1
    #print(sheetName)
    #写入目标文件
    sheet = book.get_sheet(sheetName)
    for rownum in range(1, nrows):
        row = table.row_values(rownum)
        if row:
            sheet.row(myrownum[int(sheetName)]).write(0, ndate)
            for i in range(len(colnames)):
                sheet.row(myrownum[int(sheetName)]).write(i+1, row[i])
            myrownum[int(sheetName)] += 1

def copydataPD(file='test.xls'):

    #读取源数据
    pindao = ['卫视', '经济', '都市', '影视', '少儿', '公共', '农民']
    biaotou = ['日期', '节目名称', '长度', '主视源', '磁带条码']
    df = pd.read_excel(file, sheet_name=pindao)
    for x in pindao:
        df[x] = df[x].loc[:, biaotou]
        df[x].insert(1, '频道', x)
    result = pd.concat(df)
    #写入汇总表
    result.to_excel('result.xls', index=False, sheet_name="汇总")


def main():

    tables = excel_table_byindex()
    for row in tables[-5:]:
        print(row)

    tables = excel_table_byname()
    for row in tables[-5:]:
        print(row)

def huizong():
    # 获取文件名
    documnetNames = wblist()
    # 新建workbook
    book = xlwt.Workbook(encoding='utf-8')
    # 填写表头
    documnetName = os.path.join(documnetNames[0][0], documnetNames[0][1])
    tableHeader(book, documnetName)
    # 各数据表中数据记录初始化
    myrownum = [1, 1, 1, 1, 1, 1, 1]
    # 写数据
    for d in documnetNames:
        copydata(book, d[0], d[1], myrownum)
    print(myrownum)
    inMessage = easygui.enterbox(msg='请输入要保存的文件名', title='另存为', default='2020年第一季度节目单汇总')
    if inMessage:
        pass
    else:
        inMessage = 'test'
    book.save('{}.xls'.format(inMessage))

if __name__=="__main__":
    copydataPD()