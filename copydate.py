# -*- coding: utf-8 -*-
import xdrlib,sys
import xlrd
import xlwt
import xlsxwriter
import easygui
import os
import pandas as pd
import pandas.io.formats.excel
import math
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.shared import Inches, Pt, Mm
from docx.oxml.ns import qn

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
    print('正在处理：{}'.format(name))
    #读取源数据
    data = open_excel(os.path.join(root, name))
    table = data.sheets()[0]
    nrows = table.nrows  # 行数
    ncols = table.ncols  # 列数
    colnames = table.row_values(0)  # 某一行数据
    ndate = name.split('-')[2]
    ndate = ndate.split('.')
    ndate = '2020/{}/{}'.format(ndate[0], ndate[1])
    sheetName = int(name.split('-')[0])-1
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

def huizong1():
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
    print('汇总节目数：{}'.format(myrownum))
    '''
    inMessage = easygui.enterbox(msg='请输入要保存的文件名', title='另存为', default='2020年第一季度节目单汇总')
    if inMessage:
        pass
    else:
        inMessage = 'total'
    '''
    inMessage = 'total'
    book.save('{}.xls'.format(inMessage))

def huizong(filename='./total.xlsx'):
    print('分析目录')
    files = wblist()
    pindao = ['卫视', '经济', '都市', '影视', '少儿', '公共', '农民']
    df = {}
    for i in pindao:
        df[i] = pd.DataFrame()
    for file in files:
        name = file[1]
        ndate = name.split('-')[2]
        ndate = ndate.split('.')
        ndate = '2020/{}/{}'.format(ndate[0], ndate[1])
        sheetName = int(name.split('-')[0]) - 1
        data = pd.read_excel(os.path.join(file[0], name))
        data.insert(0, '日期', ndate)
        if df[pindao[sheetName]].empty:
            df[pindao[sheetName]] = data
        else:
            df[pindao[sheetName]] = pd.concat([df[pindao[sheetName]], data])
        print(file[1])
    total_number = {}
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        for i in pindao:
            print('正在写入:{}表'.format(i))
            df[i].to_excel(writer, sheet_name=i, index=False)
            total_number[i] = df[i].shape[0]
    print('各频道节目条数：{}'.format(total_number))

def programefilter(df):
    #df = pd.read_excel('total.xls')
    #清除列
    df = df.loc[:, ~ df.columns.str.contains(':')]
    #df = df.dropna(axis=1, how='all')
    #清除行
    # 宣
    xdf = df[df['节目名称'].str.contains('宣')]
    #广告
    blAdvert = df['主视源'].str.contains('广告')
    df = df[~ blAdvert]
    #其它
    blOther = df['节目名称'].str.contains('集|宣|广告|预报|预告|ID|即播|预报|剧情|片头|片尾|剧透|招募|新闻联播|头条|京津冀|这一年|旅游|呼号|导视|多看点|欢乐送|标版|引进节目|专题|logo|LOGO|战略|气象|德龙|专临|前情回顾|先睹为快|公益|年货')
    df = df[~ blOther]
    
    #print(df.info())
    return df, xdf

def cleardata(file='filter.xlsx'):
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    with pd.ExcelWriter(file, engine='xlsxwriter') as writer:
        # Convert the dataframe to an XlsxWriter Excel object. Note that we turn off
        # the default header and skip one row to allow us to insert a user defined
        # header.
        workbook = writer.book
        pindao = ['卫视', '经济', '都市', '影视', '少儿', '公共', '农民']
        rownumber = {}
        for i in pindao:
            print('正在整理:{}'.format(i))
            df = pd.read_excel('total.xls', i,)
            df, xdf = programefilter(df)
            rownumber[i] = df.shape[0]
            df.to_excel(writer, sheet_name=i, startrow=0, na_rep='', index=False)
            worksheet = writer.sheets[i]
            #设置节目名称列列宽
            worksheet.set_column(5, 5, 35)
            xdf.to_excel(writer, sheet_name='{}宣'.format(i), startrow=0, na_rep='', index=False)

    print('初选节目数：{}'.format(rownumber))

def readtoPD(file='filter.xlsx'):
    # 读取源数据
    pindao = ['卫视', '经济', '都市', '影视', '少儿', '公共', '农民']
    biaotou = ['日期', '节目名称', '长度', '主视源', '磁带条码']
    # 表头格式清除
    # pd.io.formats.excel.header_style = None
    df = pd.read_excel(file, sheet_name=pindao)
    for x in pindao:
        df[x] = df[x].loc[:, biaotou]
        df[x].insert(1, '频道', x)
    result = pd.concat(df, ignore_index=True)
    #
    result['日期'].astype('str')
    # 处理节目名称列
    s = result['节目名称']
    s = s.str.replace('^.套', '')
    s = s.str.replace('第20', '_第20')
    s = s.str.replace('-_', '_')
    s = s.str.replace('-2020-', '_2020-')
    d = s.str.split('_', expand=True)
    d.columns = ['节目名称', '期数', '日期']
    s = d['节目名称']
    s = s.str.replace(' 20|-20|-19', '_20', regex=True)
    d['节目名称'] = s.str.split('_').str.get(0)
    result.drop(['节目名称'], axis=1, inplace=True)
    result.insert(2, '期数', d['期数'])
    result.insert(2, '节目名称', d['节目名称'])
    # result.insert(1, '日期1', d['日期'])

    result['长度'] = result['长度'].str.split('.').str.get(0)
    result.rename(columns={'主视源': '来源', '磁带条码': '责任人'}, inplace=True)
    return result

def writetoEx1(myData, file='result.xlsx', bPrint=False):
    #写入精选表
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    with pd.ExcelWriter(file, engine='xlsxwriter') as writer:
        # Convert the dataframe to an XlsxWriter Excel object. Note that we turn off
        # the default header and skip one row to allow us to insert a user defined
        # header.
        myData.to_excel(writer, sheet_name="精选", startrow=0, na_rep='')
        # Get the xlsxwriter workbook and worksheet objects.
        workbook = writer.book
        worksheet = writer.sheets['精选']
        #加序列号
        worksheet.write(0, 0, '序号')

        #worksheet = workbook.add_worksheet('Sheet1')

        # Add a header format.
        header_format = workbook.add_format({
            'font_size':12,   #字体大小
            'bold':1,   #是否粗体
            'bg_color': '0f6f32', #表格背景颜色
            'font_color': '#E2F3F6',   #字体颜色
            'align':'center',  #对齐方式，left,center,rigth,top,vcenter,bottom,vjustify
            'valign': 'vcenter',#垂直居中
            'top':1,  #上边框，后面参数是线条宽度
            'left':1, #左边框
            'right':1, #右边框
            'bottom':1, #底边框
            'border_color': 'white',
            'text_wrap':1,  #自动换行，可在文本中加 '\n'来控制换行的位置
            #'num_format':'yyyy-mm-dd' #设定格式为日期格式，如：2017-07-01
            })
        index_format = workbook.add_format({
            'font_size': 12,  # 字体大小
            'bold': 0,  # 是否粗体
            'bg_color': '#0f6f32', #表格背景颜色
            'font_color': '#E2F3F6',   #字体颜色
            'align': 'center',  # 对齐方式，left,center,rigth,top,vcenter,bottom,vjustify
            'valign': 'vcenter',  # 垂直居中
            'top': 1,  # 上边框，后面参数是线条宽度
            'left': 1,  # 左边框
            'right': 1,  # 右边框
            'bottom': 1,  # 底边框
            'border_color':'white',
            'text_wrap': 1,  # 自动换行，可在文本中加 '\n'来控制换行的位置
            # 'num_format':'yyyy-mm-dd' #设定格式为日期格式，如：2017-07-01
        })
        data_format = workbook.add_format({
            'font_size': 11,  # 字体大小
            'bold': 0,  # 是否粗体
            'bg_color': '#319455', #表格背景颜色
            'font_color': '#E2F3F6',   #字体颜色
            'align': 'center',  # 对齐方式，left,center,rigth,top,vcenter,bottom,vjustify
            'valign': 'vcenter',  # 垂直居中
            'top': 1,  # 上边框，后面参数是线条宽度
            'left': 1,  # 左边框
            'right': 1,  # 右边框
            'bottom': 1,  # 底边框
            'border_color':'white',
            'text_wrap': 1,  # 自动换行，可在文本中加 '\n'来控制换行的位置
            #'num_format': 'yyyy-mm-dd'  # 设定格式为日期格式，如：2017-07-01
            })
        data2_format = workbook.add_format({
            'font_size': 11,  # 字体大小
            'bold': 0,  # 是否粗体
            'bg_color': '#54B58A', #表格背景颜色
            'font_color': '#E2F3F6',   #字体颜色
            'align': 'center',  # 对齐方式，left,center,rigth,top,vcenter,bottom,vjustify
            'valign': 'vcenter',  # 垂直居中
            'top': 1,  # 上边框，后面参数是线条宽度
            'left': 1,  # 左边框
            'right': 1,  # 右边框
            'bottom': 1,  # 底边框
            'border_color':'white',
            'text_wrap': 1,  # 自动换行，可在文本中加 '\n'来控制换行的位置
            # 'num_format': 'yyyy-mm-dd'  # 设定格式为日期格式，如：2017-07-01
        })
        defaul_format = workbook.add_format({
            'font_size': 11,  # 字体大小
            'bold': 0,  # 是否粗体
            'bg_color': 'white',  # 表格背景颜色
            'font_color': 'black',  # 字体颜色
            'align': 'center',  # 对齐方式，left,center,rigth,top,vcenter,bottom,vjustify
            'valign': 'vcenter',  # 垂直居中
            'top': 0,  # 上边框，后面参数是线条宽度
            'left': 0,  # 左边框
            'right': 0,  # 右边框
            'bottom': 0,  # 底边框
            'text_wrap': 1,  # 自动换行，可在文本中加 '\n'来控制换行的位置
            # 'num_format': 'yyyy-mm-dd'  # 设定格式为日期格式，如：2017-07-01
        })
        if bPrint:
            header_format.set_font_color('black')
            header_format.set_fg_color('white')
            header_format.set_border_color('black')
            index_format.set_font_color('black')
            index_format.set_fg_color('white')
            index_format.set_border_color('black')
            data_format.set_font_color('black')
            data_format.set_fg_color('white')
            data_format.set_border_color('black')
            data2_format.set_font_color('black')
            data2_format.set_fg_color('white')
            data2_format.set_border_color('black')

        worksheet.write(0, 0, '序号',header_format)
        # Write the column headers with the defined format.
        colWidth = {'序号':4.54,'日期':10.21,'日期1':10.66,'频道':4.54,'节目名称':29.43,'期数':11.55,
                    '长度':8.58,'来源':6.68,'责任人':6.68}
        worksheet.set_column(0,0,colWidth['序号'])
        for col_num, value in enumerate(myData.columns.values):
            #colWidth = max(len(value), max(myData[value].astype(str).str.len()))+3
            #print(colWidth)
            worksheet.set_column(col_num+1, col_num+1, colWidth[value])
            worksheet.write(0, col_num + 1, value, header_format)
        for row_num,value in enumerate(myData.index.values):
            #worksheet.set_row(row_num+1, None, data_format)
            worksheet.write(row_num+1, 0, value+1, index_format)
        #列宽
        print('保留节目数：{}'.format(myData.shape[0]))
        worksheet.freeze_panes(1, 1)
        #worksheet.set_column(myData.shape[1]+1, 200, None, defaul_format, {'hidden': 0})
        for i in range(myData.shape[0]):
            if myData.loc[i, '来源'] == '硬盘':
                datastyle = data_format
            else:
                datastyle = data2_format
            for j in range(myData.shape[1]):
                value = str(myData.iloc[i,j])
                if value == 'nan' or value == 'None' or value == '录像机':
                    value = ""
                worksheet.write(i+1, j+1, value, datastyle)
                pass

def writetoEx(myData, file='result.xlsx', bPrint=False):
    
    #写入精选表
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    with pd.ExcelWriter(file, engine='xlsxwriter') as writer:
        # Convert the dataframe to an XlsxWriter Excel object. Note that we turn off
        # the default header and skip one row to allow us to insert a user defined
        # header.
        myData.to_excel(writer, sheet_name="精选", startrow=0, na_rep='')
        # Get the xlsxwriter workbook and worksheet objects.
        workbook = writer.book
        worksheet = writer.sheets['精选']
        #加序列号
        worksheet.write(0, 0, '序号')

        #worksheet = workbook.add_worksheet('Sheet1')

        # Add a header format.
        header_format = workbook.add_format({
            'font_size':12,   #字体大小
            'bold':1,   #是否粗体
            'bg_color': '0f6f32', #表格背景颜色
            'font_color': '#E2F3F6',   #字体颜色
            'align':'center',  #对齐方式，left,center,rigth,top,vcenter,bottom,vjustify
            'valign': 'vcenter',#垂直居中
            'top':1,  #上边框，后面参数是线条宽度
            'left':1, #左边框
            'right':1, #右边框
            'bottom':1, #底边框
            'border_color': 'white',
            'text_wrap':1,  #自动换行，可在文本中加 '\n'来控制换行的位置
            #'num_format':'yyyy-mm-dd' #设定格式为日期格式，如：2017-07-01
            })
        index_format = workbook.add_format({
            'font_size': 12,  # 字体大小
            'bold': 0,  # 是否粗体
            'bg_color': '#0f6f32', #表格背景颜色
            'font_color': '#E2F3F6',   #字体颜色
            'align': 'center',  # 对齐方式，left,center,rigth,top,vcenter,bottom,vjustify
            'valign': 'vcenter',  # 垂直居中
            'top': 1,  # 上边框，后面参数是线条宽度
            'left': 1,  # 左边框
            'right': 1,  # 右边框
            'bottom': 1,  # 底边框
            'border_color':'white',
            'text_wrap': 1,  # 自动换行，可在文本中加 '\n'来控制换行的位置
            # 'num_format':'yyyy-mm-dd' #设定格式为日期格式，如：2017-07-01
        })
        data_format = workbook.add_format({
            'font_size': 11,  # 字体大小
            'bold': 0,  # 是否粗体
            'bg_color': '#319455', #表格背景颜色
            'font_color': '#E2F3F6',   #字体颜色
            'align': 'center',  # 对齐方式，left,center,rigth,top,vcenter,bottom,vjustify
            'valign': 'vcenter',  # 垂直居中
            'top': 1,  # 上边框，后面参数是线条宽度
            'left': 1,  # 左边框
            'right': 1,  # 右边框
            'bottom': 1,  # 底边框
            'border_color':'white',
            'text_wrap': 1,  # 自动换行，可在文本中加 '\n'来控制换行的位置
            #'num_format': 'yyyy-mm-dd'  # 设定格式为日期格式，如：2017-07-01
            })
        data2_format = workbook.add_format({
            'font_size': 11,  # 字体大小
            'bold': 0,  # 是否粗体
            'bg_color': '#54B58A', #表格背景颜色
            'font_color': '#E2F3F6',   #字体颜色
            'align': 'center',  # 对齐方式，left,center,rigth,top,vcenter,bottom,vjustify
            'valign': 'vcenter',  # 垂直居中
            'top': 1,  # 上边框，后面参数是线条宽度
            'left': 1,  # 左边框
            'right': 1,  # 右边框
            'bottom': 1,  # 底边框
            'border_color':'white',
            'text_wrap': 1,  # 自动换行，可在文本中加 '\n'来控制换行的位置
            # 'num_format': 'yyyy-mm-dd'  # 设定格式为日期格式，如：2017-07-01
        })
        defaul_format = workbook.add_format({
            'font_size': 11,  # 字体大小
            'bold': 0,  # 是否粗体
            'bg_color': 'white',  # 表格背景颜色
            'font_color': 'black',  # 字体颜色
            'align': 'center',  # 对齐方式，left,center,rigth,top,vcenter,bottom,vjustify
            'valign': 'vcenter',  # 垂直居中
            'top': 0,  # 上边框，后面参数是线条宽度
            'left': 0,  # 左边框
            'right': 0,  # 右边框
            'bottom': 0,  # 底边框
            'text_wrap': 1,  # 自动换行，可在文本中加 '\n'来控制换行的位置
            # 'num_format': 'yyyy-mm-dd'  # 设定格式为日期格式，如：2017-07-01
        })
        if bPrint:
            header_format.set_font_color('black')
            header_format.set_fg_color('white')
            header_format.set_border_color('black')
            index_format.set_font_color('black')
            index_format.set_fg_color('white')
            index_format.set_border_color('black')
            data_format.set_font_color('black')
            data_format.set_fg_color('white')
            data_format.set_border_color('black')
            data2_format.set_font_color('black')
            data2_format.set_fg_color('white')
            data2_format.set_border_color('black')

        worksheet.write(0, 0, '序号',header_format)
        # Write the column headers with the defined format.
        colWidth = {'序号':4.54,'日期':10.21,'日期1':10.66,'频道':4.54,'节目名称':29.43,'期数':11.55,
                    '长度':8.58,'来源':6.68,'责任人':6.68}
        worksheet.set_column(0,0,colWidth['序号'])
        for col_num, value in enumerate(myData.columns.values):
            #colWidth = max(len(value), max(myData[value].astype(str).str.len()))+3
            #print(colWidth)
            worksheet.set_column(col_num+1, col_num+1, colWidth[value])
            worksheet.write(0, col_num + 1, value, header_format)
        for row_num,value in enumerate(myData.index.values):
            #worksheet.set_row(row_num+1, None, data_format)
            worksheet.write(row_num+1, 0, value+1, index_format)
        #列宽
        print('保留节目数：{}'.format(myData.shape[0]))
        worksheet.freeze_panes(1, 1)
        #worksheet.set_column(myData.shape[1]+1, 200, None, defaul_format, {'hidden': 0})
        for i in range(myData.shape[0]):
            if myData.loc[i, '来源'] == '硬盘':
                datastyle = data_format
            else:
                datastyle = data2_format
            for j in range(myData.shape[1]):
                value = str(myData.iloc[i,j])
                if value == 'nan' or value == 'None' or value == '录像机':
                    value = ""
                worksheet.write(i+1, j+1, value, datastyle)
                pass

def main():

    tables = excel_table_byindex()
    for row in tables[-5:]:
        print(row)

    tables = excel_table_byname()
    for row in tables[-5:]:
        print(row)

def cleardocument(documentname='常规节目主观评测表（评委）.docx'):
    document = Document(documentname)
    tempTable = document.tables
    for table in tempTable:
        for rowIndex in range(2, 20, 8):
            #tempdf = df.loc[j]
            for i in range(3):
                cell = table.cell(rowIndex, i)
                #print(cell.text)
                cell.text = ''
    document.save('常规节目主观评测表（评委）.docx')

def writedocument(file = 'freeze.xlsx', sheet='常规', blMerge=False):
    df = pd.read_excel(file, sheet_name=sheet)
    df = df.loc[:, ['序号', '节目名称']]
    df = df.dropna(how='all')
    df['序号'] = df['序号'].astype("int")

    document = Document()

    document.styles['Normal'].font.name = u'宋体'
    document.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')

    sections = document.sections
    section = sections[0]
    section.left_margin = Mm(19.1)
    section.right_margin = Mm(19.1)



    for j in range(df.shape[0]):
        if j%3 == 0:
            p = document.add_paragraph()
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run('第十三期电视节目技术质量主观评测打分表')
            r.font.size = Pt(15)
            r.bold = True
        tempTable = bulitTable(document, blMerge)
        for i in range(df.shape[1]):
            cell = tempTable.cell(2, i)
            cell.text = str(df.iloc[j, i])
            #设置中间对齐
            cell.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        print(j)
    document.save('test.docx')

def bulitTable(document, blMerge=False):

    table = document.add_table(rows=8, cols=14, style='Table Grid')
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style.font.size = Pt(10.5)#Pt(10.5)
    table.autofit = False

    for i in range(14):
        table.columns[i].width = Mm(11.5)
    table.columns[1].width = Mm(18.7)
    table.columns[2].width = Mm(15.7)
    table.columns[13].width = Mm(15.7)
    for i in range(8):
        table.rows[i].height = Mm(8.5)
    table.rows[1].height = Mm(11.4)

    table.cell(0, 0).merge(table.cell(1, 0))
    table.cell(0, 1).merge(table.cell(1, 1))
    table.cell(0, 2).merge(table.cell(1, 2))
    table.cell(0, 3).merge(table.cell(0, 9))
    table.cell(0, 10).merge(table.cell(0, 12))
    table.cell(0, 13).merge(table.cell(1, 13))
    table.cell(2, 0).merge(table.cell(7, 0))
    table.cell(2, 1).merge(table.cell(7, 1))
    table.cell(2, 3).merge(table.cell(2, 12))
    table.cell(3, 3).merge(table.cell(3, 12))
    table.cell(4, 3).merge(table.cell(4, 12))
    table.cell(5, 3).merge(table.cell(5, 12))
    table.cell(6, 3).merge(table.cell(6, 12))
    table.cell(7, 3).merge(table.cell(7, 12))
    table.cell(2, 13).merge(table.cell(7, 13))
    if blMerge:
        table.cell(2, 2).merge(table.cell(7, 2))
        table.cell(2, 3).merge(table.cell(7, 12))
    hdr_cells0 = table.rows[0].cells
    hdr_cells1 = table.rows[1].cells

    hdr_cells0[0].text ='序号'
    hdr_cells0[0].paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hdr_cells0[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    hdr_cells0[1].text = '节目名称'
    hdr_cells0[1].paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hdr_cells0[1].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    hdr_cells0[2].text = '测试点'
    hdr_cells0[2].paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hdr_cells0[2].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    hdr_cells0[3].text = '图像（70）分'
    hdr_cells0[3].paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hdr_cells0[3].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    hdr_cells0[10].text = '声音（30）分'
    hdr_cells0[10].paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hdr_cells0[10].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    hdr_cells0[13].text = '总评分'
    hdr_cells0[13].paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hdr_cells0[13].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    hdr_cells1[3].text = '杂波\r干扰'
    hdr_cells1[3].paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hdr_cells1[3].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    hdr_cells1[4].text = '清晰度'
    hdr_cells1[4].paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hdr_cells1[4].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    hdr_cells1[5].text = '亮度\r层次'
    hdr_cells1[5].paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hdr_cells1[5].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    hdr_cells1[6].text = '色彩\r保真'
    hdr_cells1[6].paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hdr_cells1[6].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    hdr_cells1[7].text = '制作\r难度'
    hdr_cells1[7].paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hdr_cells1[7].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    hdr_cells1[8].text = '素材\r资料'
    hdr_cells1[8].paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hdr_cells1[8].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    hdr_cells1[9].text = '灯光\r舞美'
    hdr_cells1[9].paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hdr_cells1[9].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    hdr_cells1[10].text = '声音\r质量'
    hdr_cells1[10].paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hdr_cells1[10].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    hdr_cells1[11].text = '声音\r音量'
    hdr_cells1[11].paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hdr_cells1[11].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    hdr_cells1[12].text = '声画\r协调'
    hdr_cells1[12].paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hdr_cells1[12].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    if blMerge:
        table.cell(2, 2).text = '综评'
        table.cell(2, 2).paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        table.cell(2, 2).vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    else:
        table.cell(7, 2).text = '综评'
        table.cell(7, 2).paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        table.cell(7, 2).vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        for i in range(5):
            table.cell(i+2, 2).text = '测试{}'.format(i+1)
            table.cell(i+2, 2).paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            table.cell(i+2, 2).vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    return table

if __name__=="__main__":
    huizong()
    cleardata()
    writetoEx(readtoPD(), bPrint=False)
    #cleardocument()
    #writedocument(sheet='宣传片', blMerge=True)


