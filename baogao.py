from docx import Document
from docx.shared import Inches, Pt, Mm
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import re
import pandas as pd
import numpy as np


def initBaogao(file='./demo.docx'):
    document = Document()
    # 设置一个空白样式
    style = document.styles['Normal']
    # 设置西文字体
    style.font.name = u'宋体'#'Times New Roman'
    # 设置中文字体
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')#u'宋体')#
    # 获取段落样式
    paragraph_format = style.paragraph_format
    # 首行缩进0.74厘米，即2个字符
    paragraph_format.first_line_indent = Mm(7.4)

    #标题级别
    heading1 = 0
    heading2 = 1
    sections = document.sections
    current_section = sections[-1]
    #第一章
    p = document.add_heading('第一章 本期评测综述', heading1)
    '''
    p1 = document.add_paragraph()
    run = p1.add_run(u'第一章 本期评测综述')
    run.font.name = u'宋体'
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), u'微软雅黑')
    '''
    p = document.add_paragraph('A plain paragraph having some ')
    p.add_run('bold').bold = True
    p.add_run(' and some ')
    p.add_run('italic.').italic = True
    heading_first = ['一、参评节目数量',
                     '二、综合得分等级及占比',
                     '三、技术质量优秀节目、不达标节目列表',
                     '四、本期亮点',
                     '五、本期报告用语说明', ]
    for heading in heading_first:
        head = document.add_heading(heading, heading2)
        #head.element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
        document.add_paragraph('please input some words.')
    p_shuoming = ['台外录制：以台外演播室录制为主，包括北京的演播室、联合录制、西院五楼新媒体演播室',
                  '台内录制：以台内演播室录制为主，包括800演播厅、4heading10演播厅、300演播厅、260演播厅、120演播室、110演播室、70演播室',
                  '台外制作：是指在台外制作、但不包括在北京制作',
                  '包装制作：在我台云平台高清制作网、由电视制作中心包装制作人员完成',
                  '高清自制：在我台云平台高清制作网、由编辑人员自行完成',
                  '120自制：在电视制作中心120演播室、由编辑人员自行完成',
                  '影视自制：在影视频道制作网、由编辑人员自行完成',
                  '少儿自制：在少儿频道制作网、由编辑人员自行完成',
                  '农民自制：在农民制作网、由编辑人员自行完成',
                  '广告自制：在广告发展公司制作']
    for p in p_shuoming:
        document.add_paragraph(p)
    #第二章
    #document.add_page_break()
    document.add_section()
    document.add_heading('第二章 节目技术质量评测结果', heading1)
    heading_first = ['一、参评节目',
                     '二、综合得分排序', ]
    for heading in heading_first:
        document.add_heading(heading, heading2)
        document.add_paragraph('please input some words.')
    #第三章
    document.add_section()
    document.add_heading('第三章  数据分析', heading1)
    heading_first = ['一、按频道分析',
                     '二、按录制地点分析',
                     '三、按制作方式分析',
                     '专家意见及建议']
    for heading in heading_first:
        document.add_heading(heading, heading2)
        document.add_paragraph('please input some words.')

    document.save(file)


def move_table_after(table, paragraph):
    tbl, p = table._tbl, paragraph._p
    p.addnext(tbl)


def canping_program(file='demo.docx'):
    df = pd.read_excel('database.xlsx')
    df = df[['序号', '节目名称', '播出时间']]
    inRow = df.shape[0] // 2 + 1
    document = Document(file)
    #将表格插入指定位置
    for p in document.paragraphs:
        if re.match("^Heading \d+$", p.style.name):
            if p.text == '一、参评节目':
                table = document.add_table(rows=inRow, cols=7, style='Table Grid')
                move_table_after(table, p)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.cell(0, 3).merge(table.cell(inRow-1, 3))
    table.autofit = False
    table.columns[0].width = Mm(12)
    table.columns[1].width = Mm(42)
    table.columns[2].width = Mm(24)
    table.columns[3].width = Mm(4)
    table.columns[4].width = Mm(12)
    table.columns[5].width = Mm(42)
    table.columns[6].width = Mm(24)
    columns = df.columns.to_list()
    for i in range(2):
        for column in columns:
            cell1 = table.cell(0, i*4+columns.index(column))
            cell1.text = column
            cell1.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    #让正确转行
    inRow -= 1
    for index, row in df.iterrows():
        for i in range(3):
            cell1 = table.cell(index % inRow + 1, index // inRow * 4 + i)
            cell1.text = str(row[columns[i]])
            cell1.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    document.save(file)

def zonghe_fen(file='demo.docx'):
    df = pd.read_excel('database.xlsx')
    df = df[['排名', '序号', '节目名称', '频道', '播出时间', '录制地点',
             '制作方式', '制片人', '主观', '客观', '总分', '等级']]
    df['主观'] = df['主观'].astype(np.int64)
    df['客观'] = df['客观'].astype(np.int64)
    df['总分'] = df['总分'].astype(np.int64)
    document = Document(file)
    #将表格插入指定位置
    for p in document.paragraphs:
        if re.match("^Heading \d+$", p.style.name):

            if p.text == '二、综合得分排序':
                print(p.text)
                table = document.add_table(rows=df.shape[0]+1, cols=df.shape[1], style='Table Grid')
                move_table_after(table, p)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.columns[0].width = Mm(7.4)
    table.columns[1].width = Mm(7.4)
    table.columns[2].width = Mm(30)
    table.columns[3].width = Mm(10.9)
    table.columns[4].width = Mm(21.5)
    table.columns[5].width = Mm(20.6)
    table.columns[6].width = Mm(18)
    table.columns[7].width = Mm(14.4)
    table.columns[8].width = Mm(7.4)
    table.columns[9].width = Mm(9.1)
    table.columns[10].width = Mm(7.4)
    table.columns[11].width = Mm(10.9)

    columns = df.columns.to_list()
    for column in columns:
        cell1 = table.cell(0, columns.index(column))
        cell1.text = column
        cell1.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        #设置字体大小
        for run in cell1.paragraphs[0].runs:
            font = run.font
            font.size = Pt(10)

    for index, row in df.iterrows():
        for i in range(len(row)):
            cell1 = table.cell(index + 1, i)
            cell1.text = str(row[columns[i]])
            cell1.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            # 设置字体大小
            for run in cell1.paragraphs[0].runs:
                font = run.font
                font.size = Pt(10)
    '''
    #设置表格宋体大小
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    font = run.font
                    font.size = Pt(10)
    '''
    document.save(file)

def fenxi_pindao(file='demo.docx'):
    df = pd.read_excel('database.xlsx')
    df = df[['频道', '节目名称', '播出时间', '录制地点',
             '制作方式', '制片人', '主观', '客观', '总分', '等级']]
    df['主观'] = df['主观'].astype(np.int64)
    df['客观'] = df['客观'].astype(np.int64)
    df['总分'] = df['总分'].astype(np.int64)
    pindao = ['卫视', '经济', '都市', '影视', '少儿', '公共', '农民']

    result ={}
    fenxi = []
    for i in pindao:
        df_temp = df[df['频道']==i]
        temp = df_temp.loc[:,'等级'].value_counts()
        temp = temp.rename(i)
        s = pd.Series([df_temp['总分'].max(), df_temp['总分'].min(), df_temp['总分'].mean()],
                      index=['最大值', '最小值', '平均值'])
        s = s.rename(i)
        temp = temp.append(s)
        print(temp)
        fenxi.append(temp)
        df_temp = df_temp.sort_values(by='总分', ascending=False)
    w = pd.DataFrame(fenxi)
    s = w.columns.to_list()
    dengji = ['优秀', '良好', '良', '及格', '不及格']
    for i in s:
        dengji.pop(i)

    document = Document(file)
    #将表格插入指定位置
    for p in document.paragraphs:
        if re.match("^Heading \d+$", p.style.name):
            if p.text == '一、按频道分析':
                print(p.text)
                table = document.add_table(rows=df.shape[0]+1, cols=df.shape[1], style='Table Grid')
                move_table_after(table, p)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    #table.columns[0].width = Mm(7.4)
    table.columns[1].width = Mm(7.4)
    table.columns[2].width = Mm(30)
    table.columns[3].width = Mm(10.9)
    table.columns[4].width = Mm(21.5)
    table.columns[5].width = Mm(20.6)
    table.columns[6].width = Mm(18)
    table.columns[7].width = Mm(14.4)
    table.columns[8].width = Mm(7.4)
    table.columns[9].width = Mm(9.1)
    table.columns[10].width = Mm(7.4)
    table.columns[11].width = Mm(10.9)

    columns = df.columns.to_list()
    for column in columns:
        cell1 = table.cell(0, columns.index(column))
        cell1.text = column
        cell1.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        #设置字体大小
        for run in cell1.paragraphs[0].runs:
            font = run.font
            font.size = Pt(10)

    for index, row in df.iterrows():
        for i in range(len(row)):
            cell1 = table.cell(index + 1, i)
            cell1.text = str(row[columns[i]])
            cell1.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            # 设置字体大小
            for run in cell1.paragraphs[0].runs:
                font = run.font
                font.size = Pt(10)
    '''
    #设置表格宋体大小
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    font = run.font
                    font.size = Pt(10)
    '''
    document.save(file)

if __name__ == '__main__':
    initBaogao()
    #canping_program()
    #zonghe_fen()

