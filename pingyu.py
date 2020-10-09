import pandas as pd
import os
from readdocx import writeExcel


def wblist(filedir='./节目单', extension='.xls'):
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
    for root, dirs, files in os.walk(filedir, topdown=False):
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

def readworkbooks(workbooksname):
    df = pd.read_excel(workbooksname)
    df = df[['序号', '节目名称', '总分', '评语']]
    df = df.dropna(axis=0, how='all')
    df.reset_index(drop=True, inplace=True)
    df.fillna('', inplace=True)
    return df

if __name__ == "__main__":
    # file = writedocument(sheet='常规', blMerge=False)
    # 处理目录
    workbooksPath = './主观评测/常规节目'
    workbooksNames = wblist(filedir=workbooksPath, extension='.xls')
    # data = initdict(os.path.join(documnetsNames[0][0], documnetsNames[0][1]))
    df = readworkbooks(os.path.join(workbooksNames[0][0], workbooksNames[0][1]))
    data = df[['序号', '节目名称']]
    data['评语'] = ''

    for workbooks in workbooksNames:
        filename = os.path.join(workbooks[0], workbooks[1])
        df = readworkbooks(filename)
        name = workbooks[1].split("-")
        # data[name[0]] = df.pop("总分")

        data['评语'] = data.pop('评语') + df['评语'] # + '(' + name[0] + ')'
        print(data['评语'][:1])
        print(workbooks[1] + '完成了，共计{}个文件，还剩下{}个文件'.format(len(workbooksNames),
                                                       len(workbooksNames) - workbooksNames.index(workbooks) - 1))

    with pd.ExcelWriter('评语汇总.xlsx', engine='xlsxwriter') as writer:
        data.to_excel(writer, sheet_name="汇总", startrow=0, index=False)