import os
from pdf2txt import pdf2txt
import xlrd  # 导入模块
from xlutils.copy import copy  # 导入copy模块

#循环解析成txt
def walkfile(file):
    rb = xlrd.open_workbook(excel_path)  # 打开weng.xls文件
    wb = copy(rb)  # 利用xlutils.copy下的copy函数复制
    ws = wb.get_sheet(1)  # 获取表单0
    x = 1
    for root, dirs, files in os.walk(file):
        # 遍历文件
        for f in files:
            file_path = os.path.join(root, f)
            file_ext = file_path.rsplit('.', maxsplit=1)
            if file_ext[1] != 'pdf':
                continue;
            print(file_path)
            txt = pdf2txt(file_path, 5, 5, 0)
            ws.write(x, 0, getcpxs(txt))  # 改变（0,0）的值
            txt_path = file_path.replace('pdf','txt')
            file = open(txt_path, 'w')
            file.write(txt)
            file.close()
            x = x + 1
    wb.save(excel_path)

#循环提取内容
def walkfiletxt(file,excel_path):
    rb = xlrd.open_workbook(excel_path)  # 打开weng.xls文件
    wb = copy(rb)  # 利用xlutils.copy下的copy函数复制
    ws = wb.get_sheet(0)  # 获取表单0
    x = 1
    for root, dirs, files in os.walk(file):
        # 遍历文件
        for f in files:
            file_path = os.path.join(root, f)
            file_ext = file_path.rsplit('.', maxsplit=1)
            qar = root.split('\\')
            if file_ext[1] != 'txt':
                continue;
            file = open(file_path, "r")
            txt = file.read()
            temp = txt.split('\n')
            ws.write(x, 0, qar[len(qar)-1])
            ws.write(x, 1, getcpxs(temp))  # 改变（0,0）的值
            ws.write(x, 2, getclcs(temp))  # 改变（0,0）的值
            ws.write(x, 3, getclph(temp))  # 改变（0,0）的值
            ws.write(x, 4, getrclzt(temp))  # 改变（0,0）的值
            ws.write(x, 5, getrclgg(temp))  # 改变（0,0）的值
            ws.write(x, 6, getcllh(temp))  # 改变（0,0）的值
            ws.write(x, 7, getrclh(temp))  # 改变（0,0）的值
            x = x + 1
    wb.save(excel_path)
                #print(os.path.join(root, f))

#获取产品形式
def getcpxs(x):
    all = ['厚板','棒材','管材','型材','锻件','铸件','sheet','Plate','Bar','Rod','Tubing','Shape','Extrusion','Forging','Casting']
    for i in x:
        if('Material Type' in i):
            for j in all:
                if(j in i):
                    return j
    return ''

#获取获取材料厂商
def getclcs(x):
    all = ['Aleris','Aleris']
    for i in x:
        for j in all:
            if (j in i):
                return j
    return ''

#获取材料牌号
def getclph(x):
    all = ['7475','6061','7050','7075']
    for i in x:
        for j in all:
            if (j in i):
                return j
    return ''

#获取热处理状态
def getrclzt(x):
    for i in x:
        if('Temper' in i):
            y = i.split('Temper')
            return y[1].replace(' ','').replace(':','').replace('“','').replace(';','')
    return ''

#获取材料规格
def getrclgg(x):
    for i in x:
        if('inch' in i):
            i = i.replace('“',' ').replace(':',' ')
            i = ' '.join(i.split())
            y = i.split(' ')
            return y[len(y)-2]
    return ''

#获取材料炉号
def getcllh(x):
    for i in x:
        if('Cast No' in i):
            y = i.split(' ')
            for j in y:
                if('-' in j):
                    return j
    return ''

#获取热处理号
def getrclh(x):
    for i in x:
        if('Batch No' in i):
            i = ' '.join(i.split())
            y = i.split(' ')
            for j in y:
                if(j.isdigit() and len(j)==10):
                    return j
    return ''

def writeexcel(excel_path):
    rb = xlrd.open_workbook(excel_path)  # 打开weng.xls文件
    wb = copy(rb)  # 利用xlutils.copy下的copy函数复制
    ws = wb.get_sheet(1)  # 获取表单0
    ws.write(7, 1, 'changed!')  # 改变（0,0）的值
    #ws.write(8, 0, label='好的')  # 增加（8,0）的值
    wb.save(excel_path)

if __name__ == '__main__':
    #pdf_path = './test_files/QAR-ZM17-0057.pdf'
    pdf_path = 'D:/数据抓取'
    excel_path = 'D:/数据抓取/爱励-数据抓取模板.xls'
    #writeexcel(excel_path)
    #print(walkfile(pdf_path))
    print(walkfiletxt(pdf_path,excel_path))