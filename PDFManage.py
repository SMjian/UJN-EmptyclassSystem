# -*- coding:utf-8 -*-


import xlrd
import xlwt
from xlutils.copy import copy
import importlib
import os
import sys
import time
import pdfplumber

Root_dir = r'./学工在线各部干事课表'
"""保存有空课的学生的名字"""
table1 = [[[], [], [], [],[]],
          [[], [], [], [], []],
          [[], [], [], [], []],
          [[], [], [], [], []]]
"""创造 0 1码"""
excel = [[0,0,0,0,0],
         [0,0,0,0,0],
         [0,0,0,0,0],
         [0,0,0,0,0]]

#创建excel表
def CreatExcel():
    To_nrows = ['星期一', '星期二', '星期三', '星期四', '星期五']
    To_ncols = ['上午12节', '上午34节', '下午12节', '下午34节']
    style = xlwt.XFStyle()  # 初始化样式
    al = xlwt.Alignment()
    al.horz = 0x02  # 设置水平居中
    al.vert = 0x01  # 设置垂直居中
    al.wrap = 1     # 自动换行
    style.alignment = al
    book = xlwt.Workbook()  # 新建工作簿
    table = book.add_sheet('main', cell_overwrite_ok=True)  # 添加主要的表格信息
    book.save(filename_or_stream='output.xls')
    data = xlrd.open_workbook('output.xls', formatting_info=True)
    excel = copy(wb=data)  # 完成xlrd对象向xlwt对象转换
    excel_table = excel.get_sheet(0)  # 获得要操作的页
    table = data.sheets()[0]
    nrows = table.nrows  # 获得行数
    ncols = table.ncols  # 获得列数
    x = y = 1
    for i in To_nrows:
        excel_table.write(0, x, i,style)
        x += 1
    for i in To_ncols:
        excel_table.write(y, 0, i,style)
        y += 1
    excel.save('output.xls')

#向Excel中写入数据
def WriteExcel(index):
    data = xlrd.open_workbook('output.xls', formatting_info=True)
    excel = copy(wb=data)  # 完成xlrd对象向xlwt对象转换
    excel_table = excel.get_sheet(0)  # 获得要操作的页
    table = data.sheets()[0]
    style = xlwt.XFStyle()  # 初始化样式
    al = xlwt.Alignment()
    al.horz = 0x02  # 设置水平居中
    al.vert = 0x01  # 设置垂直居中
    al.wrap = 1     # 自动换行
    style.alignment = al
    nrows = table.nrows  # 获得行数
    ncols = table.ncols  # 获得列数
    x = y = 1
    for num_x in index:
        for num_y in num_x:
            excel_table.write(y,x,num_y,style)
            excel_table.col(x).width = 8888
            x += 1
        x = 1
        y += 1
    excel.save('output.xls')



"""原本程序的代码"""
importlib.reload(sys)
timel = time.time()
        
"""获得根文件名"""
def fileManage(rootdir):
    files = os.listdir(rootdir)
    return files

"""深度检测是否为匹配PDF"""
def Detection(x):
    new_x = []
    for i in x:
        if i != " " and i != "\n":
            new_x.append(i)
    return new_x

#保存读出的PDF存入与该程序相同根目录的问价夹下名字可改
def Save(EXCEL,addsname,F,txtname):
    """"想要保存的文件名字"""
    output = open(txtname+'.txt','a+')
    output.write(F+' ')
    output.write(addsname+'\t')
    output.write(str(EXCEL)+'\n')
    output.close()

"""主要读取PDF函数"""
def tableToMatrix(rootdir,file,F):
    num_x = 0
    num_y = 0
    """检测是否为异常PDF     1为异常   0为非异常"""
    tap = 0
    gets_pdf = file

    """将PDF转化为01矩阵"""
    with pdfplumber.open(rootdir+'\\'+gets_pdf) as pdf:
    
        first_page = pdf.pages[0]

        text = first_page.extract_text()
        pos = text.rfind('课表')
        names = text[:pos]
        tables = first_page.extract_tables()
        for table in tables:
            for index_r, row in enumerate(table):
                if index_r < 2 or index_r > 8 or index_r % 2 ==1:
                    continue
                for index_c, cell in enumerate(row):
                    if index_c < 2 or index_c >6:
                        continue
                    if cell != "":
                        excel[num_y][num_x] = 1
                        num_x+=1
                    else:
                        excel[num_y][num_x] = 0
                        table1[num_y][num_x].append(names+'|')
                        num_x+=1
                num_x = 0
                num_y+=1
        new_table = excel
        """做1 2 3等级判断"""
        num_x = 0
        num_y = 0
        for y in new_table:
            num_x = 0
            for x in y:
                if num_y == 1 or num_y == 3:
                    continue
                else:
                    if x == 0 and new_table[num_y+1][num_x] == 1:
                        new_table[num_y][num_x] = 2
                    num_x+=1
            num_y+=1
        """判断是否为满课或是PDF异常"""
        pos = text.rfind('课表')
        addsname = text[:pos]
        new_text = Detection(text)
        if new_text[pos+32] != "时":
            return 1
        else:
            Save(excel,addsname,F,'output_0_1')                                     #调用保存函数
            Save(new_table,addsname,F,'output_0_1_2')
            return 0
        pass
    pass



if __name__=='__main__':
    ErrorManage = open('ErrorManage.txt','a+')                          #创造一个错误处理文件     已经从正常文件中调了出来
    for F in fileManage(Root_dir):
        try:
            for file in fileManage(Root_dir+'\\'+F):
                try:
                    if tableToMatrix((Root_dir+'\\'+F),file,F) == 1:
                        ErrorManage.write(F+'   '+file+'\n')
                    else:
                        pass
                except BaseException as e:
                    pass                                                    #错误的是python处理文件时出来的直接跳过
                else:
                    print(F+'   '+file+'    已完成'+'\n')                  #打点逐个输出查看错误
        except BaseException as e:
            pass
    ErrorManage.close()
    CreatExcel()
    WriteExcel(table1)
    print('全部完成')
