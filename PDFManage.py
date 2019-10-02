# -*- coding:utf-8 -*-


import xlrd
import xlwt
from xlutils.copy import copy
import importlib
import os
import sys
import time
import pdfplumber

"""***************************挺重要的文件目录所在地*********************************"""
Root_dir = r'./学工在线各部干事课表'

"""原本程序的代码"""
importlib.reload(sys)
timel = time.time()

"""创造 0 1码"""
excel = [[0,0,0,0,0],[0,0,0,0,0],[0,0,0,0,0],[0,0,0,0,0]]
        
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
    ErrorManage.close()      
    print('全部完成')
