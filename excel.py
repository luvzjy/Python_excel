# -*- coding: utf-8 -*-
import xlrd
import xlsxwriter
from datetime import date,datetime

file = "demo.xlsx"

def read_excel():
    # 打开文件
    workbook = xlrd.open_workbook(filename=file)
    # 获取所有sheet
    print(workbook.sheet_names())
    # print workbook.sheet_names() # [u'Sheet1']
    sheet2_name = workbook.sheet_names()[0]

    # 根据sheet索引或者名称获取sheet内容
    sheet2 = workbook.sheet_by_index(0) # sheet索引从0开始
    sheet2 = workbook.sheet_by_name('Sheet1')

    # sheet的名称，行数，列数
    print (sheet2.name,sheet2.nrows,sheet2.ncols)

    # 获取整行和整列的值（数组）
    rows = sheet2.row_values(0) # 获取第四行内容
    cols = sheet2.col_values(0) # 获取第三列内容
    cols1 =sheet2.col_values(1)
    # print (rows)
    # print (cols)
    # print (cols1)

    # 获取单元格内容
    # print (sheet2.cell(1,0).value.encode('utf-8'))
    # print (sheet2.cell_value(1,0).encode('utf-8'))
    # print (sheet2.row(1)[0].value.encode('utf-8'))

    # print (sheet2.cell(1,0).value)
    # print (sheet2.cell_value(1,0))
    # print (sheet2.row(1)[0].value)
    
    # 获取单元格内容的数据类型
    # print (sheet2.cell(1,0).ctype)
    # 使用if语句嵌套循环
    # for t in cols1:
    #     if t!="":
    #         for i in cols:

    #             if i==t and t!="":
    #                 print(i)
    notin=[]
    for i in cols:
        if i not in cols1 and i!="":
            notin.append(i)
    return notin
            
    


def write_excel():
    notin=read_excel()
#     # print(notin)

    workbook=xlsxwriter.Workbook('woork12.xlsx')
    
    worksheet=workbook.add_worksheet()
    worksheet.set_column("A:A",20)
    worksheet.write_column('A1', notin)
    workbook.close()


        

if __name__ == '__main__':
    write_excel()