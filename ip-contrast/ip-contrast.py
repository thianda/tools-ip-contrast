# -*- coding: utf-8 -*-
import xlrd

def contrast():
  # 打开文件
  zgg = xlrd.open_workbook(r'demo.xlsx')
  #jt = xlrd.open_workbook(r'ip-jt.xlsx')
  #print zgg.name
  print zgg.sheet_names()
  sheet2 = zgg.sheet_by_index(1) # sheet索引从0开始
  #sheet2 = zg.sheet_by_name('sheet2')
  # sheet的名称，行数，列数
  print sheet2.name,sheet2.nrows,sheet2.ncols
  # 获取整行和整列的值（数组）
  rows = sheet2.row_values(3) # 获取第四行内容
  #cols = sheet2.col_values(2) # 获取第三列内容
  print rows
  #print cols
  # 获取单元格内容
  print sheet2.cell(1,0).value.encode('utf-8')
  print sheet2.cell_value(1,0).encode('utf-8')
  print sheet2.row(1)[0].value.encode('utf-8')
if __name__ == '__main__':
  contrast()