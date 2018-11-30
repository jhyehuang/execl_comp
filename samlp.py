#!/usr/bin/env python
#encoding:utf-8
 
from xlrd import open_workbook,cellname
 
def main():
    #打开文件
    xlsfilename='c:/temp/workbook1.xls'
    book=open_workbook(xlsfilename,formatting_info=True)
    print 'book.nsheets:',book.nsheets
 
    #三种方式得到sheet对象
    #begin
    print 'begin get sheet objects:'
    #索引
    mysheets=[]
    for sheet_index in range(book.nsheets):
        sheet=book.sheet_by_index(sheet_index)
        mysheets.append(sheet)
        print sheet
 
    #名字
    print book.sheet_names()
    for sheet_name in book.sheet_names():
        print book.sheet_by_name(sheet_name)
 
    #对象集
    for sheet in book.sheets():
        print sheet
    #end
 
    cursheet=mysheets[0]
 
    cinfomap=cursheet.colinfo_map
    rinfomap=cursheet.rowinfo_map
    #print 'colinfo_map:',cinfomap
    #print 'rowinfo_map:',rinfomap
    print 'show row and column hidden attribute:'
    for key,value in cinfomap.items():
        print 'col %d hidden attribute:%d' % (key,value.hidden)
    for key,value in rinfomap.items():
        print 'col %d hidden attribute:%d' % (key,value.hidden)
 
 
    #内省一个表格，包括隐藏行列
    #begin
    print 'begin introspect a sheet(including all):'
    print 'name:',cursheet.name
    print 'nrows:',cursheet.nrows
    print 'ncols:',cursheet.ncols
 
    for row_index in range(cursheet.nrows):
        for col_index in range(cursheet.ncols):
            tcell=cursheet.cell(row_index,col_index)
            cname=cellname(row_index,col_index)
            cvalue=tcell.value
            print cname,'=',cvalue
    #end
 
    #内省一个表格，不包括隐藏行列
    #begin
    print 'begin introspect a sheet(not including hidden):'
    print 'name:',cursheet.name
    print 'nrows:',cursheet.nrows
    print 'ncols:',cursheet.ncols
 
    for row_index in range(cursheet.nrows):
        if rinfomap.get(row_index,0) and rinfomap[row_index].hidden==1:
            continue
        for col_index in range(cursheet.ncols):
            if cinfomap.get(col_index,0) and cinfomap[col_index].hidden==1:
                continue
            tcell=cursheet.cell(row_index,col_index)
            cname=cellname(row_index,col_index)
            cvalue=tcell.value
            print cname,'=',cvalue
    #end
 
 
if __name__ == '__main__':
    main()