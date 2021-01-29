# -*- coding: utf-8 -*-
"""
Created on Fri DEC 13 2020

@author: Jiangth
"""
# save xls files as xlsx in the subdirectories of raw_data 
import win32com.client as win32
import os
path = r'e:\casst\bookdown\enforcement\project2020\data'
for dirpath,dirnames,filenames in os.walk(path):
    for filename in filenames:
        fname = os.path.join(dirpath,filename)
        print(fname)
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(fname)
        wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
        wb.Close()                               #FileFormat = 56 is for .xls extension
        excel.Application.Quit()
        

"""
# fun example
import win32com.client as win32
import os
def transform(parent_path,out_path):
    fileList = os.listdir(parent_path)  #文件夹下面所有的文件
    num = len(fileList)
    for i in range(num):
        file_Name = os.path.splitext(fileList[i])   #文件和格式分开
        if file_Name[1] == '.xls':
            transfile1 = parent_path+'\\'+fileList[i]  #要转换的excel
            transfile2 = out_path+'\\'+file_Name[0]    #转换出来excel
            excel=win32.gencache.EnsureDispatch('excel.application')
            pro=excel.Workbooks.Open(transfile1)   #打开要转换的excel
            pro.SaveAs(transfile2 + ".xlsx", FileFormat=51)  # 另存为xlsx格式
            pro.Close()
            excel.Application.Quit()

if __name__=='__main__':
    path1 = r"e:\casst\bookdown\enforcement\enforcement\project2020\data\raw_data\2020\qg\a6"  #待转换文件所在目录
    path2 = r"e:\casst\bookdown\enforcement\enforcement\project2020\data\raw_data\2020\qg\a6"  #转换文件存放目录
    transform(path1, path2)
"""
