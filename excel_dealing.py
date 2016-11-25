import numpy as np
import pandas as pd
import os
import xlrd
path='C:\\Users\\tana\\Desktop\\test'
file_list=[]
for root , dirs , files in os.walk(path):
    for file in files:
        file_list.append(os.path.join(root,file))
file_number=np.shape(file_list)[0]
print('总共有%d个文件'%file_number)
data=[]
col=['Project name', '部门', '规模','2016年11月', '2016年12月', '2017年1月', '2017年2月', '2017年3月']
for file_name in file_list:
    temp1=xlrd.open_workbook(file_name)
    sheet_names=temp1.sheet_names()
    for sheet_name in sheet_names:
        if sheet_name.find('人力需求')>0:
            temp2 = pd.read_excel(file_name, sheet_name)
            temp2=temp2[col]
            temp3 = temp2.fillna(0)
            temp4=temp3.sort_values(by=[ '部门', '规模','2016年11月','2016年12月','2017年1月', '2017年2月', '2017年3月'],
                                    ascending=[1 , 1 , 0 , 0 , 0 , 0 , 0])
    data.append(temp4)
for i in range(file_number):
    data[i].to_excel('C:\\Users\\tana\\Desktop\\test\\筛选'+str(i)+'.xls')
print('筛选文件输出完毕')
