import xlrd
import os
import pandas as pd
pathroot = "C:/Users/Administrator/Desktop/yanzheng"
#xlsx=xlrd.open_workbook(pathroot,formatting_info=True)
#sheet=xlsx.sheet_by_index(0)
#print(sheet.cell(3,6).value)
data = []
list = os.listdir(pathroot) #列出文件夹下所有的目录与文件

# 这里是数据存储列表
excelBiaoTou = xlrd.open_workbook(pathroot + "/" + list[0]).sheets()[0].row_values(2)
# 这里是表头(第三行则是2)
#name = xlrd.open_workbook(path + "//" + fileListInThePath[0]).sheets()[0].cell(1,10).value
# 打开表格，读取内表1，单元格2行12列内容
sheetName = "test"
# 这里是表名

for i in range(0,len(list)):
    path = os.path.join(pathroot,list[i])
    if os.path.isfile(path):
        readExcel = xlrd.open_workbook(path)
        excleData = readExcel.sheets()[0]
        exclelNumOfHang = excleData.nrows
        # 获取表格行数（从第四行获取，range为3）
        for j in range(3,exclelNumOfHang):
            data.append(excleData.row_values(j))
            # 除去开头，逐行添加到存储列表内
newExcle = pd.DataFrame(data)
newExcle.columns = excelBiaoTou
newExcle.to_excel(pathroot + '/' + sheetName + '.xlsx',header = True,index = False)


'''
data = pd.read_excel('C:/Users/Administrator/Desktop/yanzheng\北云门镇全口径贫困户信息.xls',index_col='序号')
data[(data.人数 == 2) & (data.与户主关系=="户主")]
'''
