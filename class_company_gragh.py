import pandas as pd

df = pd.read_excel('附件2：302家无信贷记录企业的相关数据.xlsx',sheet_name=0) #读附件一的第一个表单
df2 = pd.read_excel('每月利润.xlsx',sheet_name=0) #读附件一的第二个表单
company2class = {}
allkindclass = set([])

def transToIndex(year,month):
    tmp = str(year)
    tmp = tmp + str('-')
    if month < 10:
        tmp = tmp + str('0')
        tmp = tmp + str(month)
    else:
        tmp = tmp + str(month)
    return tmp

class_month_data = {}
dict_output = {}

for i in df.index.values:
    tmp = df.loc[i,'企业代号']
    class2 = df.loc[i,'类型']
    allkindclass.add(class2)
    company2class[tmp] = class2
for i in df2.index.values:
    tmp = df2.loc[i,'公司代号']
    class2 = company2class[tmp]
    tmpclass = {}
    if class2 in class_month_data:
        tmpclass = class_month_data[class2]
    else:
        class_month_data[class2] = {}
    for year in range(2016,2021):
        for month in range(1,13):
            if year == 2020 and month > 2:
                continue
            datastr = transToIndex(year,month)
            if datastr in tmpclass:
                tmpclass[datastr] += df2.loc[i, datastr]
            else:
                tmpclass[datastr] = 0
                tmpclass[datastr] += df2.loc[i, datastr]
    class_month_data[class2] = tmpclass

# for year in range(2016,2020):
#     for month in range(1,13):
#         for class2 in allkindclass:
#             date = transToIndex(year, month)
#             dict_output[date].append(class_month_data[class2][date])


df6 = pd.DataFrame(class_month_data)
df6.to_excel('test123.xlsx', index=False)