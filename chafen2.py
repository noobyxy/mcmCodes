import pandas as pd
import copy
from pandas import DataFrame, Series

df = pd.read_excel('统计指标newb.xlsx',sheet_name=0) #读附件一的第一个表单
df2 = pd.read_excel('需要筛除的数据.xlsx',sheet_name=0)
book = {}
code1 = []
code2 = []
code3 = []
code4 = []
code5 = []
code6 = []
code7 = []
code8 = []
code9 = []
code10 = []
code11 = []

for i in df2.index.values:
    book[str(df2.iloc[i,0])] = 1
for i in df.index.values:
    id = df.iloc[i,0]
    if id not in book:
        continue
    code1.append(id)
    code2.append(df.iloc[i,1])
    code3.append(df.iloc[i,2])
    code4.append(df.iloc[i,3])
    code5.append(df.iloc[i,4])
    code6.append(df.iloc[i,5])
    code7.append(df.iloc[i,6])  
    code8.append(df.iloc[i,7])
    code9.append(df.iloc[i,8])
    code10.append(df.iloc[i,9])
    # code11.append(df.iloc[i,10])
dict1 = {}
dict1['公司代号'] = code1
dict1['流水'] = code2
dict1['净利润'] = code3
dict1['利润率'] = code4
dict1['作废发票比例'] = code5
dict1['负数发票比例'] = code6
dict1['增长率'] = code7
dict1['信誉评级'] = code8
dict1['违规记录'] = code9
dict1['企业前景'] = code10

# dict1['公司代号'] = code1
# dict1['利润率'] = code2
# dict1['现金流量'] = code3
# dict1['净利润'] = code4
# dict1['利润增长率'] = code5
# dict1['信誉评级'] = code6
# dict1['违规记录'] = code7
# dict1['作废发票比例'] = code8
# dict1['负数发票比例'] = code9
# dict1['企业前景'] = code10

df6 = pd.DataFrame(dict1)
df6.to_excel('不适用于topsis的情况new-2-b.xlsx', index=False)
print('done')