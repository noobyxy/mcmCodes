import pandas as pd
import copy
from pandas import DataFrame, Series

df = pd.read_excel('附件2：302家无信贷记录企业的相关数据.xlsx',sheet_name=0) #读附件一的第一个表单
score = {}
company_id = []
company_score = []
for i in range(0,23):
    cl = df.loc[i,'前景打分']
    num = df.loc[i,'分数']
    score[cl] = num
for i in df.index.values:
    tmp = df.loc[i,'企业代号']
    leixing = df.loc[i,'类型']
    company_id.append(tmp) #加入公司的id信息
    company_score.append(score[leixing])
dict1 = {}
dict1['公司代号'] = company_id
dict1['分数'] = company_score
df6 = pd.DataFrame(dict1)
df6.to_excel('前景得分情况.xlsx', index=False)
print('done')