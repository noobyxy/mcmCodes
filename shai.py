import pandas as pd

df = pd.read_excel('平均值分析_new.xlsx',sheet_name=0)
names = []
for i in df.index.values:
    tmp_id = df.loc[i, '公司代号']
    rate = df.loc[i,'利润率2']
    if rate > 3:
        names.append(tmp_id)
dict1 = {}
dict1['公司代号'] = names
df6 = pd.DataFrame(dict1)
df6.to_excel('需要筛除的数据.xlsx', index=False)
print('done')