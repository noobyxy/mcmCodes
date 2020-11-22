import sympy
import pandas as pd

def solve(y):
    y = 1 - y
    if y > 0.9:
        return 0.15
    x = sympy.Symbol("x",real = True)
    res = sympy.solve([-4420*x**4 + 2243*x**3  -456.5*x**2 + 47.44*x -1.287 - y],[x])
    ans = 0
    # print(res)
    for i in res:
        for j in i:
            # print(j)
            if j >= 0 and j < 0.16:
                ans = j
    if ans < 0.04:
        return 0.04
    if ans > 0.15:
        return 0.15
    return ans

df = pd.read_excel('不适用于topsis的情况new.xlsx',sheet_name=0)
companys = []
rates = []
for i in df.index.values:
    #找出非0的第一个年月和最后一个年月
    tmp_id = df.loc[i, '公司代号']
    rate = solve(df.loc[i,'归一化'])
    companys.append(tmp_id)
    rates.append(rate)
dict1 = {}
dict1['公司代号'] = companys
dict1['计算得税率'] = rates
df6 = pd.DataFrame(dict1)
df6.to_excel('公司税率映射2-new.xlsx', index=False)