import pandas as pd
import copy
from pandas import DataFrame, Series

#读利润信息文件的第一个表单
df = pd.read_excel('每月利润.xlsx',sheet_name=0) 
#得到进项价税合计的信息，此为现金流通量
df2 = pd.read_excel('进项.xlsx',sheet_name=0)
tt_moneys = []#总利润
tt_months = []#计算得到的总月份
tt_money_out = []#进项钱数
money_chuxiang = []
avg_money = []#平均利润
max_moneys = []#最大利润
min_moneys = []#最小利润
avg_moneys = []
rates = []#利润率计算方法1
rate1 = []#利润率计算方法2
error_info = []#利润率计算方法1出错时的提示信息

## 月份和年转化为字符串的方法
def transToIndex(year,month):
    tmp = str(year)
    tmp = tmp + str('-')
    if month < 10:
        tmp = tmp + str('0')
        tmp = tmp + str(month)
    else:
        tmp = tmp + str(month)
    return tmp

## 利润率计算方法1
def rate_adding(rate=None):
    global rate_error
    if out_money == 0:
        rate_error = True
        rate += 1
        return rate
    rate += tmp_monoy / out_money
    return rate

## 开始循环计算
for i in df.index.values:
    #找出非0的第一个年月和最后一个年月
    tmp_id = df.loc[i, '公司代号']
    first_not_zero_year = 0
    first_not_zero_month = 0
    last_not_zero_year = 0
    last_not_zero_month = 0
    flag = False
    #利用标记循环寻找
    for j in range(2016, 2020):
        if flag == True:
            break
        str_time = ""
        for k in range(1, 13):
            str_time = transToIndex(j,k)
            if df.loc[i,str_time] != 0:
                first_not_zero_year = j
                first_not_zero_month = k
                flag = True
                break
    flag = False
    for j in range(2020, 2015,-1):
        if flag == True:
            break
        str_time = ""
        for k in range(12, 0, -1):
            str_time = transToIndex(j,k)
            if df.loc[i,str_time] != 0:
                last_not_zero_year = j
                last_not_zero_month = k
                flag = True
                break
    #计算利润之和
    money_sum = 0
    out_sum = 0
    out_sum_truly = 0
    tt_month = 0
    max_money = -1
    min_money = float('inf')
    rate = 0
    rate_error = False

    for j in range(2016, 2021):
        for k in range(1, 13):
            str_time = transToIndex(j,k)
            out_sum_truly += df2.loc[i,str_time]
    money_chuxiang.append(out_sum_truly)

    ## 找出first和last后利用循环进行累加计算
    ## 如果年份不同时的计算方法
    if first_not_zero_year != last_not_zero_year:
        for q in range(first_not_zero_month,13):
            str_time = transToIndex(first_not_zero_year, q)
            tmp_monoy = df.loc[i,str_time]
            out_money = df2.loc[i,str_time]
            rate = rate_adding(rate)
            if tmp_monoy > max_money:
                max_money = tmp_monoy
            if tmp_monoy < min_money:
                min_money = tmp_monoy
            money_sum += tmp_monoy
            out_sum += out_money
            tt_month += 1
        for year in range(first_not_zero_year+1,last_not_zero_year):
            for q in range(1,13):
                str_time = transToIndex(year, q)
                tmp_monoy = df.loc[i,str_time]
                out_money = df2.loc[i, str_time]
                rate = rate_adding(rate)
                if tmp_monoy > max_money:
                    max_money = tmp_monoy
                if tmp_monoy < min_money:
                    min_money = tmp_monoy
                money_sum += tmp_monoy
                out_sum += out_money
                tt_month += 1
        for q in range(1,last_not_zero_month+1):
            str_time = transToIndex(last_not_zero_year, q)
            tmp_monoy = df.loc[i,str_time]
            out_money = df2.loc[i, str_time]
            rate = rate_adding(rate)
            if tmp_monoy > max_money:
                max_money = tmp_monoy
            if tmp_monoy < min_money:
                min_money = tmp_monoy
            money_sum += tmp_monoy
            out_sum += out_money
            tt_month += 1
    ## 年份相同时循环月份即可
    else:
        for q in range(first_not_zero_month,last_not_zero_month + 1):
            str_time = transToIndex(first_not_zero_year, q)
            tmp_monoy = df.loc[i,str_time]
            out_money = df2.loc[i, str_time]
            rate = rate_adding(rate)
            if tmp_monoy > max_money:
                max_money = tmp_monoy
            if tmp_monoy < min_money:
                min_money = tmp_monoy
            money_sum += tmp_monoy
            out_sum += out_money
            tt_month += 1
    tt_moneys.append(money_sum)
    tt_months.append(tt_month)
    tt_money_out.append(out_sum)
    max_moneys.append(max_money)
    min_moneys.append(min_money)
    avg_moneys.append(money_sum/tt_month)
    rates.append(rate/tt_month)
    rate1.append(money_sum/out_sum_truly)
    if rate_error == True:
        error_info.append("利润率计算结果出现问题")
    else:
        error_info.append("")
dict1 = {}
arrayname = []
## 格式化输出
for i in range(124,426):
    name = str('E') + str(i)
    arrayname.append(name)
dict1["公司代号"] = arrayname
dict1["总利润"] = tt_moneys
dict1["总支出"] = money_chuxiang
dict1["总月份"] = tt_months
dict1["最大利润"] = max_moneys
dict1["最小利润"] = min_moneys
dict1["平均利润"] = avg_moneys
dict1["利润率"] = rates
dict1["利润率2"] = rate1
dict1["备注"] = error_info
df6 = pd.DataFrame(dict1)
df6.to_excel('平均值分析_new.xlsx', index=False)