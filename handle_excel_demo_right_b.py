from typing import Dict, Any, Union

import pandas as pd
import copy
from pandas import DataFrame, Series

df = pd.read_excel('附件2：302家无信贷记录企业的相关数据.xlsx',sheet_name=0) #读附件一的第一个表单
df2 = pd.read_excel('附件2：302家无信贷记录企业的相关数据.xlsx',sheet_name=1) #读附件一的第二个表单
df3 = pd.read_excel('附件2：302家无信贷记录企业的相关数据.xlsx',sheet_name=2) #读附件一的第三个表单

#company_info = df.values

#print("获取到所有的值:\n{0}".format(company_info))#格式化输出

company_id = []
company_level = {}
company_is_break = {}

company_in_money_day = {} #计算每天得到的钱
company_out_money_day = {}  #计算每天支出的钱
company_in_money_month = {} #计算每月销项的总钱数
company_out_money_month = {} #计算每月进项的总钱数
company_failed_time = {} #计算出项或者进项中交易失败的订单数
company_failed_total_money = {} #计算出项或者进项中交易失败的订单金额
company_negative_time = {}#计算出项或者进项中被退货的订单数
company_negative_money = {}#计算出项或者进项中被退货的订单金额
company_total_time = {} #计算进项或者销项一共的次数

#将日期类型的变量转化为字符串，保留到月份
def date2month( str ):

    cnt = 0
    for index in range(len(str)):
        if str[index] == '-':
            cnt = cnt + 1
            if cnt == 2:
                return str[0:index]

#销项的操作：
#以天为单位的操作
def add_in_money_day():
    global day_money_data, tmp_money
    if tmp_id in company_in_money_day:
        day_money_data = company_in_money_day[tmp_id]
        if date in day_money_data:
            tmp_money = day_money_data[date]
            tmp_money = tmp_money + add_money
            day_money_data[date] = tmp_money
        else:
            day_money_data[date] = add_money
    else:
        day_money_data = {}
        day_money_data[date] = add_money
        company_in_money_day[tmp_id] = day_money_data
#以月份为单位的操作
def add_in_money_month():#inmoney即为企业销项发票所得的总金额
    global day_money_data, tmp_money
    if tmp_id in company_in_money_month:
        day_money_data = company_in_money_month[tmp_id]
        if month in day_money_data:
            tmp_money = day_money_data[month]
            tmp_money = tmp_money + add_money
            day_money_data[month] = tmp_money
        else:
            day_money_data[month] = add_money
    else:
        day_money_data = {}
        day_money_data[month] = add_money
        company_in_money_month[tmp_id] = day_money_data

##本程序段用于销项统计金额信息，顺带统计作废发票，负数发票的信息，并计算利润值##
##本实例中计算了销项中的特殊发票数目
##如需计算进项中的相应情况复制即可
for i in df2.index.values:
    status = df2.loc[i,'发票状态']
    tmp_id = df2.loc[i,'企业代号']
    date = df2.loc[i,'开票日期']
    month = date2month(str(date))
    add_money = df2.loc[i,'金额']
    if tmp_id in company_total_time:
        company_total_time[tmp_id] += 1
    else:
        company_total_time[tmp_id] = 1
    if status == '作废发票':
        if tmp_id in company_failed_time:
            company_failed_time[tmp_id] = company_failed_time[tmp_id] + 1
            company_failed_total_money[tmp_id] = company_failed_total_money[tmp_id] + add_money
        else:
            company_failed_time[tmp_id] = 1
            company_failed_total_money[tmp_id] = add_money
        continue
    if add_money < 0:
        if tmp_id in company_negative_time:
            company_negative_time[tmp_id] = company_negative_time[tmp_id] + 1
            company_negative_money[tmp_id] = company_negative_money[tmp_id] - add_money
        else:
            company_negative_time[tmp_id] = 1
            company_negative_money[tmp_id] = -add_money
    # add_in_money_day()
    add_in_money_month()

#进项的操作：
#以天为单位的操作
def minus_money_day():
    global day_money_data, tmp_money
    if tmp_id in company_out_money_day:
        day_money_data = company_out_money_day[tmp_id]
        if date in day_money_data:
            tmp_money = day_money_data[date]
            tmp_money = tmp_money + minus_money
            day_money_data[date] = minus_money
        else:
            day_money_data[date] = minus_money
    else:
        day_money_data = {}
        day_money_data[date] = minus_money
        company_out_money_day[tmp_id] = day_money_data
#以月份为单位的操作
def minus_money_month():
    global day_money_data, tmp_money
    if tmp_id in company_out_money_month:
        day_money_data = company_out_money_month[tmp_id]
        if month in day_money_data:
            tmp_money = day_money_data[month]
            tmp_money = tmp_money + minus_money
            day_money_data[month] = tmp_money
        else:
            day_money_data[month] = minus_money
    else:
        day_money_data = {}
        day_money_data[month] = minus_money
        company_out_money_month[tmp_id] = day_money_data

##本程序段用于进项统计价税合计的信息
for i in df3.index.values:
    status = df3.loc[i,'发票状态']
    tmp_id = df3.loc[i,'企业代号']
    date = df3.loc[i,'开票日期']
    month = date2month(str(date))
    minus_money = df3.loc[i,'金额']
    # if tmp_id in company_total_time:
    #     company_total_time[tmp_id] += 1
    # else:
    #     company_total_time[tmp_id] = 1
    # if status == '作废发票':
    #     if tmp_id in company_failed_time:
    #         company_failed_time[tmp_id] = company_failed_time[tmp_id] + 1
    #         company_failed_total_money[tmp_id] = company_failed_total_money[tmp_id] + minus_money
    #     else:
    #         company_failed_time[tmp_id] = 1
    #         company_failed_total_money[tmp_id] = minus_money
    #     continue
    # if minus_money < 0:
    #     if tmp_id in company_negative_time:
    #         company_negative_time[tmp_id] = company_negative_time[tmp_id] + 1
    #         company_negative_money[tmp_id] = company_negative_money[tmp_id] - minus_money
    #     else:
    #         company_negative_time[tmp_id] = 1
    #         company_negative_money[tmp_id] = -minus_money
    if status == '作废发票':
        continue

    # minus_money_day()
    minus_money_month()


# company_total_money_day = company_in_money_day.deepcopy()
# 深度复制进行利润差额的计算，用出项金额-进项金额
company_total_money_month = copy.deepcopy(company_in_money_month)
company_total_money_month_liushui = copy.deepcopy(company_in_money_month)
##具体操作过程
for i in company_out_money_month.keys():
    if i in company_total_money_month:
        for j in company_out_money_month[i].keys():
            if j in company_total_money_month[i]:
                m = company_total_money_month[i][j]
                company_total_money_month_liushui[i][j] = m + company_out_money_month[i][j]
                m = m - company_out_money_month[i][j]
                company_total_money_month[i][j] = m
            else:
                company_total_money_month[i][j] = -company_out_money_month[i][j]
                company_total_money_month_liushui[i][j] = company_out_money_month[i][j]
    else:
        tmp_dic = {}
        for j in company_out_money_month[i].keys():
            tmp_dic[j] = -company_out_money_month[i][j]
        company_total_money_month[i] = tmp_dic
        company_total_money_month_liushui[i] = company_out_money_month[i]

dict_to_write = {}
arrayname = []
arraymoney = []
timeline = []
## 计算附录二，计算附录一只需将范围改为range(1,124)即可
for i in range(124,426):
    name = str('E') + str(i)
    arrayname.append(name)
## 格式化输出Dataframe，利用字典和元组进行操作
dict_out = {}
dict_in = {}
dict_in["公司代号"] = arrayname
dict_out["公司代号"] = arrayname
dict_to_write["公司代号"] = arrayname
dict_to_write_liushui = {}
dict_to_write_liushui["公司代号"] = arrayname
for j in range(2016,2021):
    str_time = str(j)
    str_time = str_time + str('-')
    mem = str_time
    for k in range(1,13):
        str_time = mem
        if k < 10:
            str_time = str_time + str('0')
            str_time = str_time + str(k)
        else:
            str_time = str_time + str(k)
        timeline.append(str_time)
        arraymoney = []
        arraymoney_liushui = []
        arraymoney_out = []
        arraymoney_in = []
        for i in range(124,426):
            name = str('E') + str(i)
            if str_time in company_out_money_month[name]:
                arraymoney_out.append(company_out_money_month[name][str_time])
            else:
                arraymoney_out.append(0)

            if str_time in company_in_money_month[name]:
                arraymoney_in.append(company_in_money_month[name][str_time])
            else:
                arraymoney_in.append(0)

            if str_time in company_total_money_month[name]:
                arraymoney.append(company_total_money_month[name][str_time])
                arraymoney_liushui.append(company_total_money_month_liushui[name][str_time])
            else:
                arraymoney.append(0)
                arraymoney_liushui.append(0)
        dict_in[str_time] = arraymoney_in
        dict_out[str_time] = arraymoney_out
        dict_to_write[str_time] = arraymoney
        dict_to_write_liushui[str_time] = arraymoney_liushui
df4 = pd.DataFrame(dict_to_write)
df4.to_excel('每月利润.xlsx', index=False)#输出每月利润
df5 = pd.DataFrame(dict_to_write_liushui)
df5.to_excel('result_liushui.xlsx', index=False)#此流水量在本例程中无意义，将‘金额’改为‘价税合计’后位流水信息
df_in = pd.DataFrame(dict_in)
df_in.to_excel('出项a.xlsx',index=False)#分别输出销项，进项的金额信息，以备后用
df_out = pd.DataFrame(dict_out)
df_out.to_excel('进项a.xlsx',index=False)
failed_time_array = []
failed_money_array = []
negative_money_array = []
negative_time_array = []
total_time_array = []
for i in range(124,426):
    name = str('E') + str(i)
    if name in company_total_time:
        total_time_array.append(company_total_time[name])
    else:
        total_time_array.append(0)
    if name in company_failed_time:
        failed_time_array.append(company_failed_time[name])
        failed_money_array.append(company_failed_total_money[name])
    else:
        failed_time_array.append(0)
        failed_money_array.append(0)
    if name in company_negative_time:
        negative_time_array.append(company_negative_time[name])
        negative_money_array.append(company_negative_money[name])
    else:
        negative_time_array.append(0)
        negative_money_array.append(0)
# 输出特殊交易信息
write_negative = {}
write_negative["公司代号"] = arrayname
write_negative["失败交易次数"] = failed_time_array
write_negative["失败交易金额"] = failed_money_array
write_negative["负数交易次数"] = negative_time_array
write_negative["负数交易金额"] = negative_money_array
write_negative["总的订单数"] = total_time_array
df6 = pd.DataFrame(write_negative)
df6.to_excel('负数与失败订单统计.xlsx', index=False)
print('done')