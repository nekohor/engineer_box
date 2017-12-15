# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
from dateutil.parser import parse
import datetime


def composite_data(line, subject, MONTH_START, MONTH_END):
    df_all = pd.DataFrame()
    month_list = list(range(MONTH_START, MONTH_END + 1))
    print(month_list)
    for month in month_list:
        file_name = "d:/data_collection_monthly/%d/%d/%d_%d%s.xlsx" % (
            line, month, month, line, subject)
        df = pd.read_excel(file_name)
        df["月份"] = month
        # df = df[col_list]
        df.index = list(range(df.shape[0]))
        # if month == MONTH_START:
        #     df_all = df
        # else:
        #     df_all = df_all.append(df)
        df_all = df_all.append(df)
    df_all.index = list(range(df_all.shape[0]))
    return df_all


def selection(df, df_selection, col):

    selection = list(df_selection[col])
    df = df.loc[df[col].isin(selection)]
    df.index = list(range(df.shape[0]))
    return df


def silicon_selection(df):
    steel_grade_catego = pd.read_excel('e:/stat/data/used_grade.xlsx')
    refer = list(steel_grade_catego[u'硅钢'])
    df = df.loc[(df[u'钢种'].isin(refer))]
    df.index = list(range(df.shape[0]))
    return df


def add_shift(df):
    df.index = list(range(df.shape[0]))
    shiftArray = ("白", "白", "小", "小", "休", "大", "大", "休")
    startTime = datetime.datetime(2015, 1, 1, 0, 0)
    for i in range(df.shape[0]):
        print(df.loc[i, "热卷号"] + " " + str(i))
        if type(df.loc[i, "结束日期"]) == float:
            continue

        tmp = []
        tmp.append(df.loc[i, "结束日期"])
        tmp.append(df.loc[i, "结束时间"])
        df.loc[i, "生产时间"] = " ".join(tmp)

        productTimeArray = parse(str(df.loc[i, "生产时间"]))

        if productTimeArray.hour < 8:
            df.loc[i, "班次"] = "大"
        elif (productTimeArray.hour >= 8) & (productTimeArray.hour < 16):
            df.loc[i, "班次"] = "白"
        else:
            df.loc[i, "班次"] = "小"

        aimTime = datetime.datetime(productTimeArray.year,
                                    productTimeArray.month,
                                    productTimeArray.day)

        timeDelta = aimTime - startTime
        days = int(timeDelta.days)
        arrayPosition = days % 8

        duty = {"甲": 0, "乙": 0, "丙": 0, "丁": 0}

        duty["甲"] = shiftArray[(arrayPosition - 1) % 8]
        duty["乙"] = shiftArray[(arrayPosition + 1) % 8]
        duty["丙"] = shiftArray[(arrayPosition + 3) % 8]
        duty["丁"] = shiftArray[(arrayPosition + 5) % 8]

        duty = dict((v, k) for k, v in duty.items())
        df.loc[i, "班别"] = duty[df.loc[i, "班次"]]


def parse_time(num):
    year = int(num // 1e10)
    month = int(num // 1e8 - num // 1e10 * 1e2)
    day = int(num // 1e6 - num // 1e8 * 1e2)
    hour = int(num // 1e4 - num // 1e6 * 1e2)
    minute = int(num // 1e2 - num // 1e4 * 1e2)
    second = int(num - num // 1e2 * 1e2)
    return datetime(year, month, day, hour, minute, second)


# Binning:
def binning(col, cut_points, labels=None):
    # Define min and max values:
    minval = col.min()
    maxval = col.max()

    # create list by adding min and max to cut_points
    break_points = [minval] + cut_points + [maxval]

    # if no labels provided, use default labels 0 ... (n-1)
    if not labels:
        labels = list(range(len(cut_points) + 1))

    # Binning using cut function of pandas
    colBin = pd.cut(col, bins=break_points, labels=labels, include_lowest=True)
    return colBin

# #Binning age:
# cut_points = [90,140,190]
# labels = ["low","medium","high","very high"]
# data["LoanAmount_Bin"] = binning(data["LoanAmount"], cut_points, labels)
# print pd.value_counts(data["LoanAmount_Bin"], sort=False)


# df_selection = pd.read_excel("e:/大缺陷分析/板形质量异议/热卷信息.xlsx")
# df = composite_data(2250, "excel", MONTH_START, MONTH_END)
# df = selection(df, df_selection, "热卷号")
# print(df)
# df.to_excel("e:/boom/df_need.xlsx")

# df = pd.read_excel("e:/boom/df_need.xlsx")
# add_shift(df)
# df.to_excel("e:/boom/df_need.xlsx")

# for line in line_list:
#     df = composite_data(line, "excel", MONTH_START, MONTH_END)
#     df.to_excel("e:/beike/raw_excel_data%d.xlsx" % (line))
