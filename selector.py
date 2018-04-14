# -*- coding: utf-8 -*-
import numpy as np
import pandas as pd
import matplotlib
from dateutil.parser import parse
from datetime import datetime
import seaborn as sns
import os
import sys
import docx
from docx.shared import Inches
import openpyxl
# matplotlib.use('Agg')
import matplotlib.pyplot as plt
sns.set(color_codes=True)
sns.set(rc={'font.family': [u'Microsoft YaHei']})
sns.set(rc={'font.sans-serif': [u'Microsoft YaHei', u'Arial',
                                u'Liberation Sans', u'Bitstream Vera Sans',
                                u'sans-serif']})

# =====================function=========================


def grade_selection(grade_belongs):
    df_selection = pd.read_excel('e:/stat/data/used_grade.xlsx')
    return list(df_selection[grade_belongs])


def generate_month(start_month, end_month):
    y = start_month // 100
    m = start_month % 100
    month_list = []
    while (y * 100 + m) <= end_month:
        month_list.append((y * 100 + m))
        m = m + 1
        if m > 12:
            y = y + 1
            m = 1
    return month_list


def pre_treatment(df, line, item):
    if "evaluate" == item:
        df["热卷号"] = df["CoilID"] = df["CoilID"].map(lambda x: x.strip())
        df["钢种"] = df["AlloyName"] = df["AlloyName"].map(lambda x: x.strip())
        if line == 2250:
            df["钢种"] = [("MR" + x.strip()) if len(x) == 8 else x
                        for x in df["钢种"]]
        else:
            pass


# ====================config============================
line_list = [2250]
# item = "evaluate"
selector_dict = {
    # "板坯钢种": ["MBRYT34506", "MBRYT34515", "MBRYT50015", "MBRYT50016",
    #          "MBRYT60022", "MBRYT65001", "MBRYT70001", "MBRYT70002"]   #大梁钢
    "钢种": ["MGW600", "MGW600-H", "MGW600G", "MGW800", "MGW1300",
           "MGW1300D", "MGW1300K", "MGW700DB", "MGW600DB", "MGWLG"]
    # "SPHC-YS1"
    # "目标厚度": [3.75]
}
# to_file_name = 'e:/005统计/20180209_2017全年大梁钢统计/{line}肖肖.xlsx'
input_file_name = "e:/luna_summary/{}/{}/flatness/reasoned_data.xlsx"
to_file_name = 'e:/005统计/20180319_2250产线平直度不符原因数据汇总/{line}产线平直度不符原因数据汇总.xlsx'

# =======================logic=========================

month_list = generate_month(201707, 201802)
df_all = pd.DataFrame()
for line in line_list:
    for month in month_list:
        # ==== input ====
        # file_name = ("d:/data_collection_monthly/%d/%d/%d_%d%s.xlsx"
        #              % (line, month, month, line, item)
        #              )
        file_name = input_file_name.format(month, line)
        df = pd.read_excel(file_name)

        # ==== pre-treatment ====
        # pre_treatment(df, line, item)

        # ==== selection ====
        # for key, value in selector_dict.items():
        #     df = df.loc[df[key].isin(value)]

        # ==== appending ====
        df_all = df_all.append(df)
        print("complete! %d" % month)

    # ==== output ====
    # df_all.to_excel(
    #     to_file_name.format(line=line,
    #                         grade=selector_dict["板坯钢种"]))

    df_all.to_excel(
        to_file_name.format(line=line)
    )
