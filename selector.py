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


# ====================config============================
line_list = [1580]
selector_dict = {
    "板坯钢种": ["MBRYT34506", "MBRYT34515", "MBRYT50015", "MBRYT50016",
             "MBRYT60022", "MBRYT65001", "MBRYT70001", "MBRYT70002"]
    # "SPHC-YS1"
    # "目标厚度": [3.75]
}

to_file_name = 'e:/005统计/20180209_2017全年大梁钢统计/{line}肖肖.xlsx'

# =======================logic=========================

month_list = generate_month(201701, 201712)
df_all = pd.DataFrame()
for line in line_list:
    for month in month_list:
        file_name = ("d:/data_collection_monthly/%d/%d/%d_%dexcel.xlsx"
                     % (line, month, month, line)
                     )
        df = pd.read_excel(file_name)

        for key, value in selector_dict.items():
            df = df.loc[df[key].isin(value)]

        df_all = df_all.append(df)
        print("complete! %d" % month)

    # df_all.to_excel(
    #     to_file_name.format(line=line,
    #                         grade=selector_dict["板坯钢种"]))

    df_all.to_excel(
        to_file_name.format(line=line)
    )
