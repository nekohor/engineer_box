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
    selection = list(df_selection[grade_belongs])
    return selection


# ====================config============================
month_start = 201701
month_end = 201705

line_list = [2250]
selector_dict = {
    "钢种": ["Q235B"],
    # "SPHC-YS1"
    "目标厚度": [3.75]
}

# =======================logic=========================

month_list = range(month_start, month_end + 1)
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

df_all.to_excel('e:/analysis/thk/q235b_375.xlsx')
