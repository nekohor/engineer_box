# -*- coding: utf-8 -*-
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from dateutil.parser import parse
from datetime import datetime
import seaborn as sns
import os
import sys
import openpyxl
import seaborn as sns
sns.set(color_codes=True)
sns.set(rc={'font.family': [u'Microsoft YaHei']})
sns.set(rc={'font.sans-serif': [u'Microsoft YaHei', u'Arial',
                                u'Liberation Sans', u'Bitstream Vera Sans',
                                u'sans-serif']})


def read_excel_data(line, month, source):
    file_name = ("d:/data_collection_monthly/%d/%d/%d_%d%s.xlsx" % (
        line, month, month, line, source)
    )
    df = pd.read_excel(file_name)
    return df
