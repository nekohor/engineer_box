# -*- coding: utf-8 -*-
from jinja2 import Environment, PackageLoader
from collections import OrderedDict
import pandas as pd
import numpy as np
import matplotlib as mpl
import matplotlib.pyplot as plt
from dateutil.parser import parse
import sys
import os
import seaborn as sns
from docx.shared import Inches
sys.path.append("e:/box")

sns.set(color_codes=True)
sns.set(rc={'font.family': [u'Microsoft YaHei']})
sns.set(rc={'font.sans-serif': [u'Microsoft YaHei', u'Arial',
                                u'Liberation Sans', u'Bitstream Vera Sans',
                                u'sans-serif']})
# mpl.style.use('ggplot')


def pareto(label_array, data_array, savepath=""):
    # parameter
    show_max = 7
    plt.figure(figsize=(30, 50))
    # plt.figure()
    df = pd.DataFrame()
    df["data"] = data_array
    df.index = label_array
    df = df.sort_values(by="data", ascending=False)
    df["percent"] = df["data"].cumsum() / df["data"].sum() * 100

    # --- plot code ---
    df["data"].plot(kind='bar', color="b", align="center", rot=17)
    df["percent"].plot(color='r', secondary_y=True,
                       style='-o', linewidth=2, rot=17)
    # plt.xticks(rotation=17)

    x = np.arange(len(df.index))
    y = df["percent"]
    z = df["data"]
    for a, b in zip(x[:show_max], y[:show_max]):
        plt.text(a, b - 0.8, '%.2f%%' %
                 b, ha='center', va='bottom', fontsize=12)

    for a, b, c in zip(x, y, z):
        plt.text(a, 100 + 0.05, '%.2f' %
                 c, ha='center', va='bottom', fontsize=12, rotation=17)

    plt.ylabel("累积百分比")
    plt.xlabel("Pareto图")
    if __name__ == '__main__':
        plt.show()  # for test
    else:
        plt.savefig(savepath)
        plt.close(0)


def test():
    data = pd.DataFrame()

    data["label"] = [
        "A5（孔洞）",
        "AX（线状夹杂）",
        "D2（辊印）",
        "W压氧",
        "边部起筋",
        "边浪",
        "边损",
        "成分性能不合",
        "挫伤",
        "横折印",
        "夹杂",
        "结疤",
        "结疤  ",
        "孔洞",
        "来料凹坑",
        "来料边损",
        "来料挫伤",
        "来料辊印",
        "来料温度超差",
        "来料折叠印",
        "翘皮",
        "停车斑",
        "性能不合",
        "氧化铁皮压入",
        "一类边损",
        "折叠",
        "中浪"

    ]

    data["data"] = [
        11.779,
        15.050,
        6.540,
        258.460,
        2.268,
        41.473,
        6.230,
        76.830,
        0,
        90.205,
        337.600,
        34.860,
        13.485,
        58.735,
        8.890,
        11.220,
        22.200,
        23.980,
        34.270,
        20.195,
        988.220,
        27.500,
        488.832,
        251.040,
        26.460,
        0,
        30.949

    ]
    pareto(data["label"], data["data"])


def test2():
    data = pd.DataFrame()
    data["label"] = [
        "M610L和M510L单边浪",
        "Q500C批量浪形",
        "其它"

    ]

    data["data"] = [
        38049,
        7614.15,
        5000

    ]
    pareto(data["label"], data["data"])


if __name__ == '__main__':
    test()
