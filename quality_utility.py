# -*- coding: utf-8 -*-
import numpy as np
import scipy as sp
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

block_array = pd.Series([7,
                         4,
                         7,
                         12,
                         7,
                         13,
                         7,
                         2,
                         18,
                         8,
                         17,
                         8,
                         7,
                         10,
                         1,
                         0,
                         7,
                         5,
                         12,
                         25,
                         23,
                         17,
                         2,
                         20,
                         3,
                         4,
                         7,
                         10,
                         2,
                         7,
                         2,
                         1,
                         12,
                         10,
                         0,
                         0,
                         0])

yield_array = pd.Series([179,
                         171,
                         164,
                         166,
                         180,
                         164,
                         151,
                         177,
                         152,
                         144,
                         155,
                         170,
                         158,
                         147,
                         38,
                         0,
                         117,
                         135,
                         141,
                         137,
                         153,
                         148,
                         159,
                         123,
                         190,
                         185,
                         90,
                         160,
                         172,
                         70,
                         44,
                         181,
                         151,
                         143,
                         0,
                         0,
                         0])

# ######################## p control#####################################


def average_p_array(average_p, n):
    local_array = np.array([])
    for num in range(n):
        local_array.append(average_p)
    return local_array


def remove_zero(array):
    return np.array([x for x in array if x is not 0])


def generate_out(rate_array, ucl_array, lcl_array):
    out_array = []
    upper = 0
    lower = 0
    for i in range(len(rate_array)):
        if (rate_array[i] > ucl_array[i]):
            out_array.append(rate_array[i])
            upper += 1
        elif (rate_array[i] < lcl_array[i]):
            out_array.append(rate_array[i])
            lower += 1
        else:
            out_array.append(np.nan)
    return {
        "out_array": out_array,
        "upper": upper,
        "lower": lower
    }


def p_control_content(block_array, yield_array):
    average_p = sum(block_array) / sum(yield_array)
    binomial = average_p * (1 - average_p)

    ucl_array = average_p + 3 * np.sqrt(binomial / yield_array)
    lcl_array = average_p - 3 * np.sqrt(binomial / yield_array)

    ucl_mean = average_p + 3 * np.sqrt(binomial / np.mean(yield_array))
    lcl_mean = average_p - 3 * np.sqrt(binomial / np.mean(yield_array))

    u = 1.96  # 95%置信。alpha = 0.05
    interval_top = average_p + np.sqrt(binomial / sum(yield_array)) * u
    interval_bot = average_p - np.sqrt(binomial / sum(yield_array)) * u

    rate_array = block_array / yield_array

    upper = generate_out(rate_array, ucl_array, lcl_array)["upper"]
    lower = generate_out(rate_array, ucl_array, lcl_array)["lower"]

    for i in range(len(rate_array)):
        if np.isnan(rate_array[i]):
            rate_array[i] = 0

    rate_array = remove_zero(rate_array)

# ================================ ========================================
    std_bi = np.sqrt(len(rate_array) * binomial)
    print("average_p", average_p)
    print("正常std", np.std(rate_array))
    print("样本二项std", std_bi)
    print("样本率二项std", np.sqrt(binomial / len(rate_array)))
    print("ucl_mean", ucl_mean)
    print("lcl_mean", lcl_mean)
    print("ucl-lcl", ucl_mean - lcl_mean)
    # print(np.sqrt(binomial / len(rate_array)))
    # print("规格中心M", (ucl_mean - lcl_mean))
    print("Z", (ucl_mean - average_p) / np.std(rate_array))
    cp = (ucl_mean - average_p) / np.std(rate_array)
    print(average_p - np.sqrt(binomial / sum(yield_array)) * 1.96,
          average_p + np.sqrt(binomial / sum(yield_array)) * 1.96)
    # print(sp.stats.norm.interval(0.95, loc=average_p, scale=std_error))
# ========================================================================

    content = ("本月过程能力控制如下。总缺陷比率为{aver_p}%，UCL为{ucl}，LCL为{lcl}。"
               "在{total}组样本总量当中，超过UCL上限的有{up}组样本，超过LCL下限的有{low}组样本，"
               "一共有{up_low}组样本超出了控制范围。"
               "缺陷置信区间为[{t_top},{t_bot}]，缺陷比率的标准差σ为{sigma}，过程能力为{cp}。"
               ).format(aver_p=round(average_p * 100, 4),
                        ucl=round(ucl_mean, 6),
                        lcl=round(lcl_mean, 6),
                        total=len(yield_array),
                        up=upper, low=lower, up_low=lower + upper,
                        t_top=round(interval_top, 6),
                        t_bot=round(interval_bot, 6),
                        sigma=round(np.std(rate_array), 4),
                        cp=round(cp, 2)
                        )
    return content


def p_control_plot(block_array, yield_array, plot_file_name):
    average_p = sum(block_array) / sum(yield_array)
    binomial = average_p * (1 - average_p)
    ucl_array = average_p + 3 * np.sqrt(binomial / yield_array)
    lcl_array = average_p - 3 * np.sqrt(binomial / yield_array)
    average_p_array = [average_p] * len(yield_array)
    rate_array = block_array / yield_array

    plt.figure(0)
    label_dict = {"UCL": ucl_array,
                  "AVERAGE_P": average_p_array,
                  "LCL": lcl_array}
    for label, array in label_dict.items():
        plt.plot(array, "--", label=label)

    plt.plot(rate_array, "kp-")
    out_array = generate_out(rate_array, ucl_array, lcl_array)["out_array"]
    plt.plot(out_array, "ro")
    plt.ylabel("缺陷占总样本比率")
    plt.xlabel("样本组数")
    plt.legend()
    plt.title("P控制图")
    plt.savefig(plot_file_name)
    plt.close(0)
    return plot_file_name


def defect_rate_scatter_plot(block_array, yield_array, plot_file_name):
    rate_array = remove_zero(block_array) / remove_zero(yield_array) * 100
    plt.figure(0)
    plt.scatter(remove_zero(yield_array), rate_array, label="缺陷散点图")
    plt.ylabel("缺陷比率%")
    plt.xlabel("样本")
    plt.legend()
    plt.title("每日缺陷比率和每日总样本的散点图")
    plt.savefig(plot_file_name)
    plt.close(0)
    return plot_file_name


def defect_rate_accum_plot(block_array, yield_array, plot_file_name):
    df_tmp = pd.DataFrame()
    df_tmp["block"] = remove_zero(block_array)
    df_tmp["yield"] = remove_zero(yield_array)
    df_accum = df_tmp.cumsum()
    plt.figure(0)
    df_accum["rate"] = df_accum["block"] / df_accum["yield"] * 100
    df_accum["rate"].plot()
    plt.ylabel("累积缺陷比率%")
    plt.xlabel("样本序数")
    plt.title("累积缺陷率变化图")
    plt.savefig(plot_file_name)
    plt.close(0)
    return plot_file_name


def defect_rate_hist_plot(block_array, yield_array, plot_file_name):
    df_tmp = pd.DataFrame()
    df_tmp["block"] = remove_zero(block_array)
    df_tmp["yield"] = remove_zero(yield_array)
    df_tmp["rate"] = remove_zero(block_array) / remove_zero(yield_array) * 100
    plt.figure(0)
    df_tmp.to_excel("test.xlsx")
    plt.hist(df_tmp["rate"].dropna())
    plt.ylabel("频率")
    plt.xlabel("缺陷比率%")
    plt.title("缺陷率分布图")
    plt.savefig(plot_file_name)
    plt.close(0)
    return plot_file_name


if __name__ == '__main__':
    print("thisfile")


# ######################################################################
