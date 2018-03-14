import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import random
import collections


def aimrate(child_size, total_size):
    return round(child_size / total_size * 100, 2)


def grade_selection(grade_belongs):
    df_selection = pd.read_excel('e:/stat/data/used_grade.xlsx')
    return list(df_selection[grade_belongs])


def strip_coil_id(df):
    df["热卷号"] = df["CoilID"].map(lambda x: x.strip())


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


def read_excel_data(line, month, source):
    file_name = ("d:/data_collection_monthly/%d/%d/%d_%d%s.xlsx" % (
        line, month, month, line, source)
    )
    df = pd.read_excel(file_name)
    return df


def append_monthly_table(line, month_list, tbl_name):
    monthly_tbl_dir = "d:/data_collection_monthly/{}/{}/{}_{}{}.xlsx"
    df_all = pd.DataFrame()
    for month in month_list:
        each_monthly_tbl = pd.read_excel(
            monthly_tbl_dir.format(line, month, month, line, tbl_name)
        )
        df_all = df_all.append(each_monthly_tbl)
        print("complete! %d" % month)
    return df_all


def attach_mth_tbl(month_list, to_file_name):
    df_all = pd.DataFrame()
    for month in month_list:
        each_monthly_tbl = pd.read_excel(
            to_file_name(month), header=1
        )
        df_all = df_all.append(each_monthly_tbl)
        print("complete! %d" % month)
    return df_all


def quadrant_set(df, os_dev_col, ds_dev_col, level):
    abs_os = df[os_dev_col].apply(
        lambda x: abs(x) + random.uniform(0, 0.001))
    abs_ds = df[ds_dev_col].apply(
        lambda x: abs(x) + random.uniform(0, 0.001))
    df.loc[(abs_os >= level) & (abs_ds >= level), "quad"] = 1
    df.loc[(abs_os < level) & (abs_ds >= level), "quad"] = 2
    df.loc[(abs_os < level) & (abs_ds < level), "quad"] = 3
    df.loc[(abs_os >= level) & (abs_ds < level), "quad"] = 4
    return df


def quadrant_stat(os_dev, ds_dev, level):
    total_size = os_dev.size
    fst_quad_size = os_dev.loc[(os_dev >= level) & (ds_dev >= level)].size
    scd_quad_size = os_dev.loc[(os_dev < level) & (ds_dev >= level)].size
    thd_quad_size = os_dev.loc[(os_dev < level) & (ds_dev < level)].size
    fth_quad_size = os_dev.loc[(os_dev >= level) & (ds_dev < level)].size
    quad_dict = collections.OrderedDict()
    quad_dict["一"] = ("双超", fst_quad_size)
    quad_dict["二"] = ("单超", scd_quad_size)
    quad_dict["三"] = ("合格", thd_quad_size)
    quad_dict["四"] = ("单超", fth_quad_size)
    result_list = []
    for k, v in quad_dict.items():
        result_list.append(
            "{}象限 ({}) 命中数{}\n{}象限 ({}) 命中率{}%".format(
                k, v[0], v[1], k, v[0], aimrate(v[1], total_size)
            ))
    return result_list


def silicon_crown_stat(os_dev, ds_dev, position):
    total_size = os_dev.size
    fst_quad_size = os_dev.loc[(os_dev >= level) & (ds_dev >= level)].size
    scd_quad_size = os_dev.loc[(os_dev < level) & (ds_dev >= level)].size
    thd_quad_size = os_dev.loc[(os_dev < level) & (ds_dev < level)].size
    fth_quad_size = os_dev.loc[(os_dev >= level) & (ds_dev < level)].size
    quad_dict = collections.OrderedDict()
    quad_dict["一"] = ("双超", fst_quad_size)
    quad_dict["二"] = ("单超", scd_quad_size)
    quad_dict["三"] = ("合格", thd_quad_size)
    quad_dict["四"] = ("单超", fth_quad_size)
    result_list = []
    for k, v in quad_dict.items():
        result_list.append(
            "{}象限 ({}) 命中数{}\n{}象限 ({}) 命中率{}%".format(
                k, v[0], v[1], k, v[0], aimrate(v[1], total_size)
            ))
    return result_list


def plot_deviation_on_plate(df, os_dev_col, ds_dev_col):
    df["abs_os_dev"] = df[os_dev_col].apply(
        lambda x: abs(x) + random.uniform(0, 0.001))
    df["abs_ds_dev"] = df[ds_dev_col].apply(
        lambda x: abs(x) + random.uniform(0, 0.001))
    df.plot(kind='scatter', x="abs_os_dev", y="abs_ds_dev", alpha=0.5)
    # df.plot(kind='hexbin', x="abs_os_dev", y="abs_ds_dev", gridsize=20)
    the_range = np.arange(0.0, 0.020, 0.001)
    level_x = 0.008
    level_y = 0.010
    var_zero = [x for x in the_range]
    const_x = [level_x for x in the_range]
    const_y = [level_y for x in the_range]

    plt.plot(var_zero, const_x, linestyle='--', color="k")
    plt.plot(const_x, var_zero, linestyle='--', color="k")
    plt.plot(var_zero, const_y, linestyle='--', color="r")
    plt.plot(const_y, var_zero, linestyle='--', color="r")

    quad_list = quadrant_stat(df["abs_os_dev"], df["abs_ds_dev"], level_x)

    plt.text(0.015, 0.015, quad_list[0])
    plt.text(-0.010, 0.015, quad_list[1])
    plt.text(-0.010, -0.010, quad_list[2])
    plt.text(0.015, -0.010, quad_list[3])

    aim_list = quadrant_stat(df["abs_os_dev"], df["abs_ds_dev"], level_y)
    plt.text(0.015, 0.020, aim_list[0], color="r")
    plt.text(-0.010, 0.020, aim_list[1], color="r")
    plt.text(-0.010, -0.005, aim_list[2], color="r")
    plt.text(0.015, -0.005, aim_list[3], color="r")
    return plt
