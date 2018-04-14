import subprocess
import os
import matplotlib.pyplot as plt
import matplotlib
import pandas as pd
import numpy as np
matplotlib.style.use('ggplot')
plt.rcParams["font.sans-serif"] = ["Microsoft Yahei"]
plt.rcParams["axes.unicode_minus"] = False


def get_month(the_date):
    return the_date // 100


def selection_criteria(ss):
    head_len = 50
    tail_len = 50
    seg_len = 100
    bias = 0.02
    calc_array = ss[ss.notnull()].iloc[head_len:-tail_len]
    wedge_delta = abs(calc_array.max() - calc_array.min())
    beyond_num = 0
    total_num = calc_array.iloc[:-(seg_len - 1)].shape[0]
    for start in calc_array.iloc[:-(seg_len - 1)].index:
        seg = calc_array[start:(start + seg_len - 1)]
        if (seg.max() - seg.min()) > bias:
            beyond_num += 1
    return ["共计{:04}个100m区间".format(total_num),
            "其中{:04}个区间超差".format(beyond_num),
            "主体段偏差为{}um".format(wedge_delta * 1e6 // 1000)]


class PartTable():

    def __init__(self, line, part):
        self.table = pd.read_excel("part_table.xlsx")
        self.part_table = self.table.loc[
            (self.table["LINE"] == line) & (self.table["PART"] == part)]
        self.index = self.part_table.index[0]
        self.signal_name = self.part_table.loc[
            self.index, "SIGNAL"].replace('\\\\', '\\')
        self.single_dca_file = self.part_table.loc[
            self.index, "DCAFILE"] + "_POND.dca"

    def get_month(the_date):
        return the_date // 100

    def get_dca_file(self, data_dir, product_date, coil_id):
        return "/".join([data_dir,
                         "{}".format(get_month(product_date)),
                         "{}".format(product_date),
                         coil_id,
                         self.single_dca_file])


class PondReader():

    def __init__(self, dcafile, signalname):
        self.raw_cmd = "pndex.exe"
        self.cmd = " ".join([self.raw_cmd, dcafile, signalname])
        self.p = subprocess.Popen(
            self.cmd, shell=True, stdout=subprocess.PIPE).stdout
        self.raw_series = pd.Series(self.p.read().decode().split(","))


line = 1580
root_dir = "e:/silicon_fb"
data_dir = "i:/1580hrm"

feedback_date = 20180410
product_date = 20180408
MAX_INDEX = 1300

coil_id_col = "卷号"
coil_id_table = pd.read_excel(
    root_dir + '/{date}/{date}_热卷信息.xlsx'.format(date=feedback_date),
    header=1)
coil_id_list = coil_id_table[coil_id_col]

pt = PartTable(line, "wedge_40")
df = pd.DataFrame(index=range(0, MAX_INDEX))
for coil_id in coil_id_list:
    dca_file = pt.get_dca_file(data_dir, product_date, coil_id)
    pr = PondReader(dca_file, pt.signal_name)
    df[coil_id] = pr.raw_series

    print("Complete {}".format(coil_id))


print(df)
# df.convert_objects(convert_numeric=True)
df.to_excel("test.xlsx")
# df = pd.read_excel("test.xlsx")

dest_dir = "../%d/plot" % feedback_date
if not os.path.exists(dest_dir):
    os.makedirs(dest_dir)

for coil_id in coil_id_list:
    plt.figure(0)
    the_series = pd.to_numeric(df.loc[df[coil_id].notnull()][coil_id])
    the_series.plot()
    stat_result = selection_criteria(the_series)
    plt.title("热卷%s" % coil_id + ", ".join(stat_result))
    plt.xlabel("带钢长度")
    plt.ylabel("楔形偏差mm")
    plt.xticks(range(0, the_series.shape[0], 100), fontsize=10, rotation=17)
    plt.yticks(np.arange(-0.05, 0.05, 0.005), fontsize=10, rotation=17)

    plt.ylim(-0.05, +0.05)
    plt.savefig(
        "/".join(
            [dest_dir, coil_id + "_".join(stat_result[1:]) + ".png"]
        )
    )
    plt.close(0)
    print("complete %s" % coil_id)
