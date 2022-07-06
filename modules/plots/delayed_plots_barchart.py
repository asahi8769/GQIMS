import os
import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
from modules.plots.pipelines.weekly_data_organizer_delayed_dates import OvsDelayedDataOrganizer
from datetime import datetime
from utils_.config import CUSTOMERS
import numpy as np
from utils_.functions import make_dir
import utils_.matplotlib_config as matplotlib_config


def draw_plot4(regular_month=False, yyyymm=None):

    yyyymm = int(datetime.today().strftime("%Y%m")) if not regular_month else yyyymm

    data_organizer = OvsDelayedDataOrganizer(from_to=False, yyyymm=yyyymm)
    df = data_organizer.get_delayed_days()

    data_organizer = OvsDelayedDataOrganizer(from_to=True, yyyymm=yyyymm)
    df_acc = data_organizer.get_delayed_days()

    fig = plt.figure(constrained_layout=True, figsize=(10, 4))
    gs = GridSpec(2, 1, figure=fig)
    ax_1 = fig.add_subplot(gs[0:1, 0:1])
    ax_2 = fig.add_subplot(gs[1:2, 0:1])

    width = 0.40
    margin = 0

    null_custs = [i for i in CUSTOMERS if i not in df.index.unique().tolist()]
    null_custs_acc = [i for i in CUSTOMERS if i not in df_acc.index.unique().tolist()]

    for null_cust in null_custs:
        df.loc[null_cust] = {"해외결재소요일": 0, "국내접수소요일": 0}

    for null_cust in null_custs_acc:
        df_acc.loc[null_cust] = {"해외결재소요일": 0, "국내접수소요일": 0}

    customers = sorted(CUSTOMERS, key=lambda x: 0 if x not in df.index else df.loc[x, "해외결재소요일"].mean() + df.loc[x, "국내접수소요일"].mean(), reverse=True)
    x = customers + ['종합']
    r = np.arange(len(x))

    data_a = []
    data_b = []
    data_a_acc = []
    data_b_acc = []

    for n, customer in enumerate(x):

        if customer == "종합":
            df_cust = df
            df_cust_acc = df_acc
        else:
            df_cust = df.loc[customer]
            df_cust_acc = df_acc.loc[customer]

        data_a.append(df_cust['해외결재소요일'].mean())
        data_b.append(df_cust['국내접수소요일'].mean())
        data_a_acc.append(df_cust_acc['해외결재소요일'].mean())
        data_b_acc.append(df_cust_acc['국내접수소요일'].mean())

    ch_1 = ax_1.bar(r - width / 2 - margin, data_a, width, label='해외결재소요일', color='#AED6F1', edgecolor='grey')
    ch_2 = ax_1.bar(r + width / 2 + margin, data_b, width, label='국내접수소요일', color='#F5B7B1', edgecolor='grey')
    ch_3 = ax_2.bar(r - width / 2 - margin, data_a_acc, width, label='해외결재소요일', color='#AED6F1', edgecolor='grey')
    ch_4 = ax_2.bar(r + width / 2 + margin, data_b_acc, width, label='국내접수소요일', color='#F5B7B1', edgecolor='grey')

    for p in ch_1:
        height = int(round(p.get_height(), 0))
        label = height if height != 0 else ""
        ax_1.annotate(f'{label}', xy=(p.get_x() + p.get_width() / 2, height), xytext=(-2, 3),
                     textcoords="offset points", ha='center', va='bottom')

    for p in ch_2:
        height = int(round(p.get_height(),0))
        label = height if height != 0 else ""
        ax_1.annotate(f'{label}', xy=(p.get_x() + p.get_width() / 2, height), xytext=(2, 3), textcoords="offset points",
                     ha='center', va='bottom')

    for p in ch_3:
        height = int(round(p.get_height(),0))
        label = height if height != 0 else ""
        ax_2.annotate(f'{label}', xy=(p.get_x() + p.get_width() / 2, height), xytext=(-2, 3),
                     textcoords="offset points", ha='center', va='bottom')

    for p in ch_4:
        height = int(round(p.get_height(),0))
        label = height if height != 0 else ""
        ax_2.annotate(f'{label}', xy=(p.get_x() + p.get_width() / 2, height), xytext=(2, 3), textcoords="offset points",
                     ha='center', va='bottom')

    ax_1.legend(prop={'size': 10})
    ax_1.set_title(f"{str(yyyymm)[:4]}년도 {str(yyyymm)[4:6]}월 PPR 처리 소요일 (평균 소요 일 수)")

    ax_1.set_xticks(r)
    ax_1.set_xticklabels(x, fontsize=10)
    ax_2.set_xticks(r)
    ax_2.set_xticklabels(x, fontsize=10)
    ax_1.set_ylabel("당월")
    ax_2.set_ylabel("누적")

    ax_1.set(yticklabels=[])
    ax_1.tick_params(left=False, bottom=False)
    ax_1.spines["right"].set_visible(False)
    ax_1.spines["left"].set_visible(False)
    ax_1.spines["top"].set_visible(False)

    ax_2.set(yticklabels=[])
    ax_2.tick_params(left=False, bottom=False)
    ax_2.spines["right"].set_visible(False)
    ax_2.spines["left"].set_visible(False)
    ax_2.spines["top"].set_visible(False)

    fig.tight_layout()

    folder_name = os.path.join("spawn", "plots", datetime.now().strftime("%Y-%m-%d"))
    make_dir(folder_name)

    img_name = os.path.abspath(os.path.join(folder_name, f"{yyyymm}_stacked_PPR처리소요일_plot4.png"))
    plt.savefig(img_name)

    return plt, img_name


if __name__ == "__main__":
    os.chdir(os.pardir)
    os.chdir(os.pardir)

    yyyymm = 202111

    plt, img_name = draw_plot4(regular_month=True, yyyymm=yyyymm)

    plt.show()
