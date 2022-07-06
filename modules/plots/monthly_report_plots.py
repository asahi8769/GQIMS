from modules.ppm.variations import spq_per_plant, spq_per_month, apq_per_region
import numpy as np
import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
import matplotlib.ticker as plticker
from matplotlib import font_manager, rc
from utils_.config import CUSTOMERS
import os

try:
    font_path = "C:/Windows/Fonts/현대하모니+M.ttf"
    font = font_manager.FontProperties(fname=font_path).get_name()
except FileNotFoundError:
    font_path = os.path.join(mpl.get_data_path(), "fonts/ttf/현대하모니+M.ttf")
    font = font_manager.FontProperties(fname=font_path).get_name()

rc('font', family=font)


def draw_plot7(df_apq_per_region):
    x = df_apq_per_region['구분'].tolist()
    r = np.arange(len(x))

    var = df_apq_per_region['누적PPM'].tolist()

    fig = plt.figure(constrained_layout=True, figsize=(2, 3))
    gs = GridSpec(1, 1, figure=fig)

    fontsize = 9
    width = 0.40

    ax = fig.add_subplot(gs[0:1, 0:1])
    ax.set_xticks(r)
    ax.set_xticklabels(x, fontsize=fontsize)
    ax.tick_params(axis='x', rotation=45)
    ax.set(yticklabels=[])
    ax.tick_params(left=False, bottom=False)
    ax.set_ylim([0, max(var) * 1.15 if not max(var) == 0 else 1])
    ax.axvline(0.5, 0, 1, color='lightgray', linestyle='--', linewidth=0.5)

    aq = ax.bar(r, var, width, color='#AED6F1', edgecolor='grey', label='누적PPM')

    for n, p in enumerate(aq):
        height_tot = f"{round(var[n], 1):,}"
        label_tot = height_tot if height_tot != '0' else ""
        ax.annotate(f'{label_tot}', xy=(p.get_x() + p.get_width() / 2, round(var[n], 1)),
                    xytext=(0, 3), textcoords="offset points", ha='center', va='bottom', fontsize=fontsize)

    img_name = os.path.abspath(os.path.join("spawn", "plots", f"report_plot7.png"))

    plt.show()
    return img_name, plt


def draw_plot6(yyyymm, df_spq_per_month):
    month_range = [i for i in range(int(str(yyyymm)[:4] + '01'), int(str(yyyymm)[:4] + '13'), 1)]
    x = ['20.1분기', '20.2분기', '20.3분기', '20.4분기'] + month_range
    r = np.arange(len(x))
    aq_val = [36.7, 36.8, 25.8, 47.1] + [round(i, 1) if i !="" else 0 for i in [df_spq_per_month[df_spq_per_month["년월"] == i]['외관불량'].iloc[0] for i in month_range]]
    sq_val = [210.5, 415.4, 219.7, 300.5] + [round(i, 1) if i !="" else 0 for i in [df_spq_per_month[df_spq_per_month["년월"] == i]['기능/치수'].iloc[0] for i in month_range]]
    tot_val = [round(i + j,1) for i, j in zip(aq_val, sq_val)]

    fig = plt.figure(constrained_layout=True, figsize=(7, 3))
    gs = GridSpec(1, 1, figure=fig)

    fontsize = 9
    width = 0.40

    ax = fig.add_subplot(gs[0:1, 0:1])
    ax.set_xticks(r)
    ax.set_xticklabels(x, fontsize=fontsize)
    ax.tick_params(axis='x', rotation=45)
    ax.set(yticklabels=[])
    ax.tick_params(left=False, bottom=False)
    ax.set_ylim([0, max(tot_val) * 1.15 if not max(tot_val) == 0 else 1])
    ax.axvline(3.5, 0, 1, color='lightgray', linestyle='--', linewidth=0.5)

    aq = ax.bar(r, aq_val, width, color='#AED6F1', edgecolor='grey', label='외관')
    sq = ax.bar(r, sq_val, width, color='#F9E79F', edgecolor='grey', bottom=aq_val, label='기능/치수')

    for n, p in enumerate(aq):
        height = f"{round(aq_val[n], 1):,}"
        label = height if height != '0' else ""
        va = 'top'
        ax.annotate(f'{label}', xy=(p.get_x() + p.get_width() / 2, round(aq_val[n], 1)),
                    xytext=(0, -1), textcoords="offset points", ha='center', va=va, fontsize=fontsize)

    for n, p in enumerate(aq):
        height_tot = f"{round(tot_val[n], 1):,}"
        label_tot = height_tot if height_tot != '0' else ""
        ax.annotate(f'{label_tot}', xy=(p.get_x() + p.get_width() / 2, round(tot_val[n], 1)),
                    xytext=(0, 3), textcoords="offset points", ha='center', va='bottom', fontsize=fontsize)

    handles, labels = plt.gca().get_legend_handles_labels()
    order = [0, 1]
    ax.legend([handles[idx] for idx in order], [labels[idx] for idx in order], fontsize=fontsize, frameon=True,
              loc=1, ncol=2)

    img_name = os.path.abspath(os.path.join("spawn", "plots", f"report_plot6.png"))

    return img_name, plt


def draw_plot5(df_spq_per_plant):

    customers = df_spq_per_plant['고객사'].tolist()[1:]
    customers = sorted(customers, key=lambda x: df_spq_per_plant[df_spq_per_plant['고객사'] == x]['공급망불량'].iloc[0], reverse=True)

    x = ['20년 누적', '21년 누적'] + customers
    r = np.arange(len(x))
    aq_val = [36.9] + [df_spq_per_plant[df_spq_per_plant["고객사"] == i]['외관불량'].iloc[0] for i in ['21년도 종합'] + customers]
    sq_val = [267.0] + [df_spq_per_plant[df_spq_per_plant["고객사"] == i]['기능/치수'].iloc[0] for i in ['21년도 종합'] + customers]
    tot_val = [round(i + j,1) for i, j in zip(aq_val, sq_val)]

    fig = plt.figure(constrained_layout=True, figsize=(7, 3))
    gs = GridSpec(1, 1, figure=fig)

    fontsize = 9
    width = 0.50

    ax = fig.add_subplot(gs[0:1, 0:1])
    ax.set_xticks(r)
    ax.set_xticklabels(x, fontsize=fontsize)
    ax.tick_params(axis='x', rotation=45)
    ax.set(yticklabels=[])
    ax.tick_params(left=False, bottom=False)
    ax.set_ylim([0, max(tot_val) * 1.15 if not max(tot_val) == 0 else 1])
    ax.axvline(1.5, 0, 1, color='lightgray', linestyle='--', linewidth=0.5)

    aq = ax.bar(r, aq_val, width, color='#AED6F1', edgecolor='grey', label='외관')
    sq = ax.bar(r, sq_val, width, color='#F9E79F', edgecolor='grey', bottom=aq_val, label='기능/치수')

    for n, p in enumerate(aq):
        height = f"{round(aq_val[n], 1):,}"
        label = height if height != '0' else ""
        va = 'top' if p.get_height() > 28 else 'bottom'
        ax.annotate(f'{label}', xy=(p.get_x() + p.get_width() / 2, round(aq_val[n], 1)),
                    xytext=(0, -1), textcoords="offset points", ha='center', va=va, fontsize=fontsize)

    # print(tot_val)
    for n, p in enumerate(aq):
        height_tot = f"{round(tot_val[n], 1):,}"
        label_tot = height_tot if height_tot != '0' else ""
        ax.annotate(f'{label_tot}', xy=(p.get_x() + p.get_width() / 2, round(tot_val[n], 1)),
                    xytext=(0, 3), textcoords="offset points", ha='center', va='bottom', fontsize=fontsize)

    handles, labels = plt.gca().get_legend_handles_labels()
    order = [0, 1]
    ax.legend([handles[idx] for idx in order], [labels[idx] for idx in order], fontsize=fontsize, frameon=True,
              loc=1, ncol=2)

    img_name = os.path.abspath(os.path.join("spawn", "plots", f"report_plot5.png"))

    # plt.show()
    return img_name, plt


if __name__ == "__main__":
    import os
    os.chdir(os.pardir)
    os.chdir(os.pardir)
    os.chdir(os.pardir)

    from modules.ppm.ovs_ppm import MonthlyOVSppmCalculation
    from modules.ppm.dms_ppm import MonthlyDMSppmCalculation

    yyyymm = 202110

    ppm_ovs = MonthlyOVSppmCalculation(yyyymm)
    df_ovs = ppm_ovs.get_ppm()

    df_spq_per_plant = spq_per_plant(df_ovs)
    img_name, plt = draw_plot5(df_spq_per_plant)

    df_spq_per_month = spq_per_month(yyyymm, df_ovs)
    img_name2, plt2 = draw_plot6(yyyymm, df_spq_per_month)

    df_apq_per_region = apq_per_region(df_ovs)
    draw_plot7(df_apq_per_region)