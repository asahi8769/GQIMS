import os
from datetime import datetime
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.gridspec import GridSpec
from modules.plots.cust_plot_single_barchart import customer_ranking_by_week, customer_ranking_by_month, filter_qty_data
from utils_.functions import make_dir
import utils_.matplotlib_config as matplotlib_config


def draw_plot3(type_=("외관_불량수량", "기능_치수_불량수량"), monday_to_saturday=False, regular_month=False, yyyymm=202111):

    yyyymm = yyyymm if regular_month else int(datetime.today().strftime("%Y%m"))
    month_range = [i for i in range(int(str(yyyymm)[:4] + '01'), int(str(yyyymm)[:4] + '13'), 1)]
    last_year = int(str(yyyymm)[:4]) - 1
    df_qty_this_year, df_qty_last_year_quarters, df_weekly, weekly_dates = filter_qty_data(monday_to_saturday=monday_to_saturday, regular_month=regular_month, yyyymm=yyyymm)

    if regular_month:
        customers = customer_ranking_by_month(df_qty_this_year, yyyymm, type_="외관_불량수량")
    else:
        customers = customer_ranking_by_week(df_weekly, weekly_dates, type_="외관_불량수량")

    fontsize = 9
    width = 0.40
    frame_width = 12
    division = 9 if regular_month else 10

    fig = plt.figure(constrained_layout=True, figsize=(10, 11))
    gs = GridSpec(len(customers), 13, figure=fig)

    x = [f"'{str(last_year)[2:]}.{i+1}/4월평균" for i in range(4)] + month_range
    x_weekly = [f"{str(i[0])[5:10].replace('-', '/')} ~ {str(i[1])[5:10].replace('-', '/')}" for i in weekly_dates]
    weekly_sub_index = [f'{i.strftime("%m%d")}-{j.strftime("%m%d")}' for i, j in weekly_dates]

    for n, customer in enumerate(customers):
        weekly_qty_app = []
        weekly_qty_fnd = []
        for week in weekly_sub_index:
            try:
                weekly_qty_app.append(df_weekly.loc[(customer, week)][type_[0]])
            except KeyError:
                weekly_qty_app.append(0)
            try:
                weekly_qty_fnd.append(df_weekly.loc[(customer, week)][type_[1]])
            except KeyError:
                weekly_qty_fnd.append(0)

        weekly_qty_app += [0] * (len(x_weekly) - len(weekly_qty_app))
        weekly_qty_fnd += [0] * (len(x_weekly) - len(weekly_qty_fnd))
        bar_app = weekly_qty_app
        bar_fnd = weekly_qty_fnd

        total = [i+j for i, j in zip(bar_app, bar_fnd)]

        r = np.arange(len(x_weekly))
        ax = fig.add_subplot(gs[n:n + 1, division:frame_width])

        ax.set(yticklabels=[])
        ax.tick_params(left=False, bottom=False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_visible(False)
        ax.spines["top"].set_visible(False)
        ax.set_xticks(r)
        ax.set_ylim([0, max(total)*1.05 if not max(total) == 0 else 1])

        if n + 1 == len(customers):
            ax.set_xticklabels(x_weekly, fontsize=fontsize)
            ax.tick_params(axis='x', rotation=45)
        else:
            ax.set(xticklabels=[])
        if n == 0:
            ax.set_title(f"주간 공급망 불량수량 현황", loc="center", pad=15)

        weekly_app = ax.bar(r, bar_app, width, color='#AED6F1', edgecolor='grey')
        weekly_fnd = ax.bar(r, bar_fnd, width, color='#F9E79F', edgecolor='grey', bottom=bar_app)

        for n, p in enumerate(weekly_fnd):
            height = f"{int(round(total[n],0)):,}"
            label = height if height != '0' else ""
            ax.annotate(f'{label}', xy=(p.get_x() + p.get_width() / 2, round(total[n],0)),
                        xytext=(0, 3), textcoords="offset points", ha='center', va='bottom', fontsize=fontsize)

    for n, customer in enumerate(customers):
        try:
            this_year_app_qty = df_qty_this_year.loc[customer][type_[0]].tolist()
        except KeyError:
            this_year_app_qty = [0 for _ in range(len(month_range))]
            this_year_fnd_qty = [0 for _ in range(len(month_range))]
        else:
            this_year_fnd_qty = df_qty_this_year.loc[customer][type_[1]].tolist()

        try:
            last_year_app_qty = df_qty_last_year_quarters.loc[customer][type_[0]].tolist()
        except KeyError:
            last_year_app_qty = [0 for _ in range(int(len(month_range)/3))]
            last_year_fnd_qty = [0 for _ in range(int(len(month_range)/3))]
        else:
            last_year_fnd_qty = df_qty_last_year_quarters.loc[customer][type_[1]].tolist()

        this_year_app_qty += [0]*(len(month_range)-len(this_year_app_qty))
        this_year_fnd_qty += [0]*(len(month_range)-len(this_year_fnd_qty))
        last_year_app_qty += [0]*(int(len(month_range)/3)-len(last_year_app_qty))
        last_year_fnd_qty += [0]*(int(len(month_range)/3)-len(last_year_fnd_qty))

        total_this = [i + j for i, j in zip(this_year_app_qty, this_year_fnd_qty)]
        total_last = [i + j for i, j in zip(last_year_app_qty, last_year_fnd_qty)]
        total = total_last + total_this

        bar_app = last_year_app_qty + this_year_app_qty
        bar_fnd = last_year_fnd_qty + this_year_fnd_qty

        r = np.arange(len(x))
        ax = fig.add_subplot(gs[n:n+1, 0:division])

        ax.set(yticklabels=[])
        ax.tick_params(left=False, bottom=False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_visible(False)
        ax.spines["top"].set_visible(False)
        ax.set_xticks(r)
        ax.set_ylabel(customer)

        if n+1 == len(customers):
            ax.set_xticklabels(x, fontsize=fontsize)
            ax.tick_params(axis='x', rotation=45)
        else:
            ax.set(xticklabels=[])
        if n == 0:
            ax.set_title(f"고객사별 공급망 불량 현황", loc="center", pad=15)

        ax.axvline(3.5, 0, 1, color='lightgray', linestyle='--', linewidth=0.5)
        ax.axvline(2.5 + int(str(yyyymm)[-2:]), 0, 1, color='lightgray', linestyle='--', linewidth=0.5)
        ax.axvline(3.5 + int(str(yyyymm)[-2:]), 0, 1, color='lightgray', linestyle='--', linewidth=0.5)

        monthly_app = ax.bar(r, bar_app, width, label='외관', color="#AED6F1", edgecolor='grey')
        monthly_fnd = ax.bar(r, bar_fnd, width, label='기능/치수', color="#F9E79F", edgecolor='grey', bottom=bar_app)
        if n == 0:
            handles, labels = plt.gca().get_legend_handles_labels()
            order = [1, 0]

            ax.legend([handles[idx] for idx in order],[labels[idx] for idx in order], fontsize=7, frameon=False, bbox_to_anchor=(1, 1.6, 0, 0.2))

        for n, p in enumerate(monthly_fnd):
            height = f"{int(round(total[n],0)):,}"
            label = height if height != '0' else ""
            ax.annotate(f'{label}', xy=(p.get_x() + p.get_width() / 2, round(total[n],0)),
                        xytext=(0, 3), textcoords="offset points", ha='center', va='bottom', fontsize=fontsize)

    fig.tight_layout(pad=5.0)

    folder_name = os.path.join("spawn", "plots", datetime.now().strftime("%Y-%m-%d"))
    make_dir(folder_name)

    img_name = os.path.abspath(os.path.join(folder_name, f"{yyyymm}_stacked_종합불량수량_plot3.png"))
    plt.savefig(img_name)

    return plt, img_name


if __name__ == "__main__":


    os.chdir(os.pardir)
    os.chdir(os.pardir)

    plt, img_name = draw_plot3(type_=("외관_불량수량", "기능_치수_불량수량"), monday_to_saturday=False, regular_month=False, yyyymm=202112)
    plt.show()

