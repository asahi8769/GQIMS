import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
from utils_.config import CUSTOMERS
from modules.weekly.weekly_config import W_CUSTOMERS
import numpy as np
import os
import scipy.stats as stats
from datetime import datetime
from modules.plots.pipelines.weekly_data_organizer_qty_ppm import QtyPpmDataOrganizer
from modules.plots.pipelines.weekly_data_organizer_weekly_qty import WeeklyQtyDataOrganizer
from utils_.functions import make_dir
import utils_.matplotlib_config as matplotlib_config


def filter_qty_data(monday_to_saturday=False, regular_month=False, yyyymm=None, weekly_customer=False, weekly_type=False, n_weeks=None):
    yyyymm = yyyymm if regular_month else int(datetime.today().strftime("%Y%m"))
    data_organizer_qty = QtyPpmDataOrganizer(regular_month, yyyymm, weekly_customer=weekly_customer, weekly_type=weekly_type)
    df_qty = data_organizer_qty.get_ppm()

    last_year = int(str(yyyymm)[:4]) - 1
    month_range = [i for i in range(int(str(yyyymm)[:4] + '01'), int(str(yyyymm)[:4] + '13'), 1) if i <= yyyymm]
    df_qty_this_year = df_qty[df_qty.index.map(lambda x: x[1] in month_range)]
    df_qty_last_year_quarters = df_qty[df_qty.index.map(lambda x: str(x[1]).endswith("평균") and str(x[1]).startswith(f"'{str(last_year)[2:4]}"))]

    data_organizer_weekly = WeeklyQtyDataOrganizer(monday_to_saturday=monday_to_saturday, regular_month=regular_month,
                                                   yyyymm=yyyymm, weekly_customer=weekly_customer, weekly_type=weekly_type, n_weeks=n_weeks)
    df_weekly, weekly_dates = data_organizer_weekly.get_weekly_qty()

    return df_qty_this_year, df_qty_last_year_quarters, df_weekly, weekly_dates


def customer_ranking_by_month(df_qty_this_year, yyyymm, type_="외관_불량수량", weekly_customer=False):
    month_range = [i for i in range(int(str(yyyymm)[:4] + '01'), int(str(yyyymm)[:4] + '13'), 1) if i <= yyyymm]

    customer_dict = {}
    customer_list = W_CUSTOMERS if weekly_customer else CUSTOMERS
    for customer in customer_list:
        customer_dict[customer] = {}
        for month in month_range:
            try :
                qty = df_qty_this_year.loc[(customer, month)][type_]

            except KeyError:
                qty = 0
            customer_dict[customer][month] = qty
    return sorted(customer_list, key=lambda x: customer_dict[x][yyyymm], reverse=True) + ['종합']


def customer_ranking_by_week(df_weekly, weekly_dates, type_="외관_불량수량", weekly_customer=False):
    weekly_sub_index = [f'{i.strftime("%m%d")}-{j.strftime("%m%d")}' for i, j in weekly_dates]

    customer_dict = {}
    customer_list = W_CUSTOMERS if weekly_customer else CUSTOMERS
    for customer in customer_list:
        qty = 0
        for n, week in enumerate(weekly_sub_index):
            try :
                if type_ == "종합_불량수량":
                    qty += df_weekly.loc[(customer, week)]["외관_불량수량"] * (n + 1) / 10
                    qty += (df_weekly.loc[(customer, week)]["기능_치수_불량수량"] * (n + 1) / 10) * 0.25
                else:
                    qty += df_weekly.loc[(customer, week)][type_]*(n+1)/10
            except KeyError:
                qty += 0
            customer_dict[customer] = qty
    return sorted(customer_list, key=lambda x: customer_dict[x], reverse=True) + ['종합']


def draw_plot2(type_="외관_불량수량", monday_to_saturday=False, regular_month=False, yyyymm=None):

    yyyymm = yyyymm if regular_month else int(datetime.today().strftime("%Y%m"))
    month_range = [i for i in range(int(str(yyyymm)[:4] + '01'), int(str(yyyymm)[:4] + '13'), 1)]
    last_year = int(str(yyyymm)[:4]) - 1

    df_qty_this_year, df_qty_last_year_quarters, df_weekly, weekly_dates = filter_qty_data(
        monday_to_saturday=monday_to_saturday, regular_month=regular_month, yyyymm=yyyymm)

    if regular_month:
        customers = customer_ranking_by_month(df_qty_this_year, yyyymm, type_=type_)
    else:
        customers = customer_ranking_by_week(df_weekly, weekly_dates, type_=type_)

    fontsize = 9
    width = 0.40
    frame_width = 12
    division = 9 if regular_month else 10

    fig = plt.figure(constrained_layout=True, figsize=(10, 11))
    gs = GridSpec(len(customers), 13, figure=fig)

    x = [f"'{str(last_year)[2:]}.{i+1}/4월평균" for i in range(4)] + month_range
    x_weekly = [f"{str(i[0])[5:10].replace('-', '/')} ~ {str(i[1])[5:10].replace('-', '/')}" for i in weekly_dates]
    weekly_sub_index = [f'{i.strftime("%m%d")}-{j.strftime("%m%d").split(" ")[0]}' for i, j in weekly_dates]
    title = type_.replace("_", " ")

    for n, customer in enumerate(customers):
        weekly_qty = []
        for week in weekly_sub_index:
            try:
                weekly_qty.append(df_weekly.loc[(customer, week)][type_])
            except KeyError:
                weekly_qty.append(0)

        weekly_qty += [0]*(len(x_weekly)-len(weekly_qty))
        bar = weekly_qty

        r = np.arange(len(x_weekly))
        ax = fig.add_subplot(gs[n:n + 1, division:frame_width])

        ax.set(yticklabels=[])
        ax.tick_params(left=False, bottom=False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_visible(False)
        ax.spines["top"].set_visible(False)
        ax.set_xticks(r)
        ax.set_ylim([0, max(bar)*1.05 if not max(bar) == 0 else 1])

        if n + 1 == len(customers):
            ax.set_xticklabels(x_weekly, fontsize=fontsize)
            ax.tick_params(axis='x', rotation=45)
        else:
            ax.set(xticklabels=[])
        if n == 0:
            ax.set_title(f"주간 {title} 현황", loc="center", pad=15)

        weekly = ax.bar(r, bar, width, label='종합', color='#F9E79F', edgecolor='grey')

        for p in weekly:
            height = f"{int(p.get_height()):,}"
            label = height if height != '0' else ""
            ax.annotate(f'{label}', xy=(p.get_x() + p.get_width() / 2, p.get_height()),
                        xytext=(0, 3), textcoords="offset points", ha='center', va='bottom', fontsize=fontsize)

    for n, customer in enumerate(customers):

        this_year_app_qty = []
        for month in month_range:
            try:
                this_year_app_qty.append(df_qty_this_year.loc[(customer, month)][type_])
            except KeyError:
                this_year_app_qty.append(0)

        last_year_app_qty = []
        for quarter in [f"'{str(last_year)[-2:]}_{i}/4_평균" for i in range(1, 5)]:
            try:
                last_year_app_qty.append(df_qty_last_year_quarters.loc[(customer, quarter)][type_])
            except KeyError:
                last_year_app_qty.append(0)

        bar = last_year_app_qty + this_year_app_qty

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
            ax.set_title(f"고객사별 {title} 현황", loc="center", pad=15)

        ax.axvline(3.5, 0, 1, color='lightgray', linestyle='--', linewidth=0.5)
        ax.axvline(2.5 + int(str(yyyymm)[-2:]), 0, 1, color='lightgray', linestyle='--', linewidth=0.5)
        ax.axvline(3.5 + int(str(yyyymm)[-2:]), 0, 1, color='lightgray', linestyle='--', linewidth=0.5)

        bar_z = stats.zscore([i for n, i in enumerate(bar) if x.index(yyyymm) >= n])

        color = ['#AED6F1' if i == 0 or bar_z[n] < 1.5 else '#F1948A' for n, i in enumerate(bar)]
        monthly = ax.bar(r, bar, width, label='종합', color=color, edgecolor='grey', bottom=0)

        for p in monthly:
            height = f"{int(p.get_height()):,}"
            label = height if height != '0' else ""
            ax.annotate(f'{label}', xy=(p.get_x() + p.get_width() / 2, p.get_height()),
                        xytext=(0, 3), textcoords="offset points", ha='center', va='bottom', fontsize=fontsize)

    fig.tight_layout(pad=5.0)

    folder_name = os.path.join("spawn", "plots", datetime.now().strftime("%Y-%m-%d"))

    make_dir(folder_name)
    img_name = os.path.abspath(os.path.join(folder_name, f"{yyyymm}_{type_}_plot2.png"))
    plt.savefig(img_name)

    return plt, img_name


if __name__ == "__main__":

    os.chdir(os.pardir)
    os.chdir(os.pardir)

    yyyymm = int(datetime.today().strftime("%Y%m"))

    plt, img_name = draw_plot2(type_="외관_불량수량", monday_to_saturday=False, regular_month=False, yyyymm=202112)
    plt.show()

