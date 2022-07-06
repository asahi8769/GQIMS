import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
from utils_.functions import make_dir
from datetime import datetime
import numpy as np
from utils_.config import CUSTOMERS
from modules.weekly.weekly_config import W_CUSTOMERS
import scipy.stats as stats
from tqdm import tqdm
import os
import utils_.matplotlib_config as matplotlib_config


class CustomerBarplots:

    def __init__(self, df_qty_this_year, df_qty_last_year_quarters, df_weekly, weekly_dates,
                 regular_month=False, yyyymm=None, weekly_customer=True):
        self.df_qty_this_year = df_qty_this_year
        self.df_qty_last_year_quarters = df_qty_last_year_quarters
        self.df_weekly = df_weekly
        self.weekly_dates = weekly_dates
        self.yyyymm = yyyymm if regular_month else int(datetime.today().strftime("%Y%m"))
        self.month_range = [i for i in range(int(str(self.yyyymm)[:4] + '01'), int(str(self.yyyymm)[:4] + '13'), 1)]
        self.last_year = int(str(self.yyyymm)[:4]) - 1
        self.fontsize = 10
        self.facecolor = "white"
        self.width = 0.50
        self.weekly_customer = weekly_customer

    def chart_spine_setting(self, ax, x, r, xticklabel=False, monthly=False):
        ax.set_facecolor(self.facecolor)
        ax.set(yticklabels=[])
        ax.tick_params(left=False, bottom=False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_visible(False)
        ax.spines["top"].set_visible(False)
        ax.set_xticks(r)
        if xticklabel:
            ax.set_xticklabels(x, fontsize=self.fontsize)
            ax.tick_params(axis='x', rotation=45)
        else:
            ax.set(xticklabels=[])

        if monthly:
            ax.axvline(3.5, 0, 1, color='lightgray', linestyle='--', linewidth=0.5)
            ax.axvline(2.5 + int(str(self.yyyymm)[-2:]), 0, 1, color='lightgray', linestyle='--', linewidth=0.5)
            ax.axvline(3.5 + int(str(self.yyyymm)[-2:]), 0, 1, color='lightgray', linestyle='--', linewidth=0.5)

    def get_weekly_bar(self, weekly_sub_index, df_weekly, customer, type_="외관_불량수량"):
        weekly_bar = []
        for week in weekly_sub_index:
            try:
                weekly_bar.append(df_weekly.loc[(customer, week)][type_])
            except KeyError:
                weekly_bar.append(0)
        return [round(i) for i in weekly_bar]

    def get_monthly_bar(self, df_qty_this_year, df_qty_last_year_quarters, customer, type_="외관_불량수량"):
        df_qty_this_year.fillna(0, inplace=True)
        df_qty_last_year_quarters.fillna(0, inplace=True)

        this_year_qty = []
        for month in self.month_range:
            try:
                this_year_qty.append(df_qty_this_year.loc[(customer, month)][type_])
            except KeyError:
                this_year_qty.append(0)

        last_year_qty = []
        for quarter in [f"'{str(self.last_year)[-2:]}_{i}/4_평균" for i in range(1, 5)]:
            try:
                last_year_qty.append(df_qty_last_year_quarters.loc[(customer, quarter)][type_])
            except KeyError:
                last_year_qty.append(0)
        return [round(i) for i in last_year_qty + this_year_qty]

    def get_ax_1(self, ax_1, x, r, monthly_app_bar):
        bar_z = stats.zscore([i for n, i in enumerate(monthly_app_bar) if x.index(self.yyyymm) >= n])
        color = ['#AED6F1' if i == 0 or bar_z[n] < 1.5 else '#F1948A' for n, i in enumerate(monthly_app_bar)]
        monthly_app_bar_chart = ax_1.bar(r, monthly_app_bar, self.width, label='종합', color=color, edgecolor='grey')
        ax_1.set_ylim([0, max(monthly_app_bar) * 1.05 if not max(monthly_app_bar) == 0 else 1])
        ax_1.set_ylabel("외관불량수량")
        ax_1.set_title(f"월간불량현황", loc="center", pad=15)

        for p in monthly_app_bar_chart:
            height = f"{int(p.get_height()):,}"
            label = height if height != '0' else ""
            ax_1.annotate(f'{label}', xy=(p.get_x() + p.get_width() / 2, p.get_height()),
                          xytext=(0, 3), textcoords="offset points", ha='center', va='bottom', fontsize=self.fontsize)

    def get_ax_2(self, ax_2, r, monthly_app_bar, monthly_fnd_bar, monthly_total_bar):
        monthly_app = ax_2.bar(r, monthly_app_bar, self.width, label='외관', color="#AED6F1", edgecolor='grey')
        monthly_fnd = ax_2.bar(r, monthly_fnd_bar, self.width, label='기타', color="#F9E79F", edgecolor='grey',
                               bottom=monthly_app_bar)
        ax_2.set_ylim([0, max(monthly_total_bar) * 1.05 if not max(monthly_total_bar) == 0 else 1])
        ax_2.legend(handles=[monthly_fnd, monthly_app], prop={'size': 7}, loc='best')
        ax_2.set_ylabel("공급망불량수량")

        for n, p in enumerate(monthly_fnd):
            height = f"{int(round(monthly_total_bar[n], 0)):,}"
            label = height if height != '0' else ""
            ax_2.annotate(f'{label}', xy=(p.get_x() + p.get_width() / 2, round(monthly_total_bar[n], 0)),
                          xytext=(0, 3), textcoords="offset points", ha='center', va='bottom', fontsize=self.fontsize)

    def get_ax_3(self, ax_3, r_weekly, weekly_app_bar):
        weekly = ax_3.bar(r_weekly, weekly_app_bar, self.width, label='종합', color='#AED6F1', edgecolor='grey')
        ax_3.set_title(f"주간불량현황", loc="center", pad=15)
        ax_3.set_ylim([0, max(weekly_app_bar) * 1.05 if not max(weekly_app_bar) == 0 else 1])

        for p in weekly:
            height = f"{int(p.get_height()):,}"
            label = height if height != '0' else ""
            ax_3.annotate(f'{label}', xy=(p.get_x() + p.get_width() / 2, p.get_height()),
                          xytext=(0, 3), textcoords="offset points", ha='center', va='bottom', fontsize=self.fontsize)

    def get_ax_4(self, ax_4, r_weekly, weekly_app_bar, weekly_fnd_bar, weekly_total_bar):
        weekly_app = ax_4.bar(r_weekly, weekly_app_bar, self.width, color='#AED6F1', edgecolor='grey')
        weekly_fnd = ax_4.bar(r_weekly, weekly_fnd_bar, self.width, color='#F9E79F', edgecolor='grey', bottom=weekly_app_bar)
        ax_4.set_ylim([0, max(weekly_total_bar) * 1.05 if not max(weekly_total_bar) == 0 else 1])

        for n, p in enumerate(weekly_fnd):
            height = f"{int(round(weekly_total_bar[n], 0)):,}"
            label = height if height != '0' else ""
            ax_4.annotate(f'{label}', xy=(p.get_x() + p.get_width() / 2, round(weekly_total_bar[n], 0)),
                          xytext=(0, 3), textcoords="offset points", ha='center', va='bottom', fontsize=self.fontsize)

    def customer_barplot(self, customer):
        frame_width = 12
        division = 9

        fig = plt.figure(constrained_layout=False, figsize=(10, 6), facecolor=self.facecolor)
        gs = GridSpec(2, 13, figure=fig)

        x = [f"'{str(self.last_year)[2:]}.{i + 1}/4월평균" for i in range(4)] + self.month_range
        x_weekly = [f"{str(i[0])[5:10].replace('-', '/')} ~ {str(i[1])[5:10].replace('-', '/')}" for i in self.weekly_dates]
        weekly_sub_index = [f'{i.strftime("%m%d")}-{j.strftime("%m%d").split(" ")[0]}' for i, j in self.weekly_dates]

        r = np.arange(len(x))
        r_weekly = np.arange(len(x_weekly))

        ax_1 = fig.add_subplot(gs[0:1, 0:division])
        ax_2 = fig.add_subplot(gs[1:2, 0:division])
        ax_3 = fig.add_subplot(gs[0:1, division:frame_width])
        ax_4 = fig.add_subplot(gs[1:2, division:frame_width])

        self.chart_spine_setting(ax_1, x, r, xticklabel=False, monthly=True)
        self.chart_spine_setting(ax_2, x, r, xticklabel=True, monthly=True)
        self.chart_spine_setting(ax_3, x_weekly, r_weekly, xticklabel=False, monthly=False)
        self.chart_spine_setting(ax_4, x_weekly, r_weekly, xticklabel=True, monthly=False)

        weekly_app_bar = self.get_weekly_bar(weekly_sub_index, self.df_weekly, customer, type_="외관_불량수량")
        weekly_fnd_bar = self.get_weekly_bar(weekly_sub_index, self.df_weekly, customer, type_="기능_치수_불량수량")
        monthly_app_bar = self.get_monthly_bar(self.df_qty_this_year, self.df_qty_last_year_quarters, customer, type_="외관_불량수량")
        monthly_fnd_bar = self.get_monthly_bar(self.df_qty_this_year, self.df_qty_last_year_quarters, customer, type_="기능_치수_불량수량")
        weekly_total_bar = [i + j for i, j in zip(weekly_app_bar, weekly_fnd_bar)]
        monthly_total_bar = [i + j for i, j in zip(monthly_app_bar, monthly_fnd_bar)]

        self.get_ax_1(ax_1, x, r, monthly_app_bar)
        self.get_ax_2(ax_2, r, monthly_app_bar, monthly_fnd_bar, monthly_total_bar)
        self.get_ax_3(ax_3, r_weekly, weekly_app_bar)
        self.get_ax_4(ax_4, r_weekly, weekly_app_bar, weekly_fnd_bar, weekly_total_bar)

        fig.tight_layout(pad=1.0)

        folder_name = os.path.join("spawn", "plots", datetime.now().strftime("%Y-%m-%d"))
        make_dir(folder_name)
        img_name = os.path.abspath(os.path.join(folder_name, f"{self.yyyymm}_{customer}_weekly.png"))

        plt.rcParams["axes.edgecolor"] = "black"
        plt.rcParams["axes.linewidth"] = 1

        plt.savefig(img_name, edgecolor=fig.get_edgecolor())

        return plt, img_name

    def get_customer_barplot_set(self):
        images = []
        customer_list = W_CUSTOMERS if self.weekly_customer else CUSTOMERS
        for customer in tqdm(customer_list):
            plt, img_name = self.customer_barplot(customer)
            images.append(img_name)
        return images


if __name__ == "__main__":
    import os
    from modules.weekly.pipelines.raw_data import get_raw_ovs, get_raw_dqy
    from modules.weekly.pipelines.processed_data import OvsProcessedData

    os.chdir(os.pardir)
    os.chdir(os.pardir)

    pre_noti = False
    monday_to_saturday = False
    regular_month = False
    yyyymm = None
    weekly_customer = True
    weekly_type = True
    n_weeks = None
    n_ranks = 3

    raw_ovs_df = get_raw_ovs(pre_noti=pre_noti, regular_month=regular_month, yyyymm=yyyymm, weekly_customer=weekly_customer,
                             weekly_type=weekly_type)

    print(raw_ovs_df)

    raw_dqy_df = get_raw_dqy(regular_month=regular_month, yyyymm=yyyymm, weekly_customer=weekly_customer)

    qty_data_obj = OvsProcessedData(raw_ovs_df, raw_dqy_df, pre_noti=pre_noti, monday_to_saturday=monday_to_saturday,
                                    regular_month=regular_month, yyyymm=yyyymm, weekly_customer=weekly_customer,
                                    weekly_type=weekly_type, n_weeks=n_weeks, n_ranks=n_ranks)

    df_qty_this_year, df_qty_last_year_quarters, df_weekly, weekly_dates = qty_data_obj.get_barplot_data()

    barplot_obj = CustomerBarplots(df_qty_this_year, df_qty_last_year_quarters, df_weekly, weekly_dates,
                                   regular_month=regular_month, yyyymm=yyyymm, weekly_customer=weekly_customer)

    barplot_obj.get_customer_barplot_set()