import os
import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
from datetime import datetime
from utils_.config import CUSTOMERS
from modules.weekly.weekly_config import W_CUSTOMERS
import numpy as np
from utils_.functions import make_dir
from tqdm import tqdm
import utils_.matplotlib_config as matplotlib_config



class CustomerDelayedPlots:

    def __init__(self, df_acc, regular_month=False, yyyymm=None, weekly_customer=True):
        self.df_acc = df_acc
        self.regular_month = regular_month
        self.yyyymm = yyyymm if self.regular_month else int(datetime.today().strftime("%Y%m"))
        self.fontsize = 12
        self.weekly_customer = weekly_customer

    def customer_delayed_plot(self, customer):
        fig = plt.figure(constrained_layout=False, figsize=(4, 3))
        gs = GridSpec(1, 1, figure=fig)
        ax = fig.add_subplot(gs[0:1, 0:1])

        width = 0.10
        margin = 0.04

        null_custs_acc = [i for i in W_CUSTOMERS if i not in self.df_acc.index.unique().tolist()]

        a_mean = self.df_acc['해외결재소요일'].mean()
        b_mean = self.df_acc['국내접수소요일'].mean()

        df_len = self.df_acc.index.tolist().count(customer)

        dummy_df_acc = self.df_acc.copy()
        for null_cust in null_custs_acc:
            dummy_df_acc.loc[null_cust] = {"해외결재소요일": 0, "국내접수소요일": 0}

        x = [customer]
        r = np.arange(len(x))

        dummy_df_acc_cust = dummy_df_acc.loc[customer]
        data_a_acc = [dummy_df_acc_cust['해외결재소요일'].mean()]
        data_b_acc = [dummy_df_acc_cust['국내접수소요일'].mean()]

        ch_3 = ax.bar(r - width / 2 - margin, data_a_acc, width, label=f'해외결재소요일: {round(data_a_acc[0], 1)}',
                      color='#AED6F1', edgecolor='grey')
        ch_4 = ax.bar(r + width / 2 + margin, data_b_acc, width, label=f'국내접수소요일: {round(data_b_acc[0], 1)}',
                      color='#F5B7B1', edgecolor='grey')
        ch_5 = ax.axhline(y=a_mean, xmin=0.0, xmax=0.45, color="#5499C7", linestyle='--',
                          label=f"평균해외결재소요일: {round(a_mean, 1)}", lw=0.5)
        ch_6 = ax.axhline(y=b_mean, xmin=0.55, xmax=1.0, color="#EC7063", linestyle='--',
                          label=f"평균국내접수소요일: {round(b_mean, 1)}", lw=0.5)
        if not max([data_a_acc[0], data_b_acc[0]]) == 0:
            ax.set_ylim([0, max([data_a_acc[0], data_b_acc[0], a_mean, b_mean]) * 1.4])

        ax.legend(handles=[ch_3, ch_4], prop={'size': 10}, loc='best')

        for p in ch_3:
            height = round(p.get_height(), 1)
            label = height if height != 0 else ""
            ax.annotate(f'{label}', xy=(p.get_x() + p.get_width() / 2, 0), xytext=(-2, 3),
                        textcoords="offset points", ha='center', va='baseline', fontsize=self.fontsize)
            ax.annotate(f'평균: {round(a_mean, 1)}', xy=(p.get_x() + p.get_width() / 2, a_mean), xytext=(-2, 3),
                        textcoords="offset points", ha='center', va='baseline', fontsize=self.fontsize)

        for p in ch_4:
            height = round(p.get_height(), 1)
            label = height if height != 0 else ""
            ax.annotate(f'{label}', xy=(p.get_x() + p.get_width() / 2, 0), xytext=(2, 3), textcoords="offset points",
                        ha='center', va='baseline', fontsize=self.fontsize)
            ax.annotate(f'평균: {round(b_mean, 1)}', xy=(p.get_x() + p.get_width() / 2, b_mean), xytext=(2, 3),
                        textcoords="offset points", ha='center', va='baseline', fontsize=self.fontsize)

        ax.set_xticks(r)
        ax.set_xticklabels([], fontsize=self.fontsize)
        # ax.set_ylabel("누적")

        ax.set(yticklabels=[])
        ax.tick_params(left=False, bottom=False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_visible(False)
        ax.spines["top"].set_visible(False)
        ax.set_title(f"'{str(self.yyyymm)[2:4]}년도 업무소요일 현황 (PPR: {df_len:,}건)", loc="center", pad=15, fontsize=self.fontsize)

        fig.tight_layout()

        folder_name = os.path.join("spawn", "plots", datetime.now().strftime("%Y-%m-%d"))
        make_dir(folder_name)
        img_name = os.path.abspath(os.path.join(folder_name, f"{self.yyyymm}_{customer}_weekly_delayed.png"))
        plt.savefig(img_name)

        return plt, img_name

    def get_customer_delayed_plot_set(self):
        images = []
        customer_list = W_CUSTOMERS if self.weekly_customer else CUSTOMERS
        for customer in tqdm(customer_list):
            plt, img_name = self.customer_delayed_plot(customer)
            images.append(img_name)
        return images


if __name__ == "__main__":
    import os
    from modules.weekly.pipelines.raw_data import get_raw_ovs, get_raw_dqy
    from modules.weekly.pipelines.processed_data import OvsProcessedData

    os.chdir(os.pardir)
    os.chdir(os.pardir)

    pre_noti = True
    monday_to_saturday = False
    regular_month = False
    yyyymm = None
    weekly_customer = True
    weekly_type = True
    n_weeks = None
    n_ranks = 3

    raw_ovs_df = get_raw_ovs(pre_noti=pre_noti, regular_month=regular_month, yyyymm=yyyymm, weekly_customer=weekly_customer,
                             weekly_type=weekly_type)

    raw_dqy_df = get_raw_dqy(regular_month=regular_month, yyyymm=yyyymm, weekly_customer=weekly_customer)

    qty_data_obj = OvsProcessedData(raw_ovs_df, raw_dqy_df, pre_noti=pre_noti, monday_to_saturday=monday_to_saturday,
                                    regular_month=regular_month, yyyymm=yyyymm, weekly_customer=weekly_customer,
                                    weekly_type=weekly_type, n_weeks=n_weeks,  n_ranks=n_ranks)

    df_acc = qty_data_obj.get_delayed_days()

    delayedplot_obj = CustomerDelayedPlots(df_acc, regular_month=regular_month, yyyymm=yyyymm, weekly_customer=weekly_customer)
    delayedplot_obj.get_customer_delayed_plot_set()