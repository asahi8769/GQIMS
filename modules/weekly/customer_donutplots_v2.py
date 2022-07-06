from utils_.functions import make_dir
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
import numpy as np
from tqdm import tqdm
from utils_.config import CUSTOMERS
from modules.weekly.weekly_config import W_CUSTOMERS
import utils_.matplotlib_config as matplotlib_config
import os


class CustomerDonutPlots:

    def __init__(self, ovs_df_year_type, ovs_df_month_type, types, regular_month=False, yyyymm=None, weekly_customer=True):
        self.ovs_df_year_type = ovs_df_year_type
        self.ovs_df_month_type = ovs_df_month_type
        self.types = types
        self.regular_month = regular_month
        self.yyyymm = yyyymm if self.regular_month else int(datetime.today().strftime("%Y%m"))
        self.fontsize = 12
        self.weekly_customer = weekly_customer

    def customer_donutplot(self, customer, types, from_to=False):

        fig = plt.figure(constrained_layout=False, figsize=(3, 3))
        gs = GridSpec(1, 1, figure=fig)
        ax = fig.add_subplot(gs[0:1, 0:1])
        ovs_df_type_customer = self.ovs_df_year_type.loc[(customer,)] if from_to else self.ovs_df_month_type.loc[(customer,)]
        ovs_df_type_customer.sort_values("조치수량", inplace=True, ascending=False)
        data = ovs_df_type_customer['조치수량'].tolist()
        data_sum = int(sum(data))
        types = sorted(types, key=lambda x: ovs_df_type_customer.loc[x, '조치수량'], reverse=True)

        cmap = plt.get_cmap("tab20c")
        outer_colors = cmap(np.array([4 * i % 20 for i in range(len(types))]))

        wedges, texts = ax.pie(data, radius=1, colors=outer_colors,
                               startangle=90,
                               counterclock=False,
                               wedgeprops=dict(width=0.3, edgecolor='grey', linewidth=0.7))

        total = sum(data)
        share = [0 for _ in data] if total == 0 else [i / total * 100 for i in data]
        share = [int(round(i, 0)) for i in share]

        for index, value in enumerate(texts):
            value.set_fontsize(self.fontsize)
            text = f"{types[index]}\n({share[index]}%)" if share[index] > 5 and data[index] > 0 else ""
            value.set_text(text)

        middle_text = f"'{str(self.yyyymm)[2:4]}년도\n누적불량수량\n{data_sum:,}EA" if from_to else f"{int(str(self.yyyymm)[4:6])}월\n불량수량\n{data_sum:,}EA"
        ax.text(0, 0, middle_text, ha='center', va='center', fontsize=self.fontsize)

        folder_name = os.path.join("spawn", "plots", datetime.now().strftime("%Y-%m-%d"))
        make_dir(folder_name)

        if from_to:
            img_name = os.path.abspath(os.path.join(folder_name, f"{self.yyyymm}_{customer}_weekly_donut_year.png"))
        else:
            img_name = os.path.abspath(os.path.join(folder_name, f"{self.yyyymm}_{customer}_weekly_donut_month.png"))

        plt.savefig(img_name)
        return plt, img_name

    def get_customer_donutplot_set(self):

        images_month = []
        images_year = []
        customer_list = W_CUSTOMERS if self.weekly_customer else CUSTOMERS

        for customer in tqdm(customer_list):
            _, img_name_month = self.customer_donutplot(customer, self.types, from_to=False)
            images_month.append(img_name_month)

            _, img_name_year = self.customer_donutplot(customer, self.types, from_to=True)
            images_year.append(img_name_year)

        return images_month, images_year


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
                                    weekly_type=weekly_type, n_weeks=n_weeks, n_ranks=n_ranks)

    ovs_df_year_type, ovs_df_month_type, types = qty_data_obj.get_donutplot_data()

    donutplot_obj = CustomerDonutPlots(ovs_df_year_type, ovs_df_month_type, types, regular_month=regular_month,
                                       yyyymm=yyyymm, weekly_customer=weekly_customer)
    donutplot_obj.get_customer_donutplot_set()