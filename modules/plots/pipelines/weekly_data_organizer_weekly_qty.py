from datetime import datetime
import pandas as pd
import numpy as np
from modules.db.db_overseas import OverseasDataBase
from utils_.config import CUSTOMERS, OVS_INVALID_TYPE
from modules.weekly.weekly_config import W_CUSTOMERS, W_OVS_INVALID_TYPE
from utils_.functions import get_weekly_days_from_today


class WeeklyQtyDataOrganizer:

    def __init__(self, monday_to_saturday=False, regular_month=False, yyyymm=None, weekly_customer=False, weekly_type=False, n_weeks=None):
        self.regular_month = regular_month
        self.monday_to_saturday = monday_to_saturday
        self.yyyymm = yyyymm if self.regular_month else int(datetime.today().strftime("%Y%m"))
        self.last_year = int(str(self.yyyymm)[:4]) - 1
        self.last_year_month = int(str(self.last_year)+"01")
        self.n_weeks = n_weeks
        self.dates = self.get_four_weeks_date()
        self.weekly_customer = weekly_customer
        self.weekly_type = weekly_type

    def get_four_weeks_date(self):
        four_weeks_dates = []
        if self.n_weeks is not None:
            number_of_weeks = self.n_weeks
        else:
            number_of_weeks = 5 if self.regular_month else 4
        for i in reversed(range(1, number_of_weeks+1)):
            four_weeks_dates.append(get_weekly_days_from_today(
                -i, monday_to_saturday=self.monday_to_saturday, regular_month=self.regular_month, yyyymm=self.yyyymm))
        return four_weeks_dates

    def get_ovs_raw(self):
        ovs_obj = OverseasDataBase()
        ovs_df = ovs_obj.search(self.last_year_month, from_to=True, after=True)
        ovs_df['고객사'] = [str(i).replace("(CN)", "") for i in ovs_df['고객사']]
        customer_list = W_CUSTOMERS if self.weekly_customer else CUSTOMERS
        ovs_invalid_type = W_OVS_INVALID_TYPE if self.weekly_type else OVS_INVALID_TYPE
        return ovs_df[(~ovs_df["KDQR_Type"].isin(ovs_invalid_type)) & (ovs_df["고객사"].isin(customer_list))]

    def get_stacked_qty(self):
        ovs_df = self.get_ovs_raw()
        ovs_df['해외결재일'] = [int(i.split(' ')[0].replace("-", "")) for i in ovs_df['해외결재일']]
        dates = [(int(i.strftime("%Y%m%d")), int(j.strftime("%Y%m%d"))) for i, j in self.dates]
        df_stacked = None

        for week_range in dates:
            ovs_weekly = ovs_df[(week_range[0] <= ovs_df['해외결재일']) & (ovs_df['해외결재일'] <= week_range[1])]
            ovs_weekly['주차'] = f"{str(week_range[0])[4:]}-{str(week_range[1])[4:]}"
            ovs_weekly_pivoted = ovs_weekly.pivot_table(values='조치수량', index=['고객사', '주차'], columns=['불량구분'], aggfunc=np.sum)
            ovs_weekly_pivoted.fillna(0, inplace=True)

            if df_stacked is None:
                df_stacked = ovs_weekly_pivoted
            else:
                df_stacked = pd.concat([df_stacked, ovs_weekly_pivoted])

        df_stacked.fillna(0, inplace=True)

        df_stacked["외관_불량수량"] = df_stacked["외관"] + df_stacked["이종"]
        df_stacked["기능_불량수량"] = df_stacked["기능"] + df_stacked["포장"]
        df_stacked["치수_불량수량"] = df_stacked["치수"]
        df_stacked["기능_치수_불량수량"] = df_stacked["기능_불량수량"] + df_stacked["치수_불량수량"]
        df_stacked["종합_불량수량"] = df_stacked["외관_불량수량"] + df_stacked["기능_치수_불량수량"]
        col_to_keep = ["외관_불량수량", "기능_불량수량", "치수_불량수량", "기능_치수_불량수량", "종합_불량수량"]

        return df_stacked[col_to_keep]

    def get_weekly_qty(self):
        df_stacked = self.get_stacked_qty()
        df_stacked_no_index = df_stacked.reset_index()
        df_stacked_grouped = df_stacked_no_index.groupby(df_stacked_no_index['주차']).sum()
        df_stacked_grouped['고객사'] = "종합"
        df_stacked_grouped = df_stacked_grouped.reset_index().set_index(['고객사', '주차'])
        df_stacked = pd.concat([df_stacked, df_stacked_grouped])
        df_stacked.fillna(0, inplace=True)
        return df_stacked, self.dates


if __name__ == "__main__":
    import os
    os.chdir(os.pardir)
    os.chdir(os.pardir)
    os.chdir(os.pardir)

    data_organizer = WeeklyQtyDataOrganizer(monday_to_saturday=False, regular_month=False, yyyymm=202111, weekly_type=True)
    print(data_organizer.dates)

    # df_, dates = data_organizer.get_weekly_qty()
    # df_.reset_index(inplace=True)
    # df_ = data_organizer.get_ovs_raw()
    #
    # df_.to_excel(r'spawn\test12.xlsx', index=False)
    # os.startfile(r'spawn\test12.xlsx')

    # df_2 = data_organizer.get_stacked_qty()
    # df_2.to_excel(r'spawn\test14.xlsx', index=True)
    # os.startfile(r'spawn\test14.xlsx')