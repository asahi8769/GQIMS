from datetime import datetime
import pandas as pd
import numpy as np
from modules.db.db_overseas import OverseasDataBase
from modules.db.db_dqy import DisbursementInfo
from utils_.config import CUSTOMERS, OVS_INVALID_TYPE
from modules.weekly.weekly_config import W_CUSTOMERS, W_OVS_INVALID_TYPE


class OvsDelayedDataOrganizer:

    def __init__(self, from_to=False, yyyymm=None, weekly_customer=False, weekly_type=False):

        self.from_to = from_to
        self.yyyymm = yyyymm
        self.weekly_customer = weekly_customer
        self.weekly_type = weekly_type

    def get_ovs_raw(self):
        ovs_obj = OverseasDataBase()
        ovs_df = ovs_obj.search(self.yyyymm, from_to=self.from_to)
        ovs_df['고객사'] = [str(i).replace("(CN)", "") for i in ovs_df['고객사']]
        customer_list = W_CUSTOMERS if self.weekly_customer else CUSTOMERS
        ovs_invalid_type = W_OVS_INVALID_TYPE if self.weekly_type else OVS_INVALID_TYPE
        return ovs_df[(~ovs_df["KDQR_Type"].isin(ovs_invalid_type)) & (ovs_df["고객사"].isin(customer_list))]

    def get_dates_col(self):
        ovs_df = self.get_ovs_raw()
        col_to_keep = ['고객사', '등록일자', '해외결재일', '국내접수일', '국내결재일', '대책등록일']
        ovs_df_dates = ovs_df[col_to_keep]
        ovs_df_dates.set_index('고객사', inplace=True)
        for i in col_to_keep[1:]:
            ovs_df_dates[i] = [datetime.strptime(str(j).split(" ")[0], "%Y-%m-%d") if j is not None else None for j in
                               ovs_df_dates[i]]
        return ovs_df_dates.fillna(datetime.today())

    def get_delayed_days(self):
        ovs_df_dates = self.get_dates_col()
        ovs_df_dates['해외결재소요일'] = (ovs_df_dates['해외결재일'] - ovs_df_dates['등록일자']).dt.days
        ovs_df_dates['국내접수소요일'] = (ovs_df_dates['국내접수일'] - ovs_df_dates['해외결재일']).dt.days
        return ovs_df_dates[['해외결재소요일', '국내접수소요일']]


if __name__ == "__main__":
    import os
    os.chdir(os.pardir)
    os.chdir(os.pardir)
    os.chdir(os.pardir)

    yyyymm = int(datetime.today().strftime("%Y%m"))

    data_organizer = OvsDelayedDataOrganizer(from_to=False, yyyymm=yyyymm)
    df = data_organizer.get_delayed_days()

    data_organizer = OvsDelayedDataOrganizer(from_to=True, yyyymm=202112, weekly_customer=True)
    df_acc = data_organizer.get_delayed_days()

    # df_dwi = df_acc[df_acc['고객사']=="DWI"]
    df_acc.to_excel(r'spawn\test20.xlsx', index=True)
    os.startfile(r'spawn\test20.xlsx')