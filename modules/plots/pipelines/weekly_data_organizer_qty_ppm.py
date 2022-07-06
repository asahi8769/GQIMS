from datetime import datetime
import pandas as pd
import numpy as np
from modules.db.db_overseas import OverseasDataBase
from modules.db.db_dqy import DisbursementInfo
from utils_.config import CUSTOMERS, OVS_INVALID_TYPE, QUARTERSV2
from modules.weekly.weekly_config import W_CUSTOMERS, W_OVS_INVALID_TYPE


class QtyPpmDataOrganizer:

    def __init__(self, regular_month, yyyymm, weekly_customer=False, weekly_type=False):
        self.yyyymm = yyyymm if regular_month else int(datetime.today().strftime("%Y%m"))
        self.last_year = int(str(self.yyyymm)[:4]) - 1
        self.last_year_month = int(str(self.last_year)+"01")
        self.month_range = [i for i in range(int(str(self.yyyymm)[:4] + '01'), int(str(self.yyyymm)[:4] + '13'), 1) if i <= self.yyyymm]
        self.weekly_customer = weekly_customer
        self.weekly_type = weekly_type

    def get_ovs_raw(self):
        ovs_obj = OverseasDataBase()
        ovs_df = ovs_obj.search(self.last_year_month, from_to=True, after=True)
        ovs_df['고객사'] = [str(i).replace("(CN)", "") for i in ovs_df['고객사']]
        customer_list = W_CUSTOMERS if self.weekly_customer else CUSTOMERS
        ovs_invalid_type = W_OVS_INVALID_TYPE if self.weekly_type else OVS_INVALID_TYPE
        return ovs_df[(~ovs_df["KDQR_Type"].isin(ovs_invalid_type)) & (ovs_df["고객사"].isin(customer_list))]

    def get_dqy_raw(self):
        disburse_qty = DisbursementInfo()
        dqy_df = disburse_qty.search(self.last_year_month, from_to=True, after=True)
        customer_list = W_CUSTOMERS if self.weekly_customer else CUSTOMERS

        return dqy_df[(dqy_df["대표고객"].isin(customer_list))]

    def get_ovs_pivot(self):
        df_raw = self.get_ovs_raw()
        df_pivoted = df_raw.pivot_table(values='조치수량', index=['고객사', 'timestamp'], columns=['불량구분'], aggfunc=np.sum)
        df_pivoted.fillna(0, inplace=True)

        df_pivoted['외관_불량수량'] = df_pivoted['외관'] + df_pivoted['이종']
        df_pivoted['기능_불량수량'] = df_pivoted['기능'] + df_pivoted['포장']
        df_pivoted['치수_불량수량'] = df_pivoted['치수']
        df_pivoted['기능_치수_불량수량'] = df_pivoted['기능_불량수량'] + df_pivoted['치수_불량수량']
        df_pivoted['종합_불량수량'] = df_pivoted['외관_불량수량'] + df_pivoted['기능_치수_불량수량']
        col_to_keep = ['외관_불량수량', '기능_불량수량', '치수_불량수량', '기능_치수_불량수량', '종합_불량수량']
        return df_pivoted[col_to_keep]

    def get_dqy_pivot(self):
        df_raw = self.get_dqy_raw()
        df_pivoted = df_raw.pivot_table(values='불출수량', index=['대표고객', 'timestamp'], aggfunc=np.sum)
        df_pivoted.fillna(0, inplace=True)
        return df_pivoted

    def get_concatenated_data(self):
        df_ovs_pivoted = self.get_ovs_pivot()
        df_dqy_pivoted = self.get_dqy_pivot()
        df_concat = pd.concat([df_ovs_pivoted, df_dqy_pivoted], axis=1)
        df_concat.fillna(0, inplace=True)
        return df_concat

    def get_quarters(self):
        df_concat = self.get_concatenated_data()
        df_concat['분기'] = [f"'{str(i[1])[2:4]}_{QUARTERSV2.get(str(i[1])[-2:], i[1])}" for i in df_concat.index]
        df_concat.index = df_concat.index.set_names(['고객사', 'timestamp'])
        df_concat_no_index = df_concat.reset_index()
        df_concat_no_index['timestamp'] = df_concat_no_index['분기']
        df_concat_quarters = df_concat_no_index.pivot_table(
            values=['외관_불량수량', '기능_불량수량', '치수_불량수량', '기능_치수_불량수량', '종합_불량수량', '불출수량'], index=['고객사', 'timestamp'], aggfunc=np.sum)
        df_concat_quarters_avg = df_concat_quarters/3
        df_concat_quarters_avg.index = [(i, f"{j}_평균") for i, j in df_concat_quarters_avg.index]
        df_concat.drop(['분기'], axis=1, inplace=True)
        return pd.concat([df_concat, df_concat_quarters, df_concat_quarters_avg])

    def get_timestamp_sum(self):
        df_concat_with_quarters = self.get_quarters()
        df_concat_with_quarters_no_index = df_concat_with_quarters.reset_index()
        df_concat_grouped = df_concat_with_quarters_no_index.groupby(df_concat_with_quarters_no_index['timestamp']).sum()
        df_concat_grouped['고객사'] = "종합"
        df_concat_grouped = df_concat_grouped.reset_index().set_index(['고객사', 'timestamp'])
        df_concat_with_sum = pd.concat([df_concat_with_quarters, df_concat_grouped])
        return df_concat_with_sum

    def get_ppm(self):
        df_concat_with_sum = self.get_timestamp_sum()
        df_concat_with_sum['외관_PPM'] = round(df_concat_with_sum['외관_불량수량'] / df_concat_with_sum['불출수량'] * 1000000, 1)
        df_concat_with_sum['기능_PPM'] = round(df_concat_with_sum['기능_불량수량'] / df_concat_with_sum['불출수량'] * 1000000, 1)
        df_concat_with_sum['치수_PPM'] = round(df_concat_with_sum['치수_불량수량'] / df_concat_with_sum['불출수량'] * 1000000, 1)
        df_concat_with_sum['기능_치수_PPM'] = round(df_concat_with_sum['기능_치수_불량수량'] / df_concat_with_sum['불출수량'] * 1000000, 1)
        df_concat_with_sum['종합_PPM'] = round(df_concat_with_sum['종합_불량수량'] / df_concat_with_sum['불출수량'] * 1000000, 1)
        return df_concat_with_sum


if __name__ == "__main__":
    import os
    os.chdir(os.pardir); os.chdir(os.pardir); os.chdir(os.pardir)

    data_organizer = QtyPpmDataOrganizer(regular_month=False, yyyymm=None, weekly_customer=True)

    df_raw = data_organizer.get_ovs_raw()

    df_pv = data_organizer.get_ovs_pivot()

    df_cc = data_organizer.get_concatenated_data()

    df_ = data_organizer.get_ppm()
    df_.reset_index(inplace=True)

    df_cc.to_excel(r'spawn\test11.xlsx', index=True)
    os.startfile(r'spawn\test11.xlsx')