from datetime import datetime
from utils_.functions import get_weekly_days_from_today
from utils_.config import QUARTERSV2, APPEARANCE_TYPE, VENDOR_REG_TO_SUBSTITUTE, PART_REG_TO_SUBSTITUTE, NON_APP, CUSTOMERS, FULL_TYPE
from modules.weekly.weekly_config import W_CUSTOMERS
import re
import pandas as pd
import numpy as np
from tqdm import tqdm


class OvsProcessedData:
    def __init__(self, raw_ovs_df, raw_dqy_df, pre_noti=True, monday_to_saturday=False, regular_month=False, yyyymm=None, weekly_customer=True,
                 weekly_type=True, n_weeks=None, n_ranks=3):

        self.raw_ovs_df = raw_ovs_df
        self.raw_dqy_df = raw_dqy_df
        self.monday_to_saturday = monday_to_saturday
        self.regular_month = regular_month
        self.yyyymm = yyyymm if regular_month else int(datetime.today().strftime("%Y%m"))
        self.month_range = [i for i in range(int(str(self.yyyymm)[:4] + '01'), int(str(self.yyyymm)[:4] + '13'), 1) if i <= self.yyyymm]
        self.weekly_customer = weekly_customer
        self.weekly_type = weekly_type
        self.last_year = int(str(self.yyyymm)[:4]) - 1
        self.last_year_month = int(str(self.last_year)+"01")
        self.n_weeks = n_weeks
        self.dates = self.get_four_weeks_date()
        self.weekly_customer = weekly_customer
        self.weekly_type = weekly_type
        self.search_on = "등록일자" if pre_noti else "해외결재일"
        self.n_ranks = n_ranks

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

    def get_stacked_qty(self):
        raw_ovs_df = self.raw_ovs_df.copy()
        raw_ovs_df[self.search_on] = [int(i.split(' ')[0].replace("-", "")) for i in raw_ovs_df[self.search_on]]
        dates = [(int(i.strftime("%Y%m%d")), int(j.strftime("%Y%m%d"))) for i, j in self.dates]
        df_stacked = None

        for week_range in dates:
            ovs_weekly = raw_ovs_df[(week_range[0] <= raw_ovs_df[self.search_on]) & (raw_ovs_df[self.search_on] <= week_range[1])]
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

    def get_ovs_pivot(self):
        df_pivoted = self.raw_ovs_df.pivot_table(values='조치수량', index=['고객사', 'timestamp'], columns=['불량구분'], aggfunc=np.sum)
        df_pivoted.fillna(0, inplace=True)

        df_pivoted['외관_불량수량'] = df_pivoted['외관'] + df_pivoted['이종']
        df_pivoted['기능_불량수량'] = df_pivoted['기능'] + df_pivoted['포장']
        df_pivoted['치수_불량수량'] = df_pivoted['치수']
        df_pivoted['기능_치수_불량수량'] = df_pivoted['기능_불량수량'] + df_pivoted['치수_불량수량']
        df_pivoted['종합_불량수량'] = df_pivoted['외관_불량수량'] + df_pivoted['기능_치수_불량수량']
        col_to_keep = ['외관_불량수량', '기능_불량수량', '치수_불량수량', '기능_치수_불량수량', '종합_불량수량']
        return df_pivoted[col_to_keep]

    def get_dqy_pivot(self):
        df_pivoted = self.raw_dqy_df.pivot_table(values='불출수량', index=['대표고객', 'timestamp'], aggfunc=np.sum)
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

    def get_barplot_data(self):
        df_qty = self.get_ppm()
        df_qty_this_year = df_qty[df_qty.index.map(lambda x: x[1] in self.month_range)]
        df_qty_last_year_quarters = df_qty[df_qty.index.map(lambda x: str(x[1]).endswith("평균") and str(x[1]).startswith(f"'{str(self.last_year)[2:4]}"))]
        df_weekly, weekly_dates = self.get_weekly_qty()
        return df_qty_this_year, df_qty_last_year_quarters, df_weekly, weekly_dates

    def customer_ranking_by_week(self, df_weekly, type_="외관_불량수량"):

        dates = [self.dates[-1]]

        weekly_sub_index = [f'{i.strftime("%m%d")}-{j.strftime("%m%d")}' for i, j in dates]

        customer_dict = {}
        customer_list = W_CUSTOMERS if self.weekly_customer else CUSTOMERS
        for customer in customer_list:
            qty = 0
            for n, week in enumerate(weekly_sub_index):
                try:
                    if type_ == "종합_불량수량":
                        qty += df_weekly.loc[(customer, week)]["외관_불량수량"] * (n + 1) / 10
                        qty += (df_weekly.loc[(customer, week)]["기능_치수_불량수량"] * (n + 1) / 10) * 0.25
                    else:
                        qty += df_weekly.loc[(customer, week)][type_] * (n + 1) / 10
                except KeyError:
                    qty += 0
                customer_dict[customer] = qty
        return sorted(customer_list, key=lambda x: customer_dict[x], reverse=True)

    def get_donutplot_data(self):

        types = FULL_TYPE

        ovs_df_year = self.raw_ovs_df[self.raw_ovs_df['timestamp'].isin(self.month_range)]
        ovs_df_month = ovs_df_year[(ovs_df_year['timestamp'] == self.yyyymm)]

        ovs_df_year_type = ovs_df_year.pivot_table(values=['조치수량'], index=['고객사', '불량구분'], aggfunc=np.sum)
        ovs_df_year_type = self.refine_donutplot_data(ovs_df_year_type, types)

        ovs_df_month_type = ovs_df_month.pivot_table(values=['조치수량'], index=['고객사', '불량구분'], aggfunc=np.sum)
        ovs_df_month_type = self.refine_donutplot_data(ovs_df_month_type, types)

        return ovs_df_year_type, ovs_df_month_type, types

    def refine_donutplot_data(self, ovs_df_type, types):
        if len(ovs_df_type) == 0:
            ovs_df_type = pd.DataFrame([[0, 0, 0]], columns=['고객사', '불량구분', '조치수량'], )
            ovs_df_type.set_index(['고객사', '불량구분'], inplace=True)

        customer_list = W_CUSTOMERS if self.weekly_customer else CUSTOMERS
        for customer in customer_list:
            for type_ in types:
                try:
                    assert len(ovs_df_type.loc[(customer, type_)]) > 0
                except KeyError:
                    ovs_df_type.loc[(customer, type_), '조치수량'] = [0]

        return ovs_df_type

    def delayed_dates_col(self):
        col_to_keep = ['고객사', '등록일자', '해외결재일', '국내접수일', '국내결재일', '대책등록일']
        ovs_df_dates = self.raw_ovs_df[self.raw_ovs_df['timestamp'].isin(self.month_range)][col_to_keep]
        ovs_df_dates.set_index('고객사', inplace=True)
        for i in col_to_keep[1:]:
            ovs_df_dates[i] = [datetime.strptime(str(j).split(" ")[0], "%Y-%m-%d") if j is not None else datetime.today() for j in
                               ovs_df_dates[i]]
        return ovs_df_dates

    def get_delayed_days(self):
        ovs_df_dates = self.delayed_dates_col()
        ovs_df_dates['해외결재소요일'] = (ovs_df_dates['해외결재일'] - ovs_df_dates['등록일자']).dt.days
        ovs_df_dates['국내접수소요일'] = (ovs_df_dates['국내접수일'] - ovs_df_dates['해외결재일']).dt.days
        return ovs_df_dates[['해외결재소요일', '국내접수소요일']]

    def weekly_overseas_rank(self, weeks=1):

        dates = self.dates[-weeks:]

        week_range = [(int(i.strftime("%Y%m%d")), int(j.strftime("%Y%m%d"))) for i, j in dates]

        ovs_df = self.raw_ovs_df.copy()
        ovs_df[self.search_on] = [int(i.split(' ')[0].replace("-", "")) for i in ovs_df[self.search_on]]
        if self.regular_month:
            ovs_df = ovs_df[ovs_df['timestamp'] == self.yyyymm]
        else:
            ovs_df = ovs_df[(week_range[0][0] <= ovs_df[self.search_on]) & (ovs_df[self.search_on] <= week_range[-1][1])]
        ovs_df['품명'] = [re.sub(PART_REG_TO_SUBSTITUTE, "", i) for i in ovs_df['품명']]
        ovs_df['품명'] = [i.replace("- ", "-").replace("  ", " ").replace("\n", "").rstrip(" ").rstrip(",").rstrip("-")
                        for i in ovs_df['품명']]
        ovs_df['부품협력사'] = [re.sub(VENDOR_REG_TO_SUBSTITUTE, "", i.upper()) for i in ovs_df['부품협력사']]

        df_dict = {}

        customer_list = W_CUSTOMERS if self.weekly_customer else CUSTOMERS
        for customer in tqdm(customer_list):
            for type_ in ['외관불량', '전체불량']:
                filter_type = APPEARANCE_TYPE if type_ == "외관불량" else NON_APP

                ovs_df_cust = ovs_df[(ovs_df['고객사'] == customer) & (ovs_df['불량구분'].isin(filter_type))]

                df = ovs_df_cust.pivot_table(values=['조치수량'], index=['불량구분', '부품협력사', '품명'], aggfunc=sum)

                if len(df) == 0:
                    df_dict[f"{customer}_{type_}"] = None
                else:
                    df.sort_values(by=['조치수량'], ascending=False, inplace=True)
                    df['점유율(%)'] = [i for i in round(df['조치수량'] / df['조치수량'].sum() * 100, 1)]
                    df = df.head(self.n_ranks)
                    df.reset_index(inplace=True, drop=False)
                    df_dict[f"{customer}_{type_}"] = df

        return df_dict, week_range[0][0], week_range[-1][1]


if __name__ == "__main__":
    import os
    from modules.weekly.pipelines.raw_data import get_raw_ovs, get_raw_dqy

    os.chdir(os.pardir)
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

    qty_data_obj = OvsProcessedData(raw_ovs_df, raw_dqy_df, pre_noti=pre_noti, monday_to_saturday=monday_to_saturday, regular_month=regular_month, yyyymm=yyyymm,
                                    weekly_customer=weekly_customer, weekly_type=weekly_type, n_weeks=n_weeks, n_ranks=n_ranks)

    df_weekly, weekly_dates = qty_data_obj.get_weekly_qty()
    c_list = qty_data_obj.customer_ranking_by_week(df_weekly=df_weekly, type_="외관_불량수량")
    print(c_list, )

    # ovs_df_dates = qty_data_obj.delayed_dates_col()
    # ovs_df_dates.to_excel(r'spawn\test20.xlsx', index=True)
    # os.startfile(r'spawn\test20.xlsx')

    # df_weekly, weekly_dates = qty_data_obj.get_weekly_qty()
    # df_qty = qty_data_obj.get_ppm()
    # customer_list = qty_data_obj.customer_ranking_by_week(df_weekly, weekly_dates, type_="외관_불량수량")
    #
    # print(df_weekly)
    # print(weekly_dates)
    # print(df_qty)
    # print(customer_list)

