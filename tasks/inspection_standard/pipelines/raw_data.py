from modules.db.db_overseas import OverseasDataBase
from modules.db.db_domestic import DomesticDataBase
from modules.db.db_incoming_pn import IncomingInfoWithPN
from tasks.inspection_standard.config import *
import pandas as pd
import numpy as np
from datetime import datetime
from utils_.functions import show_elapsed_time, make_dir
import calendar


class RawDataGenerator:
    def __init__(self, pre_noti=PRE_NOTI, reference_ovs=OVS_REFERENCE, reference_dms=DMS_REFERENCE,
                 keys=KEYS):
        make_dir("tasks/inspection_standard/export")
        self.days_weight_df = pd.read_excel("tasks/inspection_standard/files/settings.xlsx", sheet_name="days_weight")
        self.days_weight_dict = dict(zip(list(zip(self.days_weight_df['이상'], self.days_weight_df['이하'])), self.days_weight_df['가중치']))
        self.type_weight_df = pd.read_excel("tasks/inspection_standard/files/settings.xlsx", sheet_name="type_weight")
        self.type_weight_df.set_index("화면", inplace=True)
        packing_codes = pd.read_excel("tasks/inspection_standard/files/codes.xlsx", sheet_name="packing")
        self.packing_dict = dict(zip(packing_codes['P_CD'], packing_codes['P_NM']))
        packing_codes = pd.read_excel("tasks/inspection_standard/files/codes.xlsx", sheet_name="line")
        self.line_dict = dict(zip(packing_codes['L_CD'], packing_codes['L_NM']))
        self.pre_noti = pre_noti
        self.reference_ovs = reference_ovs
        self.reference_dms = reference_dms
        # self.year_icm = year_icm
        self.keys = keys
        self.yyyymm = None

    def elapsed_days(self, col, date_format="%Y-%m-%d"):
        year = int(str(self.yyyymm)[:4])
        month = int(str(self.yyyymm)[4:])
        now = datetime(year, month, calendar.monthrange(year, month)[1])

        return [(now-datetime.strptime(i.split(" ")[0], date_format)).days for i in col]

    @staticmethod
    def days_weight(days, elapse_weight_dict):
        for i in elapse_weight_dict.keys():
            if i[0] <= days <= i[1]:
                return elapse_weight_dict.get(i, 0)

    def type_weight(self, type, screen, pre_shipping_inclusive=PRE_SHIPPING_INCLUSIVE):
        """

        :param type:
        :param screen: 출하, 국내, 해외
        :param pre_shipping_inclusive:
        :return:
        """
        shpping_info = "출하포함" if pre_shipping_inclusive else "출하미포함"
        refer_to = f"{type}_{shpping_info}"
        return self.type_weight_df.loc[screen][refer_to]

    def packing_supplier(self, code):
        if code in self.packing_dict.keys():
            return "포장"
        else:
            return

    @show_elapsed_time
    def get_yearly_raw_ovs(self, delta=1):
        """

        :param now: datetime(2022, 1, 31)
        :return:
        """
        ovs_obj = OverseasDataBase(pre_noti=self.pre_noti)
        ovs_one_year_df = ovs_obj.yearly_search(self.yyyymm, reference=self.reference_ovs, delta=delta)
        ovs_one_year_df = ovs_one_year_df[~ovs_one_year_df['불량구분'].isnull()]
        ovs_one_year_df['고객사'] = [str(i).replace("(CN)", "") for i in ovs_one_year_df['고객사']]
        ovs_one_year_df['경과일'] = self.elapsed_days(ovs_one_year_df[self.reference_ovs], date_format="%Y-%m-%d")
        ovs_one_year_df['불량구분'] = [ovs_one_year_df['불량구분'][n] if i not in self.packing_dict.keys() else "포장" for n, i in enumerate(ovs_one_year_df['코드.1'])]
        ovs_one_year_df['가중치_경과일'] = ovs_one_year_df['경과일'].map(lambda x: self.days_weight(x, self.days_weight_dict))
        ovs_one_year_df['가중치_경과일'].fillna(0, inplace=True)
        ovs_one_year_df['가중치_불량구분'] = ovs_one_year_df['불량구분'].map(lambda x: self.type_weight(x, "해외"))
        ovs_one_year_df['가중치_전체'] = round(ovs_one_year_df['가중치_경과일'] * ovs_one_year_df['가중치_불량구분']/100, 1)
        ovs_one_year_df['조치수량_가중치'] = round(ovs_one_year_df['조치수량'] * ovs_one_year_df['가중치_전체']/100, 1)
        ovs_one_year_df = ovs_one_year_df[
            (~ovs_one_year_df["KDQR_Type"].isin(OVS_INVALID_TYPE)) & (~ovs_one_year_df["공급법인"].isin(INVALID_PPR_SUPPLIED_FROM))]

        # ovs_one_year_df = ovs_one_year_df[(ovs_one_year_df["고객사"].isin(CUSTOMERS))]

        col_to_return = ["KDQR_Type", "대표PPR_No", "PPR_No", "공급법인", "고객사", "차종", "품번", "품명",
                         "귀책구분", "코드.1", "귀책처", "조치수량", "불량구분", self.reference_ovs, "경과일", "가중치_경과일",
                         "가중치_불량구분", "가중치_전체", "조치수량_가중치", "공급국가", "timestamp"]
        return ovs_one_year_df[col_to_return]

    @show_elapsed_time
    def get_yearly_raw_dms(self, delta=1):
        """

        :param now: datetime(2022, 1, 31)
        :return:
        """
        dms_obj = DomesticDataBase()
        dms_one_year_df = dms_obj.yearly_search(self.yyyymm, reference=self.reference_dms, delta=delta)
        dms_one_year_df = dms_one_year_df[~dms_one_year_df['불량구분'].isnull()]
        dms_one_year_df['경과일'] = self.elapsed_days(dms_one_year_df[self.reference_dms], date_format="%Y-%m-%d")
        dms_one_year_df['가중치_경과일'] = dms_one_year_df['경과일'].map(lambda x: self.days_weight(x, self.days_weight_dict))
        dms_one_year_df['가중치_경과일'].fillna(0, inplace=True)
        dms_one_year_df['가중치_불량구분'] = dms_one_year_df['불량구분'].map(lambda x: self.type_weight(x, "국내"))
        dms_one_year_df['가중치_전체'] = round(dms_one_year_df['가중치_경과일'] * dms_one_year_df['가중치_불량구분']/100, 1)
        dms_one_year_df['조치수량_가중치'] = round(dms_one_year_df['불량수량'] * dms_one_year_df['가중치_전체']/100, 1)
        dms_one_year_df = dms_one_year_df[
            (~dms_one_year_df["불량조치"].isin(DMS_INVALID_TYPE)) & (~dms_one_year_df["공급법인"].isin(INVALID_PPR_SUPPLIED_FROM))]

        # dms_one_year_df = dms_one_year_df[(dms_one_year_df["고객사"].isin(CUSTOMERS))]

        col_to_return = ["통보서No", "공급법인", "고객사", "차종", "코드", "부품협력사", "품번", "품명",
                         "불량수량", "불량구분", self.reference_dms, "불량조치", "경과일", "가중치_경과일", '가중치_불량구분',
                         '가중치_전체', '조치수량_가중치', "포장장", "timestamp"]
        return dms_one_year_df[col_to_return]

    def get_monthly_part_supply_untouched(self, delta_month=12):
        customer_codes = pd.read_excel("tasks/inspection_standard/files/codes.xlsx", sheet_name="customer")
        customer_dict = dict(zip(customer_codes['CUST_CD'], customer_codes['CUST_NM']))
        icm_obj = IncomingInfoWithPN()
        icm_2021_df = icm_obj.monthly_search(yyyymm=ICM_YEAR, delta=delta_month)
        print(icm_2021_df['timestamp'].unique())
        icm_2021_df.fillna(0, inplace=True)
        icm_2021_df['고객사'] = [customer_dict.get(i, i) for i in icm_2021_df['고객사']]
        icm_2021_df['고객사'] = [str(i).replace("(CN)", "") for i in icm_2021_df['고객사']]

        # icm_2021_df = icm_2021_df[(icm_2021_df["고객사"].isin(CUSTOMERS))]

        icm_2021_df['포장장코드'] = icm_2021_df['포장장코드'].map(lambda x: self.line_dict.get(x, x))
        icm_2021_df['입고수량'] = icm_2021_df['입고수량'].map(lambda x: int(float(x)))
        return icm_2021_df

    @show_elapsed_time
    def get_monthly_part_supply(self, delta_monthly=12):
        icm_2021_df = self.get_monthly_part_supply_untouched(delta_month=delta_monthly)
        icm_2021_df = icm_2021_df.pivot_table(values='입고수량',
                                              index=self.keys,
                                              aggfunc=np.sum)
        icm_2021_df.reset_index(inplace=True)
        icm_2021_df['품목'] = icm_2021_df["부품번호"].str.slice(0, 3)

        return icm_2021_df

    def get_one_fixed_year_part_supply_super_raw(self, delta_monthly=12):
        customer_codes = pd.read_excel("tasks/inspection_standard/files/codes.xlsx", sheet_name="customer")
        customer_dict = dict(zip(customer_codes['CUST_CD'], customer_codes['CUST_NM']))
        icm_obj = IncomingInfoWithPN()
        icm_2021_df_super_raw = icm_obj.monthly_search(yyyymm=ICM_YEAR, delta=delta_monthly)
        icm_2021_df_super_raw['고객사'] = [customer_dict.get(i, i) for i in icm_2021_df_super_raw['고객사']]
        icm_2021_df_super_raw['고객사'] = [str(i).replace("(CN)", "") for i in icm_2021_df_super_raw['고객사']]
        return icm_2021_df_super_raw


if __name__ == "__main__":

    import os
    import calendar
    os.chdir(os.pardir)
    os.chdir(os.pardir)
    os.chdir(os.pardir)

    raw_obj = RawDataGenerator()

    raw_obj.yyyymm = 202101
    df_ovs = raw_obj.get_yearly_raw_ovs()
    df_dms = raw_obj.get_yearly_raw_dms()
    df_icm = raw_obj.get_monthly_part_supply()
    icm_2021_df_super_raw = raw_obj.get_one_fixed_year_part_supply_super_raw()

    file_name = r"tasks\inspection_standard\export\raw_data.xlsx"

    with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
        workbook = writer.book
        icm_2021_df_super_raw.to_excel(writer, sheet_name="입고이력_2021", startrow=1, index=False)
        df_icm.to_excel(writer, sheet_name="입고이력_2021", startrow=1, index=False)
        df_ovs.to_excel(writer, sheet_name="해외_1년", startrow=1, index=False)
        df_dms.to_excel(writer, sheet_name="국내_1년", startrow=1, index=False)

    os.startfile(file_name)


    # icm_2021_df_super_raw.to_excel(r"spawn\test2.xlsx", index=False)
    # os.startfile(r"spawn\test2.xlsx")