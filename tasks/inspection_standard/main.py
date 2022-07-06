from tasks.inspection_standard.pipelines.processed_data import DataProcessor
from collections import Counter
from utils_.functions import show_elapsed_time
import os
import pandas as pd
from tasks.inspection_standard.config import *


class Main:
    def __init__(self):
        self.pcs_obj = DataProcessor()
        self.yyyymm = None

    def generate(self):
        self.df_ovs = self.pcs_obj.transform_ovs(self.yyyymm, delta=YEARLY_DELTA_OVS)
        self.df_dms = self.pcs_obj.transform_dms(self.yyyymm, delta=YEARLY_DELTA_DMS)
        self.df_icm = self.pcs_obj.transform_icm(self.df_ovs, self.df_dms, delta_monthly=MONTHLY_DELTA_ICM)
        self.df_icm_pmax = self.pcs_obj.part_max_icm(self.yyyymm, self.df_icm, s_data_range_yearly=S_DATA_RANGE_YEAR)
        self.df_icm_combined = self.pcs_obj.combine_icm(self.df_icm, self.df_icm_pmax)
        self.df_pac, self.df_icm_raw = self.pcs_obj.packing_center_review(self.df_icm_combined, delta_month=MONTHLY_DELTA_ICM)

        self.df_summary_1 = pd.DataFrame(list(Counter(self.df_icm_pmax['품목등급']).items()), columns=['등급', '품목수'])
        total = self.df_summary_1['품목수'].sum()
        self.df_summary_1.loc['합계'] = {'등급': "합계", "품목수": total}
        self.df_summary_1['비중(%)'] = round(self.df_summary_1['품목수']/total * 100,1)

        self.df_summary_2 = pd.DataFrame(list(Counter(self.df_icm_combined['품목등급']).items()), columns=['등급', '품번수'])
        total = self.df_summary_2['품번수'].sum()
        self.df_summary_2.loc['합계'] = {'등급': "합계", "품번수": total}
        self.df_summary_2['비중(%)'] = round(self.df_summary_2['품번수']/total * 100,1)

        self.file_name = rf"tasks\inspection_standard\export\{self.yyyymm}_processed_data_case_7.xlsx"

    @show_elapsed_time
    def export(self):
        with pd.ExcelWriter(self.file_name, engine='xlsxwriter') as writer:
            workbook = writer.book
            self.df_pac.to_excel(writer, sheet_name="입고_포장장검증", startrow=1, index=True)
            self.df_summary_1.to_excel(writer, sheet_name="입고_포장장검증", startrow=1, startcol=15, index=False)
            self.df_summary_2.to_excel(writer, sheet_name="입고_포장장검증", startrow=9, startcol=15, index=False)
            self.df_icm_raw.to_excel(writer, sheet_name="입고_로우", startrow=1, index=False)
            self.df_icm_combined.to_excel(writer, sheet_name="입고_종합", startrow=1, index=False)
            self.df_icm_pmax.to_excel(writer, sheet_name="입고_품목", startrow=1, index=False)
            self.df_icm.to_excel(writer, sheet_name="입고_변환", startrow=1, index=False)
            self.df_ovs.to_excel(writer, sheet_name="해외_변환", startrow=1, index=False)
            self.df_dms.to_excel(writer, sheet_name="국내_변환", startrow=1, index=False)

        # os.startfile(self.file_name)


if __name__ == "__main__":

    os.chdir(os.pardir)
    os.chdir(os.pardir)

    obj = Main()
    month_range = [i for i in range(202101, 202102)]
    df_stack = None
    for i in month_range:
        print(i)
        obj.yyyymm = i
        obj.generate()
        obj.export()

        df = obj.df_summary_2
        df['yyyymm'] = i

        if df_stack is None:
            df_stack = df
        else:
            df_stack = pd.concat([df_stack, df], axis=0)

        print(df_stack)

    df_stack.to_excel(r"spawn\test3.xlsx", index=False)
    # os.startfile(r"spawn\test3.xlsx")
    os.startfile(r"tasks\inspection_standard\export")
