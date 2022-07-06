import pandas as pd
from modules.db.db_domestic import DomesticDataBase
from modules.db.db_incoming import IncomingInfo
from modules.db.db_inr import IncomingResultInfo
from tqdm import tqdm
from utils_.config import *
from datetime import datetime


class MonthlyDMSppmCalculation:

    def __init__(self, yyyymm:int):
        self.yyyymm = yyyymm
        self.month_range = [i for i in range(int(str(self.yyyymm)[:4] + '01'), int(str(self.yyyymm)[:4] + '01' + '13'), 1) if i <= self.yyyymm] + ['종합']
        self.remaining_months = [i for i in range(int(str(self.yyyymm)[:4] + '01'), int(str(self.yyyymm)[:4] + '13'), 1) if i > self.yyyymm]
        self.df = None
        # self.dms_actions = DMS_VALID_TYPE #  Use negative criteria.
        self.dms_invalid_type = DMS_INVALID_TYPE
        self.customer_name_revision = CUSTOMER_NAME_REVISION
        self.customers = CUSTOMERS
        self.now = datetime.now().strftime("%Y-%m-%d_%H:%M:%S")
        self.download_date = []

    def get_ppm(self):

        self.generate_indices()

        for month in tqdm(self.month_range):
            from_to = False
            col_name = month
            if month == '종합':
                month = self.yyyymm
                from_to = True
                col_name = '종합'

            domestic = DomesticDataBase()
            dms_df = domestic.search(month, from_to=from_to)
            dms_df['공급법인'] = [i.replace('(N)', "").replace('(A)', "") for i in dms_df['공급법인']]
            dms_df['고객사'] = [str(i).replace("(CN)", "") for i in dms_df['고객사'].replace(self.customer_name_revision)]
            dms_df = dms_df[~dms_df['불량조치'].isin(self.dms_invalid_type)]

            new_date = dms_df['입력일자'].unique().tolist()
            for date in new_date:
                if date not in self.download_date:
                    self.download_date.append(date)

            incoming = IncomingInfo()
            icm_df = incoming.search(month, from_to=from_to)
            icm_df['고객사'] = [i.replace("(CN)", "").replace("(M)", "") for i in icm_df['고객사'].replace(self.customer_name_revision)]
            icm_df['국가'] = icm_df['사업장'].map(GLOBAL_CENTERS)

            results = IncomingResultInfo()
            results_df = results.search(month, from_to=from_to)
            results_df = results_df[results_df['공급법인'] == "글로비스 한국"]

            slots = ["" for _ in range(24)]

            kor_incoming = sum(icm_df[icm_df['국가'] == "한국"]['입고수량'])
            chn_incoming = sum(icm_df[icm_df['국가'] == "중국"]['입고수량'])
            ind_incoming = sum(icm_df[icm_df['국가'] == "인도"]['입고수량'])
            tur_incoming = sum(icm_df[icm_df['국가'] == "터키"]['입고수량'])
            total_incoming = kor_incoming + chn_incoming + ind_incoming + tur_incoming

            try :
                assert total_incoming == sum(icm_df['입고수량'])
            except AssertionError:
                print(icm_df[~icm_df['국가'].isin(['한국', "중국", "인도", "터키"])]['포장장_명'].iloc[0])

            kor_dms_qty = sum(dms_df[dms_df['공급법인'] == "글로비스 한국"]['불량수량'])
            chn_dms_qty = sum(dms_df[dms_df['공급법인'] == "글로비스 중국"]['불량수량'])
            ind_dms_qty = sum(dms_df[dms_df['공급법인'] == "글로비스 인도"]['불량수량'])
            tur_dms_qty = sum(dms_df[dms_df['공급법인'] == "글로비스 터키"]['불량수량'])
            total_dms = kor_dms_qty + chn_dms_qty + ind_dms_qty + tur_dms_qty

            try: assert total_dms == sum(dms_df['불량수량'])
            except AssertionError: print(total_dms, sum(dms_df['불량수량']))

            ppm_kor = round(kor_dms_qty / kor_incoming * 1000000, 1) if kor_incoming != 0 else 0
            ppm_chn = round(chn_dms_qty / chn_incoming * 1000000, 1) if chn_incoming != 0 else 0
            ppm_ind = round(ind_dms_qty / ind_incoming * 1000000, 1) if ind_incoming != 0 else 0
            ppm_tur = round(tur_dms_qty / tur_incoming * 1000000, 1) if tur_incoming != 0 else 0
            ppm_total = round(total_dms / total_incoming * 1000000, 1) if total_incoming != 0 else 0

            joined_ins_qty = sum(results_df[results_df['검사기준']=="합동검사"]['입고수량'])
            joined_def_qty = sum(results_df[results_df['검사기준'] == "합동검사"]['불량수량'])
            total_ins_qty = sum(results_df[results_df['검사기준'] == "전수검사"]['입고수량'])
            total_def_qty = sum(results_df[results_df['검사기준'] == "전수검사"]['불량수량'])
            all_ins_qty = joined_ins_qty + total_ins_qty
            all_def_qty = joined_def_qty + total_def_qty
            ppm_joined_rst = round(joined_def_qty/joined_ins_qty * 1000000, 1) if joined_ins_qty !=0 else 0
            ppm_total_rst = round(total_def_qty / total_ins_qty * 1000000, 1) if total_ins_qty != 0 else 0
            ppm_all_rst = round(all_def_qty / all_ins_qty * 1000000, 1) if all_ins_qty != 0 else 0

            slots[0] = ppm_total
            slots[1] = ppm_kor
            slots[2] = ppm_chn
            slots[3] = ppm_ind
            slots[4] = ppm_tur
            slots[5] = total_dms
            slots[6] = kor_dms_qty
            slots[7] = chn_dms_qty
            slots[8] = ind_dms_qty
            slots[9] = tur_dms_qty
            slots[10] = total_incoming
            slots[11] = kor_incoming
            slots[12] = chn_incoming
            slots[13] = ind_incoming
            slots[14] = tur_incoming
            slots[15] = ppm_joined_rst
            slots[16] = joined_def_qty
            slots[17] = joined_ins_qty
            slots[18] = ppm_total_rst
            slots[19] = total_def_qty
            slots[20] = total_ins_qty
            slots[21] = ppm_all_rst
            slots[22] = all_def_qty
            slots[23] = all_ins_qty

            self.df[col_name] = slots

        for month in self.remaining_months:
            self.df[month] = ""
        self.df.set_index(['구분1', '구분2'], drop=True, inplace=True)
        self.df = self.df[['종합'] + [i for i in self.month_range if i != "종합"] + self.remaining_months]

        self.download_date = [datetime.strptime(i, "%Y%m%d").strftime("%Y-%m-%d") for i in self.download_date]

        self.download_date = ', '.join(self.download_date)

        return self.df

    def generate_indices(self):

        first_col = []
        second_col = []

        for subject in ['국내입고불량 실적', '불량수량', '입고수량']:
            first_col += [subject for _ in range(5)]
            second_col += ["전체", "한국", "중국", "인도", "터키"]

        for subject in ['합동검사실적', '전수검사실적', '합동Q-키퍼']:
            first_col += [subject for _ in range(3)]
            second_col += ["PPM", "불량수량", "검사수량"]

        self.df = pd.DataFrame({'구분1': first_col, '구분2': second_col})


if __name__ == "__main__":
    import os
    os.chdir(os.pardir)
    os.chdir(os.pardir)

    obj = MonthlyDMSppmCalculation(202206)
    df = obj.get_ppm()

    print(len(df.index[0]))
    print(len(df.columns))
    print(len(df.loc[('국내입고불량 실적', )]))

    # df.to_excel(r'spawn\test3.xlsx', index=True)
    # os.startfile(r'spawn\test3.xlsx')
