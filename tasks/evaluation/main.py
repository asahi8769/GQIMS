from modules.db.db_overseas import OverseasDataBase
from modules.db.db_domestic import DomesticDataBase
from modules.db.db_dqy import DisbursementInfo
from modules.db.db_incoming import IncomingInfo
import pandas as pd
import numpy as np
from datetime import datetime
from tasks.evaluation.configurations import *
from dateutil.relativedelta import relativedelta
from utils_.functions import show_elapsed_time
import scipy.stats as stats
from collections import Counter
from datetime import datetime, timedelta
from pytimekr import pytimekr


class Eval:

    def __init__(self):
        self.days_weight_df = pd.read_excel("tasks/evaluation/eval_critera.xlsx", sheet_name="PPR접수소요일")
        self.response_weight_df = pd.read_excel("tasks/evaluation/eval_critera.xlsx", sheet_name="대책회신률")
        self.defect_weight_df = pd.read_excel("tasks/evaluation/eval_critera.xlsx", sheet_name="불량개선율")
        self.days_weight_dict = dict(
            zip(list(zip(self.days_weight_df['이상'], self.days_weight_df['미만'])), self.days_weight_df['점수']))

        self.response_weight_dict = dict(
            zip(list(zip(self.response_weight_df['이상'], self.response_weight_df['미만'])), self.response_weight_df['점수']))

        self.defect_weight_dict = dict(
            zip(list(zip(self.defect_weight_df['이상'], self.defect_weight_df['미만'])), self.defect_weight_df['점수']))

        self.customers = []
        self.customer_per_pic = {}

        self.pics = [i[0] for i in CUSTOMER_PICS.values()]
        self.customer_per_pic = {i: [] for i in self.pics}

    def get_defect_table(self, *yyyymm, type_='외관이종'):
        yyyymm_minus_3 = tuple(
            int((datetime.strptime(str(i), '%Y%m') - relativedelta(months=3)).strftime('%Y%m')) for i in yyyymm)
        yyyymm_minus_6 = tuple(
            int((datetime.strptime(str(i), '%Y%m') - relativedelta(months=6)).strftime('%Y%m')) for i in yyyymm)

        df_ovs_old = self.get_ovs_raw(*yyyymm_minus_3)
        df_dqy_old = self.get_dqy_raw(*yyyymm_minus_3)
        df_icm_old = self.get_icm_raw(*yyyymm_minus_3)

        df_ovs_old = df_ovs_old[df_ovs_old['불량구분_'] == type_]

        df_ovs_old_pivot = df_ovs_old.pivot_table(values='조치수량', index=['고객사', '담당자_'], columns='timestamp',
                                                  aggfunc=np.sum)
        df_dqy_old_pivot = df_dqy_old.pivot_table(values='불출수량', index=['대표고객', '담당자_'], columns='timestamp',
                                                  aggfunc=np.sum)
        df_icm_old_pivot = df_icm_old.pivot_table(values='입고수량', index=['고객사', '담당자_'], columns='timestamp',
                                                  aggfunc=np.sum)
        df_dqy_old_pivot.index.names = ['고객사', '담당자_']

        df_old = df_ovs_old_pivot.merge(df_dqy_old_pivot, how='left', on=['고객사', '담당자_'])
        df_old = df_old.merge(df_icm_old_pivot, how='left', on=['고객사', '담당자_'])
        df_old.reset_index(inplace=True)
        df_old['메이저고객사_'] = [CUSTOMER_PICS.get(i, NON_DESIGNATED_CUSTOMER_PIC)[1] for i in df_old['고객사']]

        var_col = []
        for k, j in enumerate(yyyymm_minus_3):
            df_old[f'{j}_y'] = [df_old[f'{j}_y'].iloc[n] if i == 1 else df_old[yyyymm_minus_6[k]].iloc[n] for n, i in
                                enumerate(df_old['메이저고객사_'])]
            df_old.rename(columns={f'{j}_x': f'{j}_불량수량'}, inplace=True)
            df_old.rename(columns={f'{j}_y': f'{j}_불출수량'}, inplace=True)
            df_old.drop([yyyymm_minus_6[k]], axis=1, inplace=True)

            var_col.append(f'{j}_불량수량')
            var_col.append(f'{j}_불출수량')

        df_old = df_old.pivot_table(values=var_col, index=['담당자_'], aggfunc=np.sum)  ###

        for k, j in enumerate(yyyymm_minus_3):
            df_old[f'{j}_PPM'] = round(df_old[f'{j}_불량수량'] / df_old[f'{j}_불출수량'] * 1000000, 1)

        df_old.fillna(0, inplace=True)

        df_old['MIN'] = [min(df_old[f'{yyyymm_minus_3[0]}_PPM'].iloc[i], df_old[f'{yyyymm_minus_3[1]}_PPM'].iloc[i],
                             df_old[f'{yyyymm_minus_3[2]}_PPM'].iloc[i]) for i in range(len(df_old))]
        df_old['MAX'] = [max(df_old[f'{yyyymm_minus_3[0]}_PPM'].iloc[i], df_old[f'{yyyymm_minus_3[1]}_PPM'].iloc[i],
                             df_old[f'{yyyymm_minus_3[2]}_PPM'].iloc[i]) for i in range(len(df_old))]

        df_old['과거조치수량'] = df_old[f'{yyyymm_minus_3[0]}_불량수량'] + df_old[f'{yyyymm_minus_3[1]}_불량수량'] + \
                           df_old[f'{yyyymm_minus_3[2]}_불량수량']
        df_old['과거불출수량'] = df_old[f'{yyyymm_minus_3[0]}_불출수량'] + df_old[f'{yyyymm_minus_3[1]}_불출수량'] + \
                           df_old[f'{yyyymm_minus_3[2]}_불출수량']
        df_old['이전분기종합실적'] = round(df_old['과거조치수량'] / df_old['과거불출수량'] * 1000000, 1)

        df_old['이전분기환산점수'] = round(
            100 - (2 * (df_old['이전분기종합실적'] - df_old['MIN']) / (df_old['MAX'] - df_old['MIN']) - 1),
            1)  # 100-(2*(P3-N3)/(O3-N3)-1)
        df_old.drop(['과거조치수량', '과거불출수량'], axis=1, inplace=True)

        df_ovs_new = self.get_ovs_raw(*yyyymm)
        df_dqy_new = self.get_dqy_raw(*yyyymm)
        df_icm_new = self.get_icm_raw(*yyyymm)

        df_ovs_new = df_ovs_new[df_ovs_new['불량구분_'] == type_]

        df_ovs_new_pivot = df_ovs_new.pivot_table(values='조치수량', index=['고객사', '담당자_'], aggfunc=np.sum)
        df_dqy_new_pivot = df_dqy_new.pivot_table(values='불출수량', index=['대표고객', '담당자_'], aggfunc=np.sum)
        df_icm_new_pivot = df_icm_new.pivot_table(values='입고수량', index=['고객사', '담당자_'], aggfunc=np.sum)
        df_dqy_new_pivot.index.names = ['고객사', '담당자_']

        df_new = df_ovs_new_pivot.merge(df_dqy_new_pivot, how='left', on=['고객사', '담당자_'])
        df_new = df_new.merge(df_icm_new_pivot, how='left', on=['고객사', '담당자_'])
        df_new.reset_index(inplace=True)
        df_new['메이저고객사_'] = [CUSTOMER_PICS.get(i, NON_DESIGNATED_CUSTOMER_PIC)[1] for i in df_new['고객사']]

        df_new['불출수량'] = [df_new['불출수량'].iloc[n] if i == 1 else df_new['입고수량'].iloc[n] for n, i in
                          enumerate(df_new['메이저고객사_'])]

        df_new = df_new.pivot_table(values=['조치수량', '불출수량'], index=['담당자_'], aggfunc=np.sum)  ####

        df_new['금번분기종합실적'] = round(df_new['조치수량'] / df_new['불출수량'] * 1000000, 1)

        df_new = df_new.merge(df_old[['이전분기종합실적', 'MIN', 'MAX', '이전분기환산점수']], how='left', on=['담당자_'])
        df_new['금번분기환산점수'] = round(
            100 - (2 * (df_new['금번분기종합실적'] - df_new['MIN']) / (df_new['MAX'] - df_new['MIN']) - 1), 1)

        col = ['조치수량', '불출수량', '이전분기종합실적', '금번분기종합실적', 'MIN', 'MAX', '이전분기환산점수', '금번분기환산점수']
        df_new = df_new[col]

        df_new.fillna(0, inplace=True)
        df_new['개선율(%)'] = df_new['금번분기환산점수'] - df_new['이전분기환산점수']

        df_new['등급'] = df_new['개선율(%)'].map(lambda x: self.evaluate(x, self.defect_weight_dict))

        self.customer_per_pic = {i: [] for i in self.pics}
        self.get_custmer_per_pic(df_ovs_old)

        df_old['담당고객사'] = [', '.join(self.customer_per_pic.get(j, [])) for j in df_old.index]
        df_old.reset_index(inplace=True)
        df_old = df_old[['담당고객사'] + df_old.columns[:-1].tolist()]

        self.customer_per_pic = {i: [] for i in self.pics}
        self.get_custmer_per_pic(df_ovs_new)

        df_new['담당고객사'] = [', '.join(self.customer_per_pic.get(j, [])) for j in df_new.index]
        df_new.reset_index(inplace=True)
        df_new = df_new[['담당고객사'] + df_new.columns[:-1].tolist()]
        print(Counter(df_new['등급']))

        return df_old, df_new

    def get_delayed_table(self, *yyyymm):
        df_ovs = self.get_ovs_raw(*yyyymm)
        df_ovs = df_ovs[~df_ovs['공급법인'].isin(INVALID_PPR_SUPPLIED_FROM)]
        df_dms = self.get_dms_raw(*yyyymm)
        df_dms = df_dms[~df_dms['공급법인'].isin(INVALID_PPR_SUPPLIED_FROM)]

        self.customer_per_pic = {i: [] for i in self.pics}
        self.get_custmer_per_pic(df_ovs)
        self.get_custmer_per_pic(df_dms)

        df_ovs_pivot = df_ovs.pivot_table(values='해외_담당접수소요일', index=KEYS, aggfunc=np.mean)
        df_ovs_pivot['해외점수'] = df_ovs_pivot['해외_담당접수소요일'].map(lambda x: self.evaluate(x, self.days_weight_dict))

        df_dms_pivot = df_dms.pivot_table(values='국내_담당결재소요일', index=KEYS, aggfunc=np.mean)
        df_dms_pivot['국내점수'] = df_dms_pivot['국내_담당결재소요일'].map(lambda x: self.evaluate(x, self.days_weight_dict))

        df_delayed = df_ovs_pivot.merge(df_dms_pivot, how='outer', on=KEYS)

        df_delayed['해외_담당접수소요일'] = round(df_delayed['해외_담당접수소요일'], 1)
        df_delayed['국내_담당결재소요일'] = round(df_delayed['국내_담당결재소요일'], 1)

        df_delayed['평균소요일'] = [df_delayed['해외_담당접수소요일'].iloc[i] * 0.5 + df_delayed['국내_담당결재소요일'].iloc[i] * 0.5 for i in
                               range(len(df_delayed))]
        df_delayed['평균소요일'] = [df_delayed['해외_담당접수소요일'].iloc[n] if np.isnan(i) else i for n, i in
                               enumerate(df_delayed['평균소요일'])]
        df_delayed['평균소요일'] = [df_delayed['국내_담당결재소요일'].iloc[n] if np.isnan(i) else i for n, i in
                               enumerate(df_delayed['평균소요일'])]
        df_delayed['평균소요일'] = round(df_delayed['평균소요일'], 1)
        df_delayed['종합점수'] = df_delayed['평균소요일'].map(lambda x: self.evaluate(x, self.days_weight_dict))

        index_col = [i[-1] for i in df_delayed.index] if len(KEYS) > 1 else df_delayed.index
        df_delayed['담당고객사'] = [', '.join(self.customer_per_pic.get(j, [])) for j in index_col]

        df_delayed.reset_index(inplace=True)
        df_delayed = df_delayed[['담당고객사'] + df_delayed.columns[:-1].tolist()]

        return df_delayed

    def get_custmer_per_pic(self, df):
        df_pivot_dummy = df.pivot_table(values='ID', index=['고객사', '담당자_'], columns='대책접수완료', aggfunc=np.size)
        df_pivot_dummy.reset_index(inplace=True)

        for i in self.pics:
            self.customer_per_pic[i] = list(
                set(list(df_pivot_dummy[df_pivot_dummy['담당자_'] == i]['고객사'].unique()) + self.customer_per_pic[i]))

    def get_response_rate(self, *yyyymm):
        df_ovs = self.get_ovs_raw(*yyyymm)
        df_ovs = df_ovs[~df_ovs['공급법인'].isin(INVALID_PPR_SUPPLIED_FROM)]
        df_dms = self.get_dms_raw(*yyyymm)
        df_dms = df_dms[~df_dms['공급법인'].isin(INVALID_PPR_SUPPLIED_FROM)]

        self.customer_per_pic = {i: [] for i in self.pics}
        self.get_custmer_per_pic(df_ovs)
        self.get_custmer_per_pic(df_dms)

        df_ovs_pivot = df_ovs.pivot_table(values='ID', index=KEYS, columns='대책접수완료', aggfunc=np.size)
        df_ovs_pivot.fillna(0, inplace=True)
        df_ovs_pivot.columns = ['해외_미회신', '해외_회신']
        df_ovs_pivot['해외_종합건수'] = df_ovs_pivot['해외_미회신'] + df_ovs_pivot['해외_회신']
        df_ovs_pivot['해외_회신율(%)'] = round(df_ovs_pivot['해외_회신'] / df_ovs_pivot['해외_종합건수'] * 100, 1)

        df_dms_pivot = df_dms.pivot_table(values='ID', index=KEYS, columns='대책접수완료', aggfunc=np.size)
        df_dms_pivot.fillna(0, inplace=True)
        df_dms_pivot.columns = ['국내_미회신', '국내_회신']
        df_dms_pivot['국내_종합건수'] = df_dms_pivot['국내_미회신'] + df_dms_pivot['국내_회신']
        df_dms_pivot['국내_회신율(%)'] = round(df_dms_pivot['국내_회신'] / df_dms_pivot['국내_종합건수'] * 100, 1)

        df_response = df_ovs_pivot.merge(df_dms_pivot, how='outer', on=KEYS)
        df_response['종합건수'] = df_response['해외_종합건수'] + df_response['국내_종합건수']
        df_response['종합건수'] = [df_response['해외_종합건수'].iloc[n] if np.isnan(i) else i for n, i in
                               enumerate(df_response['종합건수'])]
        df_response['종합건수'] = [df_response['국내_종합건수'].iloc[n] if np.isnan(i) else i for n, i in
                               enumerate(df_response['종합건수'])]

        df_response['종합점수(%)'] = [df_response['해외_회신율(%)'].iloc[i] * 0.5 + df_response['국내_회신율(%)'].iloc[i] * 0.5 for i
                                  in range(len(df_response))]
        df_response['종합점수(%)'] = [df_response['해외_회신율(%)'].iloc[n] if np.isnan(i) else i for n, i in
                                  enumerate(df_response['종합점수(%)'])]
        df_response['종합점수(%)'] = [df_response['국내_회신율(%)'].iloc[n] if np.isnan(i) else i for n, i in
                                  enumerate(df_response['종합점수(%)'])]

        df_response['종합점수(%)'] = round(df_response['종합점수(%)'], 1)

        df_response['Z점수'] = stats.zscore(np.array(df_response['종합건수']))
        df_response['Z점수'] = round(df_response['Z점수'], 5)
        df_response['Z점수(적용)'] = [round(i * 10, 1) if i >= 0 else 0 for i in df_response['Z점수']]

        df_response['종합점수(%)_환산'] = round(df_response['종합점수(%)'] + df_response['Z점수(적용)'], 1)
        df_response['등급'] = df_response['종합점수(%)_환산'].map(lambda x: self.evaluate(x, self.response_weight_dict))

        index_col = [i[-1] for i in df_response.index] if len(KEYS) > 1 else df_response.index
        df_response['담당고객사'] = [', '.join(self.customer_per_pic.get(j, [])) for j in index_col]

        df_response.reset_index(inplace=True)
        df_response = df_response[['담당고객사'] + df_response.columns[:-1].tolist()]

        return df_response

    @staticmethod
    def evaluate(days, elapse_weight_dict):
        for i in elapse_weight_dict.keys():
            if i[0] <= days < i[1]:
                return elapse_weight_dict.get(i, 0)

    def weekdays_delta(self, start_col, end_col):
        additional_holidays = ADDITIONAL_HOLIDAYS
        result_col = [self.count_working_days(start_col.iloc[i], end_col.iloc[i], *additional_holidays) if
                      (start_col.iloc[i] != "" and end_col.iloc[i] != "") else np.nan for i in range(len(start_col))]
        result_col = [i if (i != "" and i >= 0) else np.nan for i in result_col]

        return result_col

    @staticmethod
    def count_working_days(start, end, *additional_holidays):
        year_start = start.year
        year_end = end.year
        days_list = [start + timedelta(days=x) for x in range((end - start).days)]
        holidays = pytimekr.holidays(year=year_start) + pytimekr.holidays(year=year_end) + \
                   [datetime.strptime(i, "%Y%m%d").date() for i in additional_holidays]
        return len([i for i in days_list if i.weekday() not in [5, 6] and i not in holidays])

    def get_ovs_raw(self, *yyyymm):
        ovs_obj = OverseasDataBase(pre_noti=PRE_NOTI)
        df = ovs_obj.monthly_specific_search(*yyyymm, reference="timestamp")
        df['고객사'] = [str(i).replace("(CN)", "").replace("(M)", "") for i in df['고객사']]

        df = df[(~df["KDQR_Type"].isin(OVS_INVALID_TYPE))]
        df = df[(~df["PPR_No"].isin(PPRS_TO_EXCLUDE))]

        self.customers += df['고객사'].unique().tolist()
        self.customers = list(set(self.customers))

        df['담당자_'] = [CUSTOMER_PICS.get(i, NON_DESIGNATED_CUSTOMER_PIC)[0] for i in df['고객사']]
        df = df[~((df["담당자_"] == THREEPL_CUSTOMER_PIC) & (df['공급법인'].isin(INVALID_PPR_SUPPLIED_FROM)))]

        df['메이저고객사_'] = [CUSTOMER_PICS.get(i, NON_DESIGNATED_CUSTOMER_PIC)[1] for i in df['고객사']]
        df['불량구분_'] = ["외관이종" if i in APPEARANCE_TYPE else "기능치수" for i in df['불량구분']]
        df['대책접수완료'] = [1 if i in CM_COMPLETE else 0 for i in df['진행상태']]

        date_col = ['발생일자', '등록일자', '해외결재일', '국내접수일', '국내결재일', '대책등록일', '대책확인일']
        for i in date_col:
            df[f'{i}'] = [datetime.strptime(str(j).split(" ")[0], "%Y-%m-%d").date() if j is not None else "" for j in
                          df[i]]

        df['해외결재소요일'] = self.weekdays_delta(df['등록일자'], df['해외결재일'])
        df['해외_담당접수소요일'] = self.weekdays_delta(df['해외결재일'], df['국내접수일'])
        df['국내결재소요일'] = self.weekdays_delta(df['국내접수일'], df['국내결재일'])
        df['대책등록소요일'] = self.weekdays_delta(df['국내결재일'], df['대책등록일'])
        df['대책확인소요일'] = self.weekdays_delta(df['대책등록일'], df['대책확인일'])

        return df

    def get_dms_raw(self, *yyyymm):
        dms_obj = DomesticDataBase()
        df = dms_obj.monthly_specific_search(*yyyymm, reference="timestamp")
        df['고객사'] = [str(i).replace("(CN)", "").replace("(M)", "") for i in df['고객사']]

        df = df[(~df["불량조치"].isin(DMS_INVALID_TYPE))]

        self.customers += df['고객사'].unique().tolist()
        self.customers = list(set(self.customers))

        df['담당자_'] = [CUSTOMER_PICS.get(i, NON_DESIGNATED_CUSTOMER_PIC)[0] for i in df['고객사']]
        df = df[~((df["담당자_"] == THREEPL_CUSTOMER_PIC) & (df['공급법인'].isin(INVALID_PPR_SUPPLIED_FROM)))]

        df['메이저고객사_'] = [CUSTOMER_PICS.get(i, NON_DESIGNATED_CUSTOMER_PIC)[1] for i in df['고객사']]
        df['불량구분_'] = ["외관이종" if i in APPEARANCE_TYPE else "기능치수" for i in df['불량구분']]
        df['대책접수완료'] = [1 if i in CM_COMPLETE else 0 for i in df['진행상태']]

        date_col = ['등록일자', '검사일자', '통보일자', '접수일자', '결재일자', '대책등록일', '대책확인일']
        for i in date_col:
            df[f'{i}'] = [datetime.strptime(str(j).split(" ")[0], "%Y-%m-%d").date() if j is not None else "" for j in
                          df[i]]

        df['국내등록소요일'] = self.weekdays_delta(df['검사일자'], df['등록일자'])
        df['소장결재소요일'] = self.weekdays_delta(df['등록일자'], df['통보일자'])
        df['국내_담당결재소요일'] = self.weekdays_delta(df['통보일자'], df['결재일자'])
        df['대책등록소요일'] = self.weekdays_delta(df['결재일자'], df['대책등록일'])
        df['대책확인소요일'] = self.weekdays_delta(df['대책등록일'], df['대책확인일'])

        return df

    def get_dqy_raw(self, *yyyymm):
        dqy_obj = DisbursementInfo()
        df = dqy_obj.monthly_specific_search(*yyyymm, reference="timestamp")
        df['대표고객'] = [str(i).replace("(CN)", "").replace("(M)", "") for i in df['대표고객']]

        df = df[df['대표고객'].isin(self.customers)]

        df['담당자_'] = [CUSTOMER_PICS.get(i, NON_DESIGNATED_CUSTOMER_PIC)[0] for i in df['대표고객']]
        df['메이저고객사_'] = [CUSTOMER_PICS.get(i, NON_DESIGNATED_CUSTOMER_PIC)[1] for i in df['대표고객']]
        return df

    def get_icm_raw(self, *yyyymm):
        dqy_obj = IncomingInfo()

        yyyymm_minus_3 = tuple(
            int((datetime.strptime(str(i), '%Y%m') - relativedelta(months=3)).strftime('%Y%m')) for i in yyyymm)
        df = dqy_obj.monthly_specific_search(*yyyymm_minus_3, reference="timestamp")
        df['고객사'] = [str(i).replace("(CN)", "").replace("(M)", "") for i in df['고객사']]
        df['timestamp_대상불량년월'] = list(
            int((datetime.strptime(str(i), '%Y%m') + relativedelta(months=3)).strftime('%Y%m')) for i in
            df['timestamp'])

        df = df[df['고객사'].isin(self.customers)]

        df['담당자_'] = [CUSTOMER_PICS.get(i, NON_DESIGNATED_CUSTOMER_PIC)[0] for i in df['고객사']]
        df['메이저고객사_'] = [CUSTOMER_PICS.get(i, NON_DESIGNATED_CUSTOMER_PIC)[1] for i in df['고객사']]
        return df


if __name__ == "__main__":
    import os

    os.chdir(os.pardir)
    os.chdir(os.pardir)

    current = CURRENT_QUARTER
    past = PAST_QUARTER

    obj = Eval()
    current_ovs = obj.get_ovs_raw(*current)
    current_ovs['timestamp'] = current_ovs['timestamp'].apply(str)

    old_ovs = obj.get_ovs_raw(*past)
    old_ovs['timestamp'] = old_ovs['timestamp'].apply(str)

    current_dms = obj.get_dms_raw(*current)
    current_dms['timestamp'] = current_dms['timestamp'].apply(str)

    old_dms = obj.get_dms_raw(*past)
    old_dms['timestamp'] = old_dms['timestamp'].apply(str)

    current_dqy = obj.get_dqy_raw(*current)
    current_dqy['timestamp'] = current_dqy['timestamp'].apply(str)

    old_dqy = obj.get_dqy_raw(*past)
    old_dqy['timestamp'] = old_dqy['timestamp'].apply(str)

    current_icm = obj.get_icm_raw(*current)
    current_icm['timestamp'] = current_icm['timestamp'].apply(str)

    old_dqy_icm = obj.get_icm_raw(*past)
    old_dqy_icm['timestamp'] = old_dqy_icm['timestamp'].apply(str)

    df_response = obj.get_response_rate(*current)
    df_delayed = obj.get_delayed_table(*current)
    df_old, df_new = obj.get_defect_table(*current, type_='외관이종')
    df_old_f, df_new_f = obj.get_defect_table(*current, type_='기능치수')

    file_name = r"tasks\evaluation\test1.xlsx"

    current_pic = current_ovs.pivot_table(values=['조치수량'], index=['고객사', '담당자_'], columns=['불량구분_'], aggfunc=np.sum)

    with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
        workbook = writer.book
        df_delayed.to_excel(writer, sheet_name="업무소요일_대책회신율", startrow=1, startcol=1, index=False)
        df_response.to_excel(writer, sheet_name="업무소요일_대책회신율", startrow=15, startcol=1, index=False)

        df_old.to_excel(writer, sheet_name="불량개선율_외관이종", startrow=1, startcol=1, index=False)
        df_new.to_excel(writer, sheet_name="불량개선율_외관이종", startrow=15, startcol=1, index=False)

        df_old_f.to_excel(writer, sheet_name="불량개선율_기능치수", startrow=1, startcol=1, index=False)
        df_new_f.to_excel(writer, sheet_name="불량개선율_기능치수", startrow=15, startcol=1, index=False)

        current_ovs.to_excel(writer, sheet_name="해외_현재", startrow=1, startcol=1, index=False)
        old_ovs.to_excel(writer, sheet_name="해외_과거", startrow=1, startcol=1, index=False)
        current_dms.to_excel(writer, sheet_name="국내_현재", startrow=1, startcol=1, index=False)
        old_dms.to_excel(writer, sheet_name="국내_과거", startrow=1, startcol=1, index=False)
        current_dqy.to_excel(writer, sheet_name="불출수량_현재", startrow=1, startcol=1, index=False)
        old_dqy.to_excel(writer, sheet_name="불출수량_과거", startrow=1, startcol=1, index=False)
        current_icm.to_excel(writer, sheet_name="입고수량_현재", startrow=1, startcol=1, index=False)
        old_dqy_icm.to_excel(writer, sheet_name="입고수량_과거", startrow=1, startcol=1, index=False)
        # current_pic.to_excel(writer, sheet_name="고객,담당관계", startrow=1, index=True)

    from tasks.evaluation.format import formatting
    import openpyxl

    xfile = openpyxl.load_workbook(file_name)

    formatting(xfile, df_delayed, '업무소요일_대책회신율')
    formatting(xfile, df_response, '업무소요일_대책회신율')

    formatting(xfile, df_old, '불량개선율_외관이종')
    formatting(xfile, df_new, '불량개선율_외관이종')

    formatting(xfile, df_old_f, '불량개선율_기능치수')
    formatting(xfile, df_new_f, '불량개선율_기능치수')

    formatting(xfile, current_ovs, '해외_현재', raw=True)
    formatting(xfile, old_ovs, '해외_과거', raw=True)
    formatting(xfile, current_dms, '국내_현재', raw=True)
    formatting(xfile, old_dms, '국내_과거', raw=True)
    formatting(xfile, current_dqy, '불출수량_현재', raw=True)
    formatting(xfile, old_dqy, '불출수량_과거', raw=True)
    formatting(xfile, current_icm, '입고수량_현재', raw=True)
    formatting(xfile, old_dqy_icm, '입고수량_과거', raw=True)

    xfile.save(file_name)

    os.startfile(file_name)
