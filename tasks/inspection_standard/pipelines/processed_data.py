from tasks.inspection_standard.config import *
from tasks.inspection_standard.pipelines.raw_data import RawDataGenerator
import numpy as np
from utils_.functions import show_elapsed_time
import pandas as pd


class DataProcessor:
    def __init__(self, b_rank_criterion=RANK_CRITERION_B, n_of_d=N_OF_D, n_of_c=N_OF_C, n_of_p=N_OF_P, n_of_m=N_OF_M,
                 keys=KEYS):
        self.raw_obj = RawDataGenerator()
        self.b_rank_criterion = b_rank_criterion
        self.n_of_d = n_of_d
        self.n_of_c = n_of_c
        self.n_of_p = n_of_p
        self.n_of_m = n_of_m
        self.keys = keys
        self.key_part_category = [i if i != "부품번호" else "품목" for i in self.keys]

    def get_zero_defect_dict(self, yyyymm, delta_test=1):
        df_ovs = self.transform_ovs(yyyymm, delta=delta_test)
        df_dms = self.transform_dms(yyyymm, delta=delta_test)
        df_ovs = df_ovs[['부품번호', '해외_합계']]
        df_dms = df_dms[['부품번호', '국내_합계']]
        df_ovs['품목'] = df_ovs["부품번호"].str.slice(0, 3)
        df_dms['품목'] = df_dms["부품번호"].str.slice(0, 3)

        df_ovs.rename(columns={'해외_합계': '합계'}, inplace=True)
        df_dms.rename(columns={'국내_합계': '합계'}, inplace=True)
        df = pd.concat([df_ovs, df_dms], axis=0)

        df = df.pivot_table(values=['합계'], index='품목', aggfunc=np.sum)
        df.reset_index(inplace=True)

        return dict(zip(df['품목'], round(df['합계'],1)))

    def transform_ovs(self, yyyymm, delta=1):
        self.raw_obj.yyyymm = yyyymm
        df_ovs = self.raw_obj.get_yearly_raw_ovs(delta=delta)
        df_ovs.rename(columns={'품번': '부품번호', '코드.1': '업체코드', '귀책처': '업체명'}, inplace=True)
        df_ovs = df_ovs.pivot_table(values=['조치수량_가중치'],
                                    index=self.keys,
                                    columns=['불량구분'],
                                    aggfunc=np.sum)
        df_ovs.reset_index(inplace=True)
        df_ovs.columns = [f"해외_{i[1]}" if i[1] != '' else i[0] for i in df_ovs.columns.tolist()]

        df_ovs.fillna(0, inplace=True)
        df_ovs['해외_합계'] = df_ovs['해외_외관'] + df_ovs['해외_기능'] + df_ovs['해외_치수']

        return df_ovs

    def transform_dms(self, yyyymm, delta=1):
        self.raw_obj.yyyymm = yyyymm
        df_dms = self.raw_obj.get_yearly_raw_dms(delta=delta)
        df_dms.rename(columns={'품번': '부품번호', '코드': '업체코드', '부품협력사': '업체명'}, inplace=True)
        df_dms = df_dms.pivot_table(values=['조치수량_가중치'],
                                    index=self.keys,
                                    columns=['불량구분'],
                                    aggfunc=np.sum)
        df_dms.reset_index(inplace=True)
        df_dms.columns = [f"국내_{i[1]}" if i[1] != '' else i[0] for i in df_dms.columns.tolist()]

        df_dms.fillna(0, inplace=True)
        df_dms['국내_합계'] = df_dms['국내_외관'] + df_dms['국내_기능'] + df_dms['국내_치수']

        return df_dms

    @show_elapsed_time
    def transform_icm(self, df_ovs, df_dms, delta_monthly=12):
        # self.raw_obj.yyyymm = yyyymm
        df_icm = self.raw_obj.get_monthly_part_supply(delta_monthly=delta_monthly)

        df_icm = df_icm.merge(df_ovs, how='left', left_on=self.keys, right_on=self.keys)
        df_icm = df_icm.merge(df_dms, how='left', left_on=self.keys, right_on=self.keys)
        df_icm.fillna(0, inplace=True)
        df_icm['외관_합계'] = df_icm['해외_외관'] + df_icm['국내_외관']
        df_icm['이종_합계'] = df_icm['해외_이종'] + df_icm['국내_이종']
        df_icm['기능_합계'] = df_icm['해외_기능'] + df_icm['국내_기능']
        df_icm['치수_합계'] = df_icm['해외_치수'] + df_icm['국내_치수']
        df_icm['포장_합계'] = df_icm['해외_포장'] + df_icm['국내_포장']
        df_icm['종합_합계'] = df_icm['해외_합계'] + df_icm['국내_합계']
        df_icm.sort_values('종합_합계', ascending=False, inplace=True)

        return df_icm

    @show_elapsed_time
    def part_max_icm(self,yyyymm, df_icm, s_data_range_yearly=1):
        values = ['해외_외관', '해외_이종', '해외_기능', '해외_치수', '해외_포장', '해외_합계', '국내_외관', '국내_이종', '국내_기능',
                  '국내_치수', '국내_포장', '국내_합계', '외관_합계', '이종_합계', '기능_합계', '치수_합계', '포장_합계', '종합_합계']
        df_icm_pmax = df_icm.pivot_table(values=values,
                                         index=self.key_part_category,
                                         aggfunc=np.max)
        df_icm_pmax.reset_index(inplace=True)
        df_icm_pmax.sort_values('종합_합계', ascending=False, inplace=True)

        zero_defect_dict = self.get_zero_defect_dict(yyyymm, delta_test=s_data_range_yearly)
        zero_label = f'무불량_({s_data_range_yearly}년)'
        df_icm_pmax[zero_label] = df_icm_pmax['품목'].map(lambda x: zero_defect_dict.get(x, 0))
        df_icm_pmax = self.get_marks(df_icm_pmax)

        df_icm_pmax['품목등급'] = ['S' if df_icm_pmax[zero_label].iloc[n] == 0 else i for n, i in enumerate(df_icm_pmax['품목등급'])]

        col = self.key_part_category + values + ['AB순위', '품목등급', '포장등급', '이종등급', zero_label]

        # print(df_icm_pmax[col]['고객사'].unique())

        return df_icm_pmax[col]

    @show_elapsed_time
    def combine_icm(self, df_icm, df_icm_pmax):
        df_icm_pmax = df_icm_pmax[self.key_part_category + ['품목등급', '포장등급', '이종등급']]
        df_icm_combined = df_icm.merge(df_icm_pmax, how='left', left_on=self.key_part_category, right_on=self.key_part_category)

        return df_icm_combined

    @show_elapsed_time
    def packing_center_review(self, df_icm_combined, delta_month=12):
        df_icm_combined = df_icm_combined[self.keys + ['품목등급', '포장등급', '이종등급']]
        df_icm_raw = self.raw_obj.get_monthly_part_supply_untouched(delta_month=delta_month)

        df_icm_raw = df_icm_raw.merge(df_icm_combined, how='left', left_on=self.keys, right_on=self.keys)
        df_pac = df_icm_raw.pivot_table(values=['입고수량'],
                                        index=['포장장코드'],
                                        columns=['품목등급', '포장등급', '이종등급'],
                                        aggfunc=np.sum)
        df_pac.reset_index(inplace=True)
        df_pac.columns = ["".join(i) for i in df_pac.columns.tolist()]
        df_pac.set_index('포장장코드', inplace=True)
        df_pac.loc['합계'] = df_pac.sum(numeric_only=True, axis=0)
        df_pac.loc[:, '합계'] = df_pac.sum(numeric_only=True, axis=1)

        return df_pac, df_icm_raw

    @show_elapsed_time
    def get_marks(self, df_icm_pmax):
        df_length = len(df_icm_pmax)
        mark_list = ["" for _ in range(df_length)]

        customer = {}
        customer_points = list(zip(df_icm_pmax['고객사'], df_icm_pmax['종합_합계']))

        for n, customer_point in enumerate(customer_points):
            if customer_point[0] not in CUSTOMERS:
                mark_list[n] = "A_"
                continue

            if customer_point[1] == 0:
                mark_list[n] = "S"
                continue
            elif customer_point[0] in customer.keys():
                if customer[customer_point[0]] < self.n_of_d:
                    mark_list[n] = "D" if mark_list[n] == "" else mark_list[n]

                if self.n_of_d <= customer[customer_point[0]] < self.n_of_d + self.n_of_c:
                    mark_list[n] = "C" if mark_list[n] == "" else mark_list[n]
                    customer[customer_point[0]] += 1
                else:
                    mark_list[n] = "A" if mark_list[n] == "" else mark_list[n]
            else:
                customer[customer_point[0]] = 1
                mark_list[n] = "D" if mark_list[n] == "" else mark_list[n]

        df_icm_pmax['품목등급'] = mark_list
        df_icm_pmax.sort_values('품목등급', ascending=True, inplace=True)
        mark_list = df_icm_pmax['품목등급'].tolist()
        none_target_length = len(df_icm_pmax[df_icm_pmax['품목등급'] != 'A'])
        df_icm_pmax['AB순위'] = df_icm_pmax[df_icm_pmax['품목등급'] == 'A']['종합_합계'].rank(pct=True, method='min').tolist() + [
            0 for _ in range(none_target_length)]

        for n, rank in enumerate(df_icm_pmax['AB순위'].tolist()):
            if rank >= 1 - self.b_rank_criterion:
                mark_list[n] = 'B'
        df_icm_pmax['품목등급'] = [i.replace('_', "") for i in mark_list]

        df_icm_pmax.sort_values('포장_합계', ascending=False, inplace=True)
        p_list = ['P' for _ in range(self.n_of_p)] + ["" for _ in range(df_length - self.n_of_p)]
        df_icm_pmax['포장등급'] = p_list

        df_icm_pmax.sort_values('이종_합계', ascending=False, inplace=True)
        m_list = ['MIX' for _ in range(self.n_of_m)] + ["" for _ in range(df_length - self.n_of_m)]
        df_icm_pmax['이종등급'] = m_list

        df_icm_pmax.sort_values('종합_합계', ascending=False, inplace=True)

        # print(df_icm_pmax['고객사'].unique())

        return df_icm_pmax



