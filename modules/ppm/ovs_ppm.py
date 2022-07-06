import pandas as pd
from modules.db.db_ppr import GqmsPPRInfo
from modules.db.db_dqy import DisbursementInfo
from modules.db.db_overseas import OverseasDataBase
from tqdm import tqdm
from datetime import datetime
from dateutil.relativedelta import relativedelta
from utils_.config import *
import re


class MonthlyOVSppmCalculation:

    def __init__(self, yyyymm:int, fixed_dqy=True, disadvantage=.0):
        self.yyyymm = yyyymm
        self.month_range = [i for i in range(int(str(self.yyyymm)[:4] + '01'), int(str(self.yyyymm)[:4] + '13'), 1) if i <= self.yyyymm] + ['종합']
        self.remaining_months = [i for i in range(int(str(self.yyyymm)[:4] + '01'), int(str(self.yyyymm)[:4] + '13'), 1) if i > self.yyyymm]
        self.customers = CUSTOMERS
        self.dqy_countries= DQY_COUNTRIES
        # self.ovs_allowed_type = OVS_VALID_TYPE #  Use negative criteria.
        # self.gqms_allowed_type = GQMS_VALID_TYPE #  Use negative criteria.
        self.ovs_major_countries = OVS_COUNTRIES
        self.ovs_appe_types = APPEARANCE_TYPE
        self.defect_types = DEFECT_TYPES
        self.commercial_customers = SMALL_CUSTOMERS
        self.customer_name = CUSTOMER_NAME_REVISION
        self.ovs_invalid_type = OVS_INVALID_TYPE
        self.gqms_invalid_type = GQMS_INVALID_TYPE
        self.ovs_functional = OVS_FUNCTIONAL
        self.target = TARGET[str(yyyymm)[:4]]
        self.df = None
        self.download_date = []
        self.now = datetime.now().strftime("%Y-%m-%d_%H:%M:%S")
        self.estimation = False
        self.fixed_dqy = fixed_dqy
        self.previous_month = int((datetime.strptime(str(yyyymm), '%Y%m') - relativedelta(months=1)).strftime('%Y%m'))
        self.disadvantage = 1-disadvantage

    def get_ppm(self):

        self.generate_indices()

        last_dqy = None

        for month in tqdm(self.month_range):
            from_to = False
            col_name = month
            if month == '종합':
                month = self.yyyymm
                from_to = True
                col_name = '종합'

            monthly_ppm = []

            gqms_ppr = GqmsPPRInfo()
            gqms_df = gqms_ppr.search(month, from_to=from_to)
            gqms_df['대표고객'] = gqms_df['대표고객'].replace(self.customer_name)
            gqms_df_commercial = gqms_df[(~gqms_df['조치구분'].isin(self.gqms_invalid_type)) & (gqms_df["대표고객"].isin(self.commercial_customers))]
            gqms_df = gqms_df[(~gqms_df['조치구분'].isin(self.gqms_invalid_type)) & (gqms_df["대표고객"].isin(self.customers))]

            disburse_qty = DisbursementInfo()

            if col_name == '종합':
                if self.fixed_dqy:
                    dqy_df = disburse_qty.search(month, from_to=from_to)
                else:
                    if str(self.yyyymm)[4:] == "01":
                        dqy_df = disburse_qty.search(self.previous_month, from_to=False)
                        dqy_df['불출수량'] = round(dqy_df['불출수량'] * self.disadvantage, 0)
                    else:
                        dqy_df = disburse_qty.search(self.previous_month, from_to=from_to)
                        if self.estimation:
                            dqy_df = pd.concat([dqy_df, last_dqy], ignore_index=True)
            else:
                dqy_df = disburse_qty.search(month, from_to=from_to)
                if len(dqy_df) == 0 or (not self.fixed_dqy and col_name == self.yyyymm):
                    self.estimation = True
                    dqy_df = disburse_qty.search(self.previous_month, from_to=from_to)
                    dqy_df['불출수량'] = round(dqy_df['불출수량'] * self.disadvantage, 0)
                    last_dqy = dqy_df

            dqy_df_commercial = dqy_df[(dqy_df["대표고객"].isin(self.commercial_customers))]
            dqy_df = dqy_df[(dqy_df["대표고객"].isin(self.customers))]

            overseas = OverseasDataBase()
            ovs_df = overseas.search(month, from_to=from_to)
            ovs_df['고객사'] = [str(i).replace("(CN)", "") for i in ovs_df['고객사']]
            ovs_df = ovs_df[(~ovs_df["KDQR_Type"].isin(self.ovs_invalid_type)) & (ovs_df["고객사"].isin(self.customers))]
            ovs_df = ovs_df[(~ovs_df["PPR_No"].isin(PPRS_TO_EXCLUDE))]

            new_date = ovs_df['입력일자'].unique().tolist()
            for date in new_date:
                if date not in self.download_date:
                    self.download_date.append(date)

            for customer in self.customers:
                slots = ["" for _ in range(35)]

                process_qty = sum(gqms_df[(gqms_df["대표고객"] == customer) & (gqms_df["발생유형"] == "공정")]["불량수량"])
                sorting_qty = sum(gqms_df[(gqms_df["대표고객"] == customer) & (gqms_df["발생유형"] == "선별")]["불량수량"])
                total_qty = process_qty + sorting_qty
                assert total_qty == sum(gqms_df[(gqms_df["대표고객"] == customer)]['불량수량'])

                dqy_kor = sum(dqy_df[(dqy_df["대표고객"] == customer) & (dqy_df["국가코드"] == "Korea")]["불출수량"])
                dqy_chn = sum(dqy_df[(dqy_df["대표고객"] == customer) & (dqy_df["국가코드"] == "China")]["불출수량"])
                dqy_ind = sum(dqy_df[(dqy_df["대표고객"] == customer) & (dqy_df["국가코드"] == "India")]["불출수량"])
                dqy_ect = sum(dqy_df[(dqy_df["대표고객"] == customer) & (~dqy_df["국가코드"].isin(self.dqy_countries))]["불출수량"])

                total_dqy = dqy_kor + dqy_chn + dqy_ind + dqy_ect

                ppm_process = round(process_qty / total_dqy * 1000000, 1) if total_dqy != 0 else 0
                ppm_sorting = round(sorting_qty / total_dqy * 1000000, 1) if total_dqy != 0 else 0
                ppm_total = round(total_qty / total_dqy * 1000000, 1) if total_dqy != 0 else 0

                overseas_kor_qty = sum(ovs_df[(ovs_df["고객사"] == customer) & (ovs_df["공급국가"] == 'KR')]["조치수량"])
                overseas_chn_qty = sum(ovs_df[(ovs_df["고객사"] == customer) & (ovs_df["공급국가"] == 'CN')]["조치수량"])
                overseas_ind_qty = sum(ovs_df[(ovs_df["고객사"] == customer) & (ovs_df["공급국가"] == 'IN')]["조치수량"])
                overseas_ect_qty = sum(ovs_df[(ovs_df["고객사"] == customer) & (~ovs_df["공급국가"].isin(self.ovs_major_countries))]["조치수량"])

                total_overseas_qty = overseas_kor_qty + overseas_chn_qty + overseas_ind_qty + overseas_ect_qty
                assert total_overseas_qty == sum(ovs_df[(ovs_df["고객사"] == customer)]["조치수량"])

                total_overseas_ppm = round(total_overseas_qty / total_dqy * 1000000, 1) if total_dqy != 0 else 0
                kor_overseas_ppm = round(overseas_kor_qty / dqy_kor * 1000000, 1) if dqy_kor != 0 else 0
                chn_overseas_ppm = round(overseas_chn_qty / dqy_chn * 1000000, 1) if dqy_chn != 0 else 0
                ind_overseas_ppm = round(overseas_ind_qty / dqy_ind * 1000000, 1) if dqy_ind != 0 else 0
                ect_overseas_ppm = round(overseas_ect_qty / dqy_ect * 1000000, 1) if dqy_ect != 0 else 0

                overseas_func_qty = sum(ovs_df[(ovs_df["고객사"] == customer) & (ovs_df["불량구분"].isin(self.ovs_functional))]["조치수량"])
                overseas_dime_qty = sum(ovs_df[(ovs_df["고객사"] == customer) & (ovs_df["불량구분"] == '치수')]["조치수량"])

                func_overseas_ppm = round(overseas_func_qty / total_dqy * 1000000, 1) if total_dqy != 0 else 0
                dime_overseas_ppm = round(overseas_dime_qty / total_dqy * 1000000, 1) if total_dqy != 0 else 0

                overseas_appe_kor_qty = sum(ovs_df[(ovs_df["고객사"] == customer) & (ovs_df["공급국가"] == 'KR') & (ovs_df["불량구분"].isin(self.ovs_appe_types))]["조치수량"])
                appe_overseas_kor_ppm = round(overseas_appe_kor_qty / dqy_kor * 1000000, 1) if dqy_kor != 0 else 0
                overseas_appe_chn_qty = sum(ovs_df[(ovs_df["고객사"] == customer) & (ovs_df["공급국가"] == 'CN') & (ovs_df["불량구분"].isin(self.ovs_appe_types))]["조치수량"])
                appe_overseas_chn_ppm = round(overseas_appe_chn_qty / dqy_chn * 1000000, 1) if dqy_chn != 0 else 0
                overseas_appe_ind_qty = sum(ovs_df[(ovs_df["고객사"] == customer) & (ovs_df["공급국가"] == 'IN') & (ovs_df["불량구분"].isin(self.ovs_appe_types))]["조치수량"])
                appe_overseas_ind_ppm = round(overseas_appe_ind_qty / dqy_ind * 1000000, 1) if dqy_ind != 0 else 0
                overseas_appe_ect_qty = sum(ovs_df[(ovs_df["고객사"] == customer) & (~ovs_df["공급국가"].isin(self.ovs_major_countries)) & (ovs_df["불량구분"].isin(self.ovs_appe_types))]["조치수량"])
                appe_overseas_ect_ppm = round(overseas_appe_ect_qty / dqy_ect * 1000000, 1) if dqy_ect != 0 else 0

                overseas_appe_total_qty = overseas_appe_kor_qty + overseas_appe_chn_qty + overseas_appe_ind_qty + overseas_appe_ect_qty
                appe_overseas_total_ppm = round(overseas_appe_total_qty / total_dqy * 1000000, 1) if total_dqy != 0 else 0

                slots[0] = ppm_total
                slots[1] = total_qty
                slots[2] = ppm_process
                slots[3] = process_qty
                slots[4] = ppm_sorting
                slots[5] = sorting_qty
                slots[6] = total_overseas_ppm
                slots[7] = total_overseas_qty
                slots[8] = kor_overseas_ppm
                slots[9] = overseas_kor_qty
                slots[10] = chn_overseas_ppm
                slots[11] = overseas_chn_qty
                slots[12] = ind_overseas_ppm
                slots[13] = overseas_ind_qty
                slots[14] = ect_overseas_ppm
                slots[15] = overseas_ect_qty
                slots[16] = func_overseas_ppm
                slots[17] = overseas_func_qty
                slots[16] = func_overseas_ppm
                slots[17] = overseas_func_qty
                slots[18] = dime_overseas_ppm
                slots[19] = overseas_dime_qty
                slots[20] = appe_overseas_total_ppm
                slots[21] = overseas_appe_total_qty
                slots[22] = appe_overseas_kor_ppm
                slots[23] = overseas_appe_kor_qty
                slots[24] = appe_overseas_chn_ppm
                slots[25] = overseas_appe_chn_qty
                slots[26] = appe_overseas_ind_ppm
                slots[27] = overseas_appe_ind_qty
                slots[28] = appe_overseas_ect_ppm
                slots[29] = overseas_appe_ect_qty
                slots[30] = total_dqy
                slots[31] = dqy_kor
                slots[32] = dqy_chn
                slots[33] = dqy_ind
                slots[34] = dqy_ect

                monthly_ppm += slots

            monthly_process_qty = sum(gqms_df[(gqms_df["발생유형"] == "공정")]["불량수량"])
            monthly_sorting_qty = sum(gqms_df[(gqms_df["발생유형"] == "선별")]["불량수량"])
            monthly_total_qty = monthly_process_qty + monthly_sorting_qty
            monthly_dqy_kor = sum(dqy_df[(dqy_df["국가코드"] == "Korea")]["불출수량"])
            monthly_dqy_chn = sum(dqy_df[(dqy_df["국가코드"] == "China")]["불출수량"])
            monthly_dqy_ind = sum(dqy_df[(dqy_df["국가코드"] == "India")]["불출수량"])
            monthly_dqy_ect = sum(dqy_df[(~dqy_df["국가코드"].isin(self.dqy_countries))]["불출수량"])
            monthly_dqy_total = monthly_dqy_kor + monthly_dqy_chn + monthly_dqy_ind + monthly_dqy_ect

            monthly_ppm_process = round(monthly_process_qty / monthly_dqy_total * 1000000, 1) if monthly_dqy_total != 0 else 0
            monthly_ppm_sorting = round(monthly_sorting_qty / monthly_dqy_total * 1000000, 1) if monthly_dqy_total != 0 else 0
            monthly_ppm_total = round(monthly_total_qty / monthly_dqy_total * 1000000, 1) if monthly_dqy_total != 0 else 0

            monthly_ppm.append(monthly_ppm_total)
            monthly_ppm.append(monthly_total_qty)
            monthly_ppm.append(monthly_ppm_process)
            monthly_ppm.append(monthly_process_qty)
            monthly_ppm.append(monthly_ppm_sorting)
            monthly_ppm.append(monthly_sorting_qty)

            overseas_total_monthly_qty = sum(ovs_df["조치수량"])
            overseas_kor_monthly_qty = sum(ovs_df[(ovs_df["공급국가"] == 'KR')]["조치수량"])
            overseas_chn_monthly_qty = sum(ovs_df[(ovs_df["공급국가"] == 'CN')]["조치수량"])
            overseas_ind_monthly_qty = sum(ovs_df[(ovs_df["공급국가"] == 'IN')]["조치수량"])
            overseas_ect_monthly_qty = sum(ovs_df[(~ovs_df["공급국가"].isin(self.ovs_major_countries))]["조치수량"])
            overseas_global_monthly_qty = overseas_chn_monthly_qty + overseas_ind_monthly_qty
            overseas_korect_monthly_qty = overseas_kor_monthly_qty + overseas_ect_monthly_qty

            overseas_total_monthly_ppm = round(overseas_total_monthly_qty / monthly_dqy_total * 1000000, 1) if monthly_dqy_total != 0 else 0
            overseas_kor_monthly_ppm = round(overseas_kor_monthly_qty / monthly_dqy_kor * 1000000, 1) if monthly_dqy_kor != 0 else 0
            overseas_chn_monthly_ppm = round(overseas_chn_monthly_qty / monthly_dqy_chn * 1000000, 1) if monthly_dqy_chn != 0 else 0
            overseas_ind_monthly_ppm = round(overseas_ind_monthly_qty / monthly_dqy_ind * 1000000, 1) if monthly_dqy_ind != 0 else 0
            overseas_ect_monthly_ppm = round(overseas_ect_monthly_qty / monthly_dqy_ect * 1000000, 1) if monthly_dqy_ect != 0 else 0
            overseas_global_monthly_ppm = round(overseas_global_monthly_qty / (monthly_dqy_chn + monthly_dqy_ind) * 1000000, 1) if (monthly_dqy_chn + monthly_dqy_ind) != 0 else 0
            overseas_korect_monthly_ppm = round(overseas_korect_monthly_qty / (monthly_dqy_kor + monthly_dqy_ect) * 1000000, 1) if (monthly_dqy_kor + monthly_dqy_ect) != 0 else 0

            monthly_ppm.append(overseas_total_monthly_ppm)
            monthly_ppm.append(overseas_total_monthly_qty)
            monthly_ppm.append(overseas_kor_monthly_ppm)
            monthly_ppm.append(overseas_kor_monthly_qty)
            monthly_ppm.append(overseas_chn_monthly_ppm)
            monthly_ppm.append(overseas_chn_monthly_qty)
            monthly_ppm.append(overseas_ind_monthly_ppm)
            monthly_ppm.append(overseas_ind_monthly_qty)
            monthly_ppm.append(overseas_global_monthly_ppm)
            monthly_ppm.append(overseas_global_monthly_qty)
            monthly_ppm.append(overseas_ect_monthly_ppm)
            monthly_ppm.append(overseas_ect_monthly_qty)
            monthly_ppm.append(overseas_korect_monthly_ppm)
            monthly_ppm.append(overseas_korect_monthly_qty)

            overseas_total_func_monthly_qty = sum(ovs_df[(ovs_df["불량구분"].isin(self.ovs_functional))]["조치수량"])
            overseas_kor_func_monthly_qty = sum(ovs_df[(ovs_df["불량구분"].isin(self.ovs_functional)) & (ovs_df["공급국가"] == 'KR')]["조치수량"])
            overseas_chn_func_monthly_qty = sum(ovs_df[(ovs_df["불량구분"].isin(self.ovs_functional)) & (ovs_df["공급국가"] == 'CN')]["조치수량"])
            overseas_ind_func_monthly_qty = sum(ovs_df[(ovs_df["불량구분"].isin(self.ovs_functional)) & (ovs_df["공급국가"] == 'IN')]["조치수량"])
            overseas_ect_func_monthly_qty = sum(ovs_df[(ovs_df["불량구분"].isin(self.ovs_functional)) & (~ovs_df["공급국가"].isin(self.ovs_major_countries))]["조치수량"])
            overseas_global_func_monthly_qty = overseas_chn_func_monthly_qty + overseas_ind_func_monthly_qty

            overseas_total_func_monthly_ppm = round(overseas_total_func_monthly_qty / monthly_dqy_total * 1000000, 1) if monthly_dqy_total != 0 else 0
            overseas_kor_func_monthly_ppm = round(overseas_kor_func_monthly_qty / monthly_dqy_kor * 1000000, 1) if monthly_dqy_kor != 0 else 0
            overseas_chn_func_monthly_ppm = round(overseas_chn_func_monthly_qty / monthly_dqy_chn * 1000000, 1) if monthly_dqy_chn != 0 else 0
            overseas_ind_func_monthly_ppm = round(overseas_ind_func_monthly_qty / monthly_dqy_ind * 1000000, 1) if monthly_dqy_ind != 0 else 0
            overseas_ect_func_monthly_ppm = round(overseas_ect_func_monthly_qty / monthly_dqy_ect * 1000000, 1) if monthly_dqy_ect != 0 else 0
            overseas_global_func_monthly_ppm = round(overseas_global_func_monthly_qty / (monthly_dqy_chn + monthly_dqy_ind) * 1000000, 1) if (monthly_dqy_chn + monthly_dqy_ind) != 0 else 0

            monthly_ppm.append(overseas_total_func_monthly_ppm)
            monthly_ppm.append(overseas_total_func_monthly_qty)
            monthly_ppm.append(overseas_kor_func_monthly_ppm)
            monthly_ppm.append(overseas_kor_func_monthly_qty)
            monthly_ppm.append(overseas_chn_func_monthly_ppm)
            monthly_ppm.append(overseas_chn_func_monthly_qty)
            monthly_ppm.append(overseas_ind_func_monthly_ppm)
            monthly_ppm.append(overseas_ind_func_monthly_qty)
            monthly_ppm.append(overseas_global_func_monthly_ppm)
            monthly_ppm.append(overseas_global_func_monthly_qty)
            monthly_ppm.append(overseas_ect_func_monthly_ppm)
            monthly_ppm.append(overseas_ect_func_monthly_qty)

            overseas_total_dime_monthly_qty = sum(ovs_df[(ovs_df["불량구분"] == '치수')]["조치수량"])
            overseas_kor_dime_monthly_qty = sum(ovs_df[(ovs_df["불량구분"] == '치수') & (ovs_df["공급국가"] == 'KR')]["조치수량"])
            overseas_chn_dime_monthly_qty = sum(ovs_df[(ovs_df["불량구분"] == '치수') & (ovs_df["공급국가"] == 'CN')]["조치수량"])
            overseas_ind_dime_monthly_qty = sum(ovs_df[(ovs_df["불량구분"] == '치수') & (ovs_df["공급국가"] == 'IN')]["조치수량"])
            overseas_ect_dime_monthly_qty = sum(ovs_df[(ovs_df["불량구분"] == '치수') & (~ovs_df["공급국가"].isin(self.ovs_major_countries))]["조치수량"])
            overseas_global_dime_monthly_qty = overseas_chn_dime_monthly_qty + overseas_ind_dime_monthly_qty

            overseas_total_dime_monthly_ppm = round(overseas_total_dime_monthly_qty / monthly_dqy_total * 1000000, 1) if monthly_dqy_total != 0 else 0
            overseas_kor_dime_monthly_ppm = round(overseas_kor_dime_monthly_qty / monthly_dqy_kor * 1000000, 1) if monthly_dqy_kor != 0 else 0
            overseas_chn_dime_monthly_ppm = round(overseas_chn_dime_monthly_qty / monthly_dqy_chn * 1000000, 1) if monthly_dqy_chn != 0 else 0
            overseas_ind_dime_monthly_ppm = round(overseas_ind_dime_monthly_qty / monthly_dqy_ind * 1000000, 1) if monthly_dqy_ind != 0 else 0
            overseas_ect_dime_monthly_ppm = round(overseas_ect_dime_monthly_qty / monthly_dqy_ect * 1000000, 1) if monthly_dqy_ect != 0 else 0
            overseas_global_dime_monthly_ppm = round(overseas_global_dime_monthly_qty / (monthly_dqy_chn + monthly_dqy_ind) * 1000000, 1) if (monthly_dqy_chn + monthly_dqy_ind) != 0 else 0

            monthly_ppm.append(overseas_total_dime_monthly_ppm)
            monthly_ppm.append(overseas_total_dime_monthly_qty)
            monthly_ppm.append(overseas_kor_dime_monthly_ppm)
            monthly_ppm.append(overseas_kor_dime_monthly_qty)
            monthly_ppm.append(overseas_chn_dime_monthly_ppm)
            monthly_ppm.append(overseas_chn_dime_monthly_qty)
            monthly_ppm.append(overseas_ind_dime_monthly_ppm)
            monthly_ppm.append(overseas_ind_dime_monthly_qty)
            monthly_ppm.append(overseas_global_dime_monthly_ppm)
            monthly_ppm.append(overseas_global_dime_monthly_qty)
            monthly_ppm.append(overseas_ect_dime_monthly_ppm)
            monthly_ppm.append(overseas_ect_dime_monthly_qty)

            overseas_total_appe_monthly_qty = sum(ovs_df[(ovs_df["불량구분"].isin(self.ovs_appe_types))]["조치수량"])
            overseas_kor_appe_monthly_qty = sum(ovs_df[(ovs_df["불량구분"].isin(self.ovs_appe_types)) & (ovs_df["공급국가"] == 'KR')]["조치수량"])
            overseas_chn_appe_monthly_qty = sum(ovs_df[(ovs_df["불량구분"].isin(self.ovs_appe_types)) & (ovs_df["공급국가"] == 'CN')]["조치수량"])
            overseas_ind_appe_monthly_qty = sum(ovs_df[(ovs_df["불량구분"].isin(self.ovs_appe_types)) & (ovs_df["공급국가"] == 'IN')]["조치수량"])
            overseas_global_appe_monthly_qty = overseas_chn_appe_monthly_qty + overseas_ind_appe_monthly_qty
            overseas_ect_appe_monthly_qty = sum(ovs_df[(ovs_df["불량구분"].isin(self.ovs_appe_types)) & (~ovs_df["공급국가"].isin(self.ovs_major_countries))]["조치수량"])
            overseas_ect_appe_monthly_qty_incl =  overseas_global_appe_monthly_qty + overseas_ect_appe_monthly_qty
            overseas_korect_appe_monthly_qty = overseas_kor_appe_monthly_qty + overseas_ect_appe_monthly_qty

            overseas_total_appe_monthly_ppm = round(overseas_total_appe_monthly_qty / monthly_dqy_total * 1000000, 1) if monthly_dqy_total != 0 else 0
            overseas_kor_appe_monthly_ppm = round(overseas_kor_appe_monthly_qty / monthly_dqy_kor * 1000000, 1) if monthly_dqy_kor != 0 else 0
            overseas_chn_appe_monthly_ppm = round(overseas_chn_appe_monthly_qty / monthly_dqy_chn * 1000000, 1) if monthly_dqy_chn != 0 else 0
            overseas_ind_appe_monthly_ppm = round(overseas_ind_appe_monthly_qty / monthly_dqy_ind * 1000000, 1) if monthly_dqy_ind != 0 else 0
            overseas_ect_appe_monthly_ppm = round(overseas_ect_appe_monthly_qty / monthly_dqy_ect * 1000000, 1) if monthly_dqy_ect != 0 else 0
            overseas_global_appe_monthly_ppm = round(overseas_global_appe_monthly_qty / (monthly_dqy_chn + monthly_dqy_ind) * 1000000, 1) if (monthly_dqy_chn + monthly_dqy_ind) != 0 else 0
            overseas_global_appe_monthly_ppm_incl = round(overseas_ect_appe_monthly_qty_incl / (monthly_dqy_chn + monthly_dqy_ind + monthly_dqy_ect) * 1000000, 1) if (monthly_dqy_chn + monthly_dqy_ind + monthly_dqy_ect) != 0 else 0
            overseas_korect_appe_monthly_ppm = round(overseas_korect_appe_monthly_qty / (monthly_dqy_kor + monthly_dqy_ect) * 1000000, 1) if (monthly_dqy_kor + monthly_dqy_ect) != 0 else 0

            total_appe_score = (round (((self.target - overseas_total_appe_monthly_ppm)/self.target + 1) * 100, 1))
            kor_appe_score = (round(((self.target - overseas_kor_appe_monthly_ppm) / self.target + 1) * 100, 1))
            chn_appe_score = (round(((self.target - overseas_chn_appe_monthly_ppm) / self.target + 1) * 100, 1))
            ind_appe_score = (round(((self.target - overseas_ind_appe_monthly_ppm) / self.target + 1) * 100, 1))
            global_appe_score = (round(((self.target - overseas_global_appe_monthly_ppm) / self.target + 1) * 100, 1))
            global_incl_appe_score = (round(((self.target - overseas_global_appe_monthly_ppm_incl) / self.target + 1) * 100, 1))

            monthly_ppm.append(overseas_total_appe_monthly_ppm)
            monthly_ppm.append(overseas_total_appe_monthly_qty)
            monthly_ppm.append(total_appe_score)

            monthly_ppm.append(overseas_kor_appe_monthly_ppm)
            monthly_ppm.append(overseas_kor_appe_monthly_qty)
            monthly_ppm.append(kor_appe_score)

            monthly_ppm.append(overseas_chn_appe_monthly_ppm)
            monthly_ppm.append(overseas_chn_appe_monthly_qty)
            monthly_ppm.append(chn_appe_score)

            monthly_ppm.append(overseas_ind_appe_monthly_ppm)
            monthly_ppm.append(overseas_ind_appe_monthly_qty)
            monthly_ppm.append(ind_appe_score)

            monthly_ppm.append(overseas_global_appe_monthly_ppm)
            monthly_ppm.append(overseas_global_appe_monthly_qty)
            monthly_ppm.append(global_appe_score)

            monthly_ppm.append(overseas_ect_appe_monthly_ppm)
            monthly_ppm.append(overseas_ect_appe_monthly_qty)

            monthly_ppm.append(overseas_global_appe_monthly_ppm_incl)
            monthly_ppm.append(overseas_ect_appe_monthly_qty_incl)
            monthly_ppm.append(global_incl_appe_score)

            monthly_ppm.append(overseas_korect_appe_monthly_ppm)
            monthly_ppm.append(overseas_korect_appe_monthly_qty)

            monthly_ppm.append(monthly_dqy_total)
            monthly_ppm.append(monthly_dqy_kor)
            monthly_ppm.append(monthly_dqy_chn)
            monthly_ppm.append(monthly_dqy_ind)
            monthly_ppm.append(monthly_dqy_ect)

            for customer in self.commercial_customers:
                slots = ["" for _ in range(10)]

                process_qty = sum(gqms_df_commercial[(gqms_df_commercial["대표고객"] == customer) & (gqms_df_commercial["발생유형"] == "공정")]["불량수량"])
                sorting_qty = sum(gqms_df_commercial[(gqms_df_commercial["대표고객"] == customer) & (gqms_df_commercial["발생유형"] == "선별")]["불량수량"])
                total_qty = process_qty + sorting_qty

                dqy_kor = sum(dqy_df_commercial[(dqy_df_commercial["대표고객"] == customer) & (dqy_df_commercial["국가코드"] == "Korea")]["불출수량"])
                dqy_chn = sum(dqy_df_commercial[(dqy_df_commercial["대표고객"] == customer) & (dqy_df_commercial["국가코드"] == "China")]["불출수량"])
                dqy_ect = sum(dqy_df_commercial[(dqy_df_commercial["대표고객"] == customer) & (~dqy_df_commercial["국가코드"].isin(self.dqy_countries))]["불출수량"])
                total_dqy = dqy_kor + dqy_chn + dqy_ect

                ppm_process = round(process_qty / total_dqy * 1000000, 1) if total_dqy != 0 else 0
                ppm_sorting = round(sorting_qty / total_dqy * 1000000, 1) if total_dqy != 0 else 0
                ppm_total = round(total_qty / total_dqy * 1000000, 1) if total_dqy != 0 else 0

                slots[0] = ppm_total
                slots[1] = total_qty
                slots[2] = ppm_process
                slots[3] = process_qty
                slots[4] = ppm_sorting
                slots[5] = sorting_qty
                slots[6] = total_dqy
                slots[7] = dqy_kor
                slots[8] = dqy_chn
                slots[9] = dqy_ect

                monthly_ppm += slots

            self.df[col_name] = monthly_ppm

        for month in self.remaining_months:
            self.df[month] = ""

        self.df.set_index(['구분1', '구분2', '구분3', '구분4'], drop=True, inplace=True)
        self.df = self.df[['종합'] + [i for i in self.month_range if i != "종합"] + self.remaining_months]
        self.download_date = [datetime.strptime(i, "%Y%m%d").strftime("%Y-%m-%d") for i in self.download_date]

        self.download_date = ', '.join(self.download_date)

        return self.df

    def generate_indices(self):

        first_col = []
        second_col = []
        third_col = []
        forth_col = []

        for customer in self.customers:
            first_col += [customer for _ in range(35)]
            for type in self.defect_types:
                second_col += [type for _ in range(2)]
                third_col += ["계" for _ in range(2)]
            second_col += ["공급망 품질(종합)" for _ in range(10)]
            second_col += ["공급망 품질(기능)" for _ in range(2)]
            second_col += ["공급망 품질(치수)" for _ in range(2)]
            second_col += ["공급망 품질(외관)" for _ in range(10)]
            second_col += ["불출수량" for _ in range(5)]
            for note in ["종합", "한국", "중국", '인도', '기타', '기능', '치수', '종합', '한국', '중국', '인도', '기타']:
                third_col += [note for _ in range(2)]
            third_col += ["계", "한국", "중국", "인도", "기타"]
            forth_col += ['PPM', '불량수량'] * 15
            forth_col += ["불출수량" for _ in range(5)]

        for type in self.defect_types:
            first_col += [type for _ in range(2)]
            second_col += ["계" for _ in range(2)]
            third_col += ["계" for _ in range(2)]
            forth_col += ['PPM', '불량수량']

        first_col += ["공급망 품질" for _ in range(14)]
        second_col += ["종합" for _ in range(14)]
        for note in ["전사", "한국", "중국", '인도', '글로벌소싱(중국/인도)', '기타', '한국/기타']:
            third_col += [note for _ in range(2)]
            forth_col += ['PPM', '불량수량']

        first_col += ["공급망 품질(기능)" for _ in range(12)]
        second_col += ["종합" for _ in range(12)]
        for note in ["전사", "한국", "중국", '인도', '글로벌소싱(중국/인도)', '기타']:
            third_col += [note for _ in range(2)]
            forth_col += ['PPM', '불량수량']

        first_col += ["공급망 품질(치수)" for _ in range(12)]
        second_col += ["종합" for _ in range(12)]
        for note in ["전사", "한국", "중국", '인도', '글로벌소싱(중국/인도)', '기타']:
            third_col += [note for _ in range(2)]
            forth_col += ['PPM', '불량수량']

        first_col += ["공급망 외관품질" for _ in range(22)]
        second_col += ["종합" for _ in range(22)]
        for note in ["전사", "한국", "중국", '인도', '글로벌소싱(중국/인도)']:
            third_col += [note for _ in range(3)]
            forth_col += ['PPM', '불량수량', '달성률']
        third_col += ['기타' for _ in range(2)]
        third_col += ['글로벌소싱(중국/인도/외)' for _ in range(3)]
        third_col += ['한국/기타' for _ in range(2)]
        forth_col += ['PPM', '불량수량', 'PPM', '불량수량', '달성률', 'PPM', '불량수량']

        first_col += ["불출수량" for _ in range(5)]
        for country in ["계", "한국", "중국", "인도", "기타"]:
            second_col += [country for _ in range(1)]
            third_col += ["계" for _ in range(1)]
            forth_col += ["불출수량" for _ in range(1)]

        for customer in self.commercial_customers:
            first_col += [customer for _ in range(10)]
            second_col += ["종합" for _ in range(2)]
            second_col += ["공정" for _ in range(2)]
            second_col += ["선별" for _ in range(2)]
            second_col += ["불출수량" for _ in range(4)]
            third_col += ["계" for _ in range(6)]
            third_col += ["계", "한국", "중국", "기타"]
            forth_col += ["PPM", "불량수량"] * 3
            forth_col += ["불출수량"] * 4

        self.df = pd.DataFrame({'구분1': first_col, '구분2': second_col, '구분3': third_col, '구분4': forth_col})


if __name__ == "__main__":
    import os
    import openpyxl
    os.chdir(os.pardir)
    os.chdir(os.pardir)

    ppm = MonthlyOVSppmCalculation(202112, fixed_dqy=True, disadvantage=0.1)
    df = ppm.get_ppm()
    # df.set_index(['구분1', '구분2', '구분3', '구분4'], drop=True, inplace=True)

    writer = pd.ExcelWriter(r'spawn\test.xlsx', engine='xlsxwriter')
    workbook = writer.book
    sheet_name = "종합지수(GQMS+공급망)"

    df.to_excel(writer, sheet_name=sheet_name, startrow=3, startcol=0, index=True)
    writer.save()

    xfile = openpyxl.load_workbook(r'spawn\test.xlsx')
    sheet = xfile.get_sheet_by_name(sheet_name)

    sheet['A1'] = ppm.download_date
    sheet['A2'] = ppm.now

    sheet.column_dimensions["A"].width = 16.88
    sheet.column_dimensions["B"].width = 16.88
    sheet.column_dimensions["C"].width = 24.38
    sheet.column_dimensions["D"].width = 8.63
    sheet.column_dimensions["E"].width = 9.88
    sheet.column_dimensions["F"].width = 9.88
    sheet.column_dimensions["G"].width = 9.88
    sheet.column_dimensions["H"].width = 9.88
    sheet.column_dimensions["I"].width = 9.88
    sheet.column_dimensions["J"].width = 9.88
    sheet.column_dimensions["K"].width = 9.88
    sheet.column_dimensions["L"].width = 9.88
    sheet.column_dimensions["M"].width = 9.88
    sheet.column_dimensions["N"].width = 9.88
    sheet.column_dimensions["O"].width = 9.88
    sheet.column_dimensions["P"].width = 9.88
    sheet.column_dimensions["Q"].width = 9.88
    sheet.column_dimensions["R"].width = 9.88

    xfile.save(r'spawn\test.xlsx')
    os.startfile(r'spawn\test.xlsx')