import pandas as pd
from utils_.config import *
from modules.ppm.ovs_ppm import MonthlyOVSppmCalculation
from modules.ppm.dms_ppm import MonthlyDMSppmCalculation
from modules.ppm.ovs_ranking_v2 import overseas_rank
from modules.ppm.dms_summary import DomesticSummary
from modules.db.db_inr import IncomingResultInfo
from modules.db.db_domestic import DomesticDataBase
import re


def spq_per_plant(df):
    apq = []
    fnq_dmq = []
    spq = []

    apq.append(df.loc[('공급망 외관품질', '종합', '전사', 'PPM'), '종합'])
    fnq_dmq.append(df.loc[('공급망 품질(기능)', '종합', '전사', 'PPM'), '종합'] + df.loc[('공급망 품질(치수)', '종합', '전사', 'PPM'), '종합'])
    spq.append(df.loc[('공급망 품질', '종합', '전사', 'PPM'), '종합'])

    for customer in CUSTOMERS:
        apq.append(df.loc[(customer, '공급망 품질(외관)', '종합', 'PPM'), '종합'])
        fnq_qty = df.loc[(customer, '공급망 품질(기능)', '기능', '불량수량'), '종합'] + df.loc[
            (customer, '공급망 품질(치수)', '치수', '불량수량'), '종합']
        fnq_dmq.append(round(fnq_qty / df.loc[(customer, '불출수량', '계', '불출수량'), '종합'] * 1000000, 1))
        spq.append(df.loc[(customer, '공급망 품질(종합)', '종합', 'PPM'), '종합'])

    return pd.DataFrame({"고객사": ['금년 종합'] + CUSTOMERS, "외관불량": apq, "기능/치수": fnq_dmq, "공급망불량": spq})


def calculate_quarters(df, df_):
    quarter_info = {}

    for quarter in QUARTERS.keys():
        quarter_info[quarter] = {"외관불량수량": 0, "기능치수수량": 0, "공급망수량": 0, "불출수량": 0}

    for n, quarter in enumerate(df_['분기']):
        if quarter != "종합":
            if df_['외관불량'].tolist()[n] != "":

                month = int(df_['년월'].tolist()[n])

                aqy = df.loc[('공급망 외관품질', '종합', '전사', '불량수량'), month]
                fdy = df.loc[('공급망 품질(기능)', '종합', '전사', '불량수량'), month] + df.loc[
                    ('공급망 품질(치수)', '종합', '전사', '불량수량'), month]
                spq = df.loc[('공급망 품질', '종합', '전사', '불량수량'), month]
                dqy = df.loc[('불출수량', '계', '계', '불출수량'), month]

                quarter_info[quarter]["외관불량수량"] += aqy
                quarter_info[quarter]["기능치수수량"] += fdy
                quarter_info[quarter]["공급망수량"] += spq
                quarter_info[quarter]["불출수량"] += dqy

            else:
                break

    for quarter in quarter_info.keys():
        dqy = quarter_info[quarter]["불출수량"]
        aqy_ppm = round(quarter_info[quarter]["외관불량수량"] / dqy * 1000000, 1) if dqy != 0 else ""
        fdy_ppm = round(quarter_info[quarter]["기능치수수량"] / dqy * 1000000, 1) if dqy != 0 else ""
        spq_ppm = round(quarter_info[quarter]["공급망수량"] / dqy * 1000000, 1) if dqy != 0 else ""

        quarter_info[quarter]["외관PPM"] = aqy_ppm
        quarter_info[quarter]["기능치수PPM"] = fdy_ppm
        quarter_info[quarter]["공급망PPM"] = spq_ppm

        df_ = df_.append({"년월": "분기별",
                          "외관불량": quarter_info[quarter]["외관PPM"],
                          "기능/치수": quarter_info[quarter]["기능치수PPM"],
                          "공급망불량": quarter_info[quarter]["공급망PPM"],
                          "분기": quarter}, ignore_index=True)

    return df_


def spq_per_month(yyyymm, df):
    months = ["종합"] + list(range(int(str(yyyymm)[:4] + '01'), int(str(yyyymm)[:4] + '13')))
    apq = []
    fnq_dmq = []
    spq = []
    quarters = []

    for month in months:
        if month == "종합":
            quarters.append("종합")
        else:
            for quarter in QUARTERS.keys():
                if str(month)[4:] in QUARTERS[quarter]:
                    quarters.append(quarter)

        apq.append(df.loc[('공급망 외관품질', '종합', '전사', 'PPM'), month])
        fnq_qty = df.loc[('공급망 품질(기능)', '종합', '전사', '불량수량'), month] + df.loc[('공급망 품질(치수)', '종합', '전사', '불량수량'), month]
        try:
            fnq_dmq.append(round(fnq_qty / df.loc[('불출수량', '계', '계', '불출수량'), month] * 1000000, 1))
        except TypeError:
            fnq_dmq.append("")
        spq.append(df.loc[('공급망 품질', '종합', '전사', 'PPM'), month])

    df_result = pd.DataFrame({"년월": months, "분기": quarters, "외관불량": apq, "기능/치수": fnq_dmq, "공급망불량": spq})

    df_with_quarters = calculate_quarters(df, df_result)

    return df_with_quarters


def apq_per_region(df):
    index_ = ["전사", "한국", "글로벌", "중국", "인도"]

    ppm = []
    ppm.append(df.loc[('공급망 외관품질', '종합', '전사', 'PPM'), '종합'])
    ppm.append(df.loc[('공급망 외관품질', '종합', '한국', 'PPM'), '종합'])
    ppm.append(df.loc[('공급망 외관품질', '종합', '글로벌소싱(중국/인도)', 'PPM'), '종합'])
    ppm.append(df.loc[('공급망 외관품질', '종합', '중국', 'PPM'), '종합'])
    ppm.append(df.loc[('공급망 외관품질', '종합', '인도', 'PPM'), '종합'])

    return pd.DataFrame({"구분": index_, "누적PPM": ppm})


def apq_per_month(yyyymm, df):
    months = ["종합"] + list(range(int(str(yyyymm)[:4] + '01'), int(str(yyyymm)[:4] + '13')))
    quarters = []

    for month in months:
        if month == "종합":
            quarters.append("종합")
        else:
            for quarter in QUARTERS.keys():
                if str(month)[4:] in QUARTERS[quarter]:
                    quarters.append(quarter)

    df_ = pd.DataFrame({"년월": months, "분기": quarters})

    for region in ["한국", "중국", "인도"]:
        apq = []
        dqy = []
        ppm = []

        for month in months:
            apq.append(df.loc[('공급망 외관품질', '종합', region, '불량수량'), month])
            dqy.append(df.loc[('불출수량', region, '계', '불출수량'), month])
            ppm.append(df.loc[('공급망 외관품질', '종합', region, 'PPM'), month])
        df_[f"불량수량_{region}"] = apq
        df_[f"불출수량_{region}"] = dqy
        df_[f"PPM_{region}"] = ppm

    quarter_info = {}
    for quarter in QUARTERS.keys():
        quarter_info[quarter] = {"불량수량_한국": 0, "불출수량_한국": 0, "불량수량_중국": 0, "불출수량_중국": 0,
                                 "불량수량_인도": 0, "불출수량_인도": 0}

    for n, quarter in enumerate(df_['분기']):
        if quarter != "종합":
            if df_['불량수량_한국'].tolist()[n] != "":
                quarter_info[quarter]["불량수량_한국"] += df_['불량수량_한국'].tolist()[n]
                quarter_info[quarter]["불출수량_한국"] += df_['불출수량_한국'].tolist()[n]
                quarter_info[quarter]["불량수량_중국"] += df_['불량수량_중국'].tolist()[n]
                quarter_info[quarter]["불출수량_중국"] += df_['불출수량_중국'].tolist()[n]
                quarter_info[quarter]["불량수량_인도"] += df_['불량수량_인도'].tolist()[n]
                quarter_info[quarter]["불출수량_인도"] += df_['불출수량_인도'].tolist()[n]
            else:
                break

    for quarter in quarter_info.keys():
        dqy_kor = quarter_info[quarter]["불출수량_한국"]
        dqy_chn = quarter_info[quarter]["불출수량_중국"]
        dqy_ind = quarter_info[quarter]["불출수량_인도"]

        quarter_info[quarter]["PPM_한국"] = round(quarter_info[quarter]["불량수량_한국"] / dqy_kor * 1000000, 1) if dqy_kor != 0 else ""
        quarter_info[quarter]["PPM_중국"] = round(quarter_info[quarter]["불량수량_중국"] / dqy_chn * 1000000, 1) if dqy_chn != 0 else ""
        quarter_info[quarter]["PPM_인도"] = round(quarter_info[quarter]["불량수량_인도"] / dqy_ind * 1000000, 1) if dqy_ind != 0 else ""

        for account in quarter_info[quarter].keys():
            quarter_info[quarter][account] = "" if quarter_info[quarter][account] == 0 else quarter_info[quarter][account]

        df_ = df_.append({"년월": "분기별",
                          "분기": quarter,
                          "불량수량_한국": quarter_info[quarter]["불량수량_한국"],
                          "불출수량_한국": quarter_info[quarter]["불출수량_한국"],
                          "PPM_한국": quarter_info[quarter]["PPM_한국"],
                          "불량수량_중국": quarter_info[quarter]["불량수량_중국"],
                          "불출수량_중국": quarter_info[quarter]["불출수량_중국"],
                          "PPM_중국": quarter_info[quarter]["PPM_중국"],
                          "불량수량_인도": quarter_info[quarter]["불량수량_인도"],
                          "불출수량_인도": quarter_info[quarter]["불출수량_인도"],
                          "PPM_인도": quarter_info[quarter]["PPM_인도"],
                          }, ignore_index=True)

    return df_


def spq_ranking(yyyymm):
    return overseas_rank(yyyymm)


def dms_ppm_per_month(yyyymm, df):
    months = ["종합"] + list(range(int(str(yyyymm)[:4] + '01'), int(str(yyyymm)[:4] + '13')))
    ppm_dms = []

    for month in months:
        ppm_dms.append(df.loc[('국내입고불량 실적', '전체'), month])

    return pd.DataFrame({"년월" : months, "PPM": ppm_dms})


def dms_joined_ppm_per_month(yyyymm, df):
    months = ["종합"] + list(range(int(str(yyyymm)[:4] + '01'), int(str(yyyymm)[:4] + '13')))
    ppm_joined_dms = []

    for month in months:
        ppm_joined_dms.append(df.loc[('합동Q-키퍼', 'PPM'), month])

    return pd.DataFrame({"년월" : months, "PPM": ppm_joined_dms})


def dms_ratio_per_customer_type(df):
    customers = ["10대공장", "비계열+상용"]
    all_qty = sum(df[(df['고객사'].isin(customers))]['불량수량'])

    ratio = [round(df[(df['고객사'] == customer)]['불량수량'].values[0]/all_qty*100,1) for customer in customers]
    return pd.DataFrame({"고객사": customers, "비율(%)": ratio})


def dms_ratio_per_customer_region(df_piv):
    df_ratio = df_piv[['고객사구분', '불량수량']]
    all_qty = sum(df_ratio['불량수량'])

    df_ratio['비중'] = round(df_ratio['불량수량']/all_qty*100,1)
    df_ratio.sort_values(by="비중", ascending=False, inplace=True)
    return df_ratio


def dms_ppm_per_customer(df):
    df_customer_ppm = df[df['고객사'].isin(['비계열'] + BROADER_CUSTOMERS)][['고객사', 'PPM']]
    df_customer_ppm.sort_values(by="PPM", ascending=False, inplace=True)
    return df_customer_ppm


def dms_by_inspection_mode(df):

    df_concat = None

    df['품명'] = [re.sub(PART_REG_TO_SUBSTITUTE, "", i) for i in df['품명']]
    df['품명'] = [i.replace("- ", "-").replace("  ", " ").rstrip(" ").rstrip(",").rstrip("-") for i in df['품명']]
    df['부품협력사'] = [re.sub(VENDOR_REG_TO_SUBSTITUTE, "", i.upper()) for i in df['부품협력사']]
    df = df[df['공급법인']=="글로비스 한국"]

    for mode in INCOMING_MODE:
        df_mode = df[df['검사기준'] == mode]
        df_mode_pv = df_mode.pivot_table(values=['입고수량', '불량수량'], index=['검사기준', '포장장', '코드', '부품협력사', '고객사', '품명'], aggfunc=sum)
        try:
            df_mode_pv.sort_values(by="불량수량", ascending=False, inplace=True)
        except:
            continue
        df_mode_pv = df_mode_pv[["입고수량", "불량수량"]]
        df_mode_pv = df_mode_pv[:10]
        df_mode_pv.reset_index(inplace=True)

        if df_concat is None:
            df_concat = df_mode_pv
        else:
            df_concat = pd.concat([df_concat, df_mode_pv])

    return df_concat


def dms_by_defect_type(df):
    df['품명'] = [re.sub(PART_REG_TO_SUBSTITUTE, "", i) for i in df['품명']]
    df['품명'] = [i.replace("- ", "-").replace("  ", " ").rstrip(" ").rstrip(",").rstrip("-") for i in df['품명']]

    df = df[(df['고객사'].isin(CUSTOMERS)) & (df['공급법인'] == "글로비스 한국") & (~df['공급법인'].isin(DMS_INVALID_TYPE))]
    df_concat = None

    for type_ in APPEARANCE_TYPE:
        df_type = df[df['불량구분']==type_]
        df_mode_pv = df_type.pivot_table(values='불량수량', index=['불량구분', '코드', '부품협력사', '고객사', '품명'],
                                         aggfunc=sum)
        df_mode_pv.sort_values(by="불량수량", ascending=False, inplace=True)
        df_mode_pv['점유율(%)'] = [int(i) for i in round(df_mode_pv['불량수량'] / df_mode_pv['불량수량'].sum() * 100, 0)]
        df_mode_pv = df_mode_pv[:10]
        df_mode_pv.reset_index(inplace=True)



        if df_concat is None:
            df_concat = df_mode_pv
        else:
            df_concat = pd.concat([df_concat, df_mode_pv])

    # print(df_concat)

    return df_concat


if __name__ == "__main__":
    import os, openpyxl

    os.chdir(os.pardir)
    os.chdir(os.pardir)

    yyyymm = 202112
    fixed_dqy = True

    ppm_ovs = MonthlyOVSppmCalculation(yyyymm, fixed_dqy=fixed_dqy, disadvantage=.0)
    df_ovs = ppm_ovs.get_ppm()

    ppm_dms = MonthlyDMSppmCalculation(yyyymm)
    df_dms = ppm_dms.get_ppm()

    summary_dms = DomesticSummary(yyyymm)
    df_dms_summary, df_dms_summary_piv = summary_dms.get_summary()

    results = IncomingResultInfo()
    results_df = results.search(yyyymm, from_to=True)

    dms_defect = DomesticDataBase()
    dms_defect_df = dms_defect.search(yyyymm, from_to=False)

    df_spq_per_plant = spq_per_plant(df_ovs)
    df_spq_per_month = spq_per_month(yyyymm, df_ovs)

    df_apq_per_region = apq_per_region(df_ovs)
    df_apq_per_month = apq_per_month(yyyymm, df_ovs)
    dict_spq_ranking = spq_ranking(yyyymm)

    df_dms_ppm_per_month = dms_ppm_per_month(yyyymm, df_dms)
    df_joined_ppm_per_month = dms_joined_ppm_per_month(yyyymm, df_dms)
    df_dms_ratio_per_region = dms_ratio_per_customer_type(df_dms_summary)
    df_ratio = dms_ratio_per_customer_region(df_dms_summary_piv)
    df_customer_ppm = dms_ppm_per_customer(df_dms_summary)
    df_dms_by_inspection_mode = dms_by_inspection_mode(results_df)
    df_dms_by_defect_type = dms_by_defect_type(dms_defect_df)

    sheet_name_0 = "해외"
    sheet_name_1 = "국내"
    sheet_name_2 = "업체순위"

    writer =  pd.ExcelWriter(r'spawn\test8.xlsx', engine='xlsxwriter')
    workbook = writer.book

    df_spq_per_plant.to_excel(writer, sheet_name=sheet_name_0, startrow=2, startcol=1, index=False)
    df_spq_per_month.to_excel(writer, sheet_name=sheet_name_0, startrow=2, startcol=6, index=False)
    df_apq_per_region.to_excel(writer, sheet_name=sheet_name_0, startrow=2, startcol=12, index=False)
    df_apq_per_month.to_excel(writer, sheet_name=sheet_name_0, startrow=23, startcol=1, index=False)

    df_dms_ppm_per_month.to_excel(writer, sheet_name=sheet_name_1, startrow=2, startcol=1, index=False)
    df_joined_ppm_per_month.to_excel(writer, sheet_name=sheet_name_1, startrow=2, startcol=4, index=False)
    df_dms_ratio_per_region.to_excel(writer, sheet_name=sheet_name_1, startrow=2, startcol=7, index=False)
    df_ratio.to_excel(writer, sheet_name=sheet_name_1, startrow=2, startcol=10, index=False)
    df_customer_ppm.to_excel(writer, sheet_name=sheet_name_1, startrow=2, startcol=14, index=False)
    df_dms_by_inspection_mode.to_excel(writer, sheet_name=sheet_name_1, startrow=19, startcol=1, index=False)
    df_dms_by_defect_type.to_excel(writer, sheet_name=sheet_name_1, startrow=19, startcol=10, index=False)
    # print(df_dms_by_defect_type)

    row_n = 0
    row_anchor = 2
    col_anchor = 1
    for n, i in enumerate(dict_spq_ranking.keys()):
        ticker = n % 2
        dict_spq_ranking[i]['구분'] = i

        row = row_n * (len(dict_spq_ranking[i]) + 1) + row_anchor
        col = ticker * (len(dict_spq_ranking[i].columns) + 1) + col_anchor
        if ticker == 1:
            row_n += 1
            row_anchor += 1

        dict_spq_ranking[i].to_excel(writer, sheet_name=sheet_name_2, startrow=row, startcol=col, index=False)

    writer.save()

    os.startfile(r'spawn\test8.xlsx')
