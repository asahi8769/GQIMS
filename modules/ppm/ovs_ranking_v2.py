import pandas as pd
from modules.db.db_overseas import OverseasDataBase
from tqdm import tqdm
from utils_.config import CUSTOMERS, APPEARANCE_TYPE, OVS_INVALID_TYPE, GLOBAL_CODES, PART_REG_TO_SUBSTITUTE, \
    VENDOR_REG_TO_SUBSTITUTE, PPRS_TO_EXCLUDE
import re


def overseas_rank(month: int):

    overseas = OverseasDataBase()
    customers = CUSTOMERS
    ovs_invalid_type = OVS_INVALID_TYPE
    appearance_type = APPEARANCE_TYPE
    global_codes = GLOBAL_CODES

    df_dict = {}

    for i in tqdm(['누적', '당월']):

        if i == '누적':
            from_to = True
        else :
            from_to = False
        ovs_df = overseas.search(month, from_to=from_to)
        ovs_df['고객사'] = [str(i).replace("(CN)", "") for i in ovs_df['고객사']]
        ovs_df = ovs_df[(~ovs_df["KDQR_Type"].isin(ovs_invalid_type)) & (ovs_df["고객사"].isin(customers))]
        ovs_df['품명'] = [re.sub(PART_REG_TO_SUBSTITUTE, "", i) for i in ovs_df['품명']]
        ovs_df['품명'] = [i.replace("- ", "-").replace("  ", " ").rstrip(" ").rstrip(",").rstrip("-") for i in ovs_df['품명']]
        ovs_df['부품협력사'] = [re.sub(VENDOR_REG_TO_SUBSTITUTE, "", i.upper()) for i in ovs_df['부품협력사']]
        ovs_df = ovs_df[(~ovs_df["PPR_No"].isin(PPRS_TO_EXCLUDE))]

        for j in ['한국', '글로벌']:
            if j == '한국':
                df_country = ovs_df[ovs_df["공급국가"]=='KR']
            else :
                df_country = ovs_df[ovs_df["공급국가"].isin(global_codes)]

            for k in ['외관불량', '전체불량']:

                if k == '외관불량':
                    df_defect = df_country[ovs_df['불량구분'].isin(appearance_type)]
                else :
                    df_defect = df_country

                table_name = f"{i}_{j}_{k}"
                df = df_defect.pivot_table(values=['조치수량'], index=['코드', '부품협력사', '고객사', '품명'], aggfunc=sum)
                df.sort_values(by=['조치수량'], ascending=False, inplace=True)
                df['점유율(%)'] = [int(i) for i in round(df['조치수량'] / df['조치수량'].sum() * 100, 0)]

                df = df.head(10)
                df['순위'] = [i for i in range(1, 11)]
                df.reset_index(inplace=True)

                problem_customer = []
                problem_part = []

                for vendor in df['코드']:
                    customer = df_defect[df_defect['코드'] == vendor]['고객사'].unique().tolist()
                    customer = sorted(customer, key=lambda x: df_defect[(df_defect['코드'] == vendor) & (df_defect['고객사'] == x)]['조치수량'].sum(), reverse=True)
                    problem_customer.append(customer[0])
                    part = df_defect[df_defect['코드'] == vendor]['품명'].unique().tolist()
                    part = sorted(part, key=lambda x: df_defect[(df_defect['코드'] == vendor) & (df_defect['품명'] == x)]['조치수량'].sum(), reverse=True)
                    problem_part.append(part[0])

                # df['주고객사'] = problem_customer
                # df['주요품명'] = problem_part

                df = df[['순위', '코드', '부품협력사', '조치수량', "점유율(%)", '고객사', "품명"]]
                df_dict[table_name] = df

    return df_dict


if __name__ == "__main__":
    import os
    import openpyxl
    from modules.ppm.decorate_ppm import get_col_from_index

    os.chdir(os.pardir)
    os.chdir(os.pardir)

    df_dict = overseas_rank(202111)

    test_file = r'spawn\test4_v2.xlsx'

    writer = pd.ExcelWriter(test_file, engine='xlsxwriter')
    workbook = writer.book
    sheet_name = "요약"

    row_n = 0
    row_anchor = 2
    col_anchor = 1
    for n, i in enumerate(df_dict.keys()):
        ticker = n % 2
        row = row_n * (len(df_dict[i]) + 1) + row_anchor
        col = ticker * (len(df_dict[i].columns) + 1) + col_anchor
        if ticker == 1:
            row_n += 1
            row_anchor += 2
        df_dict[i].to_excel(writer, sheet_name=sheet_name, startrow=row, startcol=col, index=False)

    writer.save()

    xfile = openpyxl.load_workbook(test_file)
    sheet = xfile.get_sheet_by_name(sheet_name)

    row_n = 0
    row_anchor = 2
    col_anchor = 1
    for n, i in enumerate(df_dict.keys()):
        ticker = n % 2
        row = row_n * (len(df_dict[i]) + 1) + row_anchor
        col = ticker * (len(df_dict[i].columns) + 1) + col_anchor
        if ticker == 1:
            row_n += 1
            row_anchor += 2
            sheet[f'{get_col_from_index("B", len(df_dict[i].columns)+1)}{row}'] = i
        else:
            sheet[f'B{row}'] = i

    xfile.save(test_file)
    os.startfile(test_file)