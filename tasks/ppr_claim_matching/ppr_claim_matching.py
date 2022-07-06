from modules.db.db_overseas import OverseasDataBase
import pandas as pd
import math
from utils_.config import CUSTOMER_NAME_REVISION


def get_ovs(yyyymm, from_to=True, after=True):
    overseas = OverseasDataBase(pre_noti=True)
    ovs_df = overseas.search(yyyymm, from_to=from_to, after=after)
    ovs_df['고객사'] = [str(i).replace("(CN)", "") for i in ovs_df['고객사']]
    ovs_df.rename(columns={'코드.1': '귀책처코드', "timestamp": "등록월", "Monetary": "대금변제수량"}, inplace=True)
    col_to_keep = ["등록월", "진행상태", "공급법인", "KDQR_Type", "PPR_No", "KDQR_No", "담당자", "고객사", "차종", "귀책처코드", "귀책처",
                   "품번", "품명", "클레임수량", "대금변제수량"]
    # "진행상태", "대표PPR_No", "ID"
    return ovs_df[col_to_keep]


def get_ics():
    df = pd.read_excel("tasks\ppr_claim_matching\인라인클레임(3년치).xlsx")
    df['품번'] = [str(i).replace("-", "") for i in df['품번']]
    df['대표고객'] = [CUSTOMER_NAME_REVISION.get(i, i) for i in df['대표고객']]
    df = df[df["마감년월"] >= 201701]
    df.rename(columns={'해외품질정보': 'PPR_No_클레임', "차종": "차종_클레임", "코드": "귀책처코드_클레임", "귀책처": "귀책처_클레임",
                       '품번': '품번_클레임', "품명": "품명_클레임", "클레임수량": "클레임수량_클레임", "대표고객": "고객사_클레임",
                       "진행상태": "진행상태_클레임"}, inplace=True)
    # df.sort_values("마감년월")
    col_to_keep = ["No.", "마감년월", "클레임No", "클레임타입", "진행상태_클레임", "PPR_No_클레임", "고객사_클레임", "차종_클레임", "귀책처코드_클레임",
                   "귀책처_클레임", "품번_클레임", "품명_클레임", "클레임수량_클레임"]
    #  "외화금액", "통화", "환율", "원화금액"

    return df[col_to_keep]


class MatchDataFrames:
    def __init__(self):
        self.ovs_df = get_ovs(201301, from_to=True, after=True)
        self.ics_df = get_ics()

    def merge_1st(self):
        df = self.ics_df.merge(self.ovs_df, how='left', left_on=["PPR_No_클레임", "품번_클레임"], right_on=["PPR_No", "품번"])
        df["PPR/품번_매칭"] = ["정상" if not math.isnan(i) else "미확정" for i in df['등록월']]
        df["고객사_매칭"] = ["정상" if i[0] == i[1] else "미확정" for i in zip(df['고객사_클레임'], df['고객사'])]
        df["귀책처_매칭"] = ["정상" if i[0] == i[1] else "미확정" for i in zip(df['귀책처코드_클레임'], df['귀책처코드'])]
        status_col = ["PPR/품번_매칭", "고객사_매칭", "귀책처_매칭"]
        col = status_col + [i for i in df.columns.tolist() if i not in status_col]
        return df[col]

    def merge_2nd(self):
        ovs_trimmed = self.ovs_df[self.ovs_df['등록월'] >= 201701]
        df = ovs_trimmed.merge(self.ics_df, how='left', left_on=["PPR_No", "품번"], right_on=["PPR_No_클레임", "품번_클레임"])
        df["PPR/품번_매칭"] = ["정상" if not math.isnan(i) else "미확정" for i in df['No.']]
        df["고객사_매칭"] = ["정상" if i[0] == i[1] else "미확정" for i in zip(df['고객사_클레임'], df['고객사'])]
        df["귀책처_매칭"] = ["정상" if i[0] == i[1] else "미확정" for i in zip(df['귀책처코드_클레임'], df['귀책처코드'])]
        status_col = ["PPR/품번_매칭", "고객사_매칭", "귀책처_매칭"]
        col = status_col + [i for i in df.columns.tolist() if i not in status_col]
        return df[col]


if __name__ == "__main__":
    import os
    os.chdir(os.pardir)
    os.chdir(os.pardir)

    obj = MatchDataFrames()
    df_1 = obj.merge_1st()
    df_2 = obj.merge_2nd()

    file = r"tasks\ppr_claim_matching\return3.xlsx"

    with pd.ExcelWriter(file, engine='xlsxwriter') as writer:
        df_1.to_excel(writer, sheet_name="인라인클레임기준_PPR없음", startrow=1, index=False)
        df_2.to_excel(writer, sheet_name="해외품질정보기준_클레임없음", startrow=1, index=False)

    os.startfile(file)

