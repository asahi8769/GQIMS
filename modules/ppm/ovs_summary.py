import pandas as pd
from tqdm import tqdm
from utils_.config import CUSTOMERS


class OverSeasSummary:

    def __init__(self, yyyymm: int, df):
        self.df_input = df.copy()
        self.yyyymm = yyyymm
        self.month_range = [i for i in range(int(str(self.yyyymm)[:4] + '01'), int(str(self.yyyymm)[:4] + '13'), 1) if i <= self.yyyymm] + ['종합']
        self.remaining_months = [i for i in range(int(str(self.yyyymm)[:4] + '01'), int(str(self.yyyymm)[:4] + '13'), 1) if i > self.yyyymm]
        self.customers = CUSTOMERS
        self.df = None

    def get_summary(self):

        self.generate_indicies()

        for month in tqdm(self.month_range):
            slot = ["" for _ in range(6)]
            slot[0] = self.df_input.loc[('공급망 품질', '종합', '전사', 'PPM'), month]
            slot[1] = self.df_input.loc[('공급망 외관품질', '종합', '전사', 'PPM'), month]
            slot[2] = self.df_input.loc[('공급망 외관품질', '종합', '한국', 'PPM'), month]
            slot[3] = self.df_input.loc[('공급망 외관품질', '종합', '글로벌소싱(중국/인도)', 'PPM'), month]
            slot[4] = self.df_input.loc[('공급망 외관품질', '종합', '중국', 'PPM'), month]
            slot[5] = self.df_input.loc[('공급망 외관품질', '종합', '인도', 'PPM'), month]

            for customer in self.customers:
                slot.append(self.df_input.loc[(customer, '공급망 품질(종합)', '종합', 'PPM'), month])
                slot.append(self.df_input.loc[(customer, '공급망 품질(외관)', '종합', 'PPM'), month])

            self.df[month] = slot

        for month in self.remaining_months:
            self.df[month] = ""

        self.df.set_index(['구분1', '구분2', '구분3'], drop=True, inplace=True)
        self.df = self.df[['종합'] + [i for i in self.month_range if i != "종합"] + self.remaining_months]

        return self.df

    def generate_indicies(self):
        first_col = []
        second_col = []
        third_col = []

        first_col += ["종합" for _ in range(6)]
        second_col += ['공급망품질', '외관품질', '외관(한국)', '외관(글로벌)', '외관(중국)', '외관(인도)']
        third_col += ['PPM' for _ in range(6)]

        for customer in self.customers:
            first_col += [customer for _ in range(2)]
            second_col += ['공급망품질', '외관품질']
            third_col += ['PPM' for _ in range(2)]

        self.df = pd.DataFrame({'구분1': first_col, '구분2': second_col, '구분3': third_col})


if __name__ =="__main__":
    import os
    import openpyxl
    os.chdir(os.pardir)
    os.chdir(os.pardir)

    df = pd.read_excel("spawn/(202109)KD해외품질지수(GQMS기준)_작성일시_ 20211025_22_52_55.xlsx",
                       sheet_name='종합지수(GQMS+공급망)', index_col=[1,2,3,4], skiprows=3)

    obj = OverSeasSummary(202109, df=df)
    df = obj.get_summary()

    print(len(df.index[0]))
    print(len(df.columns))
    print(len(df.loc[('종합', )]))


    # writer = pd.ExcelWriter(r'spawn\test6.xlsx', engine='xlsxwriter')
    # workbook = writer.book
    # sheet_name = "요약"

    # df.to_excel(writer, sheet_name=sheet_name, startrow=2, startcol=0, index=True)
    # writer.save()

    # xfile = openpyxl.load_workbook(r'spawn\test6.xlsx')
    # # sheet = xfile.get_sheet_by_name(sheet_name)
    # sheet = xfile[sheet_name]
    # amount_width = 10
    #
    # sheet.column_dimensions["A"].width = 8
    # sheet.column_dimensions["B"].width = 12
    # sheet.column_dimensions["C"].width = 9
    # sheet.column_dimensions["D"].width = amount_width
    # sheet.column_dimensions["E"].width = amount_width
    # sheet.column_dimensions["F"].width = amount_width
    # sheet.column_dimensions["G"].width = amount_width
    # sheet.column_dimensions["H"].width = amount_width
    # sheet.column_dimensions["I"].width = amount_width
    # sheet.column_dimensions["J"].width = amount_width
    # sheet.column_dimensions["K"].width = amount_width
    # sheet.column_dimensions["L"].width = amount_width
    # sheet.column_dimensions["M"].width = amount_width
    # sheet.column_dimensions["N"].width = amount_width
    # sheet.column_dimensions["O"].width = amount_width
    # sheet.column_dimensions["P"].width = amount_width
    #
    # xfile.save(r'spawn\test6.xlsx')
    os.startfile(r'spawn\test6.xlsx')

