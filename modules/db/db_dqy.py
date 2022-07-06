import sqlite3
import pandas as pd
import warnings
from modules.db.db_domestic import DomesticDataBase
from datetime import datetime
from utils_.config import PPR_SUPPLIED_FROM
warnings.filterwarnings('ignore')


class DisbursementInfo(DomesticDataBase):

    def __init__(self):
        super().__init__()
        self.table_name = f"불출수량현황"

    def create(self):
        with sqlite3.connect(self.db) as conn:
            cur = conn.cursor()
            cur.execute(f'''CREATE TABLE "{self.table_name}" (
                            "조회년월"	INTEGER,
                         	"공급법인"	TEXT,
                         	"대표고객"	TEXT,
                         	"차종"	TEXT,
                         	"코드"	TEXT,
                         	"부품협력사"	TEXT,
                         	"품번"	TEXT,
                         	"품명"	TEXT,
                         	"포장장"	TEXT,
                         	"불출수량"	INTEGER,
                         	"국가코드"	TEXT,
                         	"ID"	INTEGER,
                         	"timestamp"	INTEGER,
                            "입력일자"	TEXT,
                         	PRIMARY KEY("timestamp","ID")
                         );'''
            )
            conn.commit()

    def preprocessing(self, yyyymm):

        df_base = None
        for corp in PPR_SUPPLIED_FROM:

            file = f"downloaded/{self.table_name}_{yyyymm}_{corp}.xlsx"

            with open(file, 'rb') as f:
                df = pd.read_excel(f, converters={'품번': lambda x: str(x)}, engine='openpyxl')

            if df_base is None:
                df_base = df
            else :
                df_base = pd.concat([df_base, df], ignore_index=True)

        df_base['ID'] = [i + 1 for i in range(len(df_base))]
        df_base['품번'] = [str(i).replace("-", "") for i in df_base['품번']]
        df_base['timestamp'] = [yyyymm for _ in range(len(df_base))]
        df_base['품명'] = [str(i).upper() for i in df_base['품명'].tolist()]
        df_base['입력일자'] = datetime.now().strftime("%Y%m%d")
        df_base['대표고객'] = [str(i).upper() for i in df_base['대표고객']]

        return df_base


if __name__ == "__main__":
    import os
    from modules.gqims import GqimsDataRetrieval

    os.chdir(os.pardir)
    os.chdir(os.pardir)

    yyyymm = 202203

    #########################GQIMS DOWNLOAD###############################
    # gqims = GqimsDataRetrieval()
    # gqims.login()
    # gqims.set_dqy_screen()
    # # gqims.download_multiple_table(yyyymm, gqims.download_ovs_table)
    # gqims.download_dqy_table(yyyymm)
    # gqims.close()
    #########################GQIMS DOWNLOAD###############################

    dqy_defect = DisbursementInfo()

    # dqy_defect.append(yyyymm=yyyymm, if_exist='overwrite')
    df = dqy_defect.search(yyyymm, from_to=True)
    print(df['timestamp'].unique())

    df.to_excel(r"spawn\KD불출수량_1.xlsx", index=False)
    os.startfile(r"spawn\KD불출수량_1.xlsx")