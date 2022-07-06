import sqlite3
import pandas as pd
import warnings
from modules.db.db_domestic import DomesticDataBase
from datetime import datetime
from utils_.config import PPR_SUPPLIED_FROM
warnings.filterwarnings('ignore')


class GqmsPPRInfo(DomesticDataBase):

    def __init__(self):
        super().__init__()
        self.table_name = f"GqmsPPR현황"
        # self.update_customer("KMS", "KASK")
        # self.update_customer("KMMG", "KAGA")
        # self.update_customer("KMM", "KMX")
        # self.update_customer("KMI", "KIN")

    def create(self):
        with sqlite3.connect(self.db) as conn:
            cur = conn.cursor()
            cur.execute(f'''CREATE TABLE "{self.table_name}" (
	                    "No"	INTEGER,
	                    "PPR_No"	TEXT,
	                    "KDQR_No"	TEXT,
	                    "공급법인"	TEXT,
	                    "대표고객"	TEXT,
	                    "차종"	TEXT,
	                    "코드"	TEXT,
	                    "부품협력사"	TEXT,
	                    "품번"	TEXT,
	                    "품명"	TEXT,
	                    "코드.1"	TEXT,
	                    "귀책처"	TEXT,
	                    "불량수량"	INTEGER,
	                    "발생일자"	TIMESTAMP,
	                    "제기일자"	TIMESTAMP,
	                    "제목"	TEXT,
	                    "제기부서"	TEXT,
	                    "작성자"	TEXT,
	                    "공급처"	TEXT,
	                    "공장"	TEXT,
	                    "Shop"	TEXT,
	                    "발생장소"	TEXT,
	                    "발생유형"	TEXT,
	                    "조치구분"	TEXT,
	                    "불량유형1"	TEXT,
	                    "불량유형2"	TEXT,
	                    "불량유형3"	TEXT,
	                    "긴급도"	TEXT,
	                    "시스템"	TEXT,
	                    "P/T기종"	TEXT,
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

        df_base.columns = [i.replace(" ", "_") for i in df_base.columns.tolist()]
        df_base['ID'] = [i + 1 for i in range(len(df_base))]
        df_base['품번'] = [str(i).replace("-", "") for i in df_base['품번']]
        df_base['timestamp'] = [yyyymm for _ in range(len(df_base))]
        df_base['품명'] = [str(i).upper() for i in df_base['품명'].tolist()]
        df_base['입력일자'] = datetime.now().strftime("%Y%m%d")
        df_base.columns = [i.replace(" ", "_") for i in df_base.columns.tolist()]
        df_base['발생일자'] = [i if i.year < 3000 else i.replace(year=int(str(yyyymm)[:4])) for i in df_base['발생일자'].tolist()]
        df_base['대표고객'] = [str(i).upper() for i in df_base['대표고객']]
        df_base.fillna("", inplace=True)
        return df_base

    def update_customer(self, old_name, new_name):
        with sqlite3.connect(self.db) as conn:
            cur = conn.cursor()
            query = rf''' UPDATE {self.table_name} SET 대표고객='{new_name}' WHERE 대표고객='{old_name}'  '''
            cur.execute(query)
            conn.commit()


if __name__ == "__main__":
    import os

    os.chdir(os.pardir)
    os.chdir(os.pardir)

    gqms_ppr = GqmsPPRInfo()

    # gqms_ppr.drop()
    gqms_ppr.append(yyyymm=202112, if_exist='overwrite')
    # gqms_ppr.store_multiple_table(202109)

    df = gqms_ppr.search(202108, from_to=False)
    print(df.head())