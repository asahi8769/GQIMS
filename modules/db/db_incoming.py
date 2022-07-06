import sqlite3
import pandas as pd
import warnings
from modules.db.db_domestic import DomesticDataBase
import time, os
warnings.filterwarnings('ignore')
from utils_.functions import make_dir


class IncomingInfo(DomesticDataBase):

    def __init__(self):
        super().__init__()
        self.table_name = f"입고내역"
        self.modified_date = None
        make_dir("downloaded/입고내역")


    def create(self):
        with sqlite3.connect(self.db) as conn:
            cur = conn.cursor()
            cur.execute(f'''CREATE TABLE "{self.table_name}" (
                            "업체코드"	TEXT,
                         	"업체명"	TEXT,
                         	"카덱스"	TEXT,
                         	"입고년월"	TEXT,
                         	"고객사"	TEXT,
                         	"차종"	TEXT,
                         	"통화"	TEXT,
                         	"사업장"	TEXT,
                         	"입고건수"	INTEGER,
                         	"입고수량"	INTEGER,
                         	"입고금액"	REAL,
                            "입고금액_KRW"	INTEGER,
                            "ID" INTEGER,
                         	"timestamp"	INTEGER,
                            "입력일자"	TEXT,
                         	PRIMARY KEY("timestamp","ID")
                         );'''
            )
            conn.commit()

    def preprocessing(self, yyyymm):
        file = f"downloaded/입고내역/{self.table_name}_{yyyymm}.xlsx"
        self.modified_date = time.ctime(os.path.getmtime(file))

        with open(file, 'rb') as f:
            df = pd.read_excel(f, converters={'업체코드': lambda x: str(x)}, engine='openpyxl')

        df.columns = [str(i).replace(" ", "_").replace("\n", "") for i in df.columns.tolist()]
        df['ID'] = [i + 1 for i in range(len(df))]
        df['timestamp'] = [yyyymm for _ in range(len(df))]
        df['입력일자'] = self.modified_date
        df['고객사'] = [str(i).upper() for i in df['고객사']]

        return df

if __name__ =="__main__":
    import os

    os.chdir(os.pardir)
    os.chdir(os.pardir)

    incoming_amount = IncomingInfo()

    # incoming_amount.drop()
    # incoming_amount.append(202101)
    incoming_amount.store_multiple_table(202012)

    df = incoming_amount.search(202112, from_to=True)
    # print(sum(df['입고수량']))

    df.to_excel(r"spawn\test2.xlsx", index=False)
    os.startfile(r"spawn\test2.xlsx")
