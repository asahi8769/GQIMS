import sqlite3
import pandas as pd
import warnings
from modules.db.db_domestic import DomesticDataBase
import time, os
warnings.filterwarnings('ignore')
# from utils_.functions import make_dir


class IncomingInfoWithPN(DomesticDataBase):

    def __init__(self):
        super().__init__()
        self.table_name = f"입고품번"
        self.modified_date = None
        # make_dir("downloaded/입고내역_품번")

    def create(self):
        with sqlite3.connect(self.db) as conn:
            cur = conn.cursor()
            cur.execute(f'''CREATE TABLE "{self.table_name}" (
                            "고객사"	TEXT,
                         	"차종"	TEXT,
                         	"부품번호"	TEXT,
                         	"부품명"	TEXT,
                         	"업체코드"	TEXT,
                         	"업체명"	TEXT,
                         	"포장장코드"	TEXT,
                         	"입고수량"	TEXT,
                            "ID"	TEXT,    
                         	"timestamp"	INTEGER,
                            "입력일자"	TEXT,
                         	PRIMARY KEY("timestamp","ID")
                         );'''
            )
            conn.commit()

    def preprocessing(self, yyyymm):
        file = f"downloaded/입고내역_품번/{self.table_name}_{yyyymm}.xls"
        self.modified_date = time.ctime(os.path.getmtime(file))

        with open(file, 'rb') as f:
            df = pd.read_excel(f, converters={'업체코드': lambda x: str(x)})

        df.columns = [str(i).replace(" ", "_").replace("\n", "") for i in df.columns.tolist()]
        df['ID'] = [i + 1 for i in range(len(df))]
        df['timestamp'] = [yyyymm for _ in range(len(df))]
        df['입력일자'] = self.modified_date
        df['고객사'] = [str(i).upper() for i in df['고객사']]

        return df

if __name__ =="__main__":
    import os
    from datetime import datetime
    import calendar

    os.chdir(os.pardir)
    os.chdir(os.pardir)

    incoming_amount = IncomingInfoWithPN()

    # incoming_amount.drop()
    # incoming_amount.append(202012)
    # incoming_amount.append(202201)
    # incoming_amount.store_multiple_table(202111)




    # df = incoming_amount.search(202109, from_to=False)
    # print(sum(df['입고수량']))
    df = incoming_amount.monthly_search(yyyymm=202101, delta=2)


    # df = incoming_amount.search(yyyymm=202112, from_to=True)
    df.to_excel(r"spawn\test2.xlsx", index=False)
    os.startfile(r"spawn\test2.xlsx")
