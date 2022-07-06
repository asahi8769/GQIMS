import sqlite3
import pandas as pd
import warnings
from modules.db.db_domestic import DomesticDataBase
from datetime import datetime
warnings.filterwarnings('ignore')


class IncomingResultInfo(DomesticDataBase):

    def __init__(self):
        super().__init__()
        self.table_name = f"입고검사결과"

    def create(self):
        pass
        with sqlite3.connect(self.db) as conn:
            cur = conn.cursor()
            cur.execute(f'''CREATE TABLE "입고검사결과" (
    	                "검사번호"	TEXT,
    	                "순번"	INTEGER,
    	                "검사일자"	TIMESTAMP,
    	                "공급법인"	TEXT,
    	                "고객사"	TEXT,
    	                "차종"	TEXT,
    	                "코드"	TEXT,
    	                "부품협력사"	TEXT,
    	                "품번"	TEXT,
    	                "품명"	TEXT,
    	                "공급국가"	TEXT,
    	                "오더번호"	TEXT,
    	                "납입묶음번호"	TEXT,
    	                "검사자"	REAL,
    	                "Lot No"	TEXT,
    	                "사업장"	TEXT,
    	                "포장장"	TEXT,
    	                "검사기준"	TEXT,
    	                "입고수량"	INTEGER,
    	                "검사결과"	TEXT,
    	                "불량수량"	INTEGER,
    	                "통보서No"	TEXT,
    	                "ID"	INTEGER,
    	                "timestamp"	INTEGER,
    	                "입력일자"	TEXT,
    	                PRIMARY KEY("timestamp","ID")
                        );'''
            )
            conn.commit()

    def preprocessing(self, yyyymm):
        inspection_mode = ["전수검사", "합동검사"]

        df_base = None
        for mode in inspection_mode:

            file = f"downloaded/{self.table_name}_{yyyymm}_{mode}.xlsx"

            with open(file, 'rb') as f:
                df = pd.read_excel(f, converters={'품번': lambda x: str(x), "불량수량": lambda x: int(x) if x != "" else 0}, engine='openpyxl')

            if df_base is None:
                df_base = df
            else :
                df_base = pd.concat([df_base, df], ignore_index=True)

        df_base['ID'] = [i + 1 for i in range(len(df_base))]
        df_base['품번'] = [str(i).replace("-", "") for i in df_base['품번']]
        df_base['timestamp'] = [yyyymm for _ in range(len(df_base))]
        df_base['품명'] = [str(i).upper() for i in df_base['품명'].tolist()]
        df_base['입력일자'] = datetime.now().strftime("%Y%m%d")
        df_base['고객사'] = [str(i).upper() for i in df_base['고객사']]

        return df_base


if __name__ == "__main__":
    import os

    os.chdir(os.pardir)
    os.chdir(os.pardir)

    inc_result = IncomingResultInfo()

    # inc_result.drop()
    # inc_result.append(yyyymm=202101, if_exist='overwrite')
    inc_result.store_multiple_table(202109)

    df = inc_result.search(202109, from_to=True)
    print(df)