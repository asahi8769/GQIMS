import sqlite3
import pandas as pd
import warnings, time, os
from modules.db.db_domestic import DomesticDataBase
warnings.filterwarnings('ignore')


class DomDetailedInfo(DomesticDataBase):

    def __init__(self):
        super().__init__()
        self.table_name = f"국내입고품질상세정보"
        self.modified_date = None

    def __str__(self):
        return self.table_name

    def create(self):
        with sqlite3.connect(self.db) as conn:
            cur = conn.cursor()
            cur.execute(f'''CREATE TABLE "{self.table_name}" (
                            	"통보서NO"	TEXT,
                            	"불량상세"	TEXT,
                            	"원인상세"	TEXT,
                            	"개선대책"	TEXT,
                                "입력일자"	TEXT,
                            	PRIMARY KEY("통보서NO")
                            );'''
            )
            conn.commit()

    def preprocessing(self):
        file = fr"downloaded/{self.table_name}.xlsx"
        self.modified_date = time.ctime(os.path.getmtime(file))

        with open(file, 'rb') as f:
            df = pd.read_excel(f, engine='openpyxl')
            df.fillna("", inplace=True)
            df['불량상세'] = [str(i).replace('\n', "") for i in df['불량상세'].tolist()]
            df['원인상세'] = [str(i).replace('\n', "") for i in df['원인상세'].tolist()]
            df['개선대책'] = [str(i).replace('\n', "") for i in df['개선대책'].tolist()]
            df['입력일자'] = self.modified_date
            df.drop("CORP_CD", axis=1, inplace=True)
            df = df.drop_duplicates(subset=['통보서NO'], keep='last')
        return df

    def append(self):

        try:
            self.create()
        except sqlite3.OperationalError:
            pass

        df = self.preprocessing()
        try:
            with sqlite3.connect(self.db) as conn:
                df.to_sql(self.table_name, con=conn, if_exists='append', index=None, index_label=None)
                print(f"{self.table_name} 삽입되었습니다.")
        except sqlite3.IntegrityError:
            print(f"{self.table_name} 이미 존재합니다, 데이터 삭제후 다시 삽입합니다.")
            self.drop()
            self.append()


class ExpDetailedInfo(DomDetailedInfo):

    def __init__(self):
        super().__init__()
        self.table_name = f"해외입고품질상세정보"

    def preprocessing(self):
        file = fr"downloaded/{self.table_name}.xlsx"
        self.modified_date = time.ctime(os.path.getmtime(file))

        with open(file, 'rb') as f:
            df = pd.read_excel(f, engine='openpyxl')
            df.fillna("", inplace=True)
            df['불량상세'] = [str(i).replace('\n', "") for i in df['불량상세'].tolist()]
            df['원인상세'] = [str(i).replace('\n', "") for i in df['원인상세'].tolist()]
            df['개선대책'] = [str(i).replace('\n', "") for i in df['개선대책'].tolist()]
            df['입력일자'] = self.modified_date
            df.drop(["CORP_CD", "REPN_PPR_NO"], axis=1, inplace=True)
            df = df.drop_duplicates(subset=['PPR_NO'], keep='last')
        return df

    def create(self):
        with sqlite3.connect(self.db) as conn:
            cur = conn.cursor()
            cur.execute(f'''CREATE TABLE "{self.table_name}" (
                            	"PPR_NO"	TEXT,
                            	"불량상세"	TEXT,
                            	"원인상세"	TEXT,
                            	"개선대책"	TEXT,
                                "입력일자"	TEXT,
                            	PRIMARY KEY("PPR_NO")
                            );'''
            )
            conn.commit()


class PackingCenterInfo(DomDetailedInfo):
    """no longer used"""

    def __init__(self):
        super().__init__()
        self.table_name = f"포장장위치구분정보"

    def preprocessing(self):
        file = fr"downloaded/{self.table_name}.xlsx"
        self.modified_date = time.ctime(os.path.getmtime(file))

        with open(file, 'rb') as f:
            df = pd.read_excel(f, engine='openpyxl', sheet_name="포장장")
            df['입력일자'] = self.modified_date
            df.fillna("", inplace=True)
        return df

    def create(self):
        with sqlite3.connect(self.db) as conn:
            cur = conn.cursor()
            cur.execute(f'''CREATE TABLE "{self.table_name}" (
                            	"포장장명"	TEXT,
                            	"사업장명"	TEXT,
                            	"포장장_위치"	TEXT,
                            	"포장장_구분"	TEXT,
                                "입력일자"	TEXT,
                            	PRIMARY KEY("포장장명")
                            );'''
            )
            conn.commit()


class CustomerLocationInfo(DomDetailedInfo):
    """no longer used"""

    def __init__(self):
        super().__init__()
        self.file_name = f"포장장위치구분정보"
        self.table_name = f"고객사위치구분정보"

    def preprocessing(self):
        file = fr"downloaded/{self.file_name}.xlsx"
        self.modified_date = time.ctime(os.path.getmtime(file))

        with open(file, 'rb') as f:
            df = pd.read_excel(f, engine='openpyxl', sheet_name="고객사")
            df['입력일자'] = self.modified_date
            df.fillna("", inplace=True)
        return df

    def create(self):
        with sqlite3.connect(self.db) as conn:
            cur = conn.cursor()
            cur.execute(f'''CREATE TABLE "{self.table_name}" (
                            	"고객사"	TEXT,
                            	"고객사_구분"	TEXT,
                                "입력일자"	TEXT,
                            	PRIMARY KEY("고객사")
                            );'''
            )
            conn.commit()


if __name__=="__main__":
    import os

    os.chdir(os.pardir)
    os.chdir(os.pardir)

    dom_info = DomDetailedInfo()
    # dom_info.drop()
    dom_info.append()

    exp_info = ExpDetailedInfo()
    # exp_info.drop()
    exp_info.append()

    # file = fr"downloaded/해외입고품질상세정보.xlsx"
    # modified_date = time.ctime(os.path.getmtime(file))
