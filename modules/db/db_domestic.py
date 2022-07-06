import sqlite3
import pandas as pd
import warnings
from datetime import datetime
from dateutil.relativedelta import relativedelta
import calendar
warnings.filterwarnings('ignore')


class DomesticDataBase:

    def __init__(self):
        self.db = "database.db"
        self.table_name = "국내품질정보현황"
        self.join_table = "국내입고품질상세정보"

    def __str__(self):
        return self.table_name

    def create(self):
        with sqlite3.connect(self.db) as conn:
            cur = conn.cursor()
            cur.execute(f'''CREATE TABLE {self.table_name} (
                "NO"	INTEGER,
	            "통보서No"	TEXT,
	            "진행상태"	TEXT,
	            "결재상태"	TEXT,
	            "담당자"	TEXT,
	            "공급법인"	TEXT,
	            "고객사"	TEXT,
	            "차종"	TEXT,
	            "코드"	TEXT,
	            "부품협력사"	TEXT,
	            "품번"	TEXT,
	            "품명"	TEXT,
	            "불량발견장소"	TEXT,
	            "불량수량"	INTEGER,
	            "불량구분"	TEXT,
	            "불량유형(대)"	TEXT,
	            "불량유형(중)"	TEXT,
	            "불량유형(소)"	TEXT,
	            "제목"	TEXT,
	            "등록일자"	TIMESTAMP,
	            "검사일자"	TIMESTAMP,
	            "통보일자"	TIMESTAMP,
	            "접수일자"	TIMESTAMP,
	            "결재일자"	TIMESTAMP,
	            "불량조치"	TEXT,
	            "사업장"	TEXT,
	            "포장장"	TEXT,
	            "검사번호"	TEXT,
	            "대책등록일"	TIMESTAMP,
	            "대책확인일"	TIMESTAMP,
	            "개선일"	TIMESTAMP,
	            "개선적용오더"	TEXT,
	            "발생원인(대)"	TEXT,
	            "발생원인(중)"	TEXT,
	            "발생원인(소)"	TEXT,
	            "최종조립처"	TEXT,
	            "UNIQ_YN"	REAL,
	            "2차사여부"	REAL,
	            "2차사명"	REAL,
	            "현재결재자"	TEXT,
            	"ID"	INTEGER,
            	"timestamp"	INTEGER,
                "입력일자"	TEXT,
            	PRIMARY KEY("timestamp","ID")
            );''')

            conn.commit()

    def preprocessing(self, yyyymm):
        file = f"downloaded/{self.table_name}_{yyyymm}.xlsx"
        with open(file, 'rb') as f:
            df = pd.read_excel(f, converters={'품번': lambda x: str(x)}, engine='openpyxl')
            df['ID'] = [i + 1 for i in range(len(df))]
            df['품번'] = [str(i).replace("-", "") for i in df['품번']]
            df['timestamp'] = [yyyymm for _ in range(len(df))]
            df['품명'] = [str(i).upper() for i in df['품명'].tolist()]
            df['입력일자'] = datetime.now().strftime("%Y%m%d")
            df.columns = [i.replace(" ", "_").replace("\n", "") for i in df.columns.tolist()]
            df['고객사'] = [str(i).upper() for i in df['고객사']]

        return df

    def search(self, yyyymm=None, from_to=False, after=False):
        year = int(str(yyyymm)[:4])
        with sqlite3.connect(self.db) as conn:
            if not from_to:
                query = f''' SELECT * FROM {self.table_name} WHERE timestamp=?'''
                df = pd.read_sql_query(query, conn, params=(yyyymm,))
            else:
                if after:
                    query = f''' SELECT * FROM {self.table_name} WHERE timestamp>=?'''
                    df = pd.read_sql_query(query, conn, params=(yyyymm,))
                else:
                    query = f''' SELECT * FROM {self.table_name} WHERE timestamp LIKE ? AND timestamp<=?'''
                    df = pd.read_sql_query(query, conn, params=(f"{year}%",yyyymm,))
            return df

    def drop(self):
        try:
            with sqlite3.connect(self.db) as conn:
                cur = conn.cursor()
                cur.execute(f'''DROP TABLE IF EXISTS {self.table_name};''')
                conn.commit()
        except sqlite3.OperationalError:
            pass

    def append(self, yyyymm, if_exist='overwrite'):

        try:
            self.create()
        except sqlite3.OperationalError:
            pass

        df = self.preprocessing(yyyymm)

        try:
            with sqlite3.connect(self.db) as conn:
                df.to_sql(self.table_name, con=conn, if_exists='append', index=None, index_label=None)
                print(f"{yyyymm} {self.table_name} 삽입되었습니다.")
        except sqlite3.IntegrityError:
            if if_exist == "overwrite":
                print(f"{yyyymm} {self.table_name} 이미 존재합니다, 데이터 삭제후 다시 삽입합니다.")
                self.delete(yyyymm)
                self.append(yyyymm, if_exist="pass")
            if if_exist == "pass":
                print(f"{yyyymm} {self.table_name} 이미 존재합니다. 데이터를 삽입하지 못했습니다.")

    def delete(self, yyyymm):
        with sqlite3.connect(self.db) as conn:
            cur = conn.cursor()
            query = f''' DELETE FROM "{self.table_name}" WHERE timestamp=? '''
            cur.execute(query, (yyyymm,))
            conn.commit()

    def store_multiple_table(self, end):
        for j in [i for i in range(int(str(end)[:4]+'01'), int(str(end)[:4]+'01'+'13'), 1) if i <= end]:
            self.append(j, if_exist="overwrite")

    def yearly_search(self, yyyymm, reference="검사일자", delta=1):
        year = int(str(yyyymm)[:4])
        month = int(str(yyyymm)[4:])
        now = datetime(year, month, calendar.monthrange(year, month)[1])
        last_year = now - relativedelta(years=delta)

        with sqlite3.connect(self.db) as conn:
            query = f''' SELECT * FROM {self.table_name} WHERE {reference} BETWEEN ? and ?'''
            df = pd.read_sql_query(query, conn, params=(last_year,now))
            return df

    def monthly_search(self, yyyymm, reference="timestamp", delta=2):
        now = datetime.strptime(str(yyyymm), "%Y%m")
        past = now - relativedelta(months=delta-1)

        with sqlite3.connect(self.db) as conn:
            query = f''' SELECT * FROM {self.table_name} WHERE {reference} BETWEEN ? and ?'''
            df = pd.read_sql_query(query, conn, params=(int(past.strftime("%Y%m")), int(now.strftime("%Y%m"))))
            return df

    def monthly_specific_search(self, *yyyymm_list, reference="timestamp"):
        with sqlite3.connect(self.db) as conn:
            query = f''' SELECT * FROM {self.table_name} WHERE {reference} in ({','.join(['?']*len(yyyymm_list))})'''
            df = pd.read_sql_query(query, conn, params=yyyymm_list)
            return df

if __name__ == "__main__":
    import os
    import calendar

    os.chdir(os.pardir)
    os.chdir(os.pardir)

    dms_defect = DomesticDataBase()
    # dms_defect.drop()
    # dms_defect.store_multiple_table(202109)
    # dms_defect.append(202112)
    month = 1
    df = dms_defect.yearly_search(yyyymm=202101, delta=1)

    df.to_excel(r"spawn\test2.xlsx", index=False)
    os.startfile(r"spawn\test2.xlsx")

    # print(df["고객사"].unique())