import sqlite3
import pandas as pd
import warnings
from datetime import datetime
warnings.filterwarnings('ignore')


class WeeklyNotesModel:

    def __init__(self):
        self.db = "database.db"
        self.table_name = "주간업무현황"

        try:
            self.create_table()
        except sqlite3.OperationalError:
            pass

    def create_table(self):
        with sqlite3.connect(self.db) as conn:
            cur = conn.cursor()
            cur.execute(f'''CREATE TABLE {self.table_name} (
                "고객사"	TEXT,
                "담당자"	TEXT,
                "주차"	TEXT,
                "실시사항"	TEXT,
                "예정사항"	TEXT,
            	PRIMARY KEY("고객사", "담당자", "주차")
            );''')

            conn.commit()

    def drop(self):
        with sqlite3.connect(self.db) as conn:
            cur = conn.cursor()
            cur.execute(f'''DROP TABLE IF EXISTS {self.table_name};''')
            conn.commit()

    def insert(self, customer="", poc="", week="", this_week_task="", next_week_task=""):
        try :
            with sqlite3.connect(self.db) as conn:
                cur = conn.cursor()
                query = f'''INSERT INTO {self.table_name} 
                            VALUES (?,?,?,?,?)'''
                cur.execute(query, (customer, poc, week, this_week_task, next_week_task))
                conn.commit()
        except sqlite3.IntegrityError:
            print(f"{customer} {week} 로우는 이미 저장되어있습니다.")

    def delete(self,  customer="", poc="", week="", this_week_task="", next_week_task=""):
        with sqlite3.connect(self.db) as conn:
            cur = conn.cursor()
            query = f''' DELETE FROM "{self.table_name}"  WHERE 고객사=? AND 담당자=? AND 주차=?'''
            cur.execute(query, (customer, poc, week,))
            conn.commit()

    def convert_to_dataframe(self):
        with sqlite3.connect(self.db) as conn:
            query = f''' SELECT * FROM {self.table_name} WHERE 고객사!=?'''
            df = pd.read_sql_query(query, conn, params=('test',))

        return df


if __name__ == "__main__":
    import os

    os.chdir(os.pardir)
    os.chdir(os.pardir)

    weekly_note = WeeklyNotesModel()

    # weekly_note.insert("test", "이일희", "12345", "ㅇㅇㅇㅇ를 했고 \n ㅇㅇㅇ를했다 근데 수정", "ㅇㅇㅇㅇ를 할거고 \n ㅇㅇㅇ를 할것이다")
    df = weekly_note.convert_to_dataframe()
    df.sort_values("주차", ascending=False, inplace=True)

    df.to_excel(r'spawn\test30.xlsx', index=False)
    os.startfile(r'spawn\test30.xlsx')
