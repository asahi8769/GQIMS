import pandas as pd
import os
import sqlite3


class MissingPi:

    def __init__(self):
        self.db = "data/database.db"
        self.master = "data/master.xlsx"

    def load(self):
        df = pd.read_excel(self.master, usecols=['고객사', '차종', 'Part No', '포장장', '납품업체'], verbose=True)
        with sqlite3.connect(self.db) as conn:
            df.to_sql("master", con=conn, if_exists='replace')



if __name__ == "__main__":
    obj = MissingPi()
    obj.load()