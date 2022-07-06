from modules.db.db_dqy import DisbursementInfo
from utils_.config import *
import pandas as pd

if __name__ == "__main__":
    import os

    os.chdir(os.pardir)

    dqy_obj = DisbursementInfo()
    df = dqy_obj.search(yyyymm=202110, from_to=True)

    df['포장장위치'] = [PACKING_CENTERS.get(i, "기타") for i in df['포장장']]

    print(df[df['포장장위치'] == "기타"]['포장장'].unique().tolist())

    raw_data = r"tasks\dqy_data.xlsx"
    with pd.ExcelWriter(raw_data, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name="불출수량현황", startrow=1, startcol=0, index=False)

    os.startfile(raw_data)