from modules.db.db_overseas import OverseasDataBase
from modules.db.db_domestic import DomesticDataBase
from tqdm import tqdm
import pandas as pd


if __name__ == "__main__":
    import os

    os.chdir(os.pardir)

    df_stacked_ovs = None
    df_stacked_dms = None
    year_range = [2019, 2020, 2021]

    ovs_defect = OverseasDataBase()
    dms_defect = DomesticDataBase()

    for year in tqdm(year_range):
        yyyymm = int(str(year)+'12')
        df_ovs = ovs_defect.search(yyyymm=yyyymm, from_to=True)
        df_dms = dms_defect.search(yyyymm=yyyymm, from_to=True)

        if df_stacked_ovs is None:
            df_stacked_ovs = df_ovs
        else:
            df_stacked_ovs = pd.concat([df_stacked_ovs, df_ovs])

        if df_stacked_dms is None:
            df_stacked_dms = df_dms
        else:
            df_stacked_dms = pd.concat([df_stacked_dms, df_dms])

    raw_data = r"tasks\raw_data.xlsx"

    with pd.ExcelWriter(raw_data, engine='xlsxwriter') as writer:
        df_stacked_ovs.to_excel(writer, sheet_name="해외품질현황", startrow=1, startcol=0, index=False)
        df_stacked_dms.to_excel(writer, sheet_name="국내품질현황", startrow=1, startcol=0, index=False)

    os.startfile(raw_data)


