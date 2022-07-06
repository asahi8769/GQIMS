import pandas as pd
from tqdm import tqdm
from utils_.config import *
from modules.db.db_domestic import DomesticDataBase
from modules.db.db_incoming import IncomingInfo


class DomesticSummary:

    def __init__(self, yyyymm:int):

        self.yyyymm = yyyymm
        # self.dms_actions = DMS_VALID_TYPE #  Use negative criteria.
        self.dms_invalid_type = DMS_INVALID_TYPE
        self.customer_name_revision = CUSTOMER_NAME_REVISION
        self.broader_customers = BROADER_CUSTOMERS
        self.commercial_customer = C_CUSTOMERS
        self.customer_location = CUSTOMER_LOCATION
        self.customer_col = None

    def get_summary(self, ):
        df = self.generate_indicies()

        domestic = DomesticDataBase()
        dms_df = domestic.search(self.yyyymm, from_to=True)
        dms_df['공급법인'] = [i.replace('(N)', "").replace('(A)', "") for i in dms_df['공급법인']]
        dms_df['고객사'] = [str(i).replace("(CN)", "") for i in dms_df['고객사'].replace(self.customer_name_revision)]
        dms_df = dms_df[~dms_df['불량조치'].isin(self.dms_invalid_type)]

        incoming = IncomingInfo()
        icm_df = incoming.search(self.yyyymm, from_to=True)
        icm_df['고객사'] = [i.replace("(CN)", "").replace("(M)", "") for i in icm_df['고객사'].replace(self.customer_name_revision)]

        qty = ["" for _ in range(len(self.customer_col))]
        icm = ["" for _ in range(len(self.customer_col))]
        ppm = ["" for _ in range(len(self.customer_col))]

        for n, customer in enumerate(self.customer_col):
            if customer in self.broader_customers or customer in self.commercial_customer:
                d_qty = sum(dms_df[dms_df['고객사'] == customer]['불량수량'])
                n_icm = sum(icm_df[icm_df['고객사'] == customer]['입고수량'])
                qty[n] = d_qty
                icm[n] = n_icm
                ppm[n] = d_qty/n_icm*1000000 if n_icm != 0 else 0

            elif customer == "10대공장":
                d_qty = sum(dms_df[dms_df['고객사'].isin(self.broader_customers)]['불량수량'])
                n_icm = sum(icm_df[icm_df['고객사'].isin(self.broader_customers)]['입고수량'])
                qty[n] = d_qty
                icm[n] = n_icm
                ppm[n] = d_qty / n_icm * 1000000 if n_icm != 0 else 0

            elif customer == "비계열":
                d_qty = sum(dms_df[~dms_df['고객사'].isin(self.broader_customers + self.commercial_customer)]['불량수량'])
                n_icm = sum(icm_df[~icm_df['고객사'].isin(self.broader_customers + self.commercial_customer)]['입고수량'])
                qty[n] = d_qty
                icm[n] = n_icm
                ppm[n] = d_qty / n_icm * 1000000 if n_icm != 0 else 0

            elif customer == "비계열+상용":
                d_qty = sum(dms_df[~dms_df['고객사'].isin(self.broader_customers + self.commercial_customer)]['불량수량']) + sum(dms_df[dms_df['고객사'].isin(self.commercial_customer)]['불량수량'])
                n_icm = sum(icm_df[~icm_df['고객사'].isin(self.broader_customers + self.commercial_customer)]['입고수량']) + sum(icm_df[icm_df['고객사'].isin(self.commercial_customer)]['입고수량'])
                qty[n] = d_qty
                icm[n] = n_icm
                ppm[n] = d_qty / n_icm * 1000000 if n_icm != 0 else 0

        df['입고수량'] = icm
        df['불량수량'] = qty
        df['PPM'] = [round(i, 1) for i in ppm]
        df['고객사구분'] = [self.customer_location.get(i, "-") if i != "비계열" else "비계열" for i in df['고객사']]

        df_piv = df[df['고객사구분'] != "-"]

        df_piv = pd.pivot_table(df_piv, values=['불량수량', "입고수량"], index=['고객사구분'], aggfunc=sum)
        df_piv['PPM'] = round(df_piv['불량수량']/df_piv['입고수량']*1000000, 1)
        df_piv.reset_index(drop=False, inplace=True)
        df_piv = df_piv[['고객사구분', '입고수량', '불량수량', 'PPM']]
        df_piv.sort_values(by='입고수량', ascending=False, inplace=True)

        return df, df_piv

    def generate_indicies(self):
        self.customer_col = self.broader_customers + self.commercial_customer + ["10대공장", '비계열', "비계열+상용"]
        return pd.DataFrame({"고객사": self.customer_col})


if __name__ == "__main__":
    os.chdir(os.pardir)
    os.chdir(os.pardir)

    summary = DomesticSummary(202108)
    df, df_piv = summary.get_summary()

    print(df)
    print(df_piv)
