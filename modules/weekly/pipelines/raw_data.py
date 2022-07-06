from modules.db.db_overseas import OverseasDataBase
from utils_.config import CUSTOMERS, OVS_INVALID_TYPE
from modules.weekly.weekly_config import W_CUSTOMERS, W_OVS_INVALID_TYPE, INVALID_PPR_SUPPLIED_FROM
from datetime import datetime
from modules.db.db_dqy import DisbursementInfo


def get_raw_ovs(pre_noti=True, regular_month=False, yyyymm=None, weekly_customer=False, weekly_type=False):
    
    yyyymm = yyyymm if regular_month else int(datetime.today().strftime("%Y%m"))
    last_year = int(str(yyyymm)[:4]) - 1
    last_year_month = int(str(last_year) + "01")

    ovs_obj = OverseasDataBase(pre_noti=pre_noti)
    ovs_df = ovs_obj.search(last_year_month, from_to=True, after=True)
    ovs_df['고객사'] = [str(i).replace("(CN)", "") for i in ovs_df['고객사']]
    customer_list = W_CUSTOMERS if weekly_customer else CUSTOMERS
    ovs_invalid_type = W_OVS_INVALID_TYPE if weekly_type else OVS_INVALID_TYPE

    return ovs_df[(~ovs_df["KDQR_Type"].isin(ovs_invalid_type)) & (ovs_df["고객사"].isin(customer_list)) & (~ovs_df["공급법인"].isin(INVALID_PPR_SUPPLIED_FROM))]


def get_raw_dqy(regular_month=False, yyyymm=None, weekly_customer=False):

    yyyymm = yyyymm if regular_month else int(datetime.today().strftime("%Y%m"))
    last_year = int(str(yyyymm)[:4]) - 1
    last_year_month = int(str(last_year) + "01")

    disburse_qty = DisbursementInfo()
    dqy_df = disburse_qty.search(last_year_month, from_to=True, after=True)
    customer_list = W_CUSTOMERS if weekly_customer else CUSTOMERS

    return dqy_df[(dqy_df["대표고객"].isin(customer_list)) & (~dqy_df["공급법인"].isin(INVALID_PPR_SUPPLIED_FROM))]


if __name__ == "__main__":
    import os
    os.chdir(os.pardir)
    os.chdir(os.pardir)
    os.chdir(os.pardir)

    pre_noti = True
    regular_month = False
    yyyymm = None
    weekly_customer = True
    weekly_type = True

    raw_ovs_df = get_raw_ovs(pre_noti=pre_noti, regular_month=regular_month, yyyymm=yyyymm, weekly_customer=weekly_customer,
                             weekly_type=weekly_type)

    print(raw_ovs_df.info())

    raw_dqy_df = get_raw_dqy(regular_month=regular_month, yyyymm=yyyymm, weekly_customer=weekly_customer)

    print(raw_dqy_df.info())