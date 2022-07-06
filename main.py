import sys, warnings, os

warnings.simplefilter("ignore", UserWarning)
sys.coinit_flags = 2

from modules.db.db_domestic import DomesticDataBase
from modules.db.db_dqy import DisbursementInfo
from modules.db.db_overseas import OverseasDataBase
from modules.db.db_ppr import GqmsPPRInfo
from modules.db.db_inr import IncomingResultInfo
from modules.gqims import GqimsDataRetrieval
from modules.ppm.ovs_ppm import MonthlyOVSppmCalculation
from modules.warehouse import store_raw_data
from modules.ppm.ovs_ranking_v2 import overseas_rank
from modules.ppm.ovs_summary import OverSeasSummary
from modules.ppm.decorate_ppm import sheet_decorate
from modules.db.db_incoming import IncomingInfo
from modules.ppm.dms_ppm import MonthlyDMSppmCalculation
from modules.ppm.dms_summary import DomesticSummary
from modules.mail.html.monthly_html import index_html
from modules.mail.email_ import AutoMailing
from modules.weekly.customer_quality_format_v2 import ExcelFormatSetter
from datetime import date
import pandas as pd
from modules.plots.ovs_plots import draw_plot1
from modules.plots.cust_plot_single_barchart import draw_plot2
from modules.plots.cust_plot_stacked_barchart import draw_plot3
from utils_.functions import make_dir
from utils_.config import *
from tqdm import tqdm


def run():
    make_dir("spawn")
    make_dir("spawn\plots")

    while True:

        answer = input("작업 선택 (1.데이터 다운로드 및 생성, 2.입고내역 데이터 생성, 3. 월간품질지수출력, 4. 주간회의자료생성) : ")

        if answer == "1":
            yyyymm = int(input("년월선택(yyyymm) : "))
            pre_noti = input("기준 선택 (1.해외결재일, 2.등록일(해외제기중포함), 기본값 1) : ")
            pre_noti = True if pre_noti == "2" else False
            answer = input("기간/해당월 선택 (1.해당월, 2.기간(해당년도1월부터)) : ")

            if answer == "1":
                download_data_single(yyyymm, pre_noti)

            if answer == "2":
                download_data_multiple(yyyymm, pre_noti)

            else:
                continue

        elif answer == "2":
            yyyymm = int(input("년월선택(yyyymm) : "))

            incoming_info = IncomingInfo()
            incoming_info.append(yyyymm)

        elif answer == "3":

            yyyymm = int(input("년월선택(yyyymm) : "))
            print("입고수량 생성하였는지 확인후 진행바랍니다.")
            fixed_dqy = True if input("확정자료입니까?(y/n, 기본값 n) : ").lower() == "y" else False

            if not fixed_dqy:
                try:
                    disadvantage = float(input("미확정월의 불출수량 디스어드밴티지를 입력하세요(-1~1, 기본값 0) : "))
                except (TypeError, ValueError):
                    disadvantage = 0
                if abs(disadvantage) > 1:
                    disadvantage = 0
            else:
                disadvantage = 0

            ovs_receive = input("송부 범위를 지정하세요.(1.나에게만, 2.팀원, 3. 팀원+해외법인, 기본값 1) : ")

            if ovs_receive == "2":
                receiver = ','.join(list(TEAM_MEMBERS_MAIL.values()))
                cc = None
            elif ovs_receive == "3":
                receiver = ','.join(list(TEAM_MEMBERS_MAIL.values()))
                cc = ','.join(list(OVERSEAS_RECIPIENTS))
            else:
                receiver = MY_MAIL
                cc = None

            notify_monthly_q_info(yyyymm, fixed_dqy, disadvantage, receiver, cc)

        elif answer == "4":
            weeks = input("품질분석 주수를 지정하세요.(1~4, 기본값 1) : ")
            weeks = 1 if weeks not in ["2", "3", "4"] else int(weeks)

            excel = ExcelFormatSetter(r_weeks=weeks)
            excel.update()
            os.startfile(excel.updated_dir)

        else:
            continue


def download_data_single(yyyymm, pre_noti):
    obj = GqimsDataRetrieval(pre_noti=pre_noti)
    obj.login()
    obj.set_dms_screen()
    obj.download_dms_table(yyyymm)
    obj.set_ovs_screen()
    obj.download_ovs_table(yyyymm)
    obj.set_ppr_screen()
    obj.download_ppr_table(yyyymm)
    obj.set_dqy_screen()
    obj.download_dqy_table(yyyymm)
    obj.set_inr_screen()
    obj.download_inr_table(yyyymm)
    obj.close()

    dms_defect = DomesticDataBase()
    dms_defect.append(yyyymm)
    ovs_defect = OverseasDataBase(pre_noti=pre_noti)
    ovs_defect.append(yyyymm)
    gqms_ppr = GqmsPPRInfo()
    gqms_ppr.append(yyyymm)
    dqy_defect = DisbursementInfo()
    dqy_defect.append(yyyymm)
    inc_result = IncomingResultInfo()
    inc_result.append(yyyymm)

    store_raw_data(str(date.today()))


def download_data_multiple(yyyymm, pre_noti):
    obj = GqimsDataRetrieval(pre_noti=pre_noti)
    obj.login()
    obj.set_dms_screen()
    obj.download_multiple_table(yyyymm, obj.download_dms_table)
    obj.set_ovs_screen()
    obj.download_multiple_table(yyyymm, obj.download_ovs_table)
    obj.set_ppr_screen()
    obj.download_multiple_table(yyyymm, obj.download_ppr_table)
    obj.set_dqy_screen()
    obj.download_multiple_table(yyyymm, obj.download_dqy_table)
    obj.set_inr_screen()
    obj.download_multiple_table(yyyymm, obj.download_inr_table)
    obj.close()

    dms_defect = DomesticDataBase()
    dms_defect.store_multiple_table(yyyymm)
    ovs_defect = OverseasDataBase(pre_noti=pre_noti)
    ovs_defect.store_multiple_table(yyyymm)
    gqms_ppr = GqmsPPRInfo()
    gqms_ppr.store_multiple_table(yyyymm)
    dqy_defect = DisbursementInfo()
    dqy_defect.store_multiple_table(yyyymm)
    inc_result = IncomingResultInfo()
    inc_result.store_multiple_table(yyyymm)
    store_raw_data(str(date.today()))


def notify_monthly_q_info(yyyymm, fixed_dqy, disadvantage, receiver, cc):
    ovs_ppm = MonthlyOVSppmCalculation(yyyymm, fixed_dqy=fixed_dqy, disadvantage=disadvantage)
    ovs_df = ovs_ppm.get_ppm()
    ovs_download_date = ovs_ppm.download_date
    ovs_now = ovs_ppm.now

    df_dict = overseas_rank(yyyymm)
    ovs_summary = OverSeasSummary(yyyymm, df=ovs_df)
    df_ovs_summary = ovs_summary.get_summary()

    dms_ppm = MonthlyDMSppmCalculation(yyyymm)
    dms_df = dms_ppm.get_ppm()
    dms_download_date = dms_ppm.download_date
    dms_now = dms_ppm.now

    dms_summary = DomesticSummary(yyyymm)
    dms_summary, dms_summary_regional = dms_summary.get_summary()

    ovs_defect = OverseasDataBase()
    dms_defect = DomesticDataBase()
    gqms_ppr = GqmsPPRInfo()
    inc_result = IncomingResultInfo()
    df_ppr = gqms_ppr.search(yyyymm=yyyymm, from_to=True)

    df_ovs = ovs_defect.search(yyyymm=yyyymm, from_to=True)
    df_ovs = df_ovs[(~df_ovs["PPR_No"].isin(PPRS_TO_EXCLUDE))]

    df_dms = dms_defect.search(yyyymm=yyyymm, from_to=True)
    df_ihr = inc_result.search(yyyymm=yyyymm, from_to=True)
    df_ppr['품명'] = df_ppr['품명'].map(lambda x: "" if x == "NAN" else x)

    if ovs_ppm.estimation or not fixed_dqy:
        file_name = rf'spawn\(예상_{yyyymm})KD종합품질지수(GQMS기준)_{ovs_ppm.now.replace(":", "_").replace(" ", "")}.xlsx'
        subject = f"(예상) {yyyymm} KD 종합품질지수 공유"
    else:
        file_name = rf'spawn\({yyyymm})KD종합품질지수(GQMS기준)_{ovs_ppm.now.replace(":", "_").replace(" ", "")}.xlsx'
        subject = f"(확정) {yyyymm} KD 종합품질지수 공유"

    file_name = os.path.abspath(file_name)

    with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
        workbook = writer.book
        sheet_name_0 = "해외품질요약"
        sheet_name_1 = "해외품질순위"
        sheet_name_2 = "해외품질지수(GQMS+공급망)"
        sheet_name_3 = "국내품질지수"
        graph_sheet = "해외품질요약(그래프)"
        sheet_name_4 = "GQMSPPR현황"
        sheet_name_5 = "해외품질현황"
        sheet_name_6 = "국내품질현황"
        sheet_name_7 = "국내입고검사결과"

        margin = 1

        df_ovs_summary.to_excel(writer, sheet_name=sheet_name_0, startrow=2, startcol=1, index=True)

        row_n = 0
        row_anchor = 2
        col_anchor = 1
        for n, i in enumerate(df_dict.keys()):
            ticker = n % 2
            row = row_n * (len(df_dict[i]) + 1) + row_anchor
            col = ticker * (len(df_dict[i].columns) + 1) + col_anchor
            if ticker == 1:
                row_n += 1
                row_anchor += 2
            df_dict[i].to_excel(writer, sheet_name=sheet_name_1, startrow=row, startcol=col, index=False)

        ovs_df.to_excel(writer, sheet_name=sheet_name_2, startrow=3, startcol=margin, index=True)
        dms_df.to_excel(writer, sheet_name=sheet_name_3, startrow=3, startcol=margin, index=True)
        dms_summary.to_excel(writer, sheet_name=sheet_name_3, startrow=3, startcol=17, index=False)
        dms_summary_regional.to_excel(writer, sheet_name=sheet_name_3, startrow=3+len(dms_summary)+3, startcol=17, index=False)

        for raw_sheet, df in tqdm(
                list(zip([sheet_name_4, sheet_name_5, sheet_name_6, sheet_name_7], [df_ppr, df_ovs, df_dms, df_ihr]))):
            df.fillna("", inplace=True)
            df.to_excel(writer, sheet_name=raw_sheet, startrow=1, startcol=0, index=False)

        writer.save()

    make_dir('spawn/plots')
    _, plot_name_1 = draw_plot1(ovs_df, yyyymm)
    _, plot_name_2 = draw_plot2(type_="외관_불량수량", monday_to_saturday=False, regular_month=True, yyyymm=yyyymm)
    _, plot_name_3 = draw_plot3(type_=("외관_불량수량", "기능_치수_불량수량"), monday_to_saturday=False,
                                   regular_month=True, yyyymm=yyyymm)

    sheet_decorate(file_name, sheet_name_0, sheet_name_1, sheet_name_2, sheet_name_3, graph_sheet, df_ovs_summary,
                   ovs_df, ovs_download_date, ovs_now, dms_download_date, dms_now, plot_name_1, df_dict, dms_df,
                   dms_summary, dms_summary_regional, sheet_name_4, sheet_name_5, sheet_name_6, sheet_name_7, df_ppr,
                   df_ovs, df_dms, df_ihr, plot_name_2, plot_name_3)
    os.startfile(file_name)

    html, _ = index_html(yyyymm, ovs_df, ovs_download_date, fixed_dqy, disadvantage)

    img = {"ovs_plot1": plot_name_1, "footer": os.path.abspath(os.path.join("images", "mail_footer.png"))}

    print("\n팀원 메일주소: ", ','.join(list(TEAM_MEMBERS_MAIL.values())))
    print("\n해외법인 메일주소: ", ','.join(list(OVERSEAS_RECIPIENTS.values())))

    mail = AutoMailing(receiver=receiver, cc=cc, subject=subject, html=html, attachments=file_name, img=img)
    mail.send()


if __name__ == "__main__":
    run()


