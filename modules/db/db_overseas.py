import sqlite3
import warnings
from modules.db.db_domestic import DomesticDataBase
warnings.filterwarnings('ignore')


class OverseasDataBase(DomesticDataBase):

    def __init__(self, pre_noti=False):
        super().__init__()
        self.pre_noti = pre_noti
        self.table_name = "해외품질정보현황_등록일" if pre_noti else "해외품질정보현황"

    def create(self):
        with sqlite3.connect(self.db) as conn:
            cur = conn.cursor()
            cur.execute(f'''CREATE TABLE {self.table_name} (
            	"No"	INTEGER,
            	"진행상태"	TEXT,
            	"결재상태"	TEXT,
            	"KDQR_Type"	TEXT,
            	"대표PPR_No"	TEXT,
            	"PPR_No"	TEXT,
            	"KDQR_No"	TEXT,
            	"담당자"	TEXT,
            	"공급법인"	TEXT,
            	"고객사"	TEXT,
            	"차종"	TEXT,
            	"코드"	TEXT,
            	"부품협력사"	TEXT,
            	"품번"	TEXT,
            	"품명"	TEXT,
            	"귀책구분"	TEXT,
            	"코드.1"	TEXT,
            	"귀책처"	TEXT,
            	"발생처"	TEXT,
            	"불량수량"	INTEGER,
            	"조치수량"	INTEGER,
            	"Monetary"	REAL,
            	"Replacement"	REAL,
            	"Exchange"	REAL,
            	"Repair"	REAL,
            	"클레임수량"	INTEGER,
            	"불량구분"	TEXT,
            	"불량유형(대)"	TEXT,
            	"불량유형(중)"	TEXT,
            	"불량유형(소)"	TEXT,
            	"제목"	TEXT,
            	"발생일자"	TIMESTAMP,
            	"등록일자"	TIMESTAMP,
            	"해외결재일"	TIMESTAMP,
            	"국내접수일"	TIMESTAMP,
            	"국내결재일"	TIMESTAMP,
            	"공급국가"	TEXT,
            	"Type"	TEXT,
            	"공장"	TEXT,
            	"중요도"	TEXT,
            	"긴급도"	TEXT,
            	"System"	TEXT,
            	"Shop"	TEXT,
            	"CC"	TEXT,
            	"Shipmode"	TEXT,
            	"P/T_Model"	TEXT,
            	"대책작성여부"	TEXT,
            	"오더번호"	TEXT,
            	"대책등록일"	TIMESTAMP,
            	"대책확인일"	TIMESTAMP,
            	"대책확인일(해외)"	TIMESTAMP,
            	"개선일"	TIMESTAMP,
            	"개선적용오더"	TEXT,
            	"발생원인(대)"	TEXT,
            	"발생원인(중)"	TEXT,
            	"발생원인(소)"	TEXT,
            	"최종조립처"	TEXT,
                "고품회수여부"	TEXT,
                "PPR_확인_일자"	TIMESTAMP,
            	"유니크여부"	REAL,
            	"2차사여부"	REAL,
            	"2차사명"	REAL,
            	"현재결재자"	TEXT,
            	"ID"	INTEGER,
            	"timestamp"	INTEGER,
                "입력일자"	TEXT,
            	PRIMARY KEY("timestamp","ID")
            );''')
            conn.commit()


if __name__ == "__main__":
    import os
    import pandas as pd
    from modules.gqims import GqimsDataRetrieval

    os.chdir(os.pardir)
    os.chdir(os.pardir)

    yyyymm = 202203
    pre_noti = False  # 해외제기중 포함


    #########################GQIMS DOWNLOAD###############################
    # gqims = GqimsDataRetrieval(pre_noti=pre_noti)
    # gqims.login()
    # gqims.set_ovs_screen()
    # gqims.download_multiple_table(yyyymm, gqims.download_ovs_table)
    # gqims.download_ovs_table(yyyymm)
    # gqims.close()
    #########################GQIMS DOWNLOAD###############################

    ovs_defect = OverseasDataBase(pre_noti=pre_noti)
    # ovs_defect.drop()
    # ovs_defect.append(yyyymm)
    # ovs_defect.store_multiple_table(yyyymm)
    # df = ovs_defect.search(yyyymm)
    # print(df['고객사'].unique())

    df = ovs_defect.search(yyyymm=yyyymm, from_to=True)

    # print(df.head())

    file_name = r'spawn\test.xlsx'
    # df_this = ovs_defect.search(yyyymm=202112, from_to=True)
    # df_last = ovs_defect.search(yyyymm=202012, from_to=True)

    # df = ovs_defect.search(yyyymm=202110, from_to=False)
    with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
        workbook = writer.book
        # icm_2021_df_super_raw.to_excel(writer, sheet_name="입고이력_2021", startrow=1, index=False)
        # df_this.to_excel(writer, sheet_name="2021", startrow=1, index=False)
        # df_last.to_excel(writer, sheet_name="2020", startrow=1, index=False)
        df.to_excel(writer, sheet_name="2022", startrow=0, index=False)

    os.startfile(file_name)