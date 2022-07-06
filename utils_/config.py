import os
from pathlib import Path

VENV_64_DIR = os.path.join(os.getcwd(), 'venv')
SCRIPTS_DIR = os.path.join(VENV_64_DIR, 'Scripts')

REPOSITORY = r'https://github.com/asahi8769/GQIMS.git'
GQIMS_ID = os.environ.get("GQIMS_ID")
GQIMS_PW = os.environ.get("GQIMS_PASS")
GC_DRIVER = 'chromedriver.exe'
URL = "https://gqims.glovis.net/index.html"
DOWNLOAD_FOLDER = str(Path.home() / "Downloads")
EMAIL_ADDRESS = os.environ.get('EMAIL_USER')
EMAIL_PASSWORD = os.environ.get('EMAIL_PASS')

MY_MAIL = "lih@glovis.net"

TEAM_MEMBERS_MAIL = {"하성훈": "hnk@glovis.net", "김병주": "bj.kim958@glovis.net", "김시동": "a54320@glovis.net",
                     "김형동": "dong2973@glovis.net", "안남규": "namgyu.an@glovis.net", "이영록": "yllee@glovis.net",
                     "주배용": "baeyong@glovis.net", "한상진": "freehsj@glovis.net", "강현": "hyun@glovis.net",
                     "김지창": "kjc@glovis.net", "김진영": "kjy5454@glovis.net", "문재연": "jym@glovis.net",
                     "이태휘": "leeth@glovis.net", "정동윤": "jdy@glovis.net", "정성원": "jsw0829@glovis.net",
                     "이연택": "eytaek@glovis.net", "이일희": "lih@glovis.net"}

TEAM_MEMBERS = sorted(list(TEAM_MEMBERS_MAIL.keys()), reverse=False)

OVERSEAS_RECIPIENTS = {"이종성 책임매니저": "jslee@glovis.net", "H138_jslee": "jslee@glovisala.com",
                       "noname_1": "kkim@glovisala.com", "김정익 책임매니저": "jikim81@glovis.net",
                       "김정익 책임": "jikim81@glovis.net", "문중만 책임": "jmoon@glovisga.com", "noname_2": "djkim@glovis.com.br",
                       "H138_한장수 메일 연락처": "jshan@glovis.sk", "H138_서동은메일연락처": "dongeun.seo@glovis.cz",
                       "최재완": "jaewan.choi@glovis.cz", "엄종율 책임": "uhmt@glovis.net", "이낙원 책임": "leenarkwon@glovis.net",
                       "김성국 책임": "sungkook@glovis.net", "장규복 책임": "gbjang@glovis.net",
                       "H138_장규복": "gbjang@glovis.sk", "noname_3": "kimym724@glovis.com.tr", "하정훈 책임": "Hateng@glovis.net",
                       "김태호 책임": "taeho@glovis.net", "H138_Taeho": "taeho@glovis.mx", "noname_4": "o.son@glovisrus.com",
                       "김기륜 책임": "kilyunj@glovis.net", "엄준용 책임": "u.junyong@glovisrus.com",
                       "추현우 책임": "hw.choo@glovis.mx", "김진수 책임": "jjinss@glovis.net", "신중섭 책임": "jsshin@glovis.net"}

# print(",".join([i for i in OVERSEAS_RECIPIENTS.values()]))

RECIPIENTS_TEST = "lih@glovis.net,asahi8769@gmail.com"
CC_TEST = "asahi8769@gmail.com"
RECIPIENTS_PRODUCTION = ",".join(list(TEAM_MEMBERS_MAIL.values()))

CC_PRODUCTION_NONE = None
CC_PRODUCTION_OVS = ",".join(list(OVERSEAS_RECIPIENTS.values()))

TARGET = {"2020": 40, "2021": 34.4, "2022": 28.3}

CUSTOMERS = ["HAOS", "HMB", "HMMA", "HMMC", "HMMR", "KAGA", "KASK", "KMX", "KIN", "HMMI"]
PERSON_IN_CHARGE = {'HAOS': '안남규', 'HMB': '김형동', 'HMMA': '김지창', 'HMMC': '정성원', 'HMMR': '김시동',
                    'KIN': '주배용', 'KMX': '이태휘', 'KAGA': '정동윤', 'KASK': '이연택', 'HMMI': '강현'}
CUSTOMER_NAME_REVISION = {'KMMG': 'KAGA', 'KMS': 'KASK', 'KMI': 'KIN', 'KMM': 'KMX'}
CUSTOMER_LOCATION = {'HAOS': '구주', 'HMB': '미주', 'HMI': '아주', 'HMMA': '미주', 'HMMC': '구주', 'HMMR': '구주',
                     'HTBC': '아주', 'KIN': '아주', 'KMX': '미주', 'KAGA': '미주', 'KASK': '구주', 'HMMI': '아주'}

SMALL_CUSTOMERS = ['HMI', 'HTBC']
BROADER_CUSTOMERS = CUSTOMERS + ["HMI"]
C_CUSTOMERS = ['HTBC']
OVS_VALID_TYPE = ['MONETARY', 'REPAIR', 'REPLACEMENT', 'INVENTORY', 'EXCHANGE']
OVS_INVALID_TYPE = ["INFORMATION", "PILOT"]

FULL_TYPE = ['기능', '외관', '이종', '치수', '포장']
OVS_FUNCTIONAL = ["기능", "포장"]
NON_APP = ["기능", "포장", "치수"]
GQMS_VALID_TYPE = ['Monetary', 'Repair', 'Replace', 'Exchange']
GQMS_INVALID_TYPE = ['Information', 'Pilot']  # 확인 필요
DEFECT_TYPES = ['종합', '공정', '선별']
OVS_COUNTRIES = ['KR', 'CN', 'IN']
DQY_COUNTRIES = ['Korea', 'China', 'India']
APPEARANCE_TYPE = ['외관', '이종']
DMS_VALID_TYPE = ['교체', '반송', '수정', "폐기", "수량과부족"]
DMS_INVALID_TYPE = ['특채', "PILOT"]
GLOBAL_CODES = ['CN', 'CZ', 'IN', 'MX', 'TR', 'US']
PPR_SUPPLIED_FROM = ["글로비스 한국", "글로비스 중국", "글로비스 인도(N)", "글로비스 인도(A)"]
# INVALID_PPR_SUPPLIED_FROM = ["글로비스 인도(N)", "글로비스 인도(A)"]

PPRS_TO_EXCLUDE = ['K201201497']  # 임시적으로 제외하는 PPR No.

DMS_COMPLETED = ["완료"]
OVS_COMPLETED = ["개선대책확인(해외)", "완료"]
DMS_STD_DATE = "검사일자"
OVD_STD_DATE = "결재일(해외)"
OVD_STD_DATE_V2 = "등록일"

PART_REG_TO_SUBSTITUTE = "(RH|LH|RR|FR|FRT)($|,|-|\s)"
VENDOR_REG_TO_SUBSTITUTE = "(\(.*\)|주식회사|㈜|PVT|LTD|\.|,|LIMITED|유한회사|CO.|PRIVATE|유한책임회사|자동차부품)($|\s|)"

INCOMING_MODE = ["전수검사", "합동검사"]

QUARTERS = {"1분기": ['01', '02', '03'], "2분기": ['04', '05', '06'], "3분기": ['07', '08', '09'], "4분기": ['10', '11', '12']}
QUARTERSV2 = {'01': '1/4', '02': '1/4', '03': '1/4', '04': '2/4', '05': '2/4', '06': '2/4', '07': '3/4', '08': '3/4',
              '09': '3/4', '10': '4/4', '11': '4/4', '12': '4/4'}


GLOBAL_CENTERS = {'울산': '한국', '아산': '한국', '아산2': '한국', '전주': '한국', 'Beijing': '중국', 'Mexico': '한국',
                  'Japan': '한국', 'India': '인도', 'GMX': '한국', 'Turkey': '터키', 'USA': '한국', 'RUSSIA': '한국'}

PACKING_CENTERS = {'직출하(울산)': '한국', '직출하(아산)': '한국', '울산 하나TPS': '한국', '울산 유니팩': '한국',
                   '울산 PLS': '한국', '울산 2포장장': '한국', '울산 1포장장': '한국', '아산 3포장장': '한국', '아산 2포장장': '한국',
                   '아산 1포장장': '한국', '대호물류산업': '한국', '뷰텍 2포장장': '한국', '광주 YSP 2포장장': '한국', '광주 YSP': '한국',
                   '경림(아산)': '한국', 'Yeomsung ATT': '중국', 'Yeomsung': '중국', 'Tianjin wheel': '중국', 'Tianjin': '중국',
                   'Sanghai': '중국', 'Qingtao': '중국', 'Mobis': '중국', 'Mexico(R3)': '한국', 'Japan': '한국', 'India': '인도',
                   'GMX포장장': '한국', 'GLS(평동)': '한국', 'GLS(첨단)': '한국', 'GIN': '인도', 'Chongqing': '중국',
                   'YOUNGSAN(TR)': '터키', 'USA': '한국', 'RUSSIA': '한국', '글로벌 소싱 FCL': '중국', 'Sichuan': '중국',
                   'CKD PACKAGING CENTRE #2': '인도', 'CKD PACKAGING CENTRE #1': '인도', '전주포장장': '한국', '경주 1포장장': '한국'}
