CUSTOMERS = ["HAOS", "HMB", "HMMA", "HMMC", "HMMR", "KAGA", "KASK", "KMX", "KIN", "HMMI", 'HMI', 'HTBC']  # 대상고객사
INVALID_PPR_SUPPLIED_FROM = ["글로비스 인도(N)", "글로비스 인도(A)"]  # 제외 공급법인

PRE_SHIPPING_INCLUSIVE = False  # 출하검사 포함여부

KEYS = ['고객사', "부품번호", '업체코드', '업체명']

OVS_INVALID_TYPE = ["INFORMATION", "PILOT"]  # 해외품질정보 제외 조건
PRE_NOTI = True  # 해외미결제건 포함여부
OVS_REFERENCE = "발생일자"  # 해외품질정보 1년치 데이터 기준일

DMS_INVALID_TYPE = ['특채', "PILOT"]  # 국내품질정보 제외 조건
DMS_REFERENCE = "검사일자"  # 국내품질정보 1년치 데이터 기준일

ICM_YEAR = 202112  # 입고기준년도 최종월, NONE
YEARLY_DELTA_OVS = 1
YEARLY_DELTA_DMS = 1
MONTHLY_DELTA_ICM = 12
S_DATA_RANGE_YEAR = 2


RANK_CRITERION_B = 0.2  # S, A, C, D 제외 후 B 결정하는 불량률 상위 %
N_OF_D = 1  # 고객사별 D 갯수
N_OF_C = 3  # 고객사별 C 갯수
N_OF_P = 3  # 워스트 P 갯수
N_OF_M = 20  # 워스트 MIX 갯수