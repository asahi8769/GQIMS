CURRENT_QUARTER = 202201, 202202, 202203
PAST_QUARTER = 202110, 202111, 202112


PRE_NOTI = False  # 해외결재중 포함여부

INVALID_PPR_SUPPLIED_FROM = ["글로비스 인도(N)", "글로비스 인도(A)", '글로비스 터키']  # 비계열에서만 제외
OVS_INVALID_TYPE = ["INFORMATION", "PILOT"]  # 해외품질정보 제외 조건
DMS_INVALID_TYPE = ['특채', "PILOT"]

CUSTOMER_PICS = {'HAOS': ['안남규', 1], 'HMB': ['김형동', 1], 'HMI': ['강현', 1], 'HMMA': ['김지창', 1], 'HMMC': ['정성원', 1],
                 'HMMI': ['강현', 1], 'HMMR': ['김시동', 1], 'HTBC': ['이영록', 1], 'KAGA': ['정동윤', 1], 'KASK': ['이연택', 1],
                 'KIN': ['주배용', 1], 'KMC': ['주배용', 0], 'KMX': ['이태휘', 1], 'DWI': ['정성원', 0]}  # 담당자, 메이저공장여부(0, 1)


THREEPL_CUSTOMER_PIC = "이영록"
# 담당자가 지정되지 않은 고객사는 비계열로 간주해서 THREEPL_CUSTOMER_PIC로 지정,
# 비계열 고객사로 간주된 건중 INVALID_PPR_SUPPLIED_FROM에 속하는 공급법인 건은 무시됨

NON_DESIGNATED_CUSTOMER_PIC = [THREEPL_CUSTOMER_PIC, 0]

APPEARANCE_TYPE = ['외관', '이종']
NON_APPEARANCE = ["기능", "포장", "치수"]

CM_COMPLETE = ['완료', '개선대책확인(해외)'] # 국내품질정보: 완료 / 해외품질정보: 완료, 개선대책확인(해외)

KEYS = ['담당자_']  # ['고객사', '담당자_']

ADDITIONAL_HOLIDAYS = ['20220222', '20220309']
# 라이브러리로 필터링 되지 않는 추가 휴일

PPRS_TO_EXCLUDE = ["K220300923", "K220300554", "K220300926"] # 인사상시평가에서만 임시적으로 제외하는 PPR No.
# K220300923, K220300926의 경우 컨테이너 이종건으로 불량 등록방안에 대해 추가 협의 필요함
# K220300554의 경우 인포메이션으로 변경 협의 완료함