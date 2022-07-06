from utils_.config import *
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException,SessionNotCreatedException, WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options as Chrome_options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pyautogui, shutil, calendar, time, os, sys
from utils_.functions import make_dir, try_until_success
import chromedriver_autoinstaller


class GqimsDataRetrieval:
    CHROME_OPTIONS = Chrome_options()
    CHROME_OPTIONS.add_argument("--start-maximized")
    # CHROME_OPTIONS.add_argument("--headless")
    # chromedriver_autoinstaller.install()

    def __init__(self, pre_noti=False):
        make_dir("downloaded")
        self.pre_noti = pre_noti
        try:
            self.driver = webdriver.Chrome(GC_DRIVER, options=self.CHROME_OPTIONS)
        except (SessionNotCreatedException, WebDriverException):
            chromedriver_autoinstaller.install()
            self.driver = webdriver.Chrome(options=self.CHROME_OPTIONS)

        self.driver.get(URL)
        self.download_folder = DOWNLOAD_FOLDER
        files_to_remove = ['OVERSEA_QUALITY_PROBLEM.xlsx', 'RECP_NG_NOTI.xlsx', 'GQMS_PPR.xlsx', '입고검사결과 현황.xlsx']

        for file in files_to_remove:
            if os.path.isfile(os.path.join(self.download_folder, file)):
                os.remove(os.path.join(self.download_folder, file))

    def close(self):
        self.driver.delete_all_cookies()
        self.driver.quit()

    def click_element_id(self, ID, sec):
        while True:
            try:
                element = WebDriverWait(self.driver, sec).until(
                    EC.element_to_be_clickable((By.ID, ID)))
                element.click()
                return element
            except (TimeoutException, ElementClickInterceptedException):
                continue

    def login(self):
        self.click_element_id ('mainframe_VFrameSet_LoginFrame_form_div_login_Edit00InputElementTextBoxElement', 5)
        pyautogui.write(GQIMS_ID)
        self.click_element_id ('mainframe_VFrameSet_LoginFrame_form_div_login_Edit01InputElementTextBoxElement', 5)
        pyautogui.write(GQIMS_PW)
        self.click_element_id('mainframe_VFrameSet_LoginFrame_form_div_login_Button00TextBoxElement', 5)

    @try_until_success
    def set_inr_screen(self):
        self.click_element_id('mainframe_VFrameSet_TopFrame_form_Menu00_RITextBoxElement', 5)
        self.click_element_id('mainframe_VFrameSet_TopFrame_form_Menu00_popupmenu_RI01TextBoxElement', 5)
        self.click_element_id('mainframe_VFrameSet_TopFrame_form_Menu00_popupmenu_popup_RI0102TextBoxElement', 5)

        mode_input = self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winRI0102_form_div_Work_DivSearch_cbo_INSP_STD_comboedit_input', 5)
        self.insured_region_mode_input(mode_input, INCOMING_MODE[1])

        # self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winRI0102_form_div_Work_DivSearch_chk_ALL_chkimg', 5)

    @try_until_success
    def set_dqy_screen(self):
        self.click_element_id('mainframe_VFrameSet_TopFrame_form_Menu00_OQTextBoxElement', 5)
        self.click_element_id('mainframe_VFrameSet_TopFrame_form_Menu00_popupmenu_OQ03TextBoxElement', 5)
        self.click_element_id('mainframe_VFrameSet_TopFrame_form_Menu00_popupmenu_popup_AD0402TextBoxElement', 5)

    @try_until_success
    def set_ppr_screen(self):
        self.click_element_id('mainframe_VFrameSet_TopFrame_form_Menu00_OQTextBoxElement', 5)
        self.click_element_id('mainframe_VFrameSet_TopFrame_form_Menu00_popupmenu_OQ03TextBoxElement', 5)
        self.click_element_id('mainframe_VFrameSet_TopFrame_form_Menu00_popupmenu_popup_AD0401TextBoxElement', 5)

    @try_until_success
    def set_dms_screen(self):
        self.click_element_id('mainframe_VFrameSet_TopFrame_form_Menu00_DQTextBoxElement', 5)
        self.click_element_id('mainframe_VFrameSet_TopFrame_form_Menu00_popupmenu_DQ01TextBoxElement', 5)
        self.click_element_id('mainframe_VFrameSet_TopFrame_form_Menu00_popupmenu_popup_DQ0101TextBoxElement', 5)
        self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winDQ0101_form_div_Work_DivSearch_chk_SEARCH_COND_chkimg', 5)

        std_date = self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winDQ0101_form_div_Work_DivSearch_cbo_DAY_comboedit_input', 5)
        self.insured_region_mode_input(std_date, DMS_STD_DATE)

    @try_until_success
    def set_ovs_screen(self):
        self.click_element_id('mainframe_VFrameSet_TopFrame_form_Menu00_OQTextBoxElement', 5)
        self.click_element_id('mainframe_VFrameSet_TopFrame_form_Menu00_popupmenu_OQ01TextBoxElement', 5)
        self.click_element_id('mainframe_VFrameSet_TopFrame_form_Menu00_popupmenu_popup_OQ0103TextBoxElement', 5)
        self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winOQ0103_form_div_Work_DivSearch00_chk_SEARCH_COND', 5)
        std_date = self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winOQ0103_form_div_Work_DivSearch00_cbo_DAY_comboedit_input', 5)

        if self.pre_noti:
            self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winOQ0103_form_div_Work_DivSearch00_chk_OSEA_APPEAL_INC_YN_chkimg',5) # do not use this line util you know how the check box filters data.
            self.insured_region_mode_input(std_date, OVD_STD_DATE_V2)
        else:
            self.insured_region_mode_input(std_date, OVD_STD_DATE)

    def download_inr_table(self, yyyymm:int):
        first_date = str(yyyymm)[0:4] + str(yyyymm)[4:] + "01"
        last_date = str(yyyymm)[0:4] + str(yyyymm)[4:] + str(calendar.monthrange(int(str(yyyymm)[0:4]), int(str(yyyymm)[4:]))[1])

        date_start = self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winRI0102_form_div_Work_DivSearch_cal_START_calendaredit_input', 5)
        self.insured_yyyymmdd_input(date_start, first_date)

        date_end = self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winRI0102_form_div_Work_DivSearch_cal_END_calendaredit_input', 5)
        self.insured_yyyymmdd_input(date_end, last_date)

        modes = ["전수검사", "합동검사"]
        for mode in modes:
            mode_input = self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winRI0102_form_div_Work_DivSearch_cbo_INSP_STD_comboedit_input',5)
            sequence = self.insured_region_mode_input(mode_input, mode)

            # print(sequence, mode)
            # if sequence == 0 and mode == "합동검사":
            #     self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winRI0102_form_div_Work_DivSearch_chk_ALL_chkimg', 5)
            if sequence == 1 and mode == "전수검사":
                self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winRI0102_form_div_Work_DivSearch_chk_ALL_chkimg', 5)

            self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winRI0102_form_div_Work_G1_btnSearch', 5)
            self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winRI0102_form_div_Work_G1_btnExcelDown',5)
            self.move_file(rf"{self.download_folder}/입고검사결과 현황.xlsx", f"입고검사결과_{yyyymm}_{mode}.xlsx")

    def download_dqy_table(self, yyyymm:int):
        month = self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winAD0402_form_div_Work_DivSearch_cal_YYYYMM_calendaredit_input', 5)
        self.insured_yyyymmdd_input(month, str(yyyymm))

        for corp in PPR_SUPPLIED_FROM:

            supplied_input = self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winAD0402_form_div_Work_DivSearch_cbo_CORP_comboedit_input', 5)
            self.insured_region_mode_input(supplied_input, corp)

            self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winAD0402_form_div_Work_G1_btnSearchTextBoxElement', 5)
            self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winAD0402_form_div_Work_G1_btnExcelDownTextBoxElement', 60)
            self.move_file(rf"{self.download_folder}/GQMS_PPR.xlsx", f"불출수량현황_{yyyymm}_{corp}.xlsx")

    def download_ppr_table(self, yyyymm:int):
        first_date = str(yyyymm)[0:4] + str(yyyymm)[4:] + "01"
        last_date = str(yyyymm)[0:4] + str(yyyymm)[4:] + str(calendar.monthrange(int(str(yyyymm)[0:4]), int(str(yyyymm)[4:]))[1])

        date_start = self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winAD0401_form_div_Work_DivSearch_cal_START_calendaredit_input', 5)
        self.insured_yyyymmdd_input(date_start, first_date)

        date_end = self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winAD0401_form_div_Work_DivSearch_cal_END_calendaredit_input', 5)
        self.insured_yyyymmdd_input(date_end, last_date)

        for corp in PPR_SUPPLIED_FROM:

            supplied_input = self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winAD0401_form_div_Work_DivSearch_cbo_CORP_CD_comboedit_input', 5)
            self.insured_region_mode_input(supplied_input, corp)

            self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winAD0401_form_div_Work_G1_btnSearchTextBoxElement', 5)
            self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winAD0401_form_div_Work_G1_btnExcelDownTextBoxElement', 60)
            self.move_file(rf"{self.download_folder}/GQMS_PPR.xlsx", f"GqmsPPR현황_{yyyymm}_{corp}.xlsx")

    def download_dms_table(self, yyyymm:int):
        first_date = str(yyyymm)[0:4] + str(yyyymm)[4:] + "01"
        last_date = str(yyyymm)[0:4] + str(yyyymm)[4:] + str(calendar.monthrange(int(str(yyyymm)[0:4]), int(str(yyyymm)[4:]))[1])

        date_start = self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winDQ0101_form_div_Work_DivSearch_cal_START_calendaredit_input', 5)
        self.insured_yyyymmdd_input(date_start, first_date)

        date_end = self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winDQ0101_form_div_Work_DivSearch_cal_END_calendaredit_input', 5)
        self.insured_yyyymmdd_input(date_end, last_date)

        self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winDQ0101_form_div_Work_G1_btnSearchTextBoxElement', 5)
        self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winDQ0101_form_div_Work_G1_btnExcelDownTextBoxElement', 60)
        self.move_file(rf"{self.download_folder}/RECP_NG_NOTI.xlsx", f"국내품질정보현황_{yyyymm}.xlsx")

    def download_ovs_table(self, yyyymm: int):
        first_date = str(yyyymm)[0:4] + str(yyyymm)[4:] + "01"
        last_date = str(yyyymm)[0:4] + str(yyyymm)[4:] + str(calendar.monthrange(int(str(yyyymm)[0:4]), int(str(yyyymm)[4:]))[1])

        date_start = self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winOQ0103_form_div_Work_DivSearch00_cal_START_calendaredit_input', 5)
        self.insured_yyyymmdd_input(date_start, first_date)

        date_end = self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winOQ0103_form_div_Work_DivSearch00_cal_END_calendaredit_input', 5)
        self.insured_yyyymmdd_input(date_end, last_date)

        self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winOQ0103_form_div_Work_G1_btnSearch', 5)
        self.click_element_id('mainframe_VFrameSet_HFrameSet_VFrameSet1_WorkFrame_winOQ0103_form_div_Work_G1_btnExcelDown', 60)

        if self.pre_noti:
            self.move_file(rf"{self.download_folder}/OVERSEA_QUALITY_PROBLEM.xlsx", f"해외품질정보현황_등록일_{yyyymm}.xlsx")
        else:
            self.move_file(rf"{self.download_folder}/OVERSEA_QUALITY_PROBLEM.xlsx", f"해외품질정보현황_{yyyymm}.xlsx")

    @staticmethod
    def insured_yyyymmdd_input(element, yyyymmdd:str):
        delay = 0
        while element.get_attribute('value').split(" ")[0].replace("-","") != yyyymmdd:
            element.send_keys(Keys.CONTROL, 'a')
            element.send_keys(Keys.DELETE)

            for char in yyyymmdd:
                element.send_keys(char)
                time.sleep(delay)
            delay += 0.1

    @staticmethod
    def insured_region_mode_input(element, text:str):
        direction = [Keys.DOWN, Keys.UP]
        n = 0
        prev_val = None

        while element.get_attribute('value') != text:
            element.send_keys(direction[n % 2])
            if prev_val == element.get_attribute('value'):
                n += 1
            prev_val = element.get_attribute('value')
        return n % 2

    def move_file(self, file_directory, newname):
        while not os.path.isfile(file_directory):
            continue
        time.sleep(1.5)
        if os.path.isfile(os.path.join('downloaded', newname)):
            os.remove(os.path.join('downloaded', newname))
        shutil.move(file_directory, f"downloaded")
        os.rename(os.path.join('downloaded', file_directory.split(r"/")[-1]), os.path.join('downloaded', newname))

    @staticmethod
    def download_multiple_table(end, function):
        for j in [i for i in range(int(str(end)[:4]+'01'), int(str(end)[:4]+'01'+'13'), 1) if i <= end]:
            function(j)


if __name__ == "__main__":
    os.chdir(os.pardir)

    obj = GqimsDataRetrieval()
    obj.login()

    # obj.set_dms_screen()
    # obj.download_dms_table(202109)
    # obj.download_multiple_table(202109, obj.download_dms_table)

    obj.set_ovs_screen(pre_noti=False)
    obj.download_ovs_table(202110)
    # obj.download_multiple_table(202109, obj.download_ovs_table)

    # obj.set_ppr_screen()
    # obj.download_ppr_table(202109)
    # obj.download_multiple_table(202102, obj.download_ppr_table)

    # obj.set_dqy_screen()
    # obj.download_dqy_table(202109)
    # obj.download_multiple_table(202109, obj.download_dqy_table)

    # obj.set_inr_screen()
    # obj.download_inr_table(202109)
    # obj.download_multiple_table(202109, obj.download_inr_table)

    obj.close()