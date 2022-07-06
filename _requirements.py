from utils_.functions import subprocess_cmd, install, force_reinstall
from utils_.config import VENV_64_DIR, SCRIPTS_DIR
import os

subprocess_cmd(rf'cd {SCRIPTS_DIR} && python -m venv {SCRIPTS_DIR} && activate')
# subprocess_cmd(f'cd {SCRIPTS_DIR} & {force_reinstall("pandas==1.2.3")}')
subprocess_cmd(f'cd {SCRIPTS_DIR} & {force_reinstall("pandas")}')
# subprocess_cmd(f'cd {SCRIPTS_DIR} & {force_reinstall("numpy==1.20.0")}')
# subprocess_cmd(f'cd {SCRIPTS_DIR} & {install("xlrd")}')
# subprocess_cmd(f'cd {SCRIPTS_DIR} & {install("openpyxl")}')
# subprocess_cmd(f'cd {SCRIPTS_DIR} & {force_reinstall("matplotlib==3.4.3")}')
# subprocess_cmd(f'cd {SCRIPTS_DIR} & {install("seaborn")}')
# subprocess_cmd(f'cd {SCRIPTS_DIR} & {install("selenium")}')
# subprocess_cmd(f'cd {SCRIPTS_DIR} & {install("pyautogui")}')
# subprocess_cmd(f'cd {SCRIPTS_DIR} & {install("tqdm")}')
# subprocess_cmd(f'cd {SCRIPTS_DIR} & {install("sendgrid")}')
# subprocess_cmd(f'cd {SCRIPTS_DIR} & {install("xlsxwriter")}')
# subprocess_cmd(f'cd {SCRIPTS_DIR} & {install("chromedriver-autoinstaller")}')
# subprocess_cmd(f'cd {SCRIPTS_DIR} & {install("scipy")}')
# subprocess_cmd(f'cd {SCRIPTS_DIR} & {install("pyinstaller")}')
# subprocess_cmd(f'cd {SCRIPTS_DIR} & {install("office365-rest-client")}')
# subprocess_cmd(f'cd {SCRIPTS_DIR} & {install("pytimekr")}')
# subprocess_cmd(f'cd {SCRIPTS_DIR} & {install("holidays")}')
# subprocess_cmd(f'cd {SCRIPTS_DIR} & {install("dash")}')
# subprocess_cmd(f'cd {SCRIPTS_DIR} & {install("jupyter-dash")}')

