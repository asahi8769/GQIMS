import os
from utils_.functions import packaging, subprocess_cmd, make_dir
from utils_.config import SCRIPTS_DIR

make_dir('package')
subprocess_cmd(rf'cd {SCRIPTS_DIR} && python -m venv {SCRIPTS_DIR} && activate')
file_name_py = os.path.join(os.getcwd(), 'main.py')
file_name_exe = os.path.join(SCRIPTS_DIR, 'dist', 'main.exe')
icon_name = os.path.join(os.getcwd(), 'app.ico')

if os.path.exists(file_name_exe):
    os.remove(file_name_exe)

freeze_command = f'pyinstaller.exe --onefile --icon={icon_name} {file_name_py}'
subprocess_cmd(f'cd {SCRIPTS_DIR} & {freeze_command}')

if os.path.exists(file_name_exe):
    subprocess_cmd(f'copy {file_name_exe} {os.getcwd()}')
    packaging('package', 'images', 'downloaded', 'templates',  'main.exe', 'chromedriver.exe', 'database.db')
    os.startfile('package')
else:
    print("pyinstaller 구동 실패")


# os.startfile('main.exe')

