from utils_.functions import subprocess_cmd
from datetime import datetime
from utils_.functions import make_dir
from utils_.config import REPOSITORY
import os

make_dir(f'clone')

pull_dir = make_dir(f'clone\clone_{datetime.now().strftime("%Y%m%d_%H%M%S")}')
subprocess_cmd(f'git config --global user.name Ilhee Lee')
subprocess_cmd(f'git config --global user.email asahi8769@gmail.com')
subprocess_cmd(f'git clone --depth=1 {REPOSITORY[:-4]} {pull_dir}')

os.startfile(pull_dir)