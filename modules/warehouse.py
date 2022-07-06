import os
import shutil
from datetime import date

from tqdm import tqdm

from utils_.functions import make_dir


def store_raw_data(new_folder):
    make_dir(f"downloaded/{new_folder}")
    files_to_maintain = ['국내입고품질상세정보.xlsx', '해외입고품질상세정보.xlsx', '포장장위치구분정보.xlsx']
    for file in tqdm(os.listdir('downloaded')):

        original = os.path.join('downloaded', file)
        if os.path.isdir(original) or file in files_to_maintain:
            continue
        if os.path.isfile(os.path.join('downloaded', new_folder, file)):
            os.remove(os.path.join('downloaded', new_folder, file))
        try:
            shutil.move(original, f"downloaded/{new_folder}")
        except PermissionError as e:
            print(e)


if __name__ == "__main__":
    os.chdir(os.pardir)
    store_raw_data(str(date.today()))
