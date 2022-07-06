import os, time
from subprocess import Popen, PIPE
from os.path import basename
from zipfile import ZipFile
from functools import reduce
import threading
import pickle
from datetime import datetime


def make_dir(dirname):
    try:
        os.mkdir(dirname)
        print("Directory ", dirname, " Created ")
        return dirname
    except FileExistsError:
        pass


def path_find(name, *paths):
    for path in paths:
        for root, dirs, files in os.walk(path):
            print(root)
            if name in files:
                return os.path.join(root, name)


def subprocess_cmd(command):
    print(command)
    try :
        process = Popen(command, stdout=PIPE, shell=True, universal_newlines=True)
        proc_stdout = process.communicate()[0].strip()
    except Exception as e:
        process = Popen(command, stdout=PIPE, shell=True, universal_newlines=False)
        proc_stdout = process.communicate()[0].strip()
    print(proc_stdout)


def packaging(dir_to_copy, *bindings):
    from tqdm import tqdm
    zipname = fr'{dir_to_copy}\package.zip'

    with ZipFile(zipname, 'w') as zipObj:
        for binding in tqdm(bindings):
            if not os.path.isdir(binding):
                zipObj.write(binding)
            else:
                for folderName, subfolders, filenames in os.walk(binding):
                    for filename in filenames:
                        fn = os.path.join(folderName, filename)
                        zipObj.write(fn, fn)
        print(f'패키징을 완료하였습니다. {zipname}')


def install(lib):
    return f'pip --trusted-host pypi.python.org --trusted-host files.pythonhosted.org --trusted-host pypi.org install {lib}'


def force_reinstall(lib):
    return f'pip --trusted-host pypi.python.org --trusted-host files.pythonhosted.org --trusted-host pypi.org install ' \
           f'"{lib}" --force-reinstall'


def venv_dir(foldername='venv'):
    return os.path.join(os.getcwd(), foldername)


def threading_timer(function, sec, Daemon=True):
    t = threading.Timer(sec, function)
    t.setDaemon(Daemon)
    t.start()
    return t


def show_elapsed_time(function):
    def wrapper(*args, **kwargs):
        start = time.time()
        returns = function(*args, **kwargs)
        print(function.__name__, 'Elapsed {0:02d}:{1:02d}'.format(*divmod(int(time.time() - start), 60)))
        return returns
    return wrapper


def try_until_success(function):
    def wrapper(*args, **kwargs):
        while True:
            try:
                returns = function(*args, **kwargs)
            except Exception as e:
                print(f"{function.__name__} failed({e}). Trying Again...")
                pass
            else:
                return returns
    return wrapper


def flatten(ls):
    """
    ls : list
    flattens multi-dimensional list
    """
    return list(set([i for i in reduce(lambda x, y: x + y, ls)]))


def remove_duplication(ls):
    """
    ls : list
    remove duplicated items while maintaining order
    """
    seen = set()
    seen_add = seen.add
    return [i for i in ls if not (i in seen or seen_add(i))]


def get_norm_slope(ls:list):

    import numpy as np
    mean = np.mean(ls)
    std = np.std(ls)
    if std == 0 :
        std += 0.00001
    normalized = [i for i in [(i - mean) / std for i in ls][0]]
    return np.polyfit([i for i in range(len(normalized))], normalized, 1)[0]


def get_weekly_days_from_today(week, monday_to_saturday=False, regular_month=False, yyyymm=None):
    """
    https://stackoverflow.com/questions/49160939/get-dates-for-last-week-in-python
    :param monday_to_saturday:
    :param week:
    :return: start date, last date
    """
    from datetime import date, timedelta
    from dateutil.relativedelta import relativedelta

    if regular_month:
        next_month = (datetime.strptime(str(yyyymm), '%Y%m') + relativedelta(months=1)).strftime('%Y%m')
        year_start = datetime.strptime(next_month+'01', '%Y%m%d')
    else:
        year_start = date.today()

    starting = 0 if regular_month else 1
    days = -year_start.isoweekday() if monday_to_saturday else starting
    week_start = year_start + timedelta(days=days, weeks=week)
    week_end = week_start + timedelta(days=6)

    return week_start, week_end