from matplotlib import font_manager, rc
import matplotlib as mpl
import os


try:
    font_path = "C:/Windows/Fonts/현대하모니+M.ttf"
    font = font_manager.FontProperties(fname=font_path).get_name()
except FileNotFoundError:
    font_path = os.path.join(mpl.get_data_path(), "fonts/ttf/현대하모니+M.ttf")
    font = font_manager.FontProperties(fname=font_path).get_name()

rc('font', family=font)