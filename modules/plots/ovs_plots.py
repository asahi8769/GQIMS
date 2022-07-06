import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
from utils_.config import CUSTOMERS, TARGET
import pandas as pd
import numpy as np
import math, os
from utils_.functions import make_dir
from datetime import datetime

# print(mpl.get_data_path())
# print(mpl.get_cachedir())
# print(mpl._get_version())

import utils_.matplotlib_config as matplotlib_config


# 컬러코드: https://htmlcolorcodes.com/
# 폰트 적용안될 시: https://jinyes-tistory.tistory.com/70

def get_ax1_data(df, yyyymm):
    x = ['전체'] + CUSTOMERS

    supply_q = ["" for _ in x]
    appea_q = ["" for _ in x]
    supply_q[0] = df.loc[('공급망 품질', '종합', '전사', 'PPM'), "종합"]
    appea_q[0] = df.loc[('공급망 외관품질', '종합', '전사', 'PPM'), "종합"]

    for n, customer in enumerate(CUSTOMERS):
        supply_q[n + 1] = df.loc[(customer, '공급망 품질(종합)', '종합', 'PPM'), "종합"]
        appea_q[n + 1] = df.loc[(customer, '공급망 품질(외관)', '종합', 'PPM'), "종합"]

    return x, supply_q, appea_q


def get_ax2_data(df, yyyymm, target):
    months = [i for i in range(int(str(yyyymm)[:4] + '01'), int(str(yyyymm)[:4] + '13'), 1)]
    x = ['누적'] + months

    total_q = ["" for _ in x]
    korea_q = ["" for _ in x]
    global_q = ["" for _ in x]

    total_q[0] = df.loc[('공급망 외관품질', '종합', '전사', 'PPM'), "종합"]
    korea_q[0] = df.loc[('공급망 외관품질', '종합', '한국', 'PPM'), "종합"]
    global_q[0] = df.loc[('공급망 외관품질', '종합', '글로벌소싱(중국/인도)', 'PPM'), "종합"]

    total_qty = df.loc[('공급망 외관품질', '종합', '전사', '불량수량'), "종합"]
    dqy_qty = df.loc[('불출수량', '계', '계', '불출수량'), "종합"]
    meet_qty = target * dqy_qty / 1000000
    exceeded = int(total_qty - meet_qty)

    # print(f"{exceeded}개만큼 불량수량이 초과되었습니다.")

    for n, month in enumerate(months):
        total = df.loc[('공급망 외관품질', '종합', '전사', 'PPM'), month]
        korea = df.loc[('공급망 외관품질', '종합', '한국', 'PPM'), month]
        global_ = df.loc[('공급망 외관품질', '종합', '글로벌소싱(중국/인도)', 'PPM'), month]
        total_q[n + 1] = total if not (type(total) == str or math.isnan(total)) else 0
        korea_q[n + 1] = korea if not (type(korea) == str or math.isnan(korea)) else 0
        global_q[n + 1] = global_ if not (type(global_) == str or math.isnan(global_)) else 0

    return x, total_q, korea_q, global_q, exceeded


def get_ax3_data(df, yyyymm):
    months = [i for i in range(int(str(yyyymm)[:4] + '01'), int(str(yyyymm)[:4] + '13'), 1)]
    x = ['누적'] + months

    korea_q = ["" for _ in x]
    china_q = ["" for _ in x]
    india_q = ["" for _ in x]

    korea_q[0] = df.loc[('공급망 외관품질', '종합', '한국', 'PPM'), "종합"]
    china_q[0] = df.loc[('공급망 외관품질', '종합', '중국', 'PPM'), "종합"]
    india_q[0] = df.loc[('공급망 외관품질', '종합', '인도', 'PPM'), "종합"]

    for n, month in enumerate(months):
        korea = df.loc[('공급망 외관품질', '종합', '한국', 'PPM'), month]
        china = df.loc[('공급망 외관품질', '종합', '중국', 'PPM'), month]
        india = df.loc[('공급망 외관품질', '종합', '인도', 'PPM'), month]
        korea_q[n + 1] = korea if not (type(korea) == str or math.isnan(korea)) else 0
        china_q[n + 1] = china if not (type(china) == str or math.isnan(china)) else 0
        india_q[n + 1] = india if not (type(india) == str or math.isnan(india)) else 0

    return x, korea_q, china_q, india_q


def get_ax4_data(df):
    app_q = df.loc[('공급망 외관품질', '종합', '전사', '불량수량'), "종합"]
    dim_q = df.loc[('공급망 품질(치수)', '종합', '전사', '불량수량'), "종합"]
    fun_q = df.loc[('공급망 품질(기능)', '종합', '전사', '불량수량'), "종합"]

    df_type = pd.DataFrame({"유형": ['외관' ,'치수', '기능'], "불량수량": [app_q, dim_q, fun_q]})
    df_type.sort_values("불량수량", inplace=True, ascending=False)

    return df_type


def get_ax5_6_data(df):
    df_dqy = pd.DataFrame({"고객사": CUSTOMERS})
    slot_all = []
    slot_app = []

    for customer in CUSTOMERS:
        slot_all.append(df.loc[(customer, '공급망 품질(종합)', '종합', '불량수량'), "종합"])
        slot_app.append(df.loc[(customer, '공급망 품질(외관)', '종합', '불량수량'), "종합"])

    df_dqy['공급망불량'] = slot_all
    df_dqy['외관불량'] = slot_app

    return df_dqy


def get_ax1(fig, gs, x, supply_q, appea_q, yyyymm):
    r = np.arange(len(x))
    ax1 = fig.add_subplot(gs[0:1, 0:3])

    ax1.set_xticks(r)
    ax1.set_xticklabels(x)
    ax1.tick_params(axis='x', rotation=45)

    ax1.set(yticklabels=[])
    ax1.tick_params(left=False, bottom=False)
    ax1.spines["right"].set_visible(False)
    ax1.spines["left"].set_visible(False)
    ax1.spines["top"].set_visible(False)

    width = 0.40
    margin = 0

    sq = ax1.bar(r - width / 2 - margin, supply_q, width, label='종합', color='#AED6F1', edgecolor='grey')
    aq = ax1.bar(r + width / 2 + margin, appea_q, width, label='외관', color='#F5B7B1', edgecolor='grey')

    for p in sq:
        height = p.get_height()
        ax1.annotate(f'{height:,}', xy=(p.get_x() + p.get_width() / 2, height), xytext=(-2, 3),
                     textcoords="offset points", ha='center', va='bottom')

    for p in aq:
        height = p.get_height()
        ax1.annotate(f'{height:,}', xy=(p.get_x() + p.get_width() / 2, height), xytext=(2, 3), textcoords="offset points",
                     ha='center', va='bottom')

    ax1.set_title(f"{str(yyyymm)[0:4]}년도 {str(yyyymm)[4:]}월 고객사별 누적 PPM", loc="left", pad=15)
    ax1.legend(prop={'size': 12})


def get_ax2(fig, gs, x, total_q, korea_q, global_q, yyyymm, target, exceeded):
    r = np.arange(len(x))
    ax2 = fig.add_subplot(gs[1:2, 0:6])

    ax2.set_xticks(r)
    ax2.set_xticklabels(x)
    ax2.tick_params(axis='x', rotation=45)

    ax2.set(yticklabels=[])
    ax2.tick_params(left=False, bottom=False)
    ax2.spines["right"].set_visible(False)
    ax2.spines["left"].set_visible(False)
    ax2.spines["top"].set_visible(False)

    width = 0.24
    margin = 0.04

    tq = ax2.bar(r - width - margin, total_q, width, label='종합', color='#AED6F1', edgecolor='grey')
    kq = ax2.bar(r, korea_q, width, label='한국', color='#ABEBC6', edgecolor='grey')
    gq = ax2.bar(r + width + margin, global_q, width, label='글로벌', color='#F9E79F', edgecolor='grey')

    for n, p in enumerate(tq):
        height = p.get_height()
        ax2.annotate(f'{height:,}' if (x[n] == "누적" or x[n] <= yyyymm) else "",
                     xy=(p.get_x() + p.get_width() / 2, height), xytext=(-3, 3), textcoords="offset points",
                     ha='center', va='bottom')

    for n, p in enumerate(kq):
        height = p.get_height()
        ax2.annotate(f'{height:,}' if (x[n] == "누적" or x[n] <= yyyymm) else "",
                     xy=(p.get_x() + p.get_width() / 2, height), xytext=(0, 3), textcoords="offset points", ha='center',
                     va='bottom')

    for n, p in enumerate(gq):
        height = p.get_height()
        ax2.annotate(f'{height:,}' if (x[n] == "누적" or x[n] <= yyyymm) else "",
                     xy=(p.get_x() + p.get_width() / 2, height), xytext=(3, 3), textcoords="offset points", ha='center',
                     va='bottom')

    color = "#EC7063" if target <= total_q[0] else "#5499C7"
    score = round(((target - total_q[0])/target+1)*100, 1)

    # if type != "y":

    ax2.axhline(y=target, color=color, linestyle='--', label="목표", lw=0.5)

    ax2.annotate(f'{target:,}\nPPM', xy=(-1, target), xytext=(0, 3), textcoords="offset points", ha='left',
                 va='bottom', color="#2C3E50", fontsize=12)
    ax2.annotate(f'{score:,}%', xy=(-1, target), xytext=(0, -7), textcoords="offset points", ha='left',
                 va='top', color="#F50707" if target <= total_q[0] else "#0712F5", fontsize=11)

    title = f"한국/글로벌 월별 외관 PPM (누적: {total_q[0]:,}PPM)" if exceeded <= 0 else f"한국/글로벌 월별 외관 PPM (누적: {total_q[0]}PPM, {exceeded:,}ea 초과)"
    ax2.set_title(title, loc="left", pad=15)
    ax2.legend(prop={'size': 10})


def get_ax3(fig, gs, x, korea_q, china_q, india_q, yyyymm):
    r = np.arange(len(x))
    ax3 = fig.add_subplot(gs[2:3, 0:6])

    ax3.set_xticks(r)
    ax3.set_xticklabels(x)
    ax3.tick_params(axis='x', rotation=45)

    ax3.set(yticklabels=[])
    ax3.tick_params(left=False, bottom=False)
    ax3.spines["right"].set_visible(False)
    ax3.spines["left"].set_visible(False)
    ax3.spines["top"].set_visible(False)

    width = 0.24
    margin = 0.04

    kq = ax3.bar(r - width - margin, korea_q, width, label='한국', color='#ABEBC6', edgecolor='grey')
    cq = ax3.bar(r, china_q, width, label='중국', color='#F5B7B1', edgecolor='grey')
    iq = ax3.bar(r + width + margin, india_q, width, label='인도', color='#F9E79F', edgecolor='grey')

    for n, p in enumerate(kq):
        height = p.get_height()
        ax3.annotate(f'{height:,}' if (x[n] == "누적" or x[n] <= yyyymm) else "",
                     xy=(p.get_x() + p.get_width() / 2, height), xytext=(-3, 3), textcoords="offset points",
                     ha='center', va='bottom')

    for n, p in enumerate(cq):
        height = p.get_height()
        ax3.annotate(f'{height:,}' if (x[n] == "누적" or x[n] <= yyyymm) else "",
                     xy=(p.get_x() + p.get_width() / 2, height), xytext=(0, 3), textcoords="offset points", ha='center',
                     va='bottom')

    for n, p in enumerate(iq):
        height = p.get_height()
        ax3.annotate(f'{height:,}' if (x[n] == "누적" or x[n] <= yyyymm) else "",
                     xy=(p.get_x() + p.get_width() / 2, height), xytext=(3, 3), textcoords="offset points", ha='center',
                     va='bottom')

    ax3.set_title(f"한국/중국/인도 월별 외관 PPM", loc="left", pad=15)
    ax3.legend(prop={'size': 10})


def get_ax4(fig, gs, df_type):
    ax4 = fig.add_subplot(gs[0:1, 3:4])

    cmap = plt.get_cmap("tab20c")
    outer_colors = cmap(np.array([4 * i % 20 for i in range(len(df_type))]))
    data = df_type['불량수량'].tolist()

    wedges, texts = ax4.pie(data, radius=1, colors=outer_colors,
                            startangle=90,
                            counterclock=False,
                            wedgeprops=dict(width=0.3, edgecolor='grey', linewidth=0.7))

    label = df_type['유형'].tolist()

    total = sum(data)
    share = [data[0]/total*100, data[1]/total*100, data[2]/total*100]
    share = [int(round(i, 0)) for i in share]

    for index, value in enumerate(texts):
        value.set_fontsize(10)
        value.set_text(f"{label[index]}\n({share[index]}%)")
    ax4.text(0, 0, f"불량구분", ha='center', va='center', fontsize=12)


def get_ax5(fig, gs, dqy_df):
    dqy_df_ = dqy_df.copy()
    dqy_df_.sort_values("공급망불량", inplace=True, ascending=False)

    data = dqy_df_['공급망불량'].tolist()
    label = dqy_df_['고객사'].tolist()
    total = sum(data)
    share = [int(round(i/total*100, 0)) for i in data]

    ax5 = fig.add_subplot(gs[0:1, 4:5])
    cmap = plt.get_cmap("tab20c")
    outer_colors = cmap(np.array([4 * i % 20 for i in range(len(label))]))

    wedges, texts = ax5.pie(data, radius=1, colors=outer_colors,
                            startangle=90,
                            counterclock=False,
                            wedgeprops=dict(width=0.3, edgecolor='grey', linewidth=0.7))

    for index, value in enumerate(texts):
        value.set_fontsize(8)
        value.set_text(f"{label[index]}\n({share[index]}%)")
        if share[index] <= 4:
            value.set_visible(False)
    ax5.text(0, 0, f"공급망불량", ha='center', va='center', fontsize=12)


def get_ax6(fig, gs, dqy_df):
    dqy_df_ = dqy_df.copy()
    dqy_df_.sort_values("외관불량", inplace=True, ascending=False)

    data = dqy_df_['외관불량'].tolist()
    label = dqy_df_['고객사'].tolist()
    total = sum(data)
    share = [int(round(i/total*100, 0)) for i in data]

    ax6 = fig.add_subplot(gs[0:1, 5:6])
    cmap = plt.get_cmap("tab20c")
    outer_colors = cmap(np.array([4 * i % 20 for i in range(len(label))]))

    wedges, texts = ax6.pie(data, radius=1, colors=outer_colors,
                            startangle=90,
                            counterclock=False,
                            wedgeprops=dict(width=0.3, edgecolor='grey', linewidth=0.7))

    for index, value in enumerate(texts):
        value.set_fontsize(7)
        value.set_text(f"{label[index]}\n({share[index]}%)")
        if share[index] <= 4:
            value.set_visible(False)

    ax6.text(0, 0, f"외관불량", ha='center', va='center', fontsize=12)


def draw_plot1(df, yyyymm):

    target = TARGET[str(yyyymm)[:4]]

    x1, supply_q, appea_q = get_ax1_data(df, yyyymm)
    x2, total_q, korea_q, global_q, exceeded = get_ax2_data(df, yyyymm, target)
    x3, korea_q, china_q, india_q = get_ax3_data(df, yyyymm)
    df_type = get_ax4_data(df)
    df_dqy = get_ax5_6_data(df)

    fig = plt.figure(constrained_layout=False, figsize=(15, 8))
    gs = GridSpec(3, 6, figure=fig)

    get_ax1(fig, gs, x1, supply_q, appea_q, yyyymm)
    get_ax2(fig, gs, x2, total_q, korea_q, global_q, yyyymm, target, exceeded)
    get_ax3(fig, gs, x3, korea_q, china_q, india_q, yyyymm)
    get_ax4(fig, gs, df_type)
    get_ax5(fig, gs, df_dqy)
    get_ax6(fig, gs, df_dqy)

    fig.tight_layout()

    folder_name = os.path.join("spawn", "plots", datetime.now().strftime("%Y-%m-%d"))
    make_dir(folder_name)

    img_name = os.path.abspath(os.path.join(folder_name, f"{yyyymm}_plot1.png"))
    plt.savefig(img_name)
    return plt, img_name


if __name__ == "__main__":
    from modules.ppm.ovs_ppm import MonthlyOVSppmCalculation

    os.chdir(os.pardir)
    os.chdir(os.pardir)

    yyyymm = 202110

    ppm_ovs = MonthlyOVSppmCalculation(yyyymm)
    df = ppm_ovs.get_ppm()

    # df = pd.read_excel("spawn/(202109)KD종합품질지수(GQMS기준)_2021-11-25_16_32_23.xlsx",
    #                    sheet_name='해외품질지수(GQMS+공급망)', index_col=[1, 2, 3, 4], skiprows=3)

    plt, _ = draw_plot1(df, yyyymm)
    plt.show()

    # xfile = openpyxl.load_workbook(r'spawn\test6.xlsx')
    # sheet = xfile.create_sheet("요약(그래프)", 0)
    # # sheet.title = "요약(그래프)"
    #
    # sheet.column_dimensions["A"].width = 2
    # sheet.sheet_view.showGridLines = False
    #
    # img = openpyxl.drawing.image.Image(f"spawn/plots/{202109}_plot1.png")
    # img.anchor = 'B2'
    # sheet.add_image(img)
    # xfile.save(r'spawn\test6.xlsx')
    #
    # os.startfile(r'spawn\test6.xlsx')


