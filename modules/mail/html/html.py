import pandas as pd
from utils_.functions import make_dir
from modules.mail.html.style import html_style
from utils_.config import TARGET
from datetime import datetime

# datetime.strptime(download_date, "%Y-%m-%d").strftime("%Y년 %m월 %d일")


def margin_for_outlook():
    """
    temporary funtion to mimic margin as outlook strips off margins.
    :return:
    """
    return '<p class="margin-outlook">&nbsp</p>'


def boilerplate_table(yyyymm:int, df, download_date, fixed_dqy):

    score = df.loc[('공급망 외관품질', '종합', '전사', '달성률'), "종합"]
    ppm = df.loc[('공급망 외관품질', '종합', '전사', 'PPM'), "종합"]
    dqy_info = "" if fixed_dqy else "  ▶  전월불출수량기준 예상실적"

    table_html = rf"""
        
        <table>
            <tr>
                <th>데이터기준일</th>
                <th colspan="8">{download_date.split(',')[-1]}{dqy_info}</th>
            </tr>
            <tr>
                <td rowspan="2">품질실적</td>
                <td rowspan="2">목표</td>
                <td rowspan="2" style="font-weight: 700;">{TARGET[str(yyyymm)[:4]]}PPM</td>
                <td rowspan="2">달성률</td>
                <td rowspan="2" style="font-weight: 700; color: {"#EC7063" if score <= 100 else "#2E86C1"};">{score}%</td>
                <td>외관품질실적</td>
                <td style="font-weight: 700;">{df.loc[('공급망 외관품질', '종합', '전사', 'PPM'), yyyymm]}PPM</td>
                <td>누적</td>
                <td style="font-weight: 700; color: {"#EC7063" if score <= 100 else "#2E86C1"};">{ppm}PPM</td>
            </tr>
            <tr>
                <td>공급망품질실적</td>
                <td style="font-weight: 700;">{df.loc[('공급망 품질', '종합', '전사', 'PPM'), yyyymm]}PPM</td>
                <td>누적</td>
                <td style="font-weight: 700;">{df.loc[('공급망 품질', '종합', '전사', 'PPM'), "종합"]}PPM</td>
            </tr>
        </table>
    """

    return table_html


def index_html(yyyymm: int, df, download_date, fixed_dqy=True):
    html = rf"""
    
        <!DOCTYPE html>
        <html lang="ko" dir="ltr">
        <head>
            {html_style()}
            <meta charset="euc-kr">
            <title>{yyyymm} 품질지수</title>
        </head>
        
        <p class="mail-text">안녕하세요, 이일희입니다.</br></br>{str(yyyymm)[:4]}년도 {str(yyyymm)[4:6]}월 해외/국내 품질지수 공유드립니다.</br></br>업무에 참고부탁 드립니다. 감사합니다.</br></p>
        {margin_for_outlook()}
        {boilerplate_table(yyyymm, df, download_date, fixed_dqy)}
        {margin_for_outlook()}
        <img id="ovs_plot1" src="cid:ovs_plot1" alt="ovs_plot1"/>
        {margin_for_outlook()}
        <img id="footer" src="cid:footer" alt="footer" />
        
        </html>
    
    """

    make_dir(r"spawn\html")
    html_file = r"spawn\html\index.html"

    with open(html_file, "w") as file:
        file.write(html)

    return html, html_file


if __name__ == "__main__":
    import os

    os.chdir(os.pardir)
    os.chdir(os.pardir)
    os.chdir(os.pardir)

    yyyymm = 202109

    img_1 = os.path.abspath(os.path.join("spawn", "plots", f"{yyyymm}_plot1.png"))
    img = [img_1]

    df = pd.read_excel("spawn/(202109)KD종합품질지수(GQMS기준)_2021-10-31_20_58_41.xlsx",
                       sheet_name='해외품질지수(GQMS+공급망)', index_col=[1, 2, 3, 4], skiprows=3)

    download_date = '2021-01-05'

    html, html_file = index_html(202109, df, download_date, fixed_dqy=False)
    os.startfile(html_file)