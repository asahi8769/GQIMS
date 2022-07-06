import pandas as pd


def html_style():
    style = """
    <style>

        @import url('https://fonts.googleapis.com/css2?family=Nanum+Gothic:wght@400;700&display=swap');
        
        .mail-text{
        font-family: 'Nanum Gothic', 'Modern H', 'Malgun Gothic', 'sans-serif';
        }
        
        .margin-outlook{
        font-family: 'Nanum Gothic', 'Modern H', 'Malgun Gothic', 'sans-serif';
        color: white;
        font-size: 3px;
        margin-top: 20px;
        }
        
        #ovs_plot1 {
        border: 1px solid #ddd;
        display: block;
        margin: auto;
        
        width: 100%;
        }
        
        #footer {
        border: 1px solid #ddd
        width: 600px;
        margin: auto;
        }
        #footer:hover {
        box-shadow: 0 0 2px 1px rgba(0, 140, 186, 0.5);
        }
        
        table, th, td {
        border-collapse: collapse;
        border: 1px solid #566573;
        border-left-width: 1px;
        font-family: 'Nanum Gothic', 'Modern H', 'Malgun Gothic', 'sans-serif';
        font-size:15px;
        text-align: center;
        }
        
        table{
        width: 70%;
        }
        
        th {
        background-color: #e6f2ff;
        }
        
        tr {
        background-color: #FEF9E7;
        height: 10px;
        }

       td:hover {background-color: #FCF3CF;}
    </style>   
        

        


    """

    return style


if __name__ == "__main__":
    import os

    os.chdir(os.pardir)
    os.chdir(os.pardir)
    os.chdir(os.pardir)

    style = html_style()