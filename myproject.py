import pandas as pd
import matplotlib.pyplot as plt
import xlwings as xw
import os

db_uri = os.environ.get('DB_URI', 'sqlite://')
query = '''
select year, quarter, metric_value from company_metric
    where ticker = '{}'
    order by year, quarter;
'''


def hello_world():
    wb = xw.Book.caller()
    wb.sheets.active.range('A1').value = 'Hello World!'


def get_metric():
    sht = xw.Book.caller().sheets['database']
    ticker = sht.range('B1').value
    df = pd.read_sql(query.format(ticker), con=db_uri)
    sht.range('A4').expand().clear_contents()
    sht.range('A4').value = df.values


@xw.func
def double_sum(x, y):
    """Return twice the sum of the two arguments"""
    return 2 * (x + y)


@xw.func
def myplot(n):
    """Sample code to plot a specified range"""
    sht = xw.Book.caller().sheets.active
    fig = plt.figure()
    plt.plot(range(int(n)))
    sht.pictures.add(fig, name='MyPlot', left=500, top=0, update=True)
    return 'Plotted with n={}'.format(n)
