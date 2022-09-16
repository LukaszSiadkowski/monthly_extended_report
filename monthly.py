import psycopg2
import pandas as pd
from datetime import datetime, timedelta
import numpy as np
import os
import time
import win32com.client as win32
import datetime
from datetime import date

con = psycopg2.connect(
    dbname='xxxxx',
    host='xxxxxxxx',
    port='xxxxxx',
    user='xxxxxxxxxxx',
    password='xxxxxxxxxxx')


class Query(str):

    
    def run(self, cursor):
        cursor.execute(self)
        return self

    
    def append(self, query_list):
        query_list.append(self)
        return self

   
    def run_append(self, cursor, query_list):
        cursor.execute(self)
        query_list.append(self)
        return self

   
    def run_store(self, cursor):
        time_init= time.time()
        cursor.execute(self)
        df = pd.DataFrame.from_records(
            iter(cursor), columns=[x[0] for x in cursor.description])
        print('execution time (minutes) =', round((time.time()-time_init)/60, 1))
        return df

   
    def run_store_append(self, cursor, query_list):
        cursor.execute(self)
        df = pd.DataFrame.from_records(
            iter(cursor), columns=[x[0] for x in cursor.description])
        query_list.append(self)
        return df



def query_to_df(query, cursor):
    return Query(query).run_store(cursor=cursor)


def inspect(df, cols):
    if len(cols) == 1 or type(cols) == str:
        return df[cols].value_counts()
    elif len(cols)>1:
        return df[cols].head()

cursor = cur = con.cursor()

query = """SELECT snapshot_date ,
empl_login ,
reports_to_login ,
xxxxxxxx ,
xxxxxxxxxxx ,
coalesce (xxxxxx, 0) as xxxxx  ,
coalesce (xxxxxxxxxxx, 0) as xxxxxxxxxxxx ,
....................
FROM xxxxxxxxxxxxxxxxx
where snapshot_date >= DATEADD(month, -1, date_trunc('month', current_date)) and snapshot_date <= date_trunc('month', current_date)-1
and xxxxxxxxxx
and (reports_to_login = 'xxxxxx' or xxxxxxxxxxxxxxx) 
and (xxxxxxxxxxxxxxxx)  
order by empl_login """

query_list = list()

df = Query(query).run_store_append(cursor=cursor,
                                       query_list=query_list)  
df.replace(0, np.nan, inplace=True)

df['xxxxxxx'] = df['xxxxxxxxxx'].fillna(0)


df.insert(loc=6, column='xxxxxxxxx', value=df['x']-df['x'])
df.insert(loc=8, column='x', value=df['x']/df['x'])
df.insert(loc=14, column='x', value=df['x']/df['x'])


df.replace(np.nan, 0, inplace=True)

cols_to_be_numeric = ['xxxxx', 'aaaaaaaa', ..........]
for col in cols_to_be_numeric:
     df[col] = pd.to_numeric(df[col])

df['delete'] = 0

cols_to_be_datetime = ['delete', 'xxxxxxx', ............]
for col in cols_to_be_datetime:
     df[col] = pd.to_datetime(df[col], unit='s').dt.floor('S')
        
cols_to_be_time = ['delete', ......]
for col in cols_to_be_time:
     df[col] = df[col].dt.time

df = df.replace((df.loc[0][34]), '')
df = df.replace(0, '')

df = df.drop(['xxxxxxxx', 'xxxxx', 'delete'] , axis=1)

df.to_csv('file_1.csv', index=False, encoding='utf-8')


cursor = cur = con.cursor()

query = """select date_trunc('day', closed_ts) as date ,
xxxxxxxx ,
xxxxxxxxx ,
case_count
from xxxxxxxxxxxxx
where (xxxxxxxxxxxxxxxxxx)
and date_trunc('day', closed_ts) >= DATEADD(month, -1, date_trunc('month', current_date)) and date_trunc('day', closed_ts) <= date_trunc('month', current_date)-1
and xxxxxx <> 'Phone' 
order by xxxxxxxxxxxx"""

query_list = list()

pano = Query(query).run_store_append(cursor=cursor,
                                       query_list=query_list) 
pano.to_csv('file_2.csv', index=False, encoding='utf-8')


cursor = cur = con.cursor()


query = """select date_trunc('day', create_date) as date ,
xxxxxxxxx ,
xxxxxxxxx ,
case_count
from xxxxxxxxxxxxx 
where date_trunc('day', create_date) >= DATEADD(month, -1, date_trunc('month', current_date)) and date_trunc('day', create_date) <= date_trunc('month', current_date)-1
and (xxxxxxxxxxxxx)"""


query_list = list()

df = Query(query).run_store_append(cursor=cursor,
                                       query_list=query_list)


df.drop_duplicates(subset =["xxxxxxxx", "login"], keep = 'last', inplace = True)
df = df.sort_values('login')
df.to_csv('file_a.csv', index=False, encoding='utf-8')

query = """select date_trunc('day', resolved_date) as date ,
xxxxxxxx ,
xxxxxxxxxx as login,
case_count
from xxxxxxxxxxx 
where date_trunc('day', resolved_date) >= DATEADD(month, -1, date_trunc('month', current_date)) and date_trunc('day', resolved_date) <= date_trunc('month', current_date)-1
and (xxxxxxxxxxxxx)"""

query_list = list()

df = Query(query).run_store_append(cursor=cursor,
                                       query_list=query_list)


df.drop_duplicates(subset =["xxxxxxxxx", "login"], keep = 'last', inplace = True)
df = df.sort_values('login')
df.to_csv('file_b.csv', index=False, encoding='utf-8')

cursor = cur = con.cursor()

query = """select survey_report_dt ,
xxxxxxxxx
xxxxxxxxxxxx
.............
from xxxxxxx
where date_trunc('day', survey_report_dt) >= DATEADD(month, -1, date_trunc('month', current_date)) and date_trunc('day', survey_report_dt) <= date_trunc('month', current_date)-1
and (xxxxxxxxxxxxxxxxxx)
order by survey_report_dt"""

query_list = list()

df = Query(query).run_store_append(cursor=cursor,
                                       query_list=query_list)  #run_store

df.to_csv('file_3.csv', index=False, encoding='utf-8')

cursor = cur = con.cursor()

query = """select survey_report_dt ,
xxxxxxxxx
xxxxxxxxxxxx
.............
from xxxxxxx
where date_trunc('day', survey_report_dt) >= DATEADD(month, -1, date_trunc('month', current_date)) and date_trunc('day', survey_report_dt) <= date_trunc('month', current_date)-1
and (xxxxxxxxxxxxxxxxxx)
order by survey_report_dt"""

query_list = list()

df = Query(query).run_store_append(cursor=cursor,
                                       query_list=query_list)  #run_store

df.to_csv('file_4.csv', index=False, encoding='utf-8')

cursor = cur = con.cursor()

query = """select survey_report_dt ,
xxxxxxxxx
xxxxxxxxxxxx
.............
from xxxxxxx
where date_trunc('day', survey_report_dt) >= DATEADD(month, -1, date_trunc('month', current_date)) and date_trunc('day', survey_report_dt) <= date_trunc('month', current_date)-1
and (xxxxxxxxxxxxxxxxxx)
order by survey_report_dt"""

query_list = list()

df = Query(query).run_store_append(cursor=cursor,
                                       query_list=query_list)  #run_store

df.to_csv('file_5.csv', index=False, encoding='utf-8')

os.system('start "excel" "xxxxxxxxxxxxxxxxxxxxxxx\\Monthly\\example.xlsm"')


today = datetime.date.today()
first = today.replace(day=1)
lastMonth = first - datetime.timedelta(days=1)

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'xxxxxxxxxxxx'
mail.Subject = f' Summary {lastMonth.strftime("%B")}' 
mail.Body = 'xxxxxxx'
mail.HTMLBody = """\
<html>
  <head></head>
  <body>
  Hello Team, <br><br>
  Hope this email finds you well. <br><br>
  
  xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx <br><br>
  
  Please contact me if you will have any questions or concerns. <br><br>
  
  Kind regards, <br>
  Lukasz Siadkowski
  </body>
</html>
"""
mail.Send()
