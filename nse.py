import tkinter as tk
import datetime
from datetime import timedelta, date
import requests
import pandas as pd
import io
from urllib.request import urlopen
from zipfile import ZipFile


HEIGHT = 500
WIDTH = 600
root = tk.Tk()
root.title='NSE EOD Downloader'

canvas = tk.Canvas(root, height=HEIGHT, width=WIDTH, background='lightgreen')
canvas.pack()


def get_weather():
    now = datetime.datetime.now()



    
    for i in range (0,10):
        date =now - datetime.timedelta(days = i)
        try:
            Date_to_Find=date.strftime("%d-%b-%Y")
            url =('https://www1.nseindia.com/content/fo/fii_stats_'+Date_to_Find+'.xls')
            res = requests.get(url)
            res.raise_for_status()
            lk=res.content
            toread = io.BytesIO()
            toread.write(lk) 
            toread.seek(0)
            df = pd.read_excel(toread)
            df.reset_index(drop=True, inplace=True)
            column_names = df.columns.values
            column_names[1:] = None
            df.columns = column_names

            df.to_excel("fii_stats_.xlsx",sheet_name='sheet1',index=False)
            
            break 

        except requests.exceptions.HTTPError:
            None
            

    Date_to_Find=date.strftime("%d%b%Y")
    dbl=date.strftime("%b")
    resp = urlopen('https://www1.nseindia.com/content/historical/EQUITIES/2020/'+dbl.upper()+'/cm'+Date_to_Find.upper()+'bhav.csv.zip')
    zipfile = ZipFile(io.BytesIO(resp.read()))
    ab=zipfile.open(zipfile.namelist()[0])
    df = pd.read_csv(ab)

    df.to_excel("volume.xlsx",sheet_name='volume',index=False) 


    Date_to_Find=date.strftime("%d%b%Y")
    dbl=date.strftime("%b")
    resp = urlopen('https://www1.nseindia.com/content/historical/DERIVATIVES/2020/'+dbl.upper()+'/fo'+Date_to_Find.upper()+'bhav.csv.zip')
    zipfile = ZipFile(io.BytesIO(resp.read()))
    ab=zipfile.open(zipfile.namelist()[0])
    df = pd.read_csv(ab)

    df.to_excel("derivatives.xlsx",sheet_name='data',index=False) 
    df.to_excel("derivatives1.xlsx",sheet_name='data',index=False) 
    df.to_excel("derivatives2.xlsx",sheet_name='data',index=False) 


    Date_to_Find=date.strftime("%d%m%Y")
    url =('https://www1.nseindia.com/content/nsccl/fao_participant_vol_'+Date_to_Find+'.csv')
    res = requests.get(url)
    res.raise_for_status()
    lk=res.content
    toread = io.BytesIO()
    toread.write(lk) 
    toread.seek(0)
    vol = pd.read_csv(toread)

    vol.to_excel("fao_participant_vol_.xlsx", sheet_name='fao_participant_vol_',index=False) 


    Date_to_Find=date.strftime("%d%m%Y")
    url =('https://www1.nseindia.com/content/nsccl/fao_top10cm_to_30042020.csv')
    res = requests.get(url)
    res.raise_for_status()
    lk=res.content
    toread = io.BytesIO()
    toread.write(lk) 
    toread.seek(0)
    df = pd.read_csv(toread)

    df.to_excel("fao_top10cm_.xlsx",sheet_name='fao_top10cm_',index=False) 


    Date_to_Find=date.strftime("%d%m%Y")
    resp = urlopen('https://www1.nseindia.com/archives/fo/mkt/fo'+Date_to_Find+'.zip')
    zipfile = ZipFile(io.BytesIO(resp.read()))
    ab=zipfile.open(zipfile.namelist()[3])
    bc=zipfile.open(zipfile.namelist()[7])
    df = pd.read_csv(ab)
    df1 = pd.read_csv(bc)

    df.reset_index().to_excel("futidx.xlsx",sheet_name='futidx',index=False) 
    df1.reset_index().to_excel("optidx.xlsx",sheet_name='optidx',index=False) 


    url =('https://www1.nseindia.com/products/content/sec_bhavdata_full.csv')
    res = requests.get(url)
    res.raise_for_status()
    lk=res.content
    toread = io.BytesIO()
    toread.write(lk) 
    toread.seek(0)
    df = pd.read_csv(toread)

    df.to_excel("price.xlsx",sheet_name='price',index=False) 
    df.to_excel("price1.xlsx",sheet_name='price1',index=False) 


    Date_to_Find=date.strftime("%d%m%Y")
    url =('https://www1.nseindia.com/content/nsccl/fao_participant_oi_'+Date_to_Find+'.csv')
    res = requests.get(url)
    res.raise_for_status()
    lk=res.content
    toread = io.BytesIO()
    toread.write(lk) 
    toread.seek(0)
    oi = pd.read_csv(toread)

    oi.to_excel("fao_participant_oi_.xlsx",sheet_name='fao_participant_oi_',index=False) 

        


frame = tk.Frame(root, bg='#80c1ff', bd=5)
frame.place(relx=0.5, rely=0.1, relwidth=0.75, relheight=0.1, anchor='n')

entry = tk.Label(frame,text="Click on download to start", font=40)
entry.place(relwidth=0.65, relheight=1)

button = tk.Button(frame, text="Download", font=40, command=lambda: get_weather())
button.place(relx=0.7, relheight=1, relwidth=0.3)

lower_frame = tk.Frame(root, bg='#80c1ff', bd=10)
lower_frame.place(relx=0.5, rely=0.25, relwidth=0.75, relheight=0.6, anchor='n')

label = tk.Label(lower_frame)
label.place(relwidth=1, relheight=1)

root.mainloop()
