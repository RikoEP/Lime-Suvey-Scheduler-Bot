# -*- coding: utf-8 -*-
"""
Created on Wed May  6 11:39:35 2020

@author: Riko EP
"""
from selenium import webdriver
import pandas as pd
import numpy as np
from datetime import datetime as dt
import time
import sqlalchemy
from sqlalchemy import create_engine
import os


def check_files():       
    # cek file ada di directory, kalau ada dihapus dahulu
    if os.path.exists('PATH_TO_EXCEL_FILE'):
        os.remove('PATH_TO_EXCEL_FILE')
        print('File has been removed')


def login(url, driver):
    # login ke Lime Survey
    driver.get(url)
    
    driver.find_element_by_id('user').send_keys('USERNAME')
    driver.find_element_by_id('password').send_keys('PASSWORD')
    driver.find_element_by_name('login_submit').click()
    
    time.sleep(1)


def download(url, driver, file_path):
    # click tombol save
    driver.get(url)
    
    driver.find_element_by_id('xls').click()
    driver.find_element_by_id('save-button').click()
    
    while not os.path.exists(file_path):
        time.sleep(1)

    
def download_data():
    # download data survey
    driver = None
    try:
        driver = webdriver.Chrome(executable_path='PATH_TO_WEBDRIVER')
        
        url_login = 'LOGIN_URL'
        url_survey = 'SURVEY_URL'
        save = 'PATH_TO_EXCEL_FILE'
        
        login(url_login, driver)
        download(url_survey, driver, save)
    finally:
        if driver is not None:
            driver.close()
            print('Download Complete')

    
def process_data():
    # membaca data excel
    df = pd.read_excel('PATH_TO_EXCEL_FILE')
    
    # collect data hanya hari ini saja
    constraint = (pd.to_datetime(df['Date submitted']).dt.date == dt.date(dt.now()))
    df = df[constraint].sort_values('Date submitted')
    
    # remove duplicate untuk yg sebelum jam 10
    constraint_duplicate = ((pd.to_datetime(df['Date submitted']).dt.hour <= 10) & (~df['Nama Lengkap Karyawan'].duplicated(keep='last'))) | (pd.to_datetime(df['Date submitted']).dt.hour > 10)
    df = df[constraint_duplicate]
    
    # remove kolom NULL
    list_drop = df.columns[46:67]
    df.drop(list_drop, axis=1, inplace=True)
    
    # rename nama kolom
    a = df.columns[:]
    b = ['B' + str(i) for i in np.arange(1, len(a) + 1)]
    columns = dict(zip(a, b))
    df = df.rename(columns=columns)

    #convert nilai kosong jadi NULL dan ubah tipe data jadi string
    df.fillna(value=sqlalchemy.sql.null(), inplace=True)
    df = df.astype(str)    
    
    print('Processed Complete')
    
    return df

    
def store_data(df):
    # store data ke database
    engine = create_engine('SQL_SERVER_ENGINE')
        
    if 'TABLE_NAME' not in engine.table_names():
        df.to_sql('TABLE_NAME', con=engine)
    elif 'TABLE_NAME' in engine.table_names():
        file = open("notification.txt","r").read()
        
        if file == "Data Stored Successfully at " + str(dt.date(dt.now())):
            df.to_sql('TABLE_NAME', con=engine, if_exists='replace')
        else:
            df.to_sql('TABLE_NAME', con=engine, if_exists='append')
    else:
        df.to_sql('TABLE_NAME', con=engine, if_exists='append')
    
    print('Data Stored Successfully')
    
    
if __name__ == '__main__':
    check_files()
    
    download_data()
    
    df = process_data()

    store_data(df)
    
    df.to_excel('Result' + str(dt.date(dt.now())) + '.xlsx', index=False)

    with open("notification.txt", "w") as file:
        file.write("Data Stored Successfully at " + str(dt.now()))

