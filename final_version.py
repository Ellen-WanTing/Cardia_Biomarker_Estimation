# -*- coding: utf-8 -*-
"""

@author: Wan-Ting Tsai, Yu Lan

ps: Create a file name (ID) on your local machine for the excel file that we export from the code.
"""

import fitparse
import matplotlib.pyplot as plt
import sys
from tkinter import filedialog
from tkinter.filedialog import asksaveasfilename
from tkinter.filedialog import askdirectory
from tkinter.filedialog import askopenfilenames
from tkinter import *
import os
import pandas as pd 
import numpy as np
import itertools
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import load_workbook
from collections import defaultdict

import datetime
import time
sys._enablelegacywindowsfsencoding() # gegen Fehlerausgabe bei ü 

import pandas as pd  
import dateutil.parser as parser

import json
from datetime import datetime, timedelta
import os

from hrvanalysis import get_time_domain_features, get_frequency_domain_features, get_poincare_plot_features
from hrvanalysis import plot_poincare
from hrvanalysis import remove_outliers, remove_ectopic_beats, interpolate_nan_values
import pyhrv.tools as tools
import pyhrv.time_domain as ts
import pyhrv.nonlinear as nl
import neurokit2 as nk
import scipy
import biosppy
import statistics


from dfply import *
import pingouin as pg
import matplotlib.pyplot as plt
import seaborn as sns
import statsmodels.formula.api as smf
import csv

import statsmodels.api as sm
from statsmodels.formula.api import ols
plt.rcParams["font.weight"] = "bold"
plt.rcParams["figure.titlesize"] ='x-large'
plt.rcParams["figure.titleweight"] ='bold'
plt.rcParams["axes.labelweight"] = "bold"
plt.rcParams["axes.titleweight"] = "bold"
plt.rcParams["axes.titlesize"] = 16
plt.rcParams['axes.labelsize'] = 14
plt.rcParams['axes.titlesize'] = 14
plt.rcParams['xtick.labelsize']=14
plt.rcParams['ytick.labelsize']=14
pd.set_option('display.max_rows',None)
pd.set_option('display.max_columns',20)

def load_paper_protocol(path):
    intervention_module = pd.read_excel(path, sheet_name='Interventionsmodul', engine='openpyxl')
    output_dataframe = intervention_module
    return output_dataframe

mypath = 'C:/Users/Ellen Tsai/venv/seminar/'
df = load_paper_protocol(mypath + 'paper_protocol_intervention_module.xlsx')

joint_string = []



print('Number of rows：',len(df))
df.drop(['Unnamed: 14'],axis=1,inplace=True)  
#df.drop(['Unnamed: 14','Unnamed: 15','Unnamed: 16','Unnamed: 17','Unnamed: 18',
#         'Unnamed: 19','Unnamed: 20','Unnamed: 21'],axis=1,inplace=True)   #drop those empty columns

Blank_data_row = 0
df_pre = pd.DataFrame()
for df_i in range(0,len(df)):
    if df.loc[df_i, 'Pulsuhr Start'] == None or \
            df.loc[df_i, 'Bemerkungen'] == 'no data in file' or \
            df.loc[df_i, 'Bemerkungen'] == 'in der Mitte abgebrochen ' or \
            df.loc[df_i, 'Bemerkungen'] == 'no data; Uhr hat sich nicht verbunden':
        #There mixed both None and NaN in the 'Bemerkungen' column blank elements
        Blank_data_row = Blank_data_row +1
        print('The %dth raw data with empty:\n'%df_i,df.loc[df_i,:])
        df.drop(df_i,0,index=None,columns=None,inplace=True)  #drop those useless rows
print('Number of useless data rows:',Blank_data_row)
df_pre = df.reset_index(drop=True)   #data cleaning
print('\nValid rows：%d'%len(df_pre))
#print('\n\nShow df_pre:\n', df_pre)

new_datalist_dataframe = pd.DataFrame(columns=['Sorting with date'])
new_list = []

for i in range(0, len(df_pre)):
    counter = 0
    container = []
    container_2 = []
    if df_pre['Pulsuhr Nr.'][i] == 'RH1':
        string = 'Rhythm 24 7002-'
        string = string + (df_pre['Datum'][i].strftime("%Y%m%d"))  # string of year, month and day
        joint_string.append(string)
        for r, d, f in os.walk(mypath + 'scosche_rri_data'):
            for file in f:
                if file.startswith(string):
                    counter = counter + 1
                    list_lst = os.path.join(r, file)
                    container.append(list_lst)  #container has list of 2 stings if there were 2 measurements on the same day
                    if len(container) > 1:
                        #print('%dth row showed a duplicated: '%(i),container)
                        for j in range(0, len(container)):
                            Zeituntershied = df_pre['Zeitunterschied'][i]  # Summer/Winter Time
                            df_pre_timetostr = df_pre['Pulsuhr Start'][i].strftime("%H%M%S")  # 'Pulsuhr Start' is type of datetime.time, cant be modified by datetime.timedelta. So transform .time->string
                            df_pre_strtodatetime = datetime.strptime(df_pre_timetostr,"%H%M%S")  # str->datetime.datetime
                            df_pre_strtodatetime = df_pre_strtodatetime - timedelta(hours=Zeituntershied)  # Modifying the ‘Pulsuhr Start’ to the correct time (thus we can correspond to .fit root)
                            temp_time_str = df_pre_strtodatetime.strftime("%H%M%S")  # After modifying, datetime->string
                            #print('temp_time_str: ',temp_time_str)
                            st_dataframe_time = str(container[j][-10:-8])
                            if temp_time_str.startswith(st_dataframe_time):
                                #print(container[j])
                                #print('container[%d]:'%(j),str(container[j][82:86]),'.startswith detected successfully!')
                                new_datalist_dataframe.loc[i,'Sorting with date'] = container[j]
                    else:
                        new_datalist_dataframe.loc[i, 'Sorting with date'] = container[0]
    elif df_pre['Pulsuhr Nr.'][i] == 'RH2 ':  #Still a space after RH2
        string = 'Rhythm 24 7048-'
        string = string + (df_pre['Datum'][i].strftime("%Y%m%d"))
        joint_string.append(string)
        for r, d, f in os.walk(mypath + 'scosche_rri_data'):
             for file_2 in f:
                 if file_2.startswith(string):
                     counter = counter + 1
                     list_lst_2 = os.path.join(r, file_2)
                     container_2.append(list_lst_2)
                     #print('container_2: ',container_2)
                     if len(container_2) > 1:
                        #print('%dth row showed a duplicated: '%(i),container_2)
                        for k in range(0, len(container_2)):
                            Zeituntershied = df_pre['Zeitunterschied'][i]  # Summer/Winter Time
                            df_pre_timetostr = df_pre['Pulsuhr Start'][i].strftime("%H%M%S")  # 'Pulsuhr Start' is type of datetime.time, cant be modified by datetime.timedelta. So transform .time->string
                            df_pre_strtodatetime = datetime.strptime(df_pre_timetostr,"%H%M%S")  # str->datetime.datetime
                            df_pre_strtodatetime = df_pre_strtodatetime - timedelta(hours=Zeituntershied)  # Modifying the ‘Pulsuhr Start’ to the correct time (thus we can correspond to .fit root)
                            temp_time_str_2 = df_pre_strtodatetime.strftime("%H%M%S")  # After modifying, datetime->string
                            st_dataframe_time_2 = str(container_2[k][-10:-8])
                            #print('line %d: %dth temp_time_str: '%(i,k),temp_time_str_2)
                            #print('line %d: %dth st_dataframe_time_2: '%(i,k),st_dataframe_time_2)
                            if temp_time_str_2.startswith(st_dataframe_time_2):
                                new_datalist_dataframe.loc[i,:] = container_2[k]
                     else:
                        new_datalist_dataframe.loc[i, 'Sorting with date'] = container_2[0]

pd.set_option('max_colwidth',100)  #no ellipsis
print(new_datalist_dataframe)
print('length of new dataframe: ',len(new_datalist_dataframe))
SK1_dataframe = pd.DataFrame()
SK2_dataframe = pd.DataFrame()
REST_dataframe = pd.DataFrame()

dd = defaultdict(list)
for k, va in [(v,i) for i, v in enumerate((df_pre['ID'].tolist()))]:
    dd[k].append(va)
print('Patients and corresponding experiments: ',dd)
pat_list = list(dd.keys())
print('Patients\' ID list； ' ,pat_list)
#list_rawdata =['Session1_time','Session1_hrv','Session1_hr','Session2_time','Session2_hrv','Session2_hr','Session3_time','Session3_hrv','Session3_hr','Session4_time','Session4_hrv','Session4_hr']
for itt_pat_id in pat_list:
    globals()['df_'+'Pat_ID_'+ str(itt_pat_id)] = pd.DataFrame(index=range(0,2500))

all_measurment =[]
for k in range(0,len(new_datalist_dataframe)):
    list_data1 = []  
    # Load the FIT file
    fitfile = fitparse.FitFile(new_datalist_dataframe['Sorting with date'][k])
    
    for record in fitfile.get_messages('record'):
    
        for record_data in record:

        # Print the records name, value and units 
            if record_data.units:
                pass
                #print(" * %s: %s %s" % (
                   # record_data.name, record_data.value, record_data.units,
                #))
            else:
                pass
               # print(" * %s: %s" % (record_data.name, record_data.value))
       
    
    messages = fitfile.messages # get all the messages and save them in a list of dicts
    
    data=[]
    for i in range(0,len(messages)):
        #print(messages[i].get_values())  
        data.append(messages[i].get_values())
    list_data1.append(data)
    
    
    list_data = [{k: v for k, v in d.items() if k == 'timestamp' or k=='time' or k=='heart_rate'} for d in list(itertools.chain(*list_data1))]
    #print(list_data)
    
    df_data =pd.DataFrame(filter(None,list_data))  #remove the empty dict in list

    df_data['combin']=df_data['time'].fillna(60/df_data['heart_rate'])
    
    df_t=df_data['combin'].dropna().explode() #hrv #s
    
    df_all= df_data.explode('combin').reset_index(drop=True)
    
    #########last few data in timestamp are nan #########
    last_valid_timestamp = 0
    NaT_counter = 0
    for itt_check in range(len(df_all) - 1, 0, -1):
        if pd.isnull(df_all.loc[itt_check, 'timestamp']):
            NaT_counter -= 1
        else:
            last_valid_timestamp = df_all.loc[itt_check, 'timestamp']
            break
    print('last_valid_timestamp: ', str(last_valid_timestamp))
    print('NaT_counter: ', NaT_counter)
    for itt_NaT in range(NaT_counter, 0):  # NaT_counter -8->-1
        print(itt_NaT)
        NaT_timestamp = len(df_all) + itt_NaT
        print(NaT_timestamp)
        Valid_timestamp = len(df_all) + itt_NaT - 1
        print(Valid_timestamp)
        df_all.loc[NaT_timestamp, 'timestamp'] = df_all.loc[Valid_timestamp, 'timestamp'] + timedelta(
            seconds=df_all.loc[Valid_timestamp, 'combin'])
    
    total_second = (df_all.loc[len(df_all)-1,'timestamp'] - df_all.loc[0,'timestamp']).seconds
    print('Total seconds from start to end: ',total_second)
    steplength_ts = total_second/(len(df_all))
    
    
    t_to_ms=[]                 # s to ms
    for t in df_t.tolist():
        s= (t*1000)
        t_to_ms.append(s)
   
    
    instant_hr =[]     
    
    for hrv in t_to_ms:
        new_hr= int(60000/hrv)   
        instant_hr.append(new_hr)
        
    timestamp_dropna=df_all['timestamp'].dropna().to_frame()
    #print(timestamp_dropna['timestamp'].iloc[1].minute)
    if timestamp_dropna['timestamp'].iloc[0].minute > timestamp_dropna['timestamp'].iloc[1].minute:  #deal with .fit file 55 
    
       date_time_ref = datetime.strptime(str(timestamp_dropna['timestamp'].iloc[1]), "%Y-%m-%d %H:%M:%S") #<class 'time.struct_time'> 
       
    else:
        date_time_ref = datetime.strptime(str(timestamp_dropna['timestamp'].iloc[0]), "%Y-%m-%d %H:%M:%S")

    
    
    #correct the day time
    Zeituntershied_k = df_pre['Zeitunterschied'][k]
    date_time_ref= date_time_ref + timedelta(hours=Zeituntershied_k)
    
        
    #find the missing date in .fit file, if there is no date time us excel protocol as benchmark
    timestamp_correction=[]
    timestamp_correction_new=[]
    if date_time_ref.date()==df_pre['Datum'][k].date(): #compare the date
        timestamp_correction=[]
        for times in df_t:
            delta =timedelta(0,times)
            date_time_ref=date_time_ref+delta
            timestamp_correction.append(date_time_ref)
        print('same')
    
    elif k==0: #deal with the first .fit file
         excel_date = datetime.strptime(str(df_pre['Datum'][k]), "%Y-%m-%d %H:%M:%S")
         excel_starttime = time.strptime(str(df_pre['Pulsuhr Start'][k]), "%H:%M:%S")
         df_ts_new = excel_date +timedelta(hours=excel_starttime.tm_hour, minutes=excel_starttime.tm_min, seconds=excel_starttime.tm_sec)
         print(df_ts_new)
         
         timestamp_correction_new=[]
         for times in df_t:
             delta =timedelta(0,times)
             df_ts_new=df_ts_new+delta
             timestamp_correction_new.append(df_ts_new)
         print('the first file')
        
    else:   
         excel_date = datetime.strptime(str(df_pre['Datum'][k]), "%Y-%m-%d %H:%M:%S")
         excel_starttime = time.strptime(str(df_pre['Pulsuhr Start'][k]), "%H:%M:%S")
         df_ts_new = excel_date +timedelta(hours=excel_starttime.tm_hour, minutes=date_time_ref.timetuple().tm_min, seconds=date_time_ref.timetuple().tm_sec)
         print(df_ts_new)
         
         timestamp_correction_new=[]
         for times in df_t:
             delta =timedelta(0,times)
             df_ts_new=df_ts_new+delta
             timestamp_correction_new.append(df_ts_new)
         print('no date')
         
     
    timestamp_new=pd.Series(timestamp_correction_new if not timestamp_correction  else timestamp_correction)
    
    #resampling the date time 
    resampled_timestamp_list = []
    for rs_i in range(0,len(df_all)):
        if rs_i == 0:
            resampled_timestamp_list.append(timestamp_new.loc[0])
        
        else:
            resampled_timestamp_list.append(resampled_timestamp_list[rs_i-1] + timedelta(seconds=steplength_ts))
            
    se_ts = pd.Series(resampled_timestamp_list) #Used to easily view its data
   
      
    data_dataframe=pd.DataFrame(list(zip(se_ts,t_to_ms,instant_hr)),columns=['timestamp','hrv[ms]','instant_hr[bpm]'])
    
    #make sure there is no outlier which will affect the data
    hrv_describe=data_dataframe['hrv[ms]'].describe()
 
     
    hrv_rmoutlier = remove_outliers(data_dataframe['hrv[ms]'],low_rri=np.percentile(data_dataframe['hrv[ms]'], 0.27).tolist() ,high_rri=np.percentile(data_dataframe['hrv[ms]'], 99.73).tolist())
  
    #print(pd.DataFrame(hrv_rmoutlier).iloc[0].isnull() == True)
   
    if pd.DataFrame(hrv_rmoutlier).iloc[0].isnull().any() == True:
        hrv_interpolated = pd.DataFrame(interpolate_nan_values(rr_intervals=hrv_rmoutlier,interpolation_method="quadratic")).fillna(method='bfill').stack().tolist() 
        print('backnan')
    #elif hrv_rmoutlier[-1] == None:
    else:
        hrv_interpolated = pd.DataFrame(interpolate_nan_values(rr_intervals=hrv_rmoutlier,interpolation_method="quadratic")).fillna(method='ffill').stack().tolist() 
        print('frontnan')
   
  
    
    nn_intervals_list = remove_ectopic_beats(rr_intervals=hrv_interpolated, method="malik")
    
    interpolated_nn_intervals = pd.DataFrame(interpolate_nan_values(rr_intervals= nn_intervals_list,interpolation_method="quadratic")).fillna(method='ffill')
    
    #plt.plot(interpolated_nn_intervals)
    data_dataframe['hrv[ms]']= interpolated_nn_intervals

     ################ Segment ################
    K1_HRV = []
    K1_HR = []
    K2_HRV = []
    K2_HR = []
    Rest_HRV = []
    Rest_HR = []
    SK1_index_list=[]
    SK2_index_list=[]
    #Rest_index_list=[]

    for itt_ts in range(0,len(data_dataframe)):
        timestamp_datetime = str(data_dataframe.loc[itt_ts,'timestamp'])
        #print('%dth timestamp'%itt_ts,timestamp_datetime)
        str_timeHM = timestamp_datetime[11:16].replace(':','')
        #print('%dth slice:'%itt_ts,str_timeHM)
        SK1 = df_pre.loc[k,'Start Konfrontation 1'].strftime("%H:%M:%S")
        SK1_str = SK1[0:5].replace(':','')
        EK1 = df_pre.loc[k,'Ende Konfrontation 1'].strftime("%H:%M:%S")
        EK1_str = EK1[0:5].replace(':', '')
        SK2 = df_pre.loc[k,'Start Konfrontation 2'].strftime("%H:%M:%S")
        SK2_str = SK2[0:5].replace(':', '')
        EK2 = df_pre.loc[k,'Ende Konfrontation 2'].strftime("%H:%M:%S")
        EK2_str = EK2[0:5].replace(':', '')
        #Segment strategy：
        if SK1_str <= str_timeHM <EK1_str:
            K1_HRV.append(data_dataframe.loc[itt_ts,'hrv[ms]'])
            K1_HR.append(data_dataframe.loc[itt_ts,'instant_hr[bpm]'])
            
            SK1_index_list.append(itt_ts)

        elif SK2_str <= str_timeHM < EK2_str:
            K2_HRV.append(data_dataframe.loc[itt_ts,'hrv[ms]'])
            K2_HR.append(data_dataframe.loc[itt_ts, 'instant_hr[bpm]'])
            
            SK2_index_list.append(itt_ts)

        else:
            Rest_HRV.append(data_dataframe.loc[itt_ts, 'hrv[ms]'])
            Rest_HR.append(data_dataframe.loc[itt_ts, 'instant_hr[bpm]'])
            
          
  
    data_dataframe.loc[0,'SK1_Start_index'] = SK1_index_list[0]
    
 
    data_dataframe.loc[0, 'SK1_End_index'] = SK1_index_list[-1]
    data_dataframe.loc[0, 'SK2_Start_index'] = SK2_index_list[0]
    data_dataframe.loc[0, 'SK2_End_index'] = SK2_index_list[-1]


    dict_index = {'Gruppe':df_pre.loc[k,'Gruppe'],'ID':df_pre.loc[k,'ID'],'Expostionsnr.':df_pre.loc[k,'Expostionsnr.']}
    dict_index_df=pd.DataFrame(dict_index,index=[0])
    dict_index_df=dict_index_df[['Gruppe','ID','Expostionsnr.']]

  
    #get domain in each segment         
    sk1_nn_interval_td=pd.DataFrame([get_time_domain_features(K1_HRV)])#[["sdnn","sdsd","rmssd","nni_50","pnni_50"]]
    sk1_nn_interval_fd=pd.DataFrame([get_frequency_domain_features(K1_HRV)])[["lf","hf","lf_hf_ratio"]]
    sk1_nn_interval_nld=pd.DataFrame([get_poincare_plot_features(K1_HRV)])
    sk1_parameters=pd.concat([sk1_nn_interval_td,sk1_nn_interval_fd,sk1_nn_interval_nld], axis=1)




    sk2_nn_interval_td=pd.DataFrame([get_time_domain_features(K2_HRV)])#[["sdnn","sdsd","rmssd","nni_50","pnni_50"]]
    sk2_nn_interval_fd=pd.DataFrame([get_frequency_domain_features(K2_HRV)])[["lf","hf","lf_hf_ratio"]]
    sk2_nn_interval_nld=pd.DataFrame([get_poincare_plot_features(K2_HRV)])   
    sk2_parameters=pd.concat([sk2_nn_interval_td,sk2_nn_interval_fd,sk2_nn_interval_nld], axis=1)



    Rest_nn_interval_td=pd.DataFrame([get_time_domain_features(Rest_HRV)])#[["sdnn","sdsd","rmssd","nni_50","pnni_50"]]
    Rest_nn_interval_fd=pd.DataFrame([get_frequency_domain_features(Rest_HRV)])[["lf","hf","lf_hf_ratio"]]
    Rest_nn_interval_nld=pd.DataFrame([get_poincare_plot_features(Rest_HRV)])    
    rest_parameters=pd.concat([Rest_nn_interval_td,Rest_nn_interval_fd,Rest_nn_interval_nld], axis=1)


    segment_parameters=pd.DataFrame(['SK1','SK2','REST'])
    segment_parameters.columns=['segment_name']
    
    hrv_parameters=pd.concat([sk1_parameters,sk2_parameters,rest_parameters],axis=0).reset_index(drop=True)
    
    data_parameters_df = pd.concat([data_dataframe,segment_parameters,hrv_parameters], axis=1)
    all_measurment.append(data_parameters_df)


    SK1_dataframe = SK1_dataframe.append(pd.concat([dict_index_df,sk1_parameters],axis=1).reset_index(drop=True))
    SK1_dataframe = SK1_dataframe.reset_index(drop=True)
    SK2_dataframe = SK2_dataframe.append(pd.concat([dict_index_df, sk2_parameters],axis=1).reset_index(drop=True))
    SK2_dataframe = SK2_dataframe.reset_index(drop=True)
    REST_dataframe = REST_dataframe.append(pd.concat([dict_index_df, rest_parameters],axis=1).reset_index(drop=True))
    REST_dataframe = REST_dataframe.reset_index(drop=True)
    
       

    
    #export a excel file for each patient
    if os.path.exists(mypath + "ID/" + str(df_pre.loc[k, 'ID']) + ".xlsx"):
        print("ID%d's xlsx had been created before" % df_pre.loc[k, 'ID'])
        wb = load_workbook(mypath + "ID/" + str(df_pre.loc[k, 'ID']) + ".xlsx")
        writer = pd.ExcelWriter(mypath + "ID/" + str(df_pre.loc[k, 'ID']) + ".xlsx", engine='openpyxl')
        writer.book = wb
        data_parameters_df.to_excel(writer, sheet_name="Expostionsnr.%d" % df_pre.loc[k, 'Expostionsnr.'], index=False)
        writer.save()

    else:
        print("Creat an excel for patient ID%d" % df_pre.loc[k, 'ID'])
        IDpath = mypath + "ID/" + str(df_pre.loc[k, 'ID']) + ".xlsx"
        excelfile = data_parameters_df.to_excel(IDpath, sheet_name="Expostionsnr.%d" % df_pre.loc[k, 'Expostionsnr.'], index=0)
    
    
    for itt_dictkey in range(len(dd)):
        if k in list(dd.values())[itt_dictkey]:
            temp_pat_id = list(dd.keys())[itt_dictkey]  #pat id
            temp_ses = str(df_pre.loc[k,'Expostionsnr.'])
            starttime=data_dataframe.loc[0,'SK1_Start_index']
            endtime=data_dataframe.loc[0,'SK2_End_index']
            temp_df = data_dataframe.loc[starttime:endtime,'timestamp':'instant_hr[bpm]']
            series_time = data_dataframe.loc[starttime:endtime,'timestamp']
            series_hrv = data_dataframe.loc[starttime:endtime,'hrv[ms]'].astype(float)
            series_hr = data_dataframe.loc[starttime:endtime,'instant_hr[bpm]'].astype(float)
            series_time.reset_index(drop=True,inplace=True)
            series_hrv.reset_index(drop=True,inplace=True)
            series_hr.reset_index(drop=True,inplace=True)
            globals()['df_' + 'Pat_ID_' + str(temp_pat_id)][str('Session' + temp_ses + '_time')] = series_time
            globals()['df_' + 'Pat_ID_' + str(temp_pat_id)][str('Session' + temp_ses + '_hrv')] = series_hrv
            globals()['df_' + 'Pat_ID_' + str(temp_pat_id)][str('Session' + temp_ses + '_instant_hr')] = series_hr
            
            break

    

######## Trendline ########
def linear_polyfit(time,data):
    z1 = np.polyfit(x=time, y=data, deg=1)
    poly_fuc1 = np.poly1d(z1)
    poly_data = poly_fuc1(time)
    poly_resample_data = poly_fuc1(list(range(0,1300)))
    return poly_data, poly_resample_data   #[0]->data [1]->resampled_data

def hrv_y_axis_range(hrvdata,valid=True):  #range of fill_betweenx function
    if valid == True:
        maxlist = []
        minlist = []
        for i in range(len(hrvdata)):
            maxlist.append(max(list(hrvdata.values())[i]) + 20)
            minlist.append(min(list(hrvdata.values())[i]) - 20)
        tempmaxvalue = max(maxlist)
        tempminvalue = min(minlist)
        return np.arange(tempminvalue, tempmaxvalue)
    elif valid == False:   #only one valid session
        maxvalue = max(list(hrvdata.values())[0].tolist()) + 20
        minvalue = min(list(hrvdata.values())[0].tolist()) - 20
        return np.arange(minvalue, maxvalue,1)

def hr_y_axis_range(hrdata,valid=True):
    if valid == True:
        maxlist = []
        minlist = []
        for i in range(len(hrdata)):
            maxlist.append(max(list(hrdata.values())[i]) + 5)
            minlist.append(min(list(hrdata.values())[i]) - 5)
        tempmaxvalue = max(maxlist)
        tempminvalue = min(minlist)
        return np.arange(tempminvalue, tempmaxvalue)
    elif valid == False:   #only onw valid session
        maxvalue = max(list(hrdata.values())[0].tolist()) + 5
        minvalue = min(list(hrdata.values())[0].tolist()) - 5
        return np.arange(minvalue, maxvalue,1)


def fig_plot(timerangelist, hrv, hrv_mean,hr, hr_mean, raw_hrv, raw_hr, pat_name,valid=True):
    
    fig, axs = plt.subplots(2, 1, figsize=(12.8,10.24))
    fig.suptitle(pat_name,fontsize=16,fontweight="bold")
    axs[0].tick_params(labelsize=14)
    axs[1].tick_params(labelsize=14)   
    axs[0].set_ylabel('NNI [ms]',fontsize=16)
    axs[1].set_ylabel('Instant HR[bpm]',fontsize=16)
    axs[1].set_xlabel('Seconds after start measurement[s]',fontsize=16)
    length = len(timerangelist)
    if valid == True:
        axs[0].fill_betweenx(y=hrv_y_axis_range(hrv), x1=0, x2=600, facecolor='xkcd:creme', linewidth=1.5,zorder=1,label='Exposure1')  #hrv plot K1 group area
        axs[0].fill_betweenx(y=hrv_y_axis_range(hrv), x1=700, x2=1400, facecolor='xkcd:ice', linewidth=1.5,zorder=1,label='Exposure2')   #hrv plot K2 group area
        axs[1].fill_betweenx(y=hr_y_axis_range(hr), x1=0, x2=600, facecolor='xkcd:creme', linewidth=1.5,zorder=1,label='Exposure1')   #hr plot K1 group area
        axs[1].fill_betweenx(y=hr_y_axis_range(hr), x1=700, x2=1400, facecolor='xkcd:ice', linewidth=1.5,zorder=1,label='Exposure2')   #hr plot K2 group area
        axs[0].plot(list(range(0,1300)), list(hrv_mean),color='xkcd:blue violet' ,linewidth=3.0, linestyle='-', label='NNI mean',zorder=5)   #mean hrv
        axs[1].plot(list(range(0,1300)), list(hr_mean), color='xkcd:blue violet',linewidth=3.0, linestyle='-', label='Instant HR mean',zorder=5)   #mean hr
        axs[0].set_ylim(hrv_y_axis_range(hrv)[0],hrv_y_axis_range(hrv)[-1])
        axs[1].set_ylim(hr_y_axis_range(hr)[0],hr_y_axis_range(hr)[-1])
    elif hrv_mean==None and hr_mean==None:
        axs[0].fill_betweenx(y=hrv_y_axis_range(hrv,False), x1=0, x2=600, facecolor='xkcd:creme', zorder=1)
        axs[0].fill_betweenx(y=hrv_y_axis_range(hrv,False), x1=700, x2=1400, facecolor='xkcd:ice', zorder=1)
        axs[1].fill_betweenx(y=hr_y_axis_range(hr,False), x1=0, x2=600, facecolor='xkcd:creme', zorder=1) #creme, pale green
        axs[1].fill_betweenx(y=hr_y_axis_range(hr,False), x1=700, x2=1400, facecolor='xkcd:ice', zorder=1) #ice, light sky blue
        axs[0].set_ylim(hrv_y_axis_range(hrv,False)[0],hrv_y_axis_range(hrv,False)[-1])
        axs[1].set_ylim(hr_y_axis_range(hr,False)[0],hr_y_axis_range(hr,False)[-1])
    else:
        pass
    for i in range(length):  #every sessions: polyfit line and scatter
        line,=axs[0].plot((list(timerangelist.values())[i]).tolist(), list(hrv.values())[i], '-', label=list(hrv.keys())[i],zorder=5)
        axs[0].scatter((list(timerangelist.values())[i]).tolist()[::200], list(hrv.values())[i][::200],zorder=2, color = line.get_color())   #scatter on linear regression line
        line1,=axs[1].plot((list(timerangelist.values())[i]).tolist(), list(hr.values())[i], '-', label=list(hr.keys())[i],zorder=5)
        axs[1].scatter((list(timerangelist.values())[i]).tolist()[::200], list(hr.values())[i][::200],zorder=2, color = line1.get_color())
        axs[0].scatter((list(timerangelist.values())[i]).tolist()[::50], list(raw_hrv.values())[i][::50],s=8,zorder=2, color = line.get_color())   #scatter of raw data
        axs[1].scatter((list(timerangelist.values())[i]).tolist()[::50], list(raw_hr.values())[i][::50],s=8,zorder=2, color = line1.get_color())
    axs[0].legend(loc=1,fontsize=14)
    axs[1].legend(loc=1,fontsize=14)
    plt.tight_layout()
    return


for id_num in list(dd.keys()):  #id_num -> Find the unique dataframe that stores raw data corresponding to each patient id

    col_index_time = []
    col_index_hrv = []
    col_index_instant_hr = []  #refresh
    df_patient = globals()['df_' + 'Pat_ID_' + str(id_num)]
    x_axis_dict = dict() #time
    hrv_axis_dict = dict() #{Session n: poly_hrv}
    hr_axis_dict = dict() #{Session n: poly_hr}


    #fetch the colmuns which belongs to timestamp in data_dataframe
    for itt_col_time in range(0,df_patient.shape[1]-2):   #df_patient.shape[1]== 3or6or9or12  itt_col:0~max11
        if itt_col_time%3 == 0:
            col_index_time.append(itt_col_time) #[0] or [0,3] or [0,3,6] or [0,3,6,9] time
            col_index_hrv.append(itt_col_time+1) #[1] or [1,4] or [1,4,7] or [1,4,7,10] hrv
            col_index_instant_hr.append(itt_col_time+2) #[2] or [2,5] or [2,5,8] or [2,5,8,11] hr
            for itt_rowinver in range(len(df_patient)-1,-1,-1):
                if pd.isnull(df_patient.iloc[itt_rowinver,itt_col_time])==False:   #NaT detect from bottom
                    # start timestamp
                    Exp_start = df_patient.iloc[0, itt_col_time]
                    # end timestamp
                    Exp_end = df_patient.iloc[itt_rowinver,itt_col_time]
                    df_patient.iloc[itt_rowinver,itt_col_time] = (Exp_end - Exp_start).seconds #all timestamps->seconds since Exp start

    df_patient.dropna(axis=0,how='all',inplace=True)
    resample_hrv_dict = {}
    resample_hr_dict = {}  #container of resampled points on polyfit line
    rawdata_hrv_dict = {}
    rawdata_hr_dict = {} #container of raw data (for scatter)
    for itt_gpcol in range(len(col_index_time)):  #itt_gpcol in [0]or[0,1]or[0,1,2]or[0,1,2,3] -> grouping columns
        time_input = df_patient.iloc[:,col_index_time[itt_gpcol]].dropna().astype(float)
        hrv_input = df_patient.iloc[:,col_index_hrv[itt_gpcol]].dropna().astype(float)
        hr_input = df_patient.iloc[:,col_index_instant_hr[itt_gpcol]].dropna().astype(float)
        lrp_hrv = linear_polyfit(time=time_input, data=list(hrv_input))  # linear regression
        lrp_hr = linear_polyfit(time=time_input, data=list(hr_input))  # linear regression
        x_axis_dict.update({str(df_patient.columns[col_index_time[itt_gpcol]]): time_input})  #time
        hrv_axis_dict.update({str(df_patient.columns[col_index_hrv[itt_gpcol]]): lrp_hrv[0]})  #hrv
        resample_hrv_dict.update({str(df_patient.columns[col_index_hrv[itt_gpcol]]): lrp_hrv[1]}) #resampled hrv  len==1300
        hr_axis_dict.update({str(df_patient.columns[col_index_instant_hr[itt_gpcol]]): lrp_hr[0]})  #hr
        resample_hr_dict.update({str(df_patient.columns[col_index_instant_hr[itt_gpcol]]): lrp_hr[1]}) #resampled instant hr  len==1300
        rawdata_hrv_dict.update({str(df_patient.columns[col_index_hrv[itt_gpcol]]):hrv_input})
        rawdata_hr_dict.update({str(df_patient.columns[col_index_hrv[itt_gpcol]]):hr_input})


    temphrv = list(resample_hrv_dict.values())
    temphr = list(resample_hr_dict.values())
    meanhrv_array=[]
    meanhr_array=[]
    if len(col_index_time) == 4:
        #temp_array= np.append(temp[0],temp[1],temp[2],temp[3],axis=1)
        temphrv_array = np.concatenate(([temphrv[0]],[temphrv[1]],[temphrv[2]],[temphrv[3]]),axis=0)
        temphr_array = np.concatenate(([temphr[0]],[temphr[1]],[temphr[2]],[temphr[3]]),axis=0)
        meanhrv_array = temphrv_array.mean(axis=0)
        meanhr_array = temphr_array.mean(axis=0)
    elif len(col_index_time) ==3:
        temphrv_array = np.concatenate(([temphrv[0]],[temphrv[1]],[temphrv[2]]),axis=0)
        temphr_array = np.concatenate(([temphr[0]],[temphr[1]],[temphr[2]]),axis=0)
        meanhrv_array = temphrv_array.mean(axis=0)
        meanhr_array = temphr_array.mean(axis=0)
    elif len(col_index_time) ==2:
        temphrv_array = np.concatenate(([temphrv[0]],[temphrv[1]]),axis=0)
        temphr_array = np.concatenate(([temphr[0]],[temphr[1]]),axis=0)
        meanhrv_array = temphrv_array.mean(axis=0)
        meanhr_array = temphr_array.mean(axis=0)
    else:
        pass


    if len(col_index_time) >1:
        fig_plot(x_axis_dict,hrv=hrv_axis_dict,hrv_mean=meanhrv_array,hr=hr_axis_dict,hr_mean=meanhr_array,raw_hrv=rawdata_hrv_dict,raw_hr=rawdata_hr_dict,pat_name='Patient ID: '+str(id_num))
    else:
        fig_plot(x_axis_dict,hrv=hrv_axis_dict,hrv_mean=None,hr=hr_axis_dict,hr_mean=None,raw_hrv=rawdata_hrv_dict,raw_hr=rawdata_hr_dict,pat_name='Patient ID: '+str(id_num),valid=False)
    plt.show()


#################################### PART 2_oversession ###############################################


parameter_list_index = ['mean_nni','sdnn','sdsd','nni_50','pnni_50','nni_20','pnni_20','rmssd','median_nni',
                        'range_nni','cvsd','cvnni','mean_hr','max_hr','min_hr','std_hr','lf','hf','lf_hf_ratio','sd1','sd2','ratio_sd2_sd1']
parameter_column_index = ['Pat-ID','Session1 Exp1','Session1 Exp2','Session2 Exp1','Session2 Exp2','Session3 Exp1','Session3 Exp2','Session4 Exp1','Session4 Exp2']

parameter_2wanova_columns = ['SS','DF1','DF2','MS','F','p-unc','np2','eps']
parameter_2wanova_rows = ['Gruppe','Session','Interaction']
twanova_result_col_list = []
for anv_col in parameter_2wanova_columns:
    for anv_row in parameter_2wanova_rows:
        temp_str = anv_col+'_'+anv_row
        twanova_result_col_list.append(temp_str)
df_2wanova = pd.DataFrame(columns=twanova_result_col_list)

#summary excel
summary_list_index = ['Pat-ID','Gruppe','Session_mean','mean_nni','sdnn','sdsd','nni_50','pnni_50','nni_20','pnni_20','rmssd','median_nni',
                        'range_nni','cvsd','cvnni','mean_hr','max_hr','min_hr','std_hr','lf','hf','lf_hf_ratio','sd1','sd2','ratio_sd2_sd1']
summary_list_index2 = ['Pat-ID','Group','Session','Exp_nr','mean_nni','sdnn','sdsd','nni_50','pnni_50','nni_20','pnni_20','rmssd','median_nni',
                        'range_nni','cvsd','cvnni','mean_hr','max_hr','min_hr','std_hr','lf','hf','lf_hf_ratio','sd1','sd2','ratio_sd2_sd1']
df_summaryexcel = pd.DataFrame(columns = summary_list_index)
df_withinsummaryexcel= pd.DataFrame(columns = summary_list_index2)

#Intergrated assumption result
# Sphericity:
twasmp_list_sph = ['Spher','Sph_W','Sph_chi2','Sph_dof','Sph_pval']
# Homogeneity assumption (Levene Test):
twasmp_list_homo = ['Levene_session1_W','Levene_session1_pval','Levene_session1_equal_var','Levene_session2_W','Levene_session2_pval','Levene_session2_equal_var','Levene_session3_W','Levene_session3_pval','Levene_session3_equal_var','Levene_session4_W','Levene_session4_pval','Levene_session4_equal_var']
# Normality assumption:
twasmp_list_nor = ['AN-Session1_mean_W','AN-Session1_mean_pval','AN-Session1_mean_normal','AN-Session2_mean_W','AN-Session2_mean_pval','AN-Session2_mean_normal','AN-Session3_mean_W','AN-Session3_mean_pval','AN-Session3_mean_normal','AN-Session4_mean_W','AN-Session4_mean_pval','AN-Session4_mean_normal',
                   'KU-Session1_mean_W','KU-Session1_mean_pval','KU-Session1_mean_normal','KU-Session2_mean_W','KU-Session2_mean_pval','KU-Session2_mean_normal','KU-Session3_mean_W','KU-Session3_mean_pval','KU-Session3_mean_normal','KU-Session4_mean_W','KU-Session4_mean_pval','KU-Session4_mean_normal']
# integrated list:
twasmp_para_list = twasmp_list_sph + twasmp_list_homo + twasmp_list_nor
df_2wasmp = pd.DataFrame(columns = twasmp_para_list)

#regression
reg_list_columns=['coef','std err','z','P>|z|','[0.025','0.975]']
GEE_list_rows=['Intercept','Session[T.Session2_mean]','Session[T.Session3_mean]','Session[T.Session4_mean]','Gruppe[T.KU]' ]
linear_list_rows=['Intercept','Session[T.Session2_mean]','Session[T.Session3_mean]','Session[T.Session4_mean]','Gruppe[T.KU]']
GEE_result_col_list = []
for anv_col in reg_list_columns:
    for anv_row in GEE_list_rows:
        temp_str = anv_col+'_'+anv_row
        GEE_result_col_list.append(temp_str)
df_GEE = pd.DataFrame(columns=GEE_result_col_list)

LMM_result_col_list = []
for anv_col in reg_list_columns:
    for anv_row in linear_list_rows:
        temp_str = anv_col+'_'+anv_row
        LMM_result_col_list.append(temp_str)
df_LMM = pd.DataFrame(columns=LMM_result_col_list)


#post hoc test
turkey_list_columns=['mean(A)','mean(B)','diff','se','T','p-tukey','hedges']
gameshowell_list_columns=['mean(A)','mean(B)','diff','se','T','df','pval','hedges']
posthoc_list_columns=['mean(A)','mean(B)','diff','se','T','p-tukey','hedges','df','pval']#,'df','pval'
posthoc_list_rows = ['Session1_mean(A)_Session2_mean(B)', 'Session1_mean(A)_Session3_mean(B)', 'Session1_mean(A)_Session4_mean(B)','Session2_mean(A)_Session3_mean(B)', 'Session2_mean(A)_Session4_mean(B)', 'Session3_mean(A)_Session4_mean(B)']

turkey_result_col_list = []
for anv_col in turkey_list_columns:
    for anv_row in posthoc_list_rows:
        temp_str = anv_col+'_'+anv_row
        turkey_result_col_list.append(temp_str)
df_turkey = pd.DataFrame(columns=turkey_result_col_list)

gameshowell_result_col_list = []
for anv_col in gameshowell_list_columns:
    for anv_row in posthoc_list_rows:
        temp_str = anv_col+'_'+anv_row
        gameshowell_result_col_list.append(temp_str)
df_gameshowell = pd.DataFrame(columns=gameshowell_result_col_list)





######## Arrange SK1_dataframe，SK2_dataframe，REST_dataframe in order#########
for itt_df_para_number in range(len(parameter_list_index)):   #for loop in 22 parameters
    globals()['df_' + str(parameter_list_index[itt_df_para_number])] = pd.DataFrame(columns=parameter_column_index)  #create multiple dataframe and str parameters separately
    for itt_pa in range(len(SK1_dataframe)):  
        if SK1_dataframe.loc[itt_pa, 'Gruppe'] == 'AN':  # Group AN in SK1
            globals()['df_' + str(parameter_list_index[itt_df_para_number])].loc[itt_pa, 'Pat-ID'] = SK1_dataframe.loc[itt_pa, 'ID']
            temp = SK1_dataframe.loc[itt_pa, 'Expostionsnr.']
            globals()['df_' + str(parameter_list_index[itt_df_para_number])].loc[itt_pa, 'Session' + str(temp) + ' Exp1'] = SK1_dataframe.loc[itt_pa, parameter_list_index[itt_df_para_number]]
        elif SK1_dataframe.loc[itt_pa, 'Gruppe'] == 'KU':  # Group KU in SK1
            globals()['df_' + str(parameter_list_index[itt_df_para_number])].loc[itt_pa, 'Pat-ID'] = SK1_dataframe.loc[
                itt_pa, 'ID']
            temp = SK1_dataframe.loc[itt_pa, 'Expostionsnr.']
            globals()['df_' + str(parameter_list_index[itt_df_para_number])].loc[itt_pa, 'Session' + str(temp) + ' Exp1'] = SK1_dataframe.loc[itt_pa, parameter_list_index[itt_df_para_number]]
    for itt_pa in range(len(SK2_dataframe)):
        if SK2_dataframe.loc[itt_pa, 'Gruppe'] == 'AN':  # Group AN in SK2
            globals()['df_' + str(parameter_list_index[itt_df_para_number])].loc[itt_pa, 'Pat-ID'] = SK2_dataframe.loc[itt_pa, 'ID']
            temp = SK2_dataframe.loc[itt_pa, 'Expostionsnr.']
            globals()['df_' + str(parameter_list_index[itt_df_para_number])].loc[itt_pa, 'Session' + str(temp) + ' Exp2'] = SK2_dataframe.loc[itt_pa, parameter_list_index[itt_df_para_number]]
        elif SK2_dataframe.loc[itt_pa, 'Gruppe'] == 'KU':  # Group KU in SK2
            globals()['df_' + str(parameter_list_index[itt_df_para_number])].loc[itt_pa, 'Pat-ID'] = SK2_dataframe.loc[
                itt_pa, 'ID']
            temp = SK2_dataframe.loc[itt_pa, 'Expostionsnr.']
            globals()['df_' + str(parameter_list_index[itt_df_para_number])].loc[
                itt_pa, 'Session' + str(temp) + ' Exp2'] = SK2_dataframe.loc[
                itt_pa, parameter_list_index[itt_df_para_number]]

    globals()['df_' + str(parameter_list_index[itt_df_para_number])].reset_index(drop=True)
    
 
    for inver_pa in reversed(range(1, len(globals()['df_' + str(parameter_list_index[itt_df_para_number])]))):  # from back to the top 
        # temp_inver = SK1_dataframe.loc[inver_pa, 'Expostionsnr.']  #temp_inver is Session number
        if globals()['df_' + str(parameter_list_index[itt_df_para_number])].loc[inver_pa - 1, 'Pat-ID'] == \
                globals()['df_' + str(parameter_list_index[itt_df_para_number])].loc[inver_pa, 'Pat-ID']:  # when the ID of the previous line is the same as ID of this line
            for itt_for_notempty in range(1, 9):  #replicate the same position on the previous line if it is not null
                if pd.notnull(globals()['df_' + str(parameter_list_index[itt_df_para_number])].iloc[inver_pa, itt_for_notempty]):  # if it's not null
                    globals()['df_' + str(parameter_list_index[itt_df_para_number])].iloc[inver_pa - 1, itt_for_notempty] = \
                        globals()['df_' + str(parameter_list_index[itt_df_para_number])].iloc[inver_pa, itt_for_notempty]  # replicate this value of this row to the corresponding position od the previous row

           
    globals()['df_' + str(parameter_list_index[itt_df_para_number])].drop_duplicates('Pat-ID',inplace=True)
    globals()['df_' + str(parameter_list_index[itt_df_para_number])].dropna(thresh=6,inplace=True)  #delete the patient who is less than three Sessions
    globals()['df_' + str(parameter_list_index[itt_df_para_number])].reset_index(drop=True,inplace=True)
    
  
 
    
    # add a column and indicate the patient is in AN or KU group
    globals()['df_' + str(parameter_list_index[itt_df_para_number])].insert(1,'Gruppe','')
    for itt_pa2 in range(len(globals()['df_' + str(parameter_list_index[itt_df_para_number])])):
        for n in range(len(df_pre)):
            if globals()['df_' + str(parameter_list_index[itt_df_para_number])].loc[itt_pa2,'Pat-ID'] == df_pre.loc[n,'ID']:
                globals()['df_' + str(parameter_list_index[itt_df_para_number])].loc[itt_pa2,'Gruppe'] = df_pre.loc[n,'Gruppe']
                continue
    

    #create a new dataframe with all the value and the corresponding patient's ID, group and Exp_nr
    globals()['df_' + str(parameter_list_index[itt_df_para_number]) + '_withinSession'] = pd.DataFrame(columns=['Pat-ID','Group','Exp_nr','Data'])
   
    for itt_row in range(len(globals()['df_' + str(parameter_list_index[itt_df_para_number])])):
        for itt_exp_columns in range(2,10):
            temp_data_new = globals()['df_' + str(parameter_list_index[itt_df_para_number])].iloc[itt_row, itt_exp_columns]
            if pd.notna(temp_data_new):
                temp_ID = globals()['df_' + str(parameter_list_index[itt_df_para_number])].loc[itt_row, 'Pat-ID']
                temp_group = globals()['df_' + str(parameter_list_index[itt_df_para_number])].loc[itt_row, 'Gruppe']
                if itt_exp_columns in [2, 4, 6, 8]:                    
                    globals()['df_' + str(parameter_list_index[itt_df_para_number]) + '_withinSession'] = globals()['df_' + str(parameter_list_index[itt_df_para_number]) + '_withinSession'].append({'Pat-ID': temp_ID, 'Group': temp_group, 'Exp_nr': 1, 'Data': temp_data_new}, ignore_index=True)
                elif itt_exp_columns in [3, 5, 7, 9]:                    
                    globals()['df_' + str(parameter_list_index[itt_df_para_number]) + '_withinSession'] = globals()['df_' + str(parameter_list_index[itt_df_para_number]) + '_withinSession'].append({'Pat-ID': temp_ID, 'Group': temp_group, 'Exp_nr': 2, 'Data': temp_data_new}, ignore_index=True)
            else:
                continue
    
            
  
    
    #export excel in within session
    globals()['df_withinsummary_' + str(parameter_list_index[itt_df_para_number])]=globals()['df_' + str(parameter_list_index[itt_df_para_number]) + '_withinSession']
       
    for temp in range(len(globals()['df_withinsummary_' + str(parameter_list_index[itt_df_para_number])])):
        df_withinsummaryexcel.loc[temp,'Pat-ID'] =globals()['df_withinsummary_' + str(parameter_list_index[itt_df_para_number])].loc[temp,'Pat-ID']#,'Gruppe','Session_mean'
        df_withinsummaryexcel.loc[temp,'Group'] =globals()['df_withinsummary_' + str(parameter_list_index[itt_df_para_number])].loc[temp,'Group']
        df_withinsummaryexcel.loc[temp,'Exp_nr'] =globals()['df_withinsummary_' + str(parameter_list_index[itt_df_para_number])].loc[temp,'Exp_nr']
    
    for temp_summary in range(len(summary_list_index)):
        df_withinsummaryexcel.loc[:,str(parameter_list_index[itt_df_para_number])] = globals()['df_withinsummary_' + str(parameter_list_index[itt_df_para_number])]['Data']#'Pat-ID','Gruppe','Session_mean',
       
    #add session     
    session=df_withinsummaryexcel['Pat-ID'].value_counts().sort_index().to_list()
    session_times = [int(x / 2) for x in session]
    sessions=[]
    for i in session_times:
        if i==3:
            number =[1,1,2,2,3,3]
            sessions.append(number)
        else:
            numbers=[1,1,2,2,3,3,4,4]
            sessions.append(numbers)
    sessions = [j for i in sessions for j in i] 
    sessions = pd.DataFrame(sessions)
    df_withinsummaryexcel['Session']=sessions
    
    #calculate mean for each session 
    globals()['df_' + str(parameter_list_index[itt_df_para_number])].insert(10,'Session1_mean',globals()['df_' + str(parameter_list_index[itt_df_para_number])][['Session1 Exp1','Session1 Exp2']].mean(axis=1))
    globals()['df_' + str(parameter_list_index[itt_df_para_number])].insert(11,'Session2_mean',globals()['df_' + str(parameter_list_index[itt_df_para_number])][['Session2 Exp1','Session2 Exp2']].mean(axis=1))
    globals()['df_' + str(parameter_list_index[itt_df_para_number])].insert(12,'Session3_mean',globals()['df_' + str(parameter_list_index[itt_df_para_number])][['Session3 Exp1','Session3 Exp2']].mean(axis=1))
    globals()['df_' + str(parameter_list_index[itt_df_para_number])].insert(13,'Session4_mean',globals()['df_' + str(parameter_list_index[itt_df_para_number])][['Session4 Exp1','Session4 Exp2']].mean(axis=1))
    
    globals()['df_new_' + str(parameter_list_index[itt_df_para_number])]=globals()['df_' + str(parameter_list_index[itt_df_para_number])].drop(['Session1 Exp1', 'Session1 Exp2','Session2 Exp1', 'Session2 Exp2','Session3 Exp1', 'Session3 Exp2','Session4 Exp1', 'Session4 Exp2'],axis=1)
    
       
    #fillnan by mean
    globals()['df_new_' + str(parameter_list_index[itt_df_para_number])].fillna( globals()['df_new_' + str(parameter_list_index[itt_df_para_number])].mean(), inplace=True)
    #rearrange data
    globals()['df_melt_' + str(parameter_list_index[itt_df_para_number])]=globals()['df_new_' + str(parameter_list_index[itt_df_para_number])].melt(id_vars=['Pat-ID','Gruppe'],var_name = 'Session', value_name = 'value')
    
    over_Session = globals()['df_melt_' + str(parameter_list_index[itt_df_para_number])]
    
    #for summary excel
    globals()['df_export_' + str(parameter_list_index[itt_df_para_number])]=globals()['df_new_' + str(parameter_list_index[itt_df_para_number])]
    globals()['df_export_' + str(parameter_list_index[itt_df_para_number])]=globals()['df_export_' + str(parameter_list_index[itt_df_para_number])].set_index(['Pat-ID','Gruppe']).stack().reset_index(name='value').rename(columns={'level_2':'Session_mean'})
    
    
    for temp in range(len(globals()['df_export_' + str(parameter_list_index[itt_df_para_number])])):
        df_summaryexcel.loc[temp,'Pat-ID'] =globals()['df_export_' + str(parameter_list_index[itt_df_para_number])].loc[temp,'Pat-ID']#,'Gruppe','Session_mean'
        df_summaryexcel.loc[temp,'Gruppe'] =globals()['df_export_' + str(parameter_list_index[itt_df_para_number])].loc[temp,'Gruppe']
        df_summaryexcel.loc[temp,'Session_mean'] =globals()['df_export_' + str(parameter_list_index[itt_df_para_number])].loc[temp,'Session_mean']
            
       
    
    
    for temp_summary in range(len(summary_list_index)):
        df_summaryexcel.loc[:,str(parameter_list_index[itt_df_para_number])] = globals()['df_export_' + str(parameter_list_index[itt_df_para_number])]['value']#'Pat-ID','Gruppe','Session_mean',
    
    
    
    
     
    #H0: AU variance =KU variance 
    #H1: AU variance != KU variance 
    
    #H0:session1 variance =session2 variance =session3 variance =session4 variance
    #H1:at least one of the session variance is not equal to others
    
    #H0:There is no interaction between the Group and Time(session)
    #H1:There is interaction between the Group and Time
    
    #a=0.05 boundaries
  
    
        
    #Homogeneity assumption (Levene Test)   
    filter1 = over_Session['Session'].isin(['Session1_mean']) #filter1 -> Exp_nr=1
    filter2 = over_Session['Session'].isin(['Session2_mean'])
    filter3 = over_Session['Session'].isin(['Session3_mean'])
    filter4 = over_Session['Session'].isin(['Session4_mean'])

    
    homo_filter1=pg.homoscedasticity(over_Session[filter1], dv='value', group='Gruppe')
    homo_filter2=pg.homoscedasticity(over_Session[filter2], dv='value', group='Gruppe')
    homo_filter3=pg.homoscedasticity(over_Session[filter3], dv='value', group='Gruppe')
    homo_filter4=pg.homoscedasticity(over_Session[filter4], dv='value', group='Gruppe')
    df_2wasmp.loc[str(parameter_list_index[itt_df_para_number]), twasmp_list_homo[0]] = homo_filter1.loc['levene','W']
    df_2wasmp.loc[str(parameter_list_index[itt_df_para_number]), twasmp_list_homo[1]] = homo_filter1.loc['levene','pval']
    df_2wasmp.loc[str(parameter_list_index[itt_df_para_number]), twasmp_list_homo[2]] = homo_filter1.loc['levene','equal_var']
    df_2wasmp.loc[str(parameter_list_index[itt_df_para_number]), twasmp_list_homo[3]] = homo_filter2.loc['levene','W']
    df_2wasmp.loc[str(parameter_list_index[itt_df_para_number]), twasmp_list_homo[4]] = homo_filter2.loc['levene','pval']
    df_2wasmp.loc[str(parameter_list_index[itt_df_para_number]), twasmp_list_homo[5]] = homo_filter2.loc['levene','equal_var']
    df_2wasmp.loc[str(parameter_list_index[itt_df_para_number]), twasmp_list_homo[6]] = homo_filter3.loc['levene','W']
    df_2wasmp.loc[str(parameter_list_index[itt_df_para_number]), twasmp_list_homo[7]] = homo_filter3.loc['levene','pval']
    df_2wasmp.loc[str(parameter_list_index[itt_df_para_number]), twasmp_list_homo[8]] = homo_filter3.loc['levene','equal_var']
    df_2wasmp.loc[str(parameter_list_index[itt_df_para_number]), twasmp_list_homo[9]] = homo_filter4.loc['levene','W']
    df_2wasmp.loc[str(parameter_list_index[itt_df_para_number]), twasmp_list_homo[10]] = homo_filter4.loc['levene','pval']
    df_2wasmp.loc[str(parameter_list_index[itt_df_para_number]), twasmp_list_homo[11]] = homo_filter4.loc['levene','equal_var']
    
    
    ###### boxplot ######
    
    plt.figure(figsize=(12.8,10.24))  #1280*1024
    group_boxplot = sns.boxplot(x="Gruppe", y="value", hue="Session", data=globals()['df_melt_' + str(parameter_list_index[itt_df_para_number])], palette="Set1").set_title(str(parameter_list_index[itt_df_para_number]))#,showmeans=True
    

    
    ###### Histogram ######
    Groupfilter1 = over_Session['Gruppe'].isin(['AN'])  # filter1 -> AN
    Groupfilter2 = over_Session['Gruppe'].isin(['KU'])  # filter2 -> KU
   
    
    fig, axes = plt.subplots(1, 2, sharex=True, figsize=(10,5))
    fig.suptitle(str(parameter_list_index[itt_df_para_number]))  
    
    pa_boxplot = sns.histplot(bins='auto',data=over_Session[filter1][Groupfilter1].iloc[:,3],alpha  = 0.5,color='red',ax=axes[0], kde=True)
    pa_boxplot = sns.histplot(bins='auto',data=over_Session[filter2][Groupfilter1].iloc[:,3],alpha  = 0.5, color='blue',ax=axes[0], kde=True)
    pa_boxplot = sns.histplot(bins='auto',data=over_Session[filter3][Groupfilter1].iloc[:,3],alpha  = 0.5, color='green',ax=axes[0], kde=True)
    pa_boxplot = sns.histplot(bins='auto',data=over_Session[filter4][Groupfilter1].iloc[:,3],alpha  = 0.5, color='black',ax=axes[0], kde=True)
    axes[0].set_xlabel(str(parameter_list_index[itt_df_para_number]))
    fig.legend(['AN_session1','AN_session2','AN_Session3','AN_Session4'],loc='upper left')
    
    pa_boxplot = sns.histplot(bins='auto',data=over_Session[filter1][Groupfilter2].iloc[:,3],alpha  = 0.5,color='red',ax=axes[1], kde=True)
    pa_boxplot = sns.histplot(bins='auto',data=over_Session[filter2][Groupfilter2].iloc[:,3],alpha  = 0.5, color='blue',ax=axes[1], kde=True)
    pa_boxplot = sns.histplot(bins='auto',data=over_Session[filter3][Groupfilter2].iloc[:,3],alpha  = 0.5, color='green',ax=axes[1], kde=True)
    pa_boxplot = sns.histplot(bins='auto',data=over_Session[filter4][Groupfilter2].iloc[:,3],alpha  = 0.5, color='black',ax=axes[1], kde=True)
    axes[1].set_xlabel(str(parameter_list_index[itt_df_para_number]))
    
    fig.legend(['KU_session1','KU_Session2','KU_Session3','KU_Session4'])
    
    
    #Sphericity assumption
    Sph_asmp = pg.sphericity(data=over_Session, dv='value', within='Session', subject='Pat-ID', method='mauchly', alpha=0.05)
    for temp_sph in range(len(twasmp_list_sph)):
        df_2wasmp.loc[str(parameter_list_index[itt_df_para_number]),twasmp_list_sph[temp_sph]] = Sph_asmp[temp_sph]
    
    #Normality assumption
    for i in range(len(over_Session)):
       over_Session.loc[i,'factor_comb'] = over_Session.loc[i,'Gruppe']+'-'+str(over_Session.loc[i,'Session']) #AN-1 AN-2 KU-1 KU-2
    Nor_asmp = pg.normality(data=over_Session,dv='value',group= 'factor_comb')
    #print(Nor_asmp)
    for temp_nor_row in ['AN-Session1_mean','AN-Session2_mean','AN-Session3_mean','AN-Session4_mean','KU-Session1_mean','KU-Session2_mean','KU-Session3_mean','KU-Session4_mean']:
        for temp_nor_col in ['W','pval','normal']:
            df_2wasmp.loc[str(parameter_list_index[itt_df_para_number]),str(temp_nor_row+'_'+temp_nor_col)] = Nor_asmp.loc[temp_nor_row,temp_nor_col]
    
   
    #two-way mixed ANOVA
    twoway_mixed_anova = pg.mixed_anova(dv='value', between='Gruppe', within='Session', subject='Pat-ID', data=globals()['df_melt_' + str(parameter_list_index[itt_df_para_number])])
    twoway_mixed_anova.set_index(['Source'], inplace=True)
    
    #append to a dataframe preparing for to_excel
    for anv_col in parameter_2wanova_columns:
        for anv_row in parameter_2wanova_rows:
            df_2wanova.loc[str(parameter_list_index[itt_df_para_number]), str(anv_col + '_' + anv_row)] = twoway_mixed_anova.loc[anv_row, anv_col]
    
    
    #post-hoc test
    if homo_filter1['equal_var'].bool() == True and homo_filter2['equal_var'].bool() == True and homo_filter3['equal_var'].bool() == True and homo_filter4['equal_var'].bool() == True :
        res_turkey=globals()['df_melt_' + str(parameter_list_index[itt_df_para_number])].pairwise_tukey(dv='value', between='Session').round(3)
        res_turkey['A_B']=res_turkey['A']+'_'+res_turkey['B']
        res_turkey=res_turkey.drop(columns=['A', 'B'])
        res_turkey.set_index(['A_B'], inplace=True)
        res_turkey=res_turkey.rename(index={'Session1_mean_Session2_mean': 'Session1_mean(A)_Session2_mean(B)','Session1_mean_Session3_mean':'Session1_mean(A)_Session3_mean(B)', 'Session1_mean_Session4_mean':'Session1_mean(A)_Session4_mean(B)','Session2_mean_Session3_mean':'Session2_mean(A)_Session3_mean(B)', 'Session2_mean_Session4_mean':'Session2_mean(A)_Session4_mean(B)', 'Session3_mean_Session4_mean':'Session3_mean(A)_Session4_mean(B)'})
        
        for anv_col in turkey_list_columns :
            for anv_row in posthoc_list_rows:
                
                df_turkey.loc[str(parameter_list_index[itt_df_para_number]), str(anv_col + '+' + anv_row)] = res_turkey.loc[anv_row, anv_col]
        
        
        
    else:
        res_gameshowell=pg.pairwise_gameshowell(data=globals()['df_melt_' + str(parameter_list_index[itt_df_para_number])],dv='value', between='Session').round(3)
        res_gameshowell['A_B']=res_gameshowell['A']+'_'+res_gameshowell['B']
        res_gameshowell=res_gameshowell.drop(columns=['A', 'B'])
        res_gameshowell.set_index(['A_B'], inplace=True)
        res_gameshowell=res_gameshowell.rename(index={'Session1_mean_Session2_mean': 'Session1_mean(A)_Session2_mean(B)','Session1_mean_Session3_mean':'Session1_mean(A)_Session3_mean(B)', 'Session1_mean_Session4_mean':'Session1_mean(A)_Session4_mean(B)','Session2_mean_Session3_mean':'Session2_mean(A)_Session3_mean(B)', 'Session2_mean_Session4_mean':'Session2_mean(A)_Session4_mean(B)', 'Session3_mean_Session4_mean':'Session3_mean(A)_Session4_mean(B)'})
        
        for anv_col in gameshowell_list_columns :
            for anv_row in posthoc_list_rows:
                df_gameshowell.loc[str(parameter_list_index[itt_df_para_number]), str(anv_col + '+' + anv_row)] = res_gameshowell.loc[anv_row, anv_col]
    df_turkey.dropna(axis=1,how='all',inplace=True)
    df_gameshowell.dropna(axis=1,how='all',inplace=True)
    df_posthoc= pd.concat([df_turkey, df_gameshowell], sort=False)
    
    
    
    #linear mixture model(statsmodels)(??use this when there is non independence https://stats.oarc.ucla.edu/other/mult-pkg/introduction-to-linear-mixed-models/ )
    globals()['df_withna_' + str(parameter_list_index[itt_df_para_number])]=globals()['df_' + str(parameter_list_index[itt_df_para_number])].drop(['Session1 Exp1', 'Session1 Exp2','Session2 Exp1', 'Session2 Exp2','Session3 Exp1', 'Session3 Exp2','Session4 Exp1', 'Session4 Exp2'],axis=1)
    globals()['df_rwithna_' + str(parameter_list_index[itt_df_para_number])]=globals()['df_withna_' + str(parameter_list_index[itt_df_para_number])].melt(id_vars=['Pat-ID','Gruppe'],var_name = 'Session', value_name = 'value')
    over_session_withna=globals()['df_rwithna_' + str(parameter_list_index[itt_df_para_number])]
    linear_statsmodels=smf.mixedlm("value ~ Session + Gruppe ",data = over_session_withna,groups=over_session_withna['Gruppe'],missing='drop').fit()
    #df = pd.concat((linear_statsmodels.params, linear_statsmodels.tvalues), axis=1)
    linear_results=getattr(linear_statsmodels.summary(),'tables')[1]
    linear_results=linear_results.rename(columns = {"Coef.": "coef","Std.Err.":"std err"})
    for anv_col in reg_list_columns:
        for anv_row in linear_list_rows:
            df_LMM.loc[str(parameter_list_index[itt_df_para_number]), str(anv_col + '_' + anv_row)] = linear_results.loc[anv_row, anv_col] 
    
    results_text=linear_statsmodels.summary()

   
    
    
    #Generalized estimating equations
    GEE_statsmodels=smf.gee("value ~ Session + Gruppe ",data = over_session_withna,groups=over_session_withna['Gruppe']).fit()
    print(GEE_statsmodels.summary())
    results=getattr(GEE_statsmodels.summary(),'tables')[1:4][0] #third[]is get the coef and intercept and so on
    data_results =[]
    for row in range(0,6):
        for col in range(0,7):
            data_results.append(getattr(results[row][col],'data'))
    chunked_list=list()
    for i in range(0, len(data_results), 7):
        chunked_list.append(data_results[i:i+7])
    GEE_result=pd.DataFrame(chunked_list,columns=[' ','coef','std err','z','P>|z|','[0.025','0.975]']).drop([0], axis=0)
    GEE_result.set_index([' '], inplace=True)
    for anv_col in reg_list_columns:
        for anv_row in GEE_list_rows:
            df_GEE.loc[str(parameter_list_index[itt_df_para_number]), str(anv_col + '_' + anv_row)] = GEE_result.loc[anv_row, anv_col]
   

    

###### Poincaré plot ###### 
nni_K1 = SK1_dataframe['mean_nni'].tolist()
peaks_k1=nk.intervals_to_peaks(nni_K1)
results_K1 = nl.poincare(nni_K1)
#hrv_non_K1 = nk.hrv_nonlinear(peaks_k1, sampling_rate=100, show=True)
nni_K2 = SK2_dataframe['mean_nni'].tolist()
peaks_k2=nk.intervals_to_peaks(nni_K2)
results_K2 = nl.poincare(nni_K2)
#hrv_non_K2 = nk.hrv_nonlinear(peaks_k1, sampling_rate=100, show=True)  

   

#outside the second big for loop

#export all the statistic results
df_2wanova.dropna(axis=1,how='all',inplace=True)
df_2wanova = df_2wanova.reset_index()
df_2wanova.to_excel(mypath+'ID/'+'oversession_Two-way ANOVA statistic.xlsx',encoding='utf-8')

df_2wasmp = df_2wasmp.reset_index()
df_2wasmp.to_excel(mypath+'ID/'+'oversession_Two-way ANOVA assumption.xlsx',encoding='utf-8')

df_LMM.dropna(axis=1,how='all',inplace=True)
df_LMM = df_LMM.reset_index()
df_LMM.to_excel(mypath+'ID/'+'oversession_LMM statistic.xlsx',encoding='utf-8')

df_GEE = df_GEE.reset_index()
df_GEE.to_excel(mypath+'ID/'+'oversession_GEE statistic.xlsx',encoding='utf-8')


df_posthoc = df_posthoc.reset_index()
df_posthoc.to_excel(mypath+'ID/'+'oversession_posthoc statistic.xlsx',encoding='utf-8')

df_summaryexcel = df_summaryexcel.reset_index()
df_summaryexcel.to_excel(mypath+'ID/'+'oversession_summaryexcel.xlsx',encoding='utf-8')

df_withinsummaryexcel = df_withinsummaryexcel.reset_index()
df_withinsummaryexcel.to_excel(mypath+'ID/'+'withinsession_summaryexcel.xlsx',encoding='utf-8')

#################################### PART 2 within_session ###############################################
#A postfix string list for creating 22 dataframes with different names of parameters: df_XXXXX
parameter_list_index = ['mean_nni','sdnn','sdsd','nni_50','pnni_50','nni_20','pnni_20','rmssd','median_nni',
                        'range_nni','cvsd','cvnni','mean_hr','max_hr','min_hr','std_hr','lf','hf','lf_hf_ratio','sd1','sd2','ratio_sd2_sd1']
#and in each dataframes we at same column names
parameter_column_index = ['Pat-ID','Session1 Exp1','Session1 Exp2','Session2 Exp1','Session2 Exp2','Session3 Exp1','Session3 Exp2','Session4 Exp1','Session4 Exp2']
#Anova result parameters' lists
#parameter_2wanova_columns = ['df','sum_sq','mean_sq','F','PR(>F)'] #in bioinfokit
#parameter_2wanova_rows = ['Intercept','C(Group)','C(Exp_nr)','C(Group):C(Exp_nr)','Residual']  # in bioinfokit
parameter_2wanova_columns = ['SS','DF1','DF2','MS','F','p-unc','np2','eps']
parameter_2wanova_rows = ['Group','Exp_nr','Interaction']
twanova_result_col_list = []
for anv_col in parameter_2wanova_columns:
    for anv_row in parameter_2wanova_rows:
        temp_str = anv_col+'_'+anv_row
        twanova_result_col_list.append(temp_str)
df_2wanova = pd.DataFrame(columns=twanova_result_col_list)   #creat dataframe of 2 way anova

#Intergrated assumption result
# Sphericity:
twasmp_list_sph = ['Spher','Sph_W','Sph_chi2','Sph_dof','Sph_pval']
# Homogeneity assumption (Levene Test):
twasmp_list_homo = ['Levene_Exp1_W','Levene_Exp1_pval','Levene_Exp1_equal_var','Levene_Exp2_W','Levene_Exp2_pval','Levene_Exp2_equal_var']
# Normality assumption:
twasmp_list_nor = ['AN-1_W','AN-1_pval','AN-1_normal','AN-2_W','AN-2_pval','AN-2_normal','KU-1_W','KU-1_pval','KU-1_normal','KU-2_W','KU-2_pval','KU-2_normal']
# integrated list:
twasmp_para_list = twasmp_list_sph + twasmp_list_homo + twasmp_list_nor
df_2wasmp = pd.DataFrame(columns = twasmp_para_list)


###catch data from SK1_dataframe and then SK2_dataframe
for itt_df_para_number in range(len(parameter_list_index)):   #iteration in 22 parameters
    globals()['df_' + str(parameter_list_index[itt_df_para_number])] = pd.DataFrame(columns=parameter_column_index)  #mass-creating 22 dataframe
    for itt_pa in range(len(SK1_dataframe)):  # iteration in 80 data from SK1
        if SK1_dataframe.loc[itt_pa, 'Gruppe'] == 'AN':  # AN group in SK1
            globals()['df_' + str(parameter_list_index[itt_df_para_number])].loc[itt_pa, 'Pat-ID'] = SK1_dataframe.loc[
                itt_pa, 'ID']
            temp = SK1_dataframe.loc[itt_pa, 'Expostionsnr.']
            globals()['df_' + str(parameter_list_index[itt_df_para_number])].loc[
                itt_pa, 'Session' + str(temp) + ' Exp1'] = SK1_dataframe.loc[
                itt_pa, parameter_list_index[itt_df_para_number]]
        elif SK1_dataframe.loc[itt_pa, 'Gruppe'] == 'KU':  # KU group in SK1
            globals()['df_' + str(parameter_list_index[itt_df_para_number])].loc[itt_pa, 'Pat-ID'] = SK1_dataframe.loc[
                itt_pa, 'ID']
            temp = SK1_dataframe.loc[itt_pa, 'Expostionsnr.']
            globals()['df_' + str(parameter_list_index[itt_df_para_number])].loc[
                itt_pa, 'Session' + str(temp) + ' Exp1'] = SK1_dataframe.loc[
                itt_pa, parameter_list_index[itt_df_para_number]]
    for itt_pa in range(len(SK2_dataframe)):
        if SK2_dataframe.loc[itt_pa, 'Gruppe'] == 'AN':  # AN group in SK2
            globals()['df_' + str(parameter_list_index[itt_df_para_number])].loc[itt_pa, 'Pat-ID'] = SK2_dataframe.loc[
                itt_pa, 'ID']
            temp = SK2_dataframe.loc[itt_pa, 'Expostionsnr.']
            globals()['df_' + str(parameter_list_index[itt_df_para_number])].loc[
                itt_pa, 'Session' + str(temp) + ' Exp2'] = SK2_dataframe.loc[
                itt_pa, parameter_list_index[itt_df_para_number]]
        elif SK2_dataframe.loc[itt_pa, 'Gruppe'] == 'KU':  # KU group in SK2
            globals()['df_' + str(parameter_list_index[itt_df_para_number])].loc[itt_pa, 'Pat-ID'] = SK2_dataframe.loc[
                itt_pa, 'ID']
            temp = SK2_dataframe.loc[itt_pa, 'Expostionsnr.']
            globals()['df_' + str(parameter_list_index[itt_df_para_number])].loc[
                itt_pa, 'Session' + str(temp) + ' Exp2'] = SK2_dataframe.loc[
                itt_pa, parameter_list_index[itt_df_para_number]]


    globals()['df_' + str(parameter_list_index[itt_df_para_number])].reset_index(drop=True)
    #find the duplicated row with same patient ID, and merge them together
    for inver_pa in reversed(range(1, len(globals()['df_' + str(parameter_list_index[itt_df_para_number])]))):  # reverse iteration
        if globals()['df_' + str(parameter_list_index[itt_df_para_number])].loc[inver_pa - 1, 'Pat-ID'] == \
                globals()['df_' + str(parameter_list_index[itt_df_para_number])].loc[inver_pa, 'Pat-ID']:  # when ID in this row == ID in the previous row -->need merging
            for itt_for_notempty in range(1, 9):  # iteration every element in this row, if there is a valid value, copy it to the same location in the previous line
                if pd.notnull(globals()['df_' + str(parameter_list_index[itt_df_para_number])].iloc[inver_pa, itt_for_notempty]):  # pd.notnull() used for checking valid data
                    globals()['df_' + str(parameter_list_index[itt_df_para_number])].iloc[inver_pa - 1, itt_for_notempty] = \
                        globals()['df_' + str(parameter_list_index[itt_df_para_number])].iloc[inver_pa, itt_for_notempty]  # copy

    print(str(parameter_list_index[itt_df_para_number])+': ')
    globals()['df_' + str(parameter_list_index[itt_df_para_number])].drop_duplicates('Pat-ID',inplace=True)
    before_dropna = len(globals()['df_' + str(parameter_list_index[itt_df_para_number])])
    globals()['df_' + str(parameter_list_index[itt_df_para_number])].dropna(thresh=6,inplace=True)  #exclude those patient data if sessions < 3
    after_dropna = len(globals()['df_' + str(parameter_list_index[itt_df_para_number])])
    print('Number of deleting rows by dropna function: ',before_dropna-after_dropna)
    globals()['df_' + str(parameter_list_index[itt_df_para_number])].reset_index(drop=True,inplace=True)


    #creat a new column to indicate the groups of patient: AN or KU
    globals()['df_' + str(parameter_list_index[itt_df_para_number])].insert(1,'Gruppe','')
    for itt_pa2 in range(len(globals()['df_' + str(parameter_list_index[itt_df_para_number])])):
        for n in range(len(df_pre)):
            if globals()['df_' + str(parameter_list_index[itt_df_para_number])].loc[itt_pa2,'Pat-ID'] == df_pre.loc[n,'ID']:
                globals()['df_' + str(parameter_list_index[itt_df_para_number])].loc[itt_pa2,'Gruppe'] = df_pre.loc[n,'Gruppe']
                continue


    #WithinSession: creat a new dataframe, include every data and corresponidng patient ID, group and Expnr.
    globals()['df_' + str(parameter_list_index[itt_df_para_number]) + '_withinSession'] = pd.DataFrame(columns=['Pat-ID','Group','Exp_nr','Data'])
    for itt_row in range(len(globals()['df_' + str(parameter_list_index[itt_df_para_number])])):
        for itt_exp_columns in range(2,10):
            temp_data_new = float(globals()['df_' + str(parameter_list_index[itt_df_para_number])].iloc[itt_row, itt_exp_columns])
            if pd.notna(temp_data_new):
                temp_ID = globals()['df_' + str(parameter_list_index[itt_df_para_number])].loc[itt_row, 'Pat-ID']
                temp_group = globals()['df_' + str(parameter_list_index[itt_df_para_number])].loc[itt_row, 'Gruppe']
                if itt_exp_columns in [2, 4, 6, 8]:
                    globals()['df_' + str(parameter_list_index[itt_df_para_number]) + '_withinSession'] = globals()['df_' + str(parameter_list_index[itt_df_para_number]) + '_withinSession'].append({'Pat-ID': temp_ID, 'Group': temp_group, 'Exp_nr': 1, 'Data': temp_data_new}, ignore_index=True)
                elif itt_exp_columns in [3, 5, 7, 9]:
                    globals()['df_' + str(parameter_list_index[itt_df_para_number]) + '_withinSession'] = globals()['df_' + str(parameter_list_index[itt_df_para_number]) + '_withinSession'].append({'Pat-ID': temp_ID, 'Group': temp_group, 'Exp_nr': 2, 'Data': temp_data_new}, ignore_index=True)
            else:
                continue

    
    ###### Box plot ######
    plt.figure(figsize=(12,10))  #1280*1024
    group_boxplot = sns.boxplot(x="Group", y="Data", hue="Exp_nr", data=globals()['df_' + str(parameter_list_index[itt_df_para_number]) + '_withinSession'], palette="Set1").set_title(str(parameter_list_index[itt_df_para_number]))
    
    #plt.xticks([0,1],['AU','CG'])
    #plt.legend(fontsize=16)
    plt.show()
    


    
    #formula = 'Data ~ C(Group) + C(Exp_nr) + C(Group)*C(Exp_nr)'

    ####### Two-way mixed model ANOVA ########
    df_within_Ses = globals()['df_' + str(parameter_list_index[itt_df_para_number]) + '_withinSession']
    # Sphericity assumption
    Sph_asmp = pg.sphericity(data=df_within_Ses, dv='Data', within='Exp_nr', subject='Pat-ID', method='mauchly',alpha=0.05)
    #print('Sphericity assumption:\n', Sph_asmp)
    # spher: True if data have the sphericity property.
    # W: Test statistic.
    # chi2: Chi-square statistic.
    # dof: Degrees of freedom.
    # pval: P-value.
    
    # Homogeneity assumption (Levene Test)
    Expfilter1 = df_within_Ses['Exp_nr'].isin([1])  # filter1 -> Exp_nr=1
    Expfilter2 = df_within_Ses['Exp_nr'].isin([2])  # filter2 -> Exp_nr=2
    homo_filter1 = pg.homoscedasticity(df_within_Ses[Expfilter1], dv='Data', group='Group')
    homo_filter2 = pg.homoscedasticity(df_within_Ses[Expfilter2], dv='Data', group='Group')
    #print('Levene Exp1: ', homo_filter1)
    #print('Levene Exp2: ', homo_filter2)
    
    
    Groupfilter1 = df_within_Ses['Group'].isin(['AN'])  # filter1 -> AN
    Groupfilter2 = df_within_Ses['Group'].isin(['KU'])  # filter2 -> KU
    plt.figure(figsize=(10,8))
    pa_histplot = sns.histplot(bins='auto', kde=True,data=df_within_Ses[Expfilter1][Groupfilter1].iloc[:,3],alpha  = 0.5,color='darkred').set_title(str(parameter_list_index[itt_df_para_number]))
    pa_histplot = sns.histplot(bins='auto', kde=True,data=df_within_Ses[Expfilter1][Groupfilter2].iloc[:,3],alpha  = 0.5,color='darkslateblue').set_title(str(parameter_list_index[itt_df_para_number]))
    pa_histplot = sns.histplot(bins='auto', kde=True,data=df_within_Ses[Expfilter2][Groupfilter1].iloc[:,3],alpha  = 0.5, color='red').set_title(str(parameter_list_index[itt_df_para_number]))
    pa_histplot = sns.histplot(bins='auto', kde=True,data=df_within_Ses[Expfilter2][Groupfilter2].iloc[:,3],alpha  = 0.5, color='lightblue').set_title(str(parameter_list_index[itt_df_para_number]))
    #plt.legend(fontsize=10)
    #plt.legend(['AN_Exp1','CG_Exp2','AN_Exp2','CG_Exp2'],prop={'size': 16}, fontsize=16)
    plt.legend(['AN_Exp1','CG_Exp2','AN_Exp2','CG_Exp2'])
    plt.xlabel(str(parameter_list_index[itt_df_para_number]))
    plt.tight_layout()
    
    
    # Normality assumption
    for i in range(len(df_within_Ses)):
        df_within_Ses.loc[i, 'factor_comb'] = df_within_Ses.loc[i, 'Group'] + '-' + str(
            df_within_Ses.loc[i, 'Exp_nr'])  # AN-1 AN-2 KU-1 KU-2
    Nor_asmp = pg.normality(data=df_within_Ses, dv='Data', group='factor_comb')
    #print('Normality assumption:\n', Nor_asmp)

    df_mixed_anova = pg.mixed_anova(dv='Data', between='Group', within='Exp_nr', subject='Pat-ID', data=df_within_Ses)
    df_mixed_anova.set_index(['Source'], inplace=True)
    #print(df_mixed_anova)
    #'Source': Names of the factor considered
    #'ddof1': Degrees of freedom (numerator)
    #'ddof2': Degrees of freedom (denominator)
    #'F': F-values
    #'p-unc': Uncorrected p-values
    #'np2': Partial eta-squared effect sizes
    #'eps': Greenhouse-Geisser epsilon factor (= index of sphericity)

    #append to a dataframe preparing for to_excel
    for anv_col in parameter_2wanova_columns:
        for anv_row in parameter_2wanova_rows:
            df_2wanova.loc[str(parameter_list_index[itt_df_para_number]), str(anv_col + '_' + anv_row)] = df_mixed_anova.loc[anv_row, anv_col]

    #fill into df_2wasmp
    # Sphericity:
    for temp_sph in range(len(twasmp_list_sph)):
        df_2wasmp.loc[str(parameter_list_index[itt_df_para_number]),twasmp_list_sph[temp_sph]] = Sph_asmp[temp_sph]
    # Homogenity (levene):
    df_2wasmp.loc[str(parameter_list_index[itt_df_para_number]), twasmp_list_homo[0]] = homo_filter1.loc['levene','W']
    df_2wasmp.loc[str(parameter_list_index[itt_df_para_number]), twasmp_list_homo[1]] = homo_filter1.loc['levene','pval']
    df_2wasmp.loc[str(parameter_list_index[itt_df_para_number]), twasmp_list_homo[2]] = homo_filter1.loc['levene','equal_var']
    df_2wasmp.loc[str(parameter_list_index[itt_df_para_number]), twasmp_list_homo[3]] = homo_filter2.loc['levene','W']
    df_2wasmp.loc[str(parameter_list_index[itt_df_para_number]), twasmp_list_homo[4]] = homo_filter2.loc['levene','pval']
    df_2wasmp.loc[str(parameter_list_index[itt_df_para_number]), twasmp_list_homo[5]] = homo_filter2.loc['levene','equal_var']
    #Normality:
    for temp_nor_row in ['AN-1','AN-2','KU-1','KU-2']:
        for temp_nor_col in ['W','pval','normal']:
            df_2wasmp.loc[str(parameter_list_index[itt_df_para_number]),str(temp_nor_row+'_'+temp_nor_col)] = Nor_asmp.loc[temp_nor_row,temp_nor_col]
    
    #df_2wasmp.dropna(axis=1,how='all',inplace=True)

    #res=stat()
    #res_turkey = globals()['df_melt_' + str(parameter_list_index[itt_df_para_number])].pairwise_tukey(dv='value',between='Exp_nr').round(3)
    #res_gameshowell = pg.pairwise_gameshowell(data=globals()['df_' + str(parameter_list_index[itt_df_para_number])], dv='value',between='Exp_nr').round(3)

    #print('\n\n')  #finish process this parameter
    


df_2wanova.dropna(axis=1,how='all',inplace=True)
df_2wanova = df_2wanova.reset_index()
df_2wanova.to_excel(mypath+'ID/'+'Two-way ANOVA statistic.xlsx',encoding='utf-8')
df_2wasmp = df_2wasmp.reset_index()
df_2wasmp.to_excel(mypath+'ID/'+'Two-way ANOVA assumption.xlsx',encoding='utf-8')



#put all the boxplot in one frame
fig, (ax1, ax2,ax3,ax4,ax5) = plt.subplots(1, 5, figsize=(16,6))
itt_df_para_number=1#1,4,7,18,19
sns.boxplot(ax=ax1,x="Group", y="Data", hue="Exp_nr", data=globals()['df_' + str(parameter_list_index[itt_df_para_number]) + '_withinSession'], palette="Set1").set_title(str(parameter_list_index[itt_df_para_number]))
ax1.set_xticklabels(['AU','CG'], fontdict=None, minor=False)
ax1.legend([],[], frameon=False)
ax1.set(ylabel='[ms]')
ax1.set_ylim((0,250))
itt_df_para_number=4
sns.boxplot(ax=ax2,x="Group", y="Data", hue="Exp_nr", data=globals()['df_' + str(parameter_list_index[itt_df_para_number]) + '_withinSession'], palette="Set1").set_title(str(parameter_list_index[itt_df_para_number]))
ax2.set_xticklabels(['AU','CG'], fontdict=None, minor=False)
ax2.legend([],[], frameon=False)
ax2.set(ylabel='[%]')
ax2.set_ylim((0,100))
itt_df_para_number=7
sns.boxplot(ax=ax3,x="Group", y="Data", hue="Exp_nr", data=globals()['df_' + str(parameter_list_index[itt_df_para_number]) + '_withinSession'], palette="Set1").set_title(str(parameter_list_index[itt_df_para_number]))
ax3.set_xticklabels(['AU','CG'], fontdict=None, minor=False)
ax3.legend([],[], frameon=False)
ax3.set(ylabel='[ms]')
ax3.set_ylim((0,140))
itt_df_para_number=18
sns.boxplot(ax=ax4,x="Group", y="Data", hue="Exp_nr", data=globals()['df_' + str(parameter_list_index[itt_df_para_number]) + '_withinSession'], palette="Set1").set_title(str(parameter_list_index[itt_df_para_number]))
ax4.set_xticklabels(['AU','CG'], fontdict=None, minor=False)
ax4.legend([],[], frameon=False)
itt_df_para_number=19
ax4.set(ylabel='[ms$^{2}$]')
ax4.set_ylim((0,10))
sns.boxplot(ax=ax5,x="Group", y="Data", hue="Exp_nr", data=globals()['df_' + str(parameter_list_index[itt_df_para_number]) + '_withinSession'], palette="Set1").set_title(str(parameter_list_index[itt_df_para_number]))
ax5.set_xticklabels(['AU','CG'], fontdict=None, minor=False)
ax5.set(ylabel='[ms]')
ax5.set_ylim((0,120))
plt.tight_layout()      
    
fig, (ax1, ax2,ax3,ax4,ax5) = plt.subplots(1, 5, figsize=(16,6))
itt_df_para_number=1
sns.boxplot(ax=ax1,x="Gruppe", y="value", hue="Session", data=globals()['df_melt_' + str(parameter_list_index[itt_df_para_number])], palette="Set1").set_title(str(parameter_list_index[itt_df_para_number]))#,showmeans=True
ax1.set_xticklabels(['AU','CG'], fontdict=None, minor=False)
ax1.legend([],[], frameon=False)
ax1.set_xlabel('Group')
ax1.set(ylabel='[ms]')
ax1.set_ylim((0,250))
itt_df_para_number=4
sns.boxplot(ax=ax2,x="Gruppe", y="value", hue="Session", data=globals()['df_melt_' + str(parameter_list_index[itt_df_para_number])], palette="Set1").set_title(str(parameter_list_index[itt_df_para_number]))#,showmeans=True
ax2.set_xticklabels(['AU','CG'], fontdict=None, minor=False)
ax2.legend([],[], frameon=False)
ax2.set_xlabel('Group')
ax2.set(ylabel='[%]')
ax2.set_ylim((0,100))
itt_df_para_number=7
sns.boxplot(ax=ax3,x="Gruppe", y="value", hue="Session", data=globals()['df_melt_' + str(parameter_list_index[itt_df_para_number])], palette="Set1").set_title(str(parameter_list_index[itt_df_para_number]))#,showmeans=True
ax3.set_xticklabels(['AU','CG'], fontdict=None, minor=False)
ax3.set_xlabel('Group')
ax3.legend([],[], frameon=False)
ax3.set(ylabel='[ms]')
ax3.set_ylim((0,140))
itt_df_para_number=18
sns.boxplot(ax=ax4,x="Gruppe", y="value", hue="Session", data=globals()['df_melt_' + str(parameter_list_index[itt_df_para_number])], palette="Set1").set_title(str(parameter_list_index[itt_df_para_number]))#,showmeans=True
ax4.legend([],[], frameon=False)
ax4.set_xlabel('Group')
itt_df_para_number=19
ax4.set_xticklabels(['AU','CG'], fontdict=None, minor=False)
ax4.set(ylabel='[ms$^{2}$]')
ax4.set_ylim((0,10))
sns.boxplot(ax=ax5,x="Gruppe", y="value", hue="Session", data=globals()['df_melt_' + str(parameter_list_index[itt_df_para_number])], palette="Set1").set_title(str(parameter_list_index[itt_df_para_number]))#,showmeans=True
ax5.set_xticklabels(['AU','CG'], fontdict=None, minor=False)
ax5.set_xlabel('Group')
ax5.set(ylabel='[ms]')
ax5.set_ylim((0,120))
plt.tight_layout()              
    
    
    
    
    
    