# -*- coding: utf-8 -*-

print(" welcome to Ceragon Daily RSL Report ")

import pandas as pd
import numpy as np
import paramiko 
import datetime
from datetime import datetime, timedelta
import win32com.client as wincl
from openpyxl import load_workbook
import os
import glob

# Auto date


do=(datetime.now() - timedelta(1)).strftime('%d_%m_%Y')
do=(datetime.now() - timedelta(1)).strftime('%Y%m%d')
dm=(datetime.now() - timedelta(1)).strftime('%m-%y')
dn=(datetime.now() - timedelta(0)).strftime('%d-%m-%Y')
dnn=(datetime.now() - timedelta(1)).strftime('%d.%m.%Y')
dnnn=(datetime.now() - timedelta(1)).strftime('%Y-%m-%Y')
d=(datetime.now() - timedelta(1)).strftime('%d%m%Y')



da7= datetime.now() - timedelta(1)
da6= datetime.now() - timedelta(2)
da5 = datetime.now() - timedelta(3)
da4 = datetime.now() - timedelta(4)
da3 = datetime.now() - timedelta(5)
da2 = datetime.now() - timedelta(6)
da1 = datetime.now() - timedelta(7)

do7= datetime.now() - timedelta(1)
do6= datetime.now() - timedelta(2)
do5 = datetime.now() - timedelta(3)
do4 = datetime.now() - timedelta(4)
do3 = datetime.now() - timedelta(5)
do2 = datetime.now() - timedelta(6)
do1 = datetime.now() - timedelta(7)

d7= datetime.now() - timedelta(1)
d6= datetime.now() - timedelta(2)
d5 = datetime.now() - timedelta(3)
d4 = datetime.now() - timedelta(4)
d3 = datetime.now() - timedelta(5)
d2 = datetime.now() - timedelta(6)
d1 = datetime.now() - timedelta(7)



da7 = datetime.strftime(da7, '%d_%m_%Y')
da6 = datetime.strftime(da6, '%d_%m_%Y')
da5 = datetime.strftime(da5, '%d_%m_%Y')
da4 = datetime.strftime(da4, '%d_%m_%Y')
da3 = datetime.strftime(da3, '%d_%m_%Y')
da2 = datetime.strftime(da2, '%d_%m_%Y')
da1 = datetime.strftime(da1, '%d_%m_%Y')

do7 = datetime.strftime(do7, '%Y%m%d')
do6 = datetime.strftime(do6, '%Y%m%d')
do5 = datetime.strftime(do5, '%Y%m%d')
do4 = datetime.strftime(do4, '%Y%m%d')
do3 = datetime.strftime(do3, '%Y%m%d')
do2 = datetime.strftime(do2, '%Y%m%d')
do1 = datetime.strftime(do1, '%Y%m%d')

d7 = datetime.strftime(d7, '%d%m%Y')
d6 = datetime.strftime(d6, '%d%m%Y')
d5 = datetime.strftime(d5, '%d%m%Y')
d4 = datetime.strftime(d4, '%d%m%Y')
d3 = datetime.strftime(d3, '%d%m%Y')
d2 = datetime.strftime(d2, '%d%m%Y')
d1 = datetime.strftime(d1, '%d%m%Y')


print(da7)

dmm=(datetime.now() - timedelta(1)).strftime('%m-%y')

print(dmm)

print('Raw files Removing....')




directory=r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera RSL\RAW\DAY1'
os.chdir(directory)
files=glob.glob('*.csv')
for filename in files:
    os.unlink(filename)



directory=r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera RSL\RAW\DAY2'
os.chdir(directory)
files=glob.glob('*.csv')
for filename in files:
    os.unlink(filename)

    

directory=r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera RSL\RAW\DAY3'
os.chdir(directory)
files=glob.glob('*.csv')
for filename in files:
    os.unlink(filename)



    
print("Removed....")



ssh3=paramiko.SSHClient()
ssh3.set_missing_host_key_policy(paramiko.AutoAddPolicy())
try:
    ssh3.connect(hostname='10.10.10.10',username='admin',password='admin',port=22)
except:
    pass
        
try:
    ssh3.connect(hostname='11.11.11.11',username='admin',password='admin',port=22)
except:
    pass

sftp_client1=ssh3.open_sftp()



sftp_client1.chdir('/opt/csvascii_nr21/cm1/UPE/')
try:
    sftp_client1.get('Full_Link_Report_'+da7+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera RSL\RAW\DAY1\U_Full_Link_Report_'+da7+'.csv')
except:
    pass

sftp_client1.chdir('/opt/csvascii_nr21/cm1/ROB/')
try:
    sftp_client1.get('Full_Link_Report_'+da7+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera RSL\RAW\DAY1\R_Full_Link_Report_'+da7+'.csv')
except:
    pass

sftp_client1.chdir('/opt/csvascii_nr21/cm1/KAR/')
try:
    sftp_client1.get('Full_Link_Report_'+da7+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera RSL\RAW\DAY1\K_Full_Link_Report_'+da7+'.csv')
except:
    pass
print("DAY1 All file downloaded")

print("SFTP Done")


ssh3=paramiko.SSHClient()
ssh3.set_missing_host_key_policy(paramiko.AutoAddPolicy())
try:
    ssh3.connect(hostname='1.1.1.1',username='admin',password='admin',port=22)
except:
    pass

try:
    ssh3.connect(hostname='2.2.2.2',username='admin',password='admin',port=22)
except:
    pass
sftp_client1=ssh3.open_sftp()


print("Downloading Cobra Master ")

sftp_client1.chdir('/opt/MyLog/TX/Master_data')
sftp_client1.get('Ceragon_planning_Data.xlsx' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera RSL\RAW\Ceragon_planning_Data.xlsx')


print("Day1 file read")


try:
    KAR = pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera RSL\RAW\DAY1\K_Full_Link_Report_'+da7+'.csv',skiprows=5)
except:
    pass
try:
    UPE = pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera RSL\RAW\DAY1\U_Full_Link_Report_'+da7+'.csv',skiprows=5,encoding='latin1')
except:
    pass
try:
    ROB = pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera RSL\RAW\DAY1\R_Full_Link_Report_'+da7+'.csv',skiprows=5,encoding='latin1')
except:
    pass


print("Reading Done")


try:
    UPE.rename(columns={'Site Z Name':'Site B Name','Site Z Physical Port':'Site B Physical Port','Site Z IP':'Site B IP','Site A Tx Power (Current) [dBm]':'Site A Tx Power Current','Site Z Tx Power (Current) [dBm]':'Site B Tx Power Current','Site A Rx Level (Current) [dBm]':'Site A Rx Level Current','Site Z Rx Level (Current) [dBm]':'Site B Rx Level Current'},inplace=True)
except:
    pass


try:
    KAR.rename(columns={'Site Z Name':'Site B Name','Site A Tx Power (Current) [dBm]':'Site A Tx Power Current','Site Z Tx Power (Current) [dBm]':'Site B Tx Power Current','Site A Rx Level (Current) [dBm]':'Site A Rx Level Current','Site Z Rx Level (Current) [dBm]':'Site B Rx Level Current'},inplace=True)
except:
    pass

try:
    ROB.rename(columns={'Site Z Name':'Site B Name','Site A Tx Power (Current) [dBm]':'Site A Tx Power Current','Site Z Tx Power (Current) [dBm]':'Site B Tx Power Current','Site A Rx Level (Current) [dBm]':'Site A Rx Level Current','Site Z Rx Level (Current) [dBm]':'Site B Rx Level Current'},inplace=True)
except:
    pass

print("Day1 column Renamed")

## SELECTED ONLY USED COILUMNS

try:
    UPE=UPE[['Site A Name','Site B Name','Site A Tx Power Current','Site B Tx Power Current','Site A Rx Level Current','Site B Rx Level Current']]
except:
    pass
try:
    KAR=KAR[['Site A Name','Site B Name','Site A Tx Power Current','Site B Tx Power Current','Site A Rx Level Current','Site B Rx Level Current']]
except:
    pass
try:
    ROB=ROB[['Site A Name','Site B Name','Site A Tx Power Current','Site B Tx Power Current','Site A Rx Level Current','Site B Rx Level Current']]
except:
    pass


## CREATE NEW SERVER COLUMN


try:
    UPE['Server']='UPE'
except:
    pass
try:
    KAR['Server']='KAR'
except:
    pass
try:
    ROB['Server']='ROB'
except:
    pass

#DAY1=pd.concat([UPE,KAR,ROB])

dataframes = []
for name in ['KAR', 'ROB', 'UPE']:
    try:
        df = globals()[name]
        if isinstance(df, pd.DataFrame):
            dataframes.append(df)
    except KeyError:
        pass

DAY1 = pd.concat(dataframes, ignore_index=True)


## removes leading and trailing whitespaces from each string in the Series
DAY1['Site A Name']=DAY1['Site A Name'].str.strip()
DAY1['Site B Name']=DAY1['Site B Name'].str.strip()

## ************ Find ROWS having space ************ ##
#rows_with_spaces = DAY1[DAY1['Site A Name'].str.contains(' ')]

## ****** Remove blanks rows having space ****** ##

DAY1_cleaned = DAY1[DAY1['Site A Name'].str.strip() != ''].copy()
DAY2_cleaned = DAY1_cleaned[DAY1_cleaned['Site B Name'].str.strip() != ''].copy()

# Rename the DataFrame to a new name, 
DAY1 = DAY2_cleaned.copy()

## UNIQ LINK

DAY1['uniq link']=np.where((DAY1['Site A Name']<DAY1['Site B Name']),(DAY1['Site A Name']+'-'+DAY1['Site B Name']),(DAY1['Site B Name']+'-'+ DAY1['Site A Name']))

## removes leading and trailing whitespaces from each string in the Series
#['uniq link']=DAY1['uniq link'].str.strip()

# Circle*******************************


DAY1.loc[DAY1['uniq link'].str.contains('IDDL|INDL',na=False),'Circle']='DEL'
DAY1.loc[DAY1['uniq link'].str.contains('IDUW|INUW',na=False),'Circle']='UPW'
DAY1.loc[DAY1['uniq link'].str.contains('IDOD|INOD',na=False),'Circle']='ODI'
DAY1.loc[DAY1['uniq link'].str.contains('IDKL|INKL',na=False),'Circle']='KEL'
DAY1.loc[DAY1['uniq link'].str.contains('IDAS|IDNE|INAS|INNE', na=False), 'Circle'] = 'ASM'
DAY1.loc[DAY1['uniq link'].str.contains('IDUE|INUE|AZMG',na=False),'Circle']='UPE'
DAY1.loc[DAY1['uniq link'].str.contains('IDKA|INKA|MYS0|MYS9',na=False),'Circle']='KAR'
DAY1.loc[DAY1['uniq link'].str.contains('IDWB|IINW|INEW|INWB',na=False),'Circle']='ROB'
DAY1.loc[DAY1['uniq link'].str.contains(' INB|BBSN|BCHN|BDAR|BLXM|BMRU|bnir|BPIA|BR10|BSAS|BTOD|BUGR|IDB0|IDBR|INBR|JBKU|KOLA|BN2083|BPPK',na=False),'Circle']='BIH'
DAY1.loc[DAY1['uniq link'].str.contains('ARJ0|Bhaw|CHK0|CNR0|DMP0|HWY1|IDJK|INJK|JMU0|jmu1|JMU2|NAG0|RAJ0|SRN0|SRR1|VIJ0',na=False),'Circle']='JNK'

try:
    DAY1['Circle'] = DAY1['Circle'].replace('nan','')
except:
    pass

DAY1['Circle'] = DAY1['Circle'].replace('',None)
DAY1['Circle'] = DAY1['Circle'].fillna(DAY1['Server'])

# In[323]:


dff=DAY1.copy()

## Remove Blanks rows from columns

dff.dropna(subset=['Site A Tx Power Current', 'Site B Tx Power Current', 'Site A Rx Level Current', 'Site B Rx Level Current'], inplace=True)

# Rename the DataFrame to a new name, 
day1 = dff.copy()

#day1=day1[~((day1['Server']== 'ROB')& (day1['Circle'] == 'BIH'))]   # Delete BIH from ROB


#selected_rows = day1[day1['Circle'].isna()]

print("RAW Data done")

## SELECTED ONLY USED COLUMNS
day1=day1[['uniq link','Circle','Site A Tx Power Current','Site B Tx Power Current','Site A Rx Level Current','Site B Rx Level Current']]

day1=day1.drop_duplicates(subset='uniq link', keep='first')

#day1.to_excel(r'D:\DEEPAK-Office\Cera RSL\Cera RSL\RSL_Report_Ceragon.xlsx')




## -----------------------------------******************************************************* LB Data ************************************************
#



print(" LB reading start ")

lb_data=pd.read_excel(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera RSL\RAW\Ceragon_planning_Data.xlsx')

print(" LB read done ")


##RENAME
lb_data.rename(columns={'A-B':'uniq link'},inplace=True)

# Merging

df=pd.merge(lb_data,day1,on='uniq link',how='left')

df.rename(columns={'uniq link':'A-B'},inplace=True)

#df.to_excel(r'D:\DEEPAK-Office\Cera RSL\Cera RSL\RSL_Report_Ceragon.xlsx')

#df = df.drop(columns=['B-A'])

new_df = df[df['Site B Rx Level Current'].isna()] # select NAN values from df dataframe which is not lookup with LB data

new_df = new_df.drop(columns=['Site A Tx Power Current'])
new_df = new_df.drop(columns=['Site B Tx Power Current'])
new_df = new_df.drop(columns=['Site A Rx Level Current'])
new_df = new_df.drop(columns=['Site B Rx Level Current'])
new_df = new_df.drop(columns=['Circle'])

df = df.dropna(subset=['Site B Rx Level Current']) # drop Nan values in df data frame because we already take it into new_df

## Again Rename
new_df.rename(columns={'B-A':'uniq link'},inplace=True)

df2=pd.merge(new_df,day1,on='uniq link',how='left')
df2.rename(columns={'uniq link':'B-A'},inplace=True)

#new_df1 = df2[df2['Site B Rx Level Current'].isna()]  #check #NA count


final=pd.concat([df,df2])

df = final.sort_values(by='A-B', ascending=True)

#df =final # rename dataframe

df.rename(columns={'MW Node name in NMS_A':'Source NE Name','MW Node name in NMS_B':'Sink Name'},inplace=True)


df.loc[(df['A-B'].str.contains('INUE',na=False)),'Circle']='UPE'
df.loc[(df['A-B'].str.contains('IDUE',na=False)),'Circle']='UPE'
df.loc[(df['A-B'].str.contains('IDDL',na=False)),'Circle']='DEL'
df.loc[(df['A-B'].str.contains('INDL',na=False)),'Circle']='DEL'
df.loc[(df['A-B'].str.contains('IDUW',na=False)),'Circle']='UPW'
df.loc[(df['A-B'].str.contains('INUW',na=False)),'Circle']='UPW'
df.loc[(df['A-B'].str.contains('IDJK',na=False)),'Circle']='JNK'
df.loc[(df['A-B'].str.contains('INJK',na=False)),'Circle']='JNK'
df.loc[(df['A-B'].str.contains('IDAS',na=False)),'Circle']='ASM'
df.loc[(df['A-B'].str.contains('INNE',na=False)),'Circle']='ASM'
df.loc[(df['A-B'].str.contains('INAS',na=False)),'Circle']='ASM'
df.loc[(df['A-B'].str.contains('IDNE',na=False)),'Circle']='ASM'
df.loc[(df['A-B'].str.contains('IDBR',na=False)),'Circle']='BIH'
df.loc[(df['A-B'].str.contains('INBR',na=False)),'Circle']='BIH'
df.loc[(df['A-B'].str.contains('INOD',na=False)),'Circle']='ODI'
df.loc[(df['A-B'].str.contains('IDOD',na=False)),'Circle']='ODI'
df.loc[(df['A-B'].str.contains('INKA',na=False)),'Circle']='KAR'
df.loc[(df['A-B'].str.contains('IDKA',na=False)),'Circle']='KAR'
df.loc[(df['A-B'].str.contains('INKL',na=False)),'Circle']='KEL'
df.loc[(df['A-B'].str.contains('IDKL',na=False)),'Circle']='KEL'
df.loc[(df['A-B'].str.contains('INWB',na=False)),'Circle']='ROB'
df.loc[(df['A-B'].str.contains('IDKO',na=False)),'Circle']='ROB'
df.loc[(df['A-B'].str.contains('INKO',na=False)),'Circle']='ROB'

df.loc[(df['A-B'].str.contains('JMU0',na=False)),'Circle']='JNK'
df.loc[(df['A-B'].str.contains('PAR0',na=False)),'Circle']='JNK'
df.loc[(df['A-B'].str.contains('JMU1',na=False)),'Circle']='JNK'
df.loc[(df['A-B'].str.contains('JNK1',na=False)),'Circle']='JNK'
df.loc[(df['A-B'].str.contains('DRM0',na=False)),'Circle']='JNK'
df.loc[(df['A-B'].str.contains('MRS0',na=False)),'Circle']='JNK'
df.loc[(df['A-B'].str.contains('SCN0',na=False)),'Circle']='JNK'
df.loc[(df['A-B'].str.contains('KTR0',na=False)),'Circle']='JNK'
df.loc[(df['A-B'].str.contains('PNT0',na=False)),'Circle']='JNK'
df.loc[(df['A-B'].str.contains('RAJ0',na=False)),'Circle']='JNK'
df.loc[(df['A-B'].str.contains('SKJ0',na=False)),'Circle']='JNK'
df.loc[(df['A-B'].str.contains('SRR0',na=False)),'Circle']='JNK'
df.loc[(df['A-B'].str.contains('CDH0',na=False)),'Circle']='JNK'
df.loc[(df['A-B'].str.contains('BAN0',na=False)),'Circle']='JNK'
df.loc[(df['A-B'].str.contains('UDH0',na=False)),'Circle']='JNK'
df.loc[(df['A-B'].str.contains('AWP0',na=False)),'Circle']='JNK'


df.loc[(df['B-A'].str.contains('IDJK',na=False)),'Circle']='JNK'
df.loc[(df['B-A'].str.contains('INJK',na=False)),'Circle']='JNK'
df.loc[(df['B-A'].str.contains('SRR2',na=False)),'Circle']='JNK'


# Calculation

df['Tx Power_A[Planned-NMS]']=df['Main Tx Power (dBm)'] - df['Site A Tx Power Current']
df['Tx Power_B[Planned-NMS]']=df['Main Tx Power (dBm)'] - df['Site B Tx Power Current']


df['RSL_A[Planned-NMS]']=df['Rx Level (dBm)'] - df['Site A Rx Level Current']
df['RSL_B[Planned-NMS]']=df['Rx Level (dBm)'] - df['Site B Rx Level Current']


df['Delta_A']=df['Tx Power_A[Planned-NMS]'] - df['RSL_B[Planned-NMS]']
df['Delta_B']=df['Tx Power_B[Planned-NMS]'] - df['RSL_A[Planned-NMS]']


# Delta A

def check(value):
    if pd.isna(value) or value == '':
        return 'N/A'
    if value > 3:
        return 'High RSL'
    elif value < -3:
        return 'Low RSL'
    else:
        return 'OK'

df['Remarks_A'] = df['Delta_A'].apply(check)


'''
def check(value):
    if value>3:
        return 'High RSL'
    elif value<-3:
        return 'Low RSL'
    else:
        return 'OK'
df['Remarks_A']=df['Delta_A'].apply(check)
'''

# Delta B


def check(value):
    if pd.isna(value) or value == '':
        return 'N/A'
    if value > 3:
        return 'High RSL'
    elif value < -3:
        return 'Low RSL'
    else:
        return 'OK'

df['Remarks_B']=df['Delta_B'].apply(check) 

'''
def check(value):
    if value>3:
        return 'High RSL'
    elif value<-3:
        return 'Low RSL'
    else:
        return 'OK'
df['Remarks_B']=df['Delta_B'].apply(check)  
'''


#df['Tx Power/RSL Deviation Remarks']=df.apply(lambda row: 'Tx Power/RSL Deviation' if (row['Remarks_A']=='High RSL' or row['Remarks_A']=='Low RSL' or row['Remarks_B']=='High RSL' or row['Remarks_B']=='Low RSL') else 'OK', axis=1)

df['Min value of both Delta'] =df[['Delta_A', 'Delta_B']].min(axis=1) 

# using function menthod
def result(row):
    if row['Remarks_A']=='N/A' and row['Remarks_B']=='N/A':
        return 'N/A'        
    elif row['Remarks_A']=='High RSL' and row['Remarks_B']=='High RSL':
        return 'High RSL'
    elif row['Remarks_A']=='Low RSL' and row['Remarks_B']=='Low RSL':
        return 'Low RSL'
    elif row['Remarks_A']=='OK' and row['Remarks_B']=='OK':
        return 'OK'
    elif row['Remarks_A']=='OK' and row['Remarks_B']=='Low RSL':
        return 'Low RSL'
    elif row['Remarks_A']=='Low RSL' and row['Remarks_B']=='OK':
        return 'Low RSL'
    elif row['Remarks_A']=='Low RSL' and row['Remarks_B']=='High RSL':
        return 'Low RSL'
    elif row['Remarks_A']=='High RSL' and row['Remarks_B']=='Low RSL':
        return 'Low RSL'
    else:
        return 'High RSL'        
df['Low/High RSL Remarks']= df.apply(result, axis=1)



# using pandas/numpy lib
conditions = [
    (df['Low/High RSL Remarks'] == 'OK'),
    (df['Low/High RSL Remarks'] == 'High RSL'),
    (df['Min value of both Delta'] <= -10),
    (df['Min value of both Delta'] <= -5) & (df['Min value of both Delta'] > -10),
    (df['Min value of both Delta'] <= -3) & (df['Min value of both Delta'] > -5),
    (df['Min value of both Delta'] > -3) & (df['Min value of both Delta'] < 0),
     
]
choices=['OK', 'High RSL','P1 (<-10 dB)', 'P2(-5dB to >- 10dB)', 'P3(-3dB to  >- 5dB)','N/A']
df['Low RSL Deviations (Priority)']=np.select(conditions, choices, default='N/A')

#####********Rearrange column header*********


df = df.reindex(columns=['Circle', 'A-B','B-A', 'Source NE Name', 'Sink Name','Main Tx Power (dBm)', 'Rx Level (dBm)' , 'Site A Tx Power Current',
                         'Site A Rx Level Current', 'Site B Tx Power Current','Site B Rx Level Current','Tx Power_A[Planned-NMS]','RSL_A[Planned-NMS]',
                         'Delta_A','Remarks_A','Tx Power_B[Planned-NMS]','RSL_B[Planned-NMS]','Delta_B','Remarks_B','Min value of both Delta',
                         'Low RSL Deviations (Priority)','Low/High RSL Remarks'])

df = df.dropna(subset=['Circle'])


#df.to_excel(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera RSL\Output\RSL_Report_Ceragon.xlsx')


print(" Pivot Start ")


pivot_df = pd.pivot_table(df, values=None, index='Circle', columns='Low RSL Deviations (Priority)', aggfunc='size', fill_value=0)
pivot_df= pivot_df.reset_index()

pivot_df = pivot_df._append(pivot_df.sum(axis=0),ignore_index=True)
pivot_df['Circle'] = np.where(pivot_df['Circle'].str.contains('ASMBIH', na=False), 'Grand Total', pivot_df['Circle']) # Column Sum

print (pivot_df)

pivot_df['Total'] = pivot_df[['High RSL', 'N/A','OK', 'P1 (<-10 dB)', 'P2(-5dB to >- 10dB)','P3(-3dB to  >- 5dB)']].sum(axis=1) # Row Sum

print (pivot_df)


#pivot_df = pd.pivot_table(df, values='Count', index='Circle', columns='Low RSL Deviations (Priority)', aggfunc='sum', fill_value=0)

#pivot_df['Grand Total'] = pivot_df.sum(axis=1) # Add a Grand Total column (sum of rows)


#pivot_df.loc['Grand Total'] = pivot_df.sum(axis=0) # Add a Grand Total row (sum of columns)


print("writing")

writer = pd.ExcelWriter(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera RSL\Output\TXN_VIL_PAN_INDIA_CERA_RSL_DEVIATION_'+d7+'_DAILY.xlsx')
pivot_df.to_excel(writer, sheet_name='Summary',index=False)
#final.to_excel(writer, sheet_name='Details',index=False)
df.to_excel(writer, sheet_name='Details',index=False)
writer.close()

print("done")


#df.to_excel(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera RSL\RSL_Report_Ceragon.xlsx')



                                                    # ** Reports Uploading **   

print(" ** Reports Uploading start on cobra and BI Portal -->> /home/snenrc/VIL_IDEA_REPORTS/TX_REPORT ** ")

print(" ** CERA RSL Uploading Start ** ")


ssh3=paramiko.SSHClient()
ssh3.set_missing_host_key_policy(paramiko.AutoAddPolicy())
try:
    ssh3.connect(hostname='10.10.10.10',username='admin',password='admin',port=22)
except:
    pass

try:
    ssh3.connect(hostname='11.11.11.11',username='admin',password='admin',port=22)
except:
    pass

sftp_client1=ssh3.open_sftp()



try:
    sftp_client1.chdir('/opt/TX/RSL_Dev_Reports_Daily')
    sftp_client1.put(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera RSL\Output\TXN_VIL_PAN_INDIA_CERA_RSL_DEVIATION_'+d7+'_DAILY.xlsx','TXN_VIL_PAN_INDIA_CERA_RSL_DEVIATION_'+d7+'_DAILY.xlsx')   
except:
    pass

sftp_client1.close
ssh3.close




print(" ** CERA RSL Uploading Start ** ")

ssh0=paramiko.SSHClient()
ssh0.set_missing_host_key_policy(paramiko.AutoAddPolicy())
ssh0.connect(hostname='11.11.11.11',username='root',password='root',port=22)
sftp_client0=ssh0.open_sftp()


try:
    sftp_client0.chdir('/home/snenrc/VIL_IDEA_REPORTS/TX_REPORT')
    sftp_client0.put(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Cera RSL\Output\TXN_VIL_PAN_INDIA_CERA_RSL_DEVIATION_'+d7+'_DAILY.xlsx','TXN_VIL_PAN_INDIA_CERA_RSL_DEVIATION_'+d7+'_DAILY.xlsx')   
except:
    pass

sftp_client0.close
ssh0.close


print(" ********* Congratulations RSL Report Successfully Uploaded *********** ")
