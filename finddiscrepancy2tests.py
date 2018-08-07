# -*- coding: utf-8 -*-
"""
Created on Mon May 21 09:10:54 2018

@author: mattr
"""

import pandas as pd

from savReaderWriter import SavReader


'************************************************'
full='GPS'
part='GPS'
AQ_TQ='TQ'
STAA_STAO='STAO'



#read spss.sav file
with SavReader('X:\{}\{}\{}\{}_{}_{}_BACKGROUND.sav'.format(STAA_STAO,AQ_TQ,full,STAA_STAO,part,AQ_TQ),ioUtf8=True,returnHeader=True,idVar='Q1') as reader:
    records=reader.all()
    
    
#the first item of records list is the columns of the dataframe
columns=[records[0]]
columns=columns[0]
#delete that first item
del records[0]
#convert records to a dataframe
spss=pd.DataFrame(records)
#rename columns
spss.columns=columns
    
excel=pd.read_excel('K:\{}\RPO\Dispatch\Letter 3 files\{}\Finish {} letter 3 - RM.xlsx'.format(STAA_STAO,full,full),sheet_name='Sheet1')
excel[AQ_TQ]=excel[AQ_TQ].apply(str)

#STAOREADtest1={1:'RAB34A',2:'RAB35B',3:'RAB34B',4:'RAB35C',5:'RAB34C',6:'RAB35D',7:'RAB34D',8:'RABXAT',9:'RAB35A',10:'More than one box ticked'}
#STAOREAD
#STAAREADtest1={1:'RPA23a',2:'RPA23b',3:'RPA23c',4:'RPA24a',5:'RPA24b',6:'RPA24c',7:'RPA1AT'}
#STAAREADtest2={1:'RABX01',2:'RABX02',3:'RABX03',4:'RABW01',5:'RABW02',6:'RABW03',7:'RAB1AT'}

test1={1:'G2V001',2:'G2V006',3:'G2V002',4:'G2V007',5:'G2V003',6:'G2V008',7:'G2V004',8:'G2VXAT',9:'G2V005',10:'More than one box ticked'}
test2={1:'G2V001',2:'G2V002',3:'G2V003',4:'G2V004',5:'G2V005',6:'G2V006',7:'G2V007',8:'G2V008'}


def find_discrepancy2tests(AQorTQ,excelt1name, excelt2name, test1q, test2q, outputfile):
    excelt1=[]
    excelt2=[]
    spsst1=[]
    spsst2=[]
    nferno=[]
    pupilid=[]

    for i in range(len(spss)):
        s_pupilid=spss['pupilid'][i]
        s_nferno=spss['nferno'][i]
        e_pupilid=(excel[AQ_TQ][excel[AQ_TQ].str.contains(s_pupilid)==True]).iloc[0]
        e_index=(excel.index[excel[AQ_TQ]==e_pupilid].tolist())[0]
        e_nferno=excel['NFERNo'][e_index]
        exceltest1=excel[excelt1name][e_index]
        exceltest2=excel[excelt2name][e_index]
        # if pupil ids and nfer bis dont match
        if ((e_pupilid != s_pupilid) and (e_nferno != s_nferno)) and (s_pupilid not in e_pupilid):
            print(s_pupilid,' not found in excel')
            continue
        try:
            spsstest1=test1[spss[test1q][i]]
            t1=1
        except KeyError:
            print('there is no spsstest1, pupilid:', s_pupilid)
            t1=0
        try:
            spsstest2=test2[spss[test2q][i]]
            t2=1
        except KeyError:
            print('there is no spsstest2, pupilid:', s_pupilid)
            t2=0
        if (t1==1) and (type(exceltest2)==float):
            continue
        elif (t1==1) and (t2==0):
            spsst1.append(spsstest1)
            spsst2.append('n/a')
            excelt1.append(exceltest1)
            excelt2.append(exceltest2)
            nferno.append(e_nferno)
            pupilid.append(e_pupilid)
            continue
        elif (t1==0) and (t2==1):
            spsst1.append('n/a')
            spsst2.append(spsstest2)
            excelt1.append(exceltest1)
            excelt2.append(exceltest2)
            nferno.append(e_nferno)
            pupilid.append(e_pupilid)
            continue        
        elif (t1==0) and (t2==0):
            spsst1.append('n/a')
            spsst2.append('n/a')
            excelt1.append(exceltest1)
            excelt2.append(exceltest2)
            nferno.append(e_nferno)
            pupilid.append(e_pupilid)
            continue              
        if (exceltest1 != spsstest1) and (exceltest2 == spsstest2):
            print('the test1 is in the wrong order, pupilid:',s_pupilid)
            spsst1.append(('wrong order ',spsstest1))
            spsst2.append(spsstest2)
            excelt1.append(exceltest1)
            excelt2.append(exceltest2)
            nferno.append(e_nferno)
            pupilid.append(e_pupilid)
            continue
        elif (exceltest1 == spsstest1) and (exceltest2 != spsstest2):
            print('the test2 is in the wrong order, pupilid:', s_pupilid)
            spsst1.append(spsstest1)
            spsst2.append(('wrong order ',spsstest2))
            excelt1.append(exceltest1)
            excelt2.append(exceltest2)
            nferno.append(e_nferno)
            pupilid.append(e_pupilid)           
    
    #save to DQA file
    discrepancies=pd.DataFrame({'Nfer No':nferno,'Pupil ID':pupilid,'excelt1':excelt1,'excelt2':excelt2,
                                'spsst1':spsst1, 'spsst2':spsst2})
    
    discrepancies.to_excel(outputfile,index =False,engine='xlsxwriter')
    
#find_discrepancy('excelt1name', 'excelt2name', 'test1q', 'test2q','outputfile')
    
find_discrepancy2tests(AQ_TQ,'LC1_Instrument1','LC1_Instrument2','Q1','Q11','X:\STAO\TQ\GPS\Disrepancy_STAO_GPS_TQ.xlsx')
    
'''
excelt1name='LC1_Instrument2'
excelt2name='LC1_Instrument4'
test1q= 'Q1'
test2q='Q19'

excelt1=[]
excelt2=[]
spsst1=[]
spsst2=[]
nferno=[]
pupilid=[]

for i in range(10):
    s_pupilid=spss['pupilid'][i]
    s_nferno=spss['nferno'][i]
    e_pupilid=(excel[AQ_TQ][excel[AQ_TQ].str.contains(s_pupilid)==True]).iloc[0]
    e_index=(excel.index[excel[AQ_TQ]==e_pupilid].tolist())[0]
    e_nferno=excel['NFERNo'][e_index]
    exceltest1=excel[excelt1name][e_index]
    exceltest2=excel[excelt2name][e_index]
    # if pupil ids and nfer bis dont match
    if ((e_pupilid != s_pupilid) and (e_nferno != s_nferno)) and (s_pupilid not in e_pupilid):
        print(s_pupilid,' not found in excel')
        continue
    try:
        spsstest1=test1[spss[test1q][i]]
        t1=1
    except KeyError:
        print('there is no spsstest1, pupilid:', s_pupilid)
        t1=0
    try:
        spsstest2=test2[spss[test2q][i]]
        t2=1
    except KeyError:
        print('there is no spsstest2, pupilid:', s_pupilid)
        t2=0
    if (t1==1) and (t2==0):
        spsst1.append(spsstest1)
        spsst2.append('n/a')
        excelt1.append(exceltest1)
        excelt2.append(exceltest2)
        nferno.append(e_nferno)
        pupilid.append(e_pupilid)
        continue
    elif (t1==0) and (t2==1):
        spsst1.append('n/a')
        spsst2.append(spsstest2)
        excelt1.append(exceltest1)
        excelt2.append(exceltest2)
        nferno.append(e_nferno)
        pupilid.append(e_pupilid)
        continue        
    elif (t1==0) and (t2==0):
        spsst1.append('n/a')
        spsst2.append('n/a')
        excelt1.append(exceltest1)
        excelt2.append(exceltest2)
        nferno.append(e_nferno)
        pupilid.append(e_pupilid)
        continue               
    if (exceltest1 != spsstest1) and (exceltest2 == spsstest2):
        print('the test1 is in the wrong order, pupilid:',s_pupilid)
        spsst1.append(('wrong order ',spsstest1))
        spsst2.append(spsstest2)
        excelt1.append(exceltest1)
        excelt2.append(exceltest2)
        nferno.append(e_nferno)
        pupilid.append(e_pupilid)
        continue
    elif (exceltest1 == spsstest1) and (exceltest2 != spsstest2):
        print('the test2 is in the wrong order, pupilid:', s_pupilid)
        spsst1.append(spsstest1)
        spsst2.append(('wrong order ',spsstest2))
        excelt1.append(exceltest1)
        excelt2.append(exceltest2)
        nferno.append(e_nferno)
        pupilid.append(e_pupilid)
   '''