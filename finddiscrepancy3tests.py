# -*- coding: utf-8 -*-
"""
Created on Tue May 22 15:32:10 2018

@author: mattr
"""

# -*- coding: utf-8 -*-
"""
Created on Mon May 21 09:10:54 2018

@author: mattr
"""

import pandas as pd

from savReaderWriter import SavReader


'************************************************'
full='Maths'
part='Maths'
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

test1={1:'M2AR01',2:'M2AR07',3:'M2AR02',4:'M2AR08',5:'M2AR03',6:'M2AR09',7:'M2AR04',8:'M2AR10',9:'M2AR05',10:'M2ARAT',11:'M2AR06',12:'More than one box ticked'}
test2={1:'N6S8P2',2:'N2S2P2',3:'N8S5P2',4:'N3SAP2',5:'NBS6P2',6:'N7S7P2',7:'N5SBP2',8:'N1S1P2',9:'NAS4P2',10:'NASAP2',11:'N4S3P2',12:'More than one box ticked'}
test3={1:'N1SAP3',2:'N7SBP3',3:'N2S4P3',4:'N8S7P3',5:'N3S2P3',6:'NAS3P3',7:'N4S1P3',8:'NBS5P3',9:'N5S8P3',10:'NBSBP3',11:'N6S6P3',12:'More than one box ticked'}

excelt1=[]
excelt2=[]
excelt3=[]
spsst1=[]
spsst2=[]
spsst3=[]
nferno=[]
pupilid=[]


def find_discrepancy3tests(AQorTQ,excelt1name, excelt2name,excelt3name, test1q, test2q, test3q, outputfile):

    for i in range(len(spss)):
        s_pupilid=spss['pupilid'][i]
        s_nferno=spss['nferno'][i]
        e_pupilid=(excel[AQ_TQ][excel[AQ_TQ].str.contains(s_pupilid)==True]).iloc[0]
        e_index=(excel.index[excel[AQ_TQ]==e_pupilid].tolist())[0]
        e_nferno=excel['NFERNo'][e_index]
        exceltest1=excel[excelt1name][e_index]
        exceltest2=excel[excelt2name][e_index]
        exceltest3=excel[excelt3name][e_index]
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
        try:
            spsstest3=test3[spss[test3q][i]]
            t3=1
        except KeyError:
            print('there is no spsstest3, pupilid:', s_pupilid)
            t3=0
        if (t1==1) and (type(exceltest2)==float):
            continue
        elif (t1==1) and (t2==0) and (t3==0):
            spsst1.append(spsstest1)
            spsst2.append('n/a')
            spsst3.append('n/a')
            excelt3.append(exceltest3)
            excelt1.append(exceltest1)
            excelt2.append(exceltest2)
            nferno.append(e_nferno)
            pupilid.append(e_pupilid)
            continue
        elif (t1==1) and (t2==0) and (t3==1):
            spsst1.append(spsstest1)
            spsst2.append('n/a')
            spsst3.append(spsstest3)
            excelt3.append(exceltest3)
            excelt1.append(exceltest1)
            excelt2.append(exceltest2)
            nferno.append(e_nferno)
            pupilid.append(e_pupilid)
            continue
        elif (t1==1) and (t2==1) and (t3==0):
            spsst1.append(spsstest1)
            spsst2.append(spsstest2)
            spsst3.append('n/a')
            excelt3.append(exceltest3)
            excelt1.append(exceltest1)
            excelt2.append(exceltest2)
            nferno.append(e_nferno)
            pupilid.append(e_pupilid)
            continue
        elif (t1==0) and (t2==1) and (t3==0):
            spsst1.append('n/a')
            spsst2.append(spsstest2)
            spsst3.append('n/a')
            excelt3.append(exceltest3)
            excelt1.append(exceltest1)
            excelt2.append(exceltest2)
            nferno.append(e_nferno)
            pupilid.append(e_pupilid)
            continue        
        elif (t1==0) and (t2==1) and (t3==1):
            spsst1.append('n/a')
            spsst2.append(spsstest2)
            spsst3.append(spsstest3)
            excelt3.append(exceltest3)
            excelt1.append(exceltest1)
            excelt2.append(exceltest2)
            nferno.append(e_nferno)
            pupilid.append(e_pupilid)
            continue      
        elif (t1==0) and (t2==0) and (t3==1):
            spsst1.append('n/a')
            spsst2.append('n/a')
            spsst3.append(spsstest3)
            excelt3.append(exceltest3)
            excelt1.append(exceltest1)
            excelt2.append(exceltest2)
            nferno.append(e_nferno)
            pupilid.append(e_pupilid)
            continue    
        elif (t1==0) and (t2==0) and (t3==0):
            spsst1.append('n/a')
            spsst2.append('n/a')
            spsst3.append('n/a')
            excelt3.append(exceltest3)
            excelt1.append(exceltest1)
            excelt2.append(exceltest2)
            nferno.append(e_nferno)
            pupilid.append(e_pupilid)
            continue    
        if (exceltest1 != spsstest1) and (exceltest2 == spsstest2) and (exceltest3 == spsstest3):
            print('test1 is in the wrong order, pupilid:',s_pupilid)
            spsst1.append(('wrong order ',spsstest1))
            spsst2.append(spsstest2)
            spsst3.append(spsstest3)
            excelt1.append(exceltest1)
            excelt2.append(exceltest2)
            excelt3.append(exceltest3)
            nferno.append(e_nferno)
            pupilid.append(e_pupilid)
            continue
        elif (exceltest1 != spsstest1) and (exceltest2 == spsstest2) and (exceltest3!=spsstest3):
            print(' test1 and test3 is in the wrong order, pupilid:', s_pupilid)
            spsst1.append(('wrong order ',spsstest1))
            spsst2.append(spsstest2)
            spsst3.append(('wrong order', spsstest3))
            excelt1.append(exceltest1)
            excelt2.append(exceltest2)
            excelt3.append(exceltest3)
            nferno.append(e_nferno)
            pupilid.append(e_pupilid)           
        elif (exceltest1 != spsstest1) and (exceltest2 != spsstest2) and (exceltest3==spsstest3):
            print('test1 and test2 is in the wrong order, pupilid:', s_pupilid)
            spsst1.append(('wrong order ',spsstest1))
            spsst2.append(('wrong order ',spsstest2))
            spsst3.append(spsstest3)
            excelt1.append(exceltest1)
            excelt2.append(exceltest2)
            excelt3.append(exceltest3)
            nferno.append(e_nferno)
            pupilid.append(e_pupilid)
        elif (exceltest1 == spsstest1) and (exceltest2 == spsstest2) and (exceltest3!=spsstest3):
            print(' test3 is in the wrong order, pupilid:', s_pupilid)
            spsst1.append(spsstest1)
            spsst2.append(spsstest2)
            spsst3.append(('wrong order',spsstest3))
            excelt1.append(exceltest1)
            excelt2.append(exceltest2)
            excelt3.append(exceltest3)
            nferno.append(e_nferno)
            pupilid.append(e_pupilid)
        elif (exceltest1 == spsstest1) and (exceltest2 != spsstest2) and (exceltest3==spsstest3):
            print('the test2 is in the wrong order, pupilid:', s_pupilid)
            spsst1.append(spsstest1)
            spsst2.append(('wrong order',spsstest2))
            spsst3.append(spsstest3)
            excelt1.append(exceltest1)
            excelt2.append(exceltest2)
            excelt3.append(exceltest3)
            nferno.append(e_nferno)
            pupilid.append(e_pupilid) 
        elif (exceltest1 == spsstest1) and (exceltest2 != spsstest2) and (exceltest3!=spsstest3):
            print('test2 and test3 is in the wrong order, pupilid:', s_pupilid)
            spsst1.append(spsstest1)
            spsst2.append(('wrong order',spsstest2))
            spsst3.append(('wrong order',spsstest3))
            excelt1.append(exceltest1)
            excelt2.append(exceltest2)
            excelt3.append(exceltest3)
            nferno.append(e_nferno)
            pupilid.append(e_pupilid)             
        elif (exceltest1 != spsstest1) and (exceltest2 != spsstest2) and (exceltest3 !=spsstest3):
            print('all is in the wrong order, pupilid:', s_pupilid)
            spsst1.append(('wrong order',spsstest1))
            spsst2.append(('wrong order ',spsstest2))
            spsst3.append(('wrong order ',spsstest3))
            excelt1.append(exceltest1)
            excelt2.append(exceltest2)
            excelt3.append(exceltest3)
            nferno.append(e_nferno)
            pupilid.append(e_pupilid)           
    
    #save to DQA file
    discrepancies=pd.DataFrame({'Nfer No':nferno,'Pupil ID':pupilid,'excelt1':excelt1,'excelt2':excelt2,
                                'excelt3':excelt3,'spsst1':spsst1, 'spsst2':spsst2, 'spsst3':spsst3})
    
    discrepancies.to_excel(outputfile,index =False,engine='xlsxwriter')
    
#find_discrepancy('excelt1name', 'excelt2name','excelt3name' ,'test1q', 'test2q','test3q','outputfile')
    
find_discrepancy3tests(AQ_TQ,'LC1_Instrument1','LC1_Instrument2','LC1_Instrument3','Q1','Q7','Q14','X:\STAO\TQ\Maths\Disrepancy_STAO_MATHS_TQ.xlsx')



    
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
        spsst1.append('n/a')
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