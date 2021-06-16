# -*- coding: utf-8 -*-
"""
Created on Thu May 02 13:48:06 2019

@author: qsb15202
"""

import win32com.client
import scipy.io
import numpy as np
import pandas as pd
import opendssdirect as dss
from opendssdirect.utils import run_command
import csv
import random
import cmath
import math

##### Load in the Harmonic Profiles ########
#derated
#rated_cv
rated_cc=pd.read_excel('rated_cc_stats.xlsx', sheet_name=None)
del rated_cc['BMW_3ph_CC']
del rated_cc['Zoe_3ph_CC']
del rated_cc['DC_CC']
g55lims=pd.read_csv('g55limits.csv')

EV_power=pd.Series(index=rated_cc.keys(),data=[6.6,6.6,7.2,7.2,7.2])
####### OpenDSS Initialisation #########

dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")
dssText = dssObj.Text
DSSCircuit = dssObj.ActiveCircuit
DSSLoads=DSSCircuit.Loads;

###### Source Impedence is adjusted for Urban/Rural networks


####------Build Lines
M=5  ##--Number of EVs
B=20 ##--Number of Buses

T='R'
f_Rsc=pd.Series(dtype=float,index=[15,33])
f_Rsc[15]=1.75  #Line length adjustment to give Ifmin=480A corresponding to Rsc=15 at the end of the feeder
f_Rsc[33]=0.56   #Line length adjustment to give Ifmin=1056A corresponding to Rsc=15 at the end of the feeder

def lineBuilder(type,B,f_Rsc):
    Lines=pd.read_csv('Lines.txt',delimiter=' ', names=['New','Line','Bus1','Bus2','phases','Linecode','Length','Units'])
    Lines['Length'][0]='Length='+str(f_Rsc*1/(B))
    
    for L in range(1,B):
        Lines=Lines.append(Lines.loc[0], ignore_index=True)
        Lines['Line'][L]='Line.LINE'+str(L+1)
        Lines['Bus1'][L]='Bus1='+str(L+2)
        Lines['Bus2'][L]='Bus2='+str(L+3)
     
    Lines.to_csv('Lines_'+str(T)+'.txt', sep=" ", index=False, header=False)

####---- Create Spectrum CSVs
def create_spectrum(rated_cc):
    THD_C=pd.Series(dtype=float,index=list(rated_cc.keys()))
    for i in list(rated_cc.keys()):
        spectrum=pd.DataFrame()
        spectrum['h']=rated_cc[i]['Harmonic order']
        spectrum['mag']=rated_cc[i]['Ih mag 75% percentile']/rated_cc[i]['Ih mag 75% percentile'][0]*100
        spectrum['ang']=rated_cc[i]['Ih phase mean w.r.t. L1_Ih1']
        spectrum.to_csv('Spectrum'+i+'.csv', header=False, index=False)

        THD_C[i]=(sum(spectrum['mag'][1:]**2))**0.5
        
    return THD_C

def export_loadflow():
    dssText.Command="export voltages"
    dssText.Command="export currents"
    dssText.Command="export powers"
    dssText.Command="export loads"
    dssText.Command="export losses"
    
    currents=pd.read_csv('LVTest_EXP_CURRENTS.csv')
    powers=pd.read_csv('LVTest_EXP_POWERS.csv')
    voltages=pd.read_csv('LVTest_EXP_VOLTAGES.csv')
    loads_out=pd.read_csv('LVTest_EXP_LOADS.csv')
    losses=pd.read_csv('LVTest_EXP_LOSSES.csv')
    
    return currents, powers, voltages, loads_out, losses

def fault_seqz():
    dssText.Command="Solve Mode=Faultstudy"
    dssText.Command="export Faultstudy"
    dssText.Command="export seqz"
    
    seqz=pd.read_csv('LVTest_EXP_SEQZ.csv')
    faults=pd.read_csv('LVTest_EXP_FAULTS.csv')
    
    return seqz,faults
  
THD_C=create_spectrum(rated_cc)

Loads=pd.read_csv('Loads.txt', delimiter=' ', names=['New','Load','Phases','Bus1','kV','kW','PF','spectrum'])


Vh_percent={}
Pass={}
All_Pass={}
pflowRes={}
Ch_Ratios=pd.DataFrame()
VoltageMin={}

for f in [15,33]:
    pflowRes[f]={}
    Vh_percent[f]=pd.DataFrame()
    Pass[f]=pd.DataFrame()
    All_Pass[f]=pd.DataFrame(index=range(1,M+1))
    VoltageMin[f]=pd.DataFrame(index=range(1,M+1))
    
    lineBuilder(T, B,f_Rsc[f])
    for r in range(1,10):
        All_Pass[f][r]=range(1,M+1)
        VoltageMin[f][r]=range(1,M+1)
        for n in range(1,M+1):
            #i=random.choice(list(rated_cc.keys()))
            i=list(rated_cc.keys())[0]
            ###--- Add loads
            Loads=pd.read_csv('Loads.txt', delimiter=' ', names=['New','Load','Phases','Bus1','kV','kW','PF','spectrum'])
            Loads['spectrum'][0]='spectrum='+str(i)
            Loads['kW'][0]='kW='+str(EV_power[i])
            b0=random.choice(range(2,B))
            Loads['Bus1'][0]='Bus1='+str()+str(b0)+'.1'
            for p in range(2,4):
                Loads=Loads.append(Loads.loc[0], ignore_index=True)
                Loads['Bus1'][p-1]='Bus1='+str(b0)+'.'+str(p)
                Loads['Load'][p-1]='Load.LOAD'+str(p)
            c=3
            for k in range(1,n):
                b=random.choice(range(2,B))
                for p in range(1,4):
                    Loads=Loads.append(Loads.loc[0], ignore_index=True)
                    Loads['Bus1'][c]='Bus1='+str(b)+'.'+str(p)
                    Loads['Load'][c]='Load.LOAD'+str(c+1)
                    Loads['kW'][c]='kW='+str(EV_power[i])
                    c=c+1
                Loads['spectrum'][k]='spectrum='+str(i)
            
            Loads.to_csv('Loads_'+T+'.txt', sep=" ", index=False, header=False)
            
            dssObj.ClearAll() 
            dssText.Command="Compile E:/PNDC/TN-006Harmonic/2Bus/Master_R.dss"
            dssText.Command="Solve"
            dssText.Command="Solve Mode=harmonics NeglectLoadY=Yes"
            dssText.Command="export monitors m1"
            res_Reactor=pd.read_csv('LVTest_Mon_m1_1.csv')
            Vh_ratios=pd.DataFrame()
            Vh_ratios['h']=res_Reactor[' Harmonic']
            Vh_ratios['V']=res_Reactor[' V1']
            Vh_ratios['V_ratio']=res_Reactor[' V1']/res_Reactor[' V1'][0]*100
            Vh_ratios['Lims']=g55lims['L']
            Ch_Ratios['h']=res_Reactor[' Harmonic']
            Ch_Ratios[n]=res_Reactor[' I1']/res_Reactor[' I1'][0]*100
            Vh_percent[f]['h']=res_Reactor[' Harmonic']
            
            Vh_percent[f][n]=Vh_ratios['V_ratio']
            Vh_percent[f][n][0] = sum(Vh_ratios['V_ratio'][1:]**2)**0.5
            Pass[f]['h']=res_Reactor[' Harmonic']
            Pass[f][n]=Vh_percent[f][n]<Vh_ratios['Lims']
             
            dssObj.ClearAll() 
            dssText.Command="Compile E:/PNDC/TN-006Harmonic/2Bus/Master_R.dss"
            dssText.Command="Solve"
            pflowRes[f]['currents'], pflowRes[f]['powers'], pflowRes[f]['voltages'], pflowRes[f]['loads_out'], pflowRes[f]['losses'] = export_loadflow()
            VoltageMin[f][r][n]=pflowRes[f]['voltages'][' Magnitude1'][-1:].values[0]
            
        All_Pass[f][r]=Pass[f].iloc[:,1:].sum()==30
    Vh_percent[f]['h'][0]='THD'
    seqz,faults=fault_seqz()
    
    print(f,faults)
    #print(i,All_Pass[f])

def validate():
    #### Validate results
    r=pd.Series(dtype=float,index=res_Reactor.index)
    q=pd.Series(dtype=float,index=res_Reactor.index)
    for i in res_Reactor.index:
        r[i]=res_Reactor[' I1'][i]*math.cos(res_Reactor[' IAngle1'][i])
        q[i]=res_Reactor[' I1'][i]*math.sin(res_Reactor[' IAngle1'][i])
    r=pd.Series(dtype=float,index=res_Reactor.index)
    q=pd.Series(dtype=float,index=res_Reactor.index)
    icplx=pd.Series(dtype=float,index=res_Reactor.index)
    
    for i in res_Reactor.index:
        r[i]=res_Reactor[' I1'][i]*math.cos(res_Reactor[' IAngle1'][i])
        q[i]=res_Reactor[' I1'][i]*math.sin(res_Reactor[' IAngle1'][i])
        icplx[i]=complex(r[i],q[i])
        
    z=complex(0.053,0.009)
    
    vcplx=icplx*z
    
    p=pd.Series(dtype=float,index=res_Reactor.index)
    theta=pd.Series(dtype=float,index=res_Reactor.index)
    for i in res_Reactor.index:
        p[i],theta[i]=cmath.polar(vcplx[i])
    
    res_Reactor['Vmag']=p
    res_Reactor['angle']=theta