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
import matplotlib.pyplot as plt
import pickle
from matplotlib.ticker import (MultipleLocator,
                               FormatStrFormatter,
                               AutoMinorLocator)
##### Load in the Harmonic Profiles ########
#derated
#rated_cv
case='CC_Unbalanced_Urban'
cm='CC'
g55lims=pd.read_csv('g55limits.csv')

if cm=='CV' or cm=='CC':
    rated_cc=pd.read_excel('rated_'+str(cm)+'_stats.xlsx', sheet_name=None)
    del rated_cc['BMW_3ph_'+str(cm)]
    del rated_cc['Zoe_3ph_'+str(cm)]
    del rated_cc['DC_'+str(cm)]
    EV_power=pd.Series(index=rated_cc.keys(),data=[6.6,6.6,7.2,7.2,7.2])
    
if cm=='de':
    rated_cc=pd.read_excel('derated_stats.xlsx', sheet_name=None)
    del rated_cc['bmw_3ph_6A']
    del rated_cc['bmw_3ph_9A']
    del rated_cc['bmw_3ph_12A']
    del rated_cc['bmw_3ph_15A']
    del rated_cc['zoe_3ph_6A']
    del rated_cc['zoe_3ph_12A']
    del rated_cc['zoe_3ph_18A']
    del rated_cc['zoe_3ph_24A']
    EV_power=pd.Series(index=rated_cc.keys(),data=[1.38,2.76,4.14,5.52,1.38,2.76,4.14,5.52,1.38,2.76,4.14,5.52,1.38,2.76,4.14,5.52,1.38,2.76,4.14,5.52])
    
    # ###--- Code to only include <12A charging
    # for i in list(rated_cc.keys()):
    #     if i[-2:] =='4A' or i[-2:] =='8A':
    #         del rated_cc[i]
    #         del EV_power[i]

########---------- Get rid of ZOE and BMW--------
# del rated_cc['Zoe_1ph_CC']
# del EV_power['Zoe_1ph_CC']
# del rated_cc['BMW_1ph_CC']
# del EV_power['BMW_1ph_CC']

####### OpenDSS Initialisation #########

dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")
dssText = dssObj.Text
DSSCircuit = dssObj.ActiveCircuit
DSSLoads=DSSCircuit.Loads;

###### Source Impedence is adjusted for Urban/Rural networks

M=40  ##--Number of EVs
B=20 ##--Number of Buses
R=100##--Number of Runs

RSCs=['WPD_Zmax',15,33,66]   ##--- FOr Urban where WPD ZMax is much higher than corresponding RSC
##RSCs=[15,33,66]  ###--- For Rural where WPD ZMax and RSC=15 are similar

#f_Rsc=pd.Series(dtype=float,index=RSCs,data=[1.6,0.62,0.21]) #for 185 mm - RURAL
f_Rsc=pd.Series(dtype=float,index=RSCs,data=[0.77,1.7,0.72,0.305]) #for 185 mm - URBAN



####------Build Lines between B buses
def lineBuilder(B,f_Rsc):
    Lines=pd.read_csv('Lines.txt',delimiter=' ', names=['New','Line','Bus1','Bus2','phases','Linecode','Length','Units'])
    Lines['Length'][0]='Length='+str(f_Rsc*1/(B))
    
    for L in range(1,B):
        Lines=Lines.append(Lines.loc[0], ignore_index=True)
        Lines['Line'][L]='Line.LINE'+str(L+1)
        Lines['Bus1'][L]='Bus1='+str(L+2)
        Lines['Bus2'][L]='Bus2='+str(L+3)
     
    Lines.to_csv('Lines_R.txt', sep=" ", index=False, header=False)

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
    dssObj.ClearAll() 
    dssText.Command="Compile E:/PNDC/TN-006Harmonic/Harmonic_OpenDSS/Master_R.dss"
    dssText.Command="Solve"


    dssText.Command="export voltages"
    dssText.Command="export currents"
    dssText.Command="export powers"
    
    currents=pd.read_csv('LVTest_EXP_CURRENTS.csv')
    powers=pd.read_csv('LVTest_EXP_POWERS.csv')
    voltages=pd.read_csv('LVTest_EXP_VOLTAGES.csv')

    VoltageMin=voltages[' Magnitude1'][-1:].values[0]
    
    return currents, powers, voltages, VoltageMin

def fault_seqz():
    dssText.Command="Solve Mode=Faultstudy"
    dssText.Command="export Faultstudy"
    dssText.Command="export seqz"
    
    seqz=pd.read_csv('LVTest_EXP_SEQZ.csv')
    faults=pd.read_csv('LVTest_EXP_FAULTS.csv')
    
    return seqz,faults
  
THD_C=create_spectrum(rated_cc)

def load_builder_balanced(n,i):
    Loads=pd.read_csv('Loads.txt', delimiter=' ', names=['New','Load','Phases','Bus1','kV','kW','PF','spectrum'])
    Loads['spectrum'][0]='spectrum='+str(i)
    Loads['kW'][0]='kW='+str(EV_power[i])
    b0=random.choice(range(2,B+1))
    Loads['Bus1'][0]='Bus1='+str()+str(b0)+'.1'
    for p in range(2,4):
        Loads=Loads.append(Loads.loc[0], ignore_index=True)
        Loads['Bus1'][p-1]='Bus1='+str(b0)+'.'+str(p)
        Loads['Load'][p-1]='Load.LOAD'+str(p)
    c=3
    for k in range(1,n):
        b=random.choice(range(2,B+1))
        for p in range(1,4):
            Loads=Loads.append(Loads.loc[0], ignore_index=True)
            Loads['Bus1'][c]='Bus1='+str(b)+'.'+str(p)
            Loads['Load'][c]='Load.LOAD'+str(c+1)
            Loads['kW'][c]='kW='+str(EV_power[i])
            c=c+1
        Loads['spectrum'][k]='spectrum='+str(i)
    
    Loads.to_csv('Loads_R.txt', sep=" ", index=False, header=False)
    
def load_builder_unbalanced(n,i):
    Loads=pd.read_csv('Loads.txt', delimiter=' ', names=['New','Load','Phases','Bus1','kV','kW','PF','spectrum'])
    i=random.choice(list(rated_cc.keys()))
    admd_factor=1
    if n>=20:
        n=int(n/2)
    Loads['spectrum'][0]='spectrum='+str(i)
    Loads['kW'][0]='kW='+str(EV_power[i]*admd_factor)
    b0=random.choice(range(2,B+1))
    Loads['Bus1'][0]='Bus1='+str(b0)+'.1'
    for k in range(1,n):
        i=random.choice(list(rated_cc.keys()))
        b=random.choice(range(2,B+1))
        p=random.choice(range(1,4))
        Loads=Loads.append(Loads.loc[0], ignore_index=True)
        Loads['Bus1'][k]='Bus1='+str(b)+'.'+str(p)
        Loads['Load'][k]='Load.LOAD'+str(k+1)
        Loads['kW'][k]='kW='+str(EV_power[i]*admd_factor)
        Loads['spectrum'][k]='spectrum='+str(i)
    
    Loads.to_csv('Loads_R.txt', sep=" ", index=False, header=False)

def harmonic_modeller(M,B,i):
    Ch_Ratios=pd.DataFrame()
    Vh_percent=pd.DataFrame()
    Pass=pd.DataFrame()
    THD_Pass=pd.DataFrame()
    All_Pass=pd.DataFrame(index=range(1,M+1))    
    All_THDs=pd.DataFrame(index=range(1,M+1))
    Full_H_pass={}
    failers=[]    
    for r in range(1,R):
        print('Run',r)
        All_Pass[r]=range(1,M+1)
        All_THDs[r]=range(1,M+1)
        for n in range(1,M+1):
            if r==1:
                Full_H_pass[n]=pd.DataFrame()
            ###--- Add loads
            load_builder_unbalanced(n,i)
            dssObj.ClearAll() 
            dssText.Command="Compile E:/PNDC/TN-006Harmonic/Harmonic_OpenDSS/Master_R.dss"
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
            Vh_percent['h']=res_Reactor[' Harmonic']
            Vh_percent[n]=Vh_ratios['V_ratio']
            Vh_percent[n][0] = sum(Vh_ratios['V_ratio'][1:]**2)**0.5
            Pass['h']=res_Reactor[' Harmonic']
            Pass[n]=Vh_percent[n]<Vh_ratios['Lims']
            Full_H_pass[n][r]=Pass[n]
            
        All_Pass[r]=Pass.iloc[:,1:].sum()==30
        All_THDs[r]=Vh_percent.iloc[0][1:]
        THD_Pass[r]=Pass.iloc[0][1:]
    for n in range(1,M+1):
        failers.append((Full_H_pass[n].sum(axis=1)).index[Full_H_pass[n].sum(axis=1)<r].values)
    failers=np.unique(np.concatenate(failers))
    
    for n in range(1,M+1):
        Full_H_pass[n]=Full_H_pass[n].loc[failers]
        

            
    Vh_percent['h'][0]='THD'
    return All_Pass,Vh_percent,Pass,THD_Pass,All_THDs,Full_H_pass,failers
    #print(i,All_Pass)

def balanced_run():
    AllPass={}
    AllVh_percent={}
    VoltageMin={}
    Pass={}
    V_Min_Av={}
    THD_Pass={}
    
    for i in list(rated_cc.keys()):
        AllPass[i],AllVh_percent[i],Pass[i],THD_Pass[i]=harmonic_modeller(M,B,i)

    for f in RSCs:
        lineBuilder(B,f_Rsc[f])
        VoltageMin[f]={}
        V_Min_Av[f]=pd.DataFrame(columns=rated_cc.keys())
        for i in list(rated_cc.keys()): 
            VoltageMin[f][i]=pd.DataFrame(index=range(1,M+1))
            for r in range(1,R):
                VoltageMin[f][i][r]=range(1,M+1)
                for n in range(1,M+1):
                    load_builder_unbalanced(n,i)
                    currents, powers, voltages, VoltageMin[f][i][r][n] =export_loadflow()
                
        V_Min_Av[f][i]=VoltageMin[f][i].mean(axis=1).values
        V_Min_Av[f].index=range(1,M+1)
                    
    # seqz,faults=fault_seqz()
    # print(faults)
    
#######---------Run Unbalanced -----------########
    
VoltageMin={}
V_Min_Av={}
i='dummy'
Unbalanced={}
Unbalanced_fullH={}
perH={}
failers={}
for f in RSCs:
    print(f)
    lineBuilder(B,f_Rsc[f])
    Unbalanced['AllPass_'+str(f)],Unbalanced['AllVh_percent_'+str(f)],Unbalanced['Pass_'+str(f)],Unbalanced['THD_Pass_'+str(f)],Unbalanced['All_THDs_'+str(f)],Unbalanced_fullH[f],failers[f]=harmonic_modeller(M,B,'dummy')
    
    perH[f]=pd.DataFrame(index=Unbalanced_fullH[list(Unbalanced_fullH.keys())[0]][1].index,columns=Unbalanced_fullH[list(Unbalanced_fullH.keys())[0]].keys())
    perH[f]['h']=Unbalanced['Pass_'+str(f)]['h'][Unbalanced_fullH[list(Unbalanced_fullH.keys())[0]][1].index]
    for j in Unbalanced_fullH[list(Unbalanced_fullH.keys())[0]].keys():
        perH[f][j]=1-Unbalanced_fullH[f][j].sum(axis=1)/Unbalanced_fullH[f][j].count(axis=1)

V_Min_Av=pd.DataFrame(columns=RSCs)
seqz={}
faults={}
for f in RSCs:
    lineBuilder(B,f_Rsc[f])
    VoltageMin[f]=pd.DataFrame(index=range(1,M+1))
    for r in range(1,R):
        print('Vmin Run ', r)
        VoltageMin[f][r]=range(1,M+1)
        for n in range(1,M+1):
            load_builder_unbalanced(n,i)
            currents, powers, voltages, VoltageMin[f][r][n] =export_loadflow()
            
    V_Min_Av[f]=VoltageMin[f].mean(axis=1).values
    V_Min_Av.index=range(1,M+1)
                
    seqz[f],faults[f]=fault_seqz()
    print(faults)
    

allfails=[]
for f in RSCs:
    allfails.append(failers[f])
allfails=np.unique(np.concatenate(allfails))
#i=random.choice(list(rated_cc.keys()))

pickle_out = open('results/Summary_'+case+'.pickle', "wb")
pickle.dump(Unbalanced, pickle_out)
pickle_out.close()

pickle_out = open('results/AllHarmonics_'+case+'.pickle', "wb")
pickle.dump(perH, pickle_out)
pickle_out.close()

pickle_out = open('results/Vmin_'+case+'.pickle', "wb")
pickle.dump(perH, pickle_out)
pickle_out.close()


styles=pd.Series(data=[':','-.','-','--'],index=RSCs)

fig, (ax1, ax2) = plt.subplots(2, sharex=False)
for f in RSCs:
    ax1.plot(100-(Unbalanced['AllPass_'+str(f)].sum(axis=1)/Unbalanced['AllPass_'+str(f)].count(axis=1))*100,label='RSC'+str(f),linestyle=styles[f])
    ax2.plot(Unbalanced['All_THDs_'+str(f)].index,Unbalanced['All_THDs_'+str(f)].max(axis=1), label='RSC'+str(f),linestyle=styles[f])
    ax1.xaxis.set_major_formatter(FormatStrFormatter('% 1.0f'))
    ax2.xaxis.set_major_formatter(FormatStrFormatter('% 1.0f'))

ax1.set_ylabel('% Probability of Failure')
ax1.legend()
ax1.grid(linewidth=0.2)
ax1.set_xlim(1,M)
ax1.set_ylim(0,100)
ax2.set_ylabel('Maximum THD')
ax2.legend()
ax2.set_xlabel('Number of EVs')
ax2.grid(linewidth=0.2)
ax2.set_xlim(1,M)
ax2.set_ylim(0,5)
print('Max prob of THD Failure', (100-((Unbalanced['THD_Pass_'+str(f)]==True).sum(axis=1)/Unbalanced['THD_Pass_'+str(f)].count(axis=1))*100).max())
plt.tight_layout()

plt.figure('Specific Harmonics',figsize=(5, 8))
c=1
for pl in allfails:
    ax=plt.subplot(len(allfails),1, c)
    ax.set_ylabel('% Failure')
    ax.text(.5,.8,'h='+str(int(perH[f]['h'][pl])),
        horizontalalignment='left',
        transform=ax.transAxes)
    for f in RSCs:
        if len(perH[f])>0:
            print(len(perH[f]))
            if c<len(perH[f]):
                ax.plot(perH[f].loc[pl][1:M]*100,linestyle=styles[f])
            if c==len(perH[f]):
               ax.plot(perH[f].loc[pl][1:M]*100, label='RSC='+str(f),linestyle=styles[f])
            c=c+1
    if c>len(allfails):
        ax.legend()
    ax.xaxis.set_major_formatter(FormatStrFormatter('% 1.0f'))
    plt.grid(linewidth=0.2)
    ax.set_xlim(1,M)
    ax.set_ylim(0,100)
    ax.xaxis.set_major_formatter(FormatStrFormatter('% 1.0f'))
    plt.tight_layout()


plt.figure()
for f in RSCs:
    plt.plot(V_Min_Av.index, V_Min_Av[f].values, linestyle=styles[f], label='Vmin with Rsc='+str(f), linewidth=1)
plt.ylabel('Voltage (V)')
plt.plot([1,M],[216,216],color='Black',linestyle=":", linewidth=0.5, label='Statutory Min')
plt.xlabel('Number of EVs')
plt.legend()

