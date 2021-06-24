# -*- coding: utf-8 -*-
"""
Created on Thu May 02 13:48:06 2019

@author: qsb15202
"""

import win32com.client
import scipy.io
import numpy as np
import pandas as pd
import csv
import random
import cmath
import math
import matplotlib.pyplot as plt
import pickle
from matplotlib.ticker import (MultipleLocator,
                               FormatStrFormatter,
                               AutoMinorLocator)
import seaborn as sns
##### Load in the Harmonic Profiles ########
#derated
#rated_cv
case='Urban_CC_100EVs_Diversity'
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
DSSLoads=DSSCircuit.Loads
DSSElement=DSSCircuit.ActiveElement

###### Source Impedence is adjusted for Urban/Rural networks

M=5 ##--Number of EVs
B=5  ##--Number of Buses
R=2 ##--Number of Runs

#RSCs=[33,66]   ##--- FOr Urban where WPD ZMax is much higher than corresponding RSC
RSCs=[33,66]  ###--- For Rural where WPD ZMax and RSC=15 are similar

#f_Rsc=pd.Series(dtype=float,index=RSCs,data=[1.6,0.62,0.21]) #for 185 mm - RURAL
#f_Rsc=pd.Series(dtype=float,index=RSCs,data=[0.77,1.7,0.72,0.305]) #for 185 mm - URBAN
f_Rsc=pd.Series(dtype=float,index=RSCs,data=[0.72,0.305]) #for 185 mm - URBAN
#f_Rsc=pd.Series(dtype=float,index=RSCs,data=[0.62,0.21]) #for 185 mm - RURAL
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
    dssText.Command="Compile Master_R.dss"
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
    n=int(round((0.36+0.64/(n**0.5))*n,0))
    Loads['spectrum'][0]='spectrum='+str(i)
    Loads['kW'][0]='kW='+str(EV_power[i])
    b0=random.choice(range(2,B+1))
    Loads['Bus1'][0]='Bus1='+str(b0)+'.1'
    for k in range(1,n):
        i=random.choice(list(rated_cc.keys()))
        b=random.choice(range(2,B+1))
        p=random.choice(range(1,4))
        Loads=Loads.append(Loads.loc[0], ignore_index=True)
        Loads['Bus1'][k]='Bus1='+str(b)+'.'+str(p)
        Loads['Load'][k]='Load.LOAD'+str(k+1)
        Loads['kW'][k]='kW='+str(EV_power[i])
        Loads['spectrum'][k]='spectrum='+str(i)
    
    Loads.to_csv('Loads_R.txt', sep=" ", index=False, header=False)

def harmonic_modeller(M,B,i):
    Ch_Ratios=pd.DataFrame()
    THD_Pass=np.zeros(shape=(R,M),dtype=bool)
    All_Pass=np.zeros(shape=(R,M),dtype=bool)
    All_THDs=np.zeros(shape=(R,M),dtype=float)
    VoltageMin=np.zeros(shape=(R,M),dtype=float)  
    Pass={} 
    for r in range(1,R+1):
        print('Run',r)
        for n in range(1,M+1):
            Vh_ratios=pd.DataFrame()
            if r==1:
                Pass[n]=np.zeros(shape=(30,R),dtype=bool)
            ###--- Add loads
            load_builder_unbalanced(n,i)
            
            dssObj.ClearAll() 
            dssText.Command="Compile Master_R.dss"
            dssText.Command="Solve"
            bvs = list(DSSCircuit.AllBusVMag)
            Voltages = bvs[0::3], bvs[1::3], bvs[2::3]
            VoltArray = np.zeros((len(Voltages[0]), 3))
            for i in range(0, 3):
                VoltArray[:, i] = np.array(Voltages[i], dtype=float)
            VoltageMin[(r-1),(n-1)]=VoltArray[-1:].mean()
             ###--- Solve Harmonics
            dssText.Command="Solve Mode=harmonics NeglectLoadY=Yes"
            dssText.Command="export monitors m1"
            res_Reactor=pd.read_csv('LVTest_Mon_m1_1.csv')
            Vh_ratios['h']=res_Reactor[' Harmonic']
            Vh_ratios['V_ratio1']=res_Reactor[' V1']/res_Reactor[' V1'][0]*100
            Vh_ratios['V_ratio2']=res_Reactor[' V2']/res_Reactor[' V2'][0]*100
            Vh_ratios['V_ratio3']=res_Reactor[' V3']/res_Reactor[' V3'][0]*100
            Vh_ratios['Lims']=g55lims['L']
            Ch_Ratios['h']=res_Reactor[' Harmonic']
            Ch_Ratios[n]=res_Reactor[' I1']/res_Reactor[' I1'][0]*100
            Vh_ratios['V_ratio1'][0] = sum(Vh_ratios['V_ratio1'][1:]**2)**0.5
            Vh_ratios['V_ratio2'][0] = sum(Vh_ratios['V_ratio2'][1:]**2)**0.5
            Vh_ratios['V_ratio3'][0] = sum(Vh_ratios['V_ratio3'][1:]**2)**0.5
            p1=Vh_ratios['V_ratio1']<Vh_ratios['Lims']
            p2=Vh_ratios['V_ratio2']<Vh_ratios['Lims']
            p3=Vh_ratios['V_ratio3']<Vh_ratios['Lims']
            for l in p1.index:
                Pass[n][l,(r-1)]=p1[l] and p2[l] and p3[l] 
            
            All_THDs[(r-1),(n-1)]=max(Vh_ratios['V_ratio1'].iloc[0],Vh_ratios['V_ratio2'].iloc[0],Vh_ratios['V_ratio3'].iloc[0])
    for n in range(1,M+1):
        All_Pass[:,n-1]=Pass[n].sum(axis=0)==30
        THD_Pass[:,n-1]=Pass[n][0,:]
    harmonic_index=res_Reactor[' Harmonic'].astype(int)
    return All_Pass,Pass,THD_Pass,All_THDs,harmonic_index,VoltageMin
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
allfails=[]
AllH={}
seqz={}
faults={}
for f in RSCs:
    print(f)
    lineBuilder(B,f_Rsc[f])
    Unbalanced['AllPass_'+str(f)],Unbalanced['Pass_'+str(f)],Unbalanced['THD_Pass_'+str(f)],Unbalanced['All_THDs_'+str(f)],h_index,VoltageMin[f]=harmonic_modeller(M,B,'dummy')
    AllH[f]=pd.DataFrame(index=h_index, columns=range(1,M+1),dtype=float) 
    for n in range(1,M+1):
        AllH[f].loc[:,n]=Unbalanced['Pass_'+str(f)][n].sum(axis=1)
    allfails=allfails + list(AllH[f][AllH[f].sum(axis=1)/(M*R)<1].index)
    seqz[f],faults[f]=fault_seqz()
    print(faults[f])
    V_Min_Av[f]=VoltageMin[f].mean(axis=0)


allfails=np.unique(allfails)

for f in RSCs:
    perH[f]=AllH[f].loc[allfails]

#i=random.choice(list(rated_cc.keys()))

pickle_out = open('results/Summary_'+case+'.pickle', "wb")
pickle.dump(Unbalanced, pickle_out)
pickle_out.close()

pickle_out = open('results/AllHarmonics_'+case+'.pickle', "wb")
pickle.dump(perH, pickle_out)
pickle_out.close()

pickle_out = open('results/Vmin_'+case+'.pickle', "wb")
pickle.dump(V_Min_Av, pickle_out)
pickle_out.close()


pick_in = open('results/Summary_'+case+'.pickle', "rb")
Unbalanced = pickle.load(pick_in)

pick_in = open('results/AllHarmonics_'+case+'.pickle', "rb")
perH = pickle.load(pick_in)

pick_in = open('results/Vmin_'+case+'.pickle', "rb")
V_Min_Av = pickle.load(pick_in)

#styles=pd.Series(data=[':','-.','-','--'],index=RSCs)
styles=pd.Series(data=[':','-'],index=RSCs)
cols=pd.Series(data=['tab:green','tab:red'],index=RSCs)
fig, (ax1, ax2) = plt.subplots(2, sharex=False)
for f in RSCs:
    ax1.plot(100-(Unbalanced['AllPass_'+str(f)].sum(axis=0)/R*100),label='RSC'+str(f),linestyle=styles[f], color=cols[f])
    ax2.plot(Unbalanced['All_THDs_'+str(f)].max(axis=0), label='RSC'+str(f),linestyle=styles[f], color=cols[f])
    ax1.xaxis.set_major_formatter(FormatStrFormatter('% 1.0f'))
    ax2.xaxis.set_major_formatter(FormatStrFormatter('% 1.0f'))

ax1.set_ylabel('% Failure')
ax1.legend()
ax1.grid(linewidth=0.2)
ax1.set_xlim(1,M-1)
ax1.set_xticks(range(0,M))
ax1.set_xticklabels(range(1,(M+1)))
ax1.set_ylim(0,100)
ax2.set_ylabel('Maximum THD')
ax2.legend()
ax2.set_xlabel('Number of EVs')
ax2.grid(linewidth=0.2)
ax2.set_xlim(1,M-1)
ax2.set_ylim(0,5)
ax2.set_xticks(range(0,M))
ax2.set_xticklabels(range(1,(M+1)))
print('Max prob of THD Failure', (100-((Unbalanced['THD_Pass_'+str(f)]==True).sum(axis=0)/R*100).max()))
plt.tight_layout()


if len(allfails)>0:
    plt.figure('Specific Harmonics 1',figsize=(5, 8))
    c=1
    for pl in allfails[:4]:
        ax=plt.subplot(len(allfails[:4]),1, c)
        ax.set_ylabel('% Failure')
        ax.text(.5,.8,'h='+str(pl),
            horizontalalignment='left',
            transform=ax.transAxes)
        for f in RSCs:
            ax.plot(100-perH[f].loc[pl][1:M]/R*100,linestyle=styles[f], color=cols[f])
            ax.plot(100-perH[f].loc[pl][1:M]/R*100, label='RSC='+str(f),linestyle=styles[f], color=cols[f])
        ax.legend()
        ax.xaxis.set_major_formatter(FormatStrFormatter('% 1.0f'))
        plt.grid(linewidth=0.2)
        ax.set_xlim(0,M-1)
        ax.set_xticks(range(0,M))
        ax.set_xticklabels(range(1,(M+1)))
        ax.set_ylim(0,100)
        ax.xaxis.set_major_formatter(FormatStrFormatter('% 1.0f'))
        plt.tight_layout()
        c=c+1

if len(allfails)>5:
    plt.figure('Specific Harmonics 2',figsize=(5, 8))
    c=1
    for pl in allfails[5:]:
        ax=plt.subplot(len(allfails[5:]),1, c)
        ax.set_ylabel('% Failure')
        ax.text(.5,.8,'h='+str(pl),
            horizontalalignment='left',
            transform=ax.transAxes)
        for f in RSCs:
            ax.plot(100-perH[f].loc[pl][1:M]/R*100,linestyle=styles[f], color=cols[f])
            ax.plot(100-perH[f].loc[pl][1:M]/R*100, label='RSC='+str(f),linestyle=styles[f], color=cols[f])
        ax.legend()
        ax.xaxis.set_major_formatter(FormatStrFormatter('% 1.0f'))
        plt.grid(linewidth=0.2)
        ax.set_xlim(0,M-1)
        ax.set_xticks(range(0,(M)))
        ax.set_xticklabels(range(1,(M+1)))
        ax.set_ylim(0,100)
        ax.xaxis.set_major_formatter(FormatStrFormatter('% 1.0f'))
        plt.tight_layout()
        c=c+1 

plt.figure()
for f in RSCs:
    plt.plot(V_Min_Av[f], linestyle=styles[f], label='Vmin with Rsc='+str(f), linewidth=1, color=cols[f])
plt.ylabel('Voltage (V)')
plt.plot([0,M],[216,216],color='Black',linestyle=":", linewidth=0.5, label='Statutory Min')
plt.xticks(ticks=range(0,(M)),labels=range(1,(M+1)))
plt.xlabel('Number of EVs')
plt.xlim(0,M-1)
plt.legend()
