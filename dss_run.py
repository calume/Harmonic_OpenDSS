# -*- coding: utf-8 -*-
"""
Created on Wed Jun 23 04:21:39 2021

@author: CalumEdmunds
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
import os

dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")
dssText = dssObj.Text
DSSCircuit = dssObj.ActiveCircuit
DSSLoads=DSSCircuit.Loads
DSSLines=DSSCircuit.Lines
DSSElement=DSSCircuit.ActiveElement
DSSTrans=DSSCircuit.Transformers
dssObj.Start(0)
dssObj.AllowForms=False
##### Load in the Harmonic Profiles ########

cm='CC'
net_type='urban'


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
 
###### Source Impedence is adjusted for Urban/Rural networks

M=5 ##--Number of EVs
B=5 ##--Number of Buses
R=2 ##--Number of Runs

case=str(cm)+'_Unbalanced_'+str(net_type)+str(M)+'EVs_'+str(R)+'Runs_'

##RSCs=['WPD_Zmax',15,33,66]   ##--- FOr Urban where WPD ZMax is much higher than corresponding RSC

RSCs=[33,66]  ###--- For Rural where WPD ZMax and RSC=15 are similar
if net_type=='urban':
    f_Rsc=pd.Series(dtype=float,index=RSCs,data=[0.72,0.305])   ##--- FOr Urban where WPD ZMax is much higher than corresponding RSC

if net_type=='rural':
    f_Rsc=pd.Series(dtype=float,index=RSCs,data=[0.62,0.21])    

##f_Rsc=pd.Series(dtype=float,index=RSCs,data=[1.6,0.62,0.21]) #for 185 mm - RURAL
##f_Rsc=pd.Series(dtype=float,index=RSCs,data=[0.77,1.7,0.72,0.305]) #for 185 mm - URBAN
 #for 185 mm - URBAN

####------Build Lines between B buses

i=random.choice(list(rated_cc.keys()))
Ch_Ratios={}
Pass={}
THD_Pass={}
All_Pass={}    
All_THDs={}
Full_H_pass={}
failers={}
perH={}
Summary={}
VoltageMin={}
V_Min_Av=pd.DataFrame(columns=RSCs)
seqz={}
faults={}
AllH={}
allfails=[]
for f in RSCs:
    print(f)
    Ch_Ratios=pd.DataFrame()
    THD_Pass[f]=np.zeros(shape=(R,M),dtype=bool)
    All_Pass[f]=np.zeros(shape=(R,M),dtype=bool)
    All_THDs[f]=np.zeros(shape=(R,M),dtype=float)
    VoltageMin[f]=np.zeros(shape=(R,M),dtype=float)  
    Pass[f]={}
    for r in range(1,R+1):
        print('Run',r)
        for n in range(1,M+1):
            if r==1:
                Pass[f][n]=np.zeros(shape=(30,R),dtype=bool)
            Vh_ratios=pd.DataFrame()
            dssObj.ClearAll() 
            dssText.Command="redirect Master_S.dss"
            ##------URBAN SETUP ----- For Min Source Impedence @ Secondary Tx of 0.022 +0.024j (WPD EV Emmissions report Table 2-7)
            if net_type=='urban':
                dssText.Command ="Edit Vsource.Source BasekV=11 Phases=3 pu=1.00 ISC3=3000 ISC1=2500 "
                dssText.Command ="Edit Transformer.TR1 Buses=[SourceBus 1] Conns=[Delta Wye] kVs=[11 0.415] kVAs=[500 500] XHL=6.15 ppm=0 tap=1.000"
                DSSTrans.First
                DSSTrans.Wdg=1
                DSSTrans.R=3.1
                DSSTrans.Wdg=2
                DSSTrans.R=3.1
            #--- Add Lines
            for L in range(1,B):
                dssText.Command ="New Line.LINE"+str(L)+" Bus1="+str(L+1)+" Bus2="+str(L+2)+" phases=3 Linecode=D2 Length="+str(f_Rsc[f]/B)+" Units=km"
            n_d=1
            ##--- Add loads
            if cm=='CC' or cm=='CV':
                n_d=int(round((0.36+0.64/(n**0.5))*n,0))  ###--- Diversity factor for EV loads
            
            for k in range(1,n_d+1):
                i=random.choice(list(rated_cc.keys()))
                b=random.choice(range(2,B+1))
                p=random.choice(range(1,4))  
                dssText.Command = "New Load.LOAD"+str(k)+" Phases=1 Status=1 Bus1="+str(b)+"."+str(p)+" kV=0.230 kW="+str(EV_power[i])+" PF=1 spectrum="+str(i)
            
            dssText.Command="New monitor.M"+str(n)+" Reactor.R1 Terminal=2"
            ###--- Solve Load Flow (and record Vmin)
            dssText.Command="Solve"
            bvs = list(DSSCircuit.AllBusVMag)
            Voltages = bvs[0::3], bvs[1::3], bvs[2::3]
            VoltArray = np.zeros((len(Voltages[0]), 3))
            iLoad=DSSLoads.First
            for i in range(0, 3):
                VoltArray[:, i] = np.array(Voltages[i], dtype=float)
            VoltageMin[f][(r-1),(n-1)]=VoltArray[-1:].mean() # Minimum voltage is at the end of the feeder (Average of 3 phases)

            ###--- Solve Harmonics
            dssText.Command="Solve Mode=harmonics NeglectLoadY=Yes"
            dssText.Command="export monitors m"+str(n)
            res_Reactor=pd.read_csv('LVTest_Mon_m'+str(n)+'_1.csv')
            os.remove('LVTest_Mon_m'+str(n)+'_1.csv')
            Vh_ratios['h']=res_Reactor[' Harmonic']
            Vh_ratios['V_ratio1']=res_Reactor[' V1']/res_Reactor[' V1'][0]*100   ### Convert from V to % of Fundamental (phase A,B,C)
            Vh_ratios['V_ratio2']=res_Reactor[' V2']/res_Reactor[' V2'][0]*100  
            Vh_ratios['V_ratio3']=res_Reactor[' V3']/res_Reactor[' V3'][0]*100   
            Vh_ratios['Lims']=g55lims['L']
            Ch_Ratios['h']=res_Reactor[' Harmonic']
            Ch_Ratios[n]=res_Reactor[' I1']/res_Reactor[' I1'][0]*100
            Vh_ratios['V_ratio1'][0] = sum(Vh_ratios['V_ratio1'][1:]**2)**0.5    ### Calculate THD (Phase A,B,C)
            Vh_ratios['V_ratio2'][0] = sum(Vh_ratios['V_ratio2'][1:]**2)**0.5
            Vh_ratios['V_ratio3'][0] = sum(Vh_ratios['V_ratio3'][1:]**2)**0.5
            p1=Vh_ratios['V_ratio1']<Vh_ratios['Lims'] ### Pass Fail against G5/5 Limits (Phase A,B,C)
            p2=Vh_ratios['V_ratio2']<Vh_ratios['Lims']
            p3=Vh_ratios['V_ratio3']<Vh_ratios['Lims']
            for l in p1.index:
                Pass[f][n][l,(r-1)]=p1[l] and p2[l] and p3[l] ### All 3 Phases must Pass

            All_THDs[f][(r-1),(n-1)]=max(Vh_ratios['V_ratio1'].iloc[0],Vh_ratios['V_ratio2'].iloc[0],Vh_ratios['V_ratio3'].iloc[0]) ### Max THD from all 3 Phases
    
    for n in range(1,M+1): ### Record if all harmonics Pass and if only THD Passes
        All_Pass[f][:,n-1]=Pass[f][n].sum(axis=0)==30
        THD_Pass[f][:,n-1]=Pass[f][n][0,:]
   
    h_index=res_Reactor[' Harmonic'].astype(int)
        
    dssText.Command="Solve Mode=Faultstudy"   ### Run a Fault Study to return Fault Currents and System Impedances
    dssText.Command="export Faultstudy"
    dssText.Command="export seqz"
    
    seqz[f]=pd.read_csv('LVTest_EXP_SEQZ.csv')
    faults[f]=pd.read_csv('LVTest_EXP_FAULTS.csv')
    print(faults[f])
    
    AllH[f]=pd.DataFrame(index=h_index, columns=range(1,M+1),dtype=float) 
    for n in range(1,M+1):
        AllH[f].loc[:,n]=Pass[f][n].sum(axis=1)
    allfails=allfails + list(AllH[f][AllH[f].sum(axis=1)/(M*R)<1].index)
    
    V_Min_Av[f]=VoltageMin[f].mean(axis=0)  ###

allfails=np.unique(allfails)
Summary['All_Pass']={}
Summary['Pass']={}
Summary['THD_Pass']={}
Summary['All_THDs']={}
for f in RSCs:
    Summary['All_Pass'][f]=All_Pass[f]
    Summary['Pass'][f]=Pass[f]
    Summary['THD_Pass'][f]=THD_Pass[f]
    Summary['All_THDs'][f]=All_THDs[f]
    perH[f]=AllH[f].loc[allfails]

pickle_out = open('results/Summary_'+case+'.pickle', "wb")
pickle.dump(Summary, pickle_out)
pickle_out.close()

pickle_out = open('results/AllHarmonics_'+case+'.pickle', "wb")
pickle.dump(perH, pickle_out)
pickle_out.close()

pickle_out = open('results/Vmin_'+case+'.pickle', "wb")
pickle.dump(V_Min_Av, pickle_out)
pickle_out.close()


# # pick_in = open('results/Summary_'+case+'.pickle', "rb")
# # Summary = pickle.load(pick_in)

# # pick_in = open('results/AllHarmonics_'+case+'.pickle', "rb")
# # perH = pickle.load(pick_in)

# # pick_in = open('results/Vmin_'+case+'.pickle', "rb")
# # V_Min_Av = pickle.load(pick_in)

# styles=pd.Series(data=[':','-.','-','--'],index=RSCs)


#styles=pd.Series(data=[':','-.','-','--'],index=RSCs)
styles=pd.Series(data=[':','-'],index=RSCs)
cols=pd.Series(data=['tab:green','tab:red'],index=RSCs)
fig, (ax1, ax2) = plt.subplots(2, sharex=False)
for f in RSCs:
    ax1.plot(100-(Summary['All_Pass'][f].sum(axis=0)/R*100),label='RSC'+str(f),linestyle=styles[f], color=cols[f])
    ax2.plot(Summary['All_THDs'][f].max(axis=0), label='RSC'+str(f),linestyle=styles[f], color=cols[f])
    ax1.xaxis.set_major_formatter(FormatStrFormatter('% 1.0f'))
    ax2.xaxis.set_major_formatter(FormatStrFormatter('% 1.0f'))
    print('Max prob of THD Failure RSC='+str(f), (100-((Summary['THD_Pass'][f]==True).sum(axis=0)/R*100).max()))

ax1.set_ylabel('% Failure')
ax1.legend()
ax1.grid(linewidth=0.2)
ax1.set_xlim(1,M-1)
ax1.set_xticks(range(0,M,10))
#ax1.set_xticklabels(range(1,(M+1),10))
ax1.set_ylim(0,100)
ax2.set_ylabel('Maximum THD')
ax2.legend()
ax2.set_xlabel('Number of EVs')
ax2.grid(linewidth=0.2)
ax2.set_xlim(1,M-1)
ax2.set_ylim(0,5)
ax2.set_xticks(range(0,M,10))
#ax2.set_xticklabels(range(1,(M+1),10))
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
            ax.plot(100-perH[f].loc[pl]/R*100,linestyle=styles[f], color=cols[f])
            ax.plot(100-perH[f].loc[pl]/R*100, label='RSC='+str(f),linestyle=styles[f], color=cols[f])
        ax.legend()
        ax.xaxis.set_major_formatter(FormatStrFormatter('% 1.0f'))
        plt.grid(linewidth=0.2)
        ax.set_xlim(1,M)
        ax.set_xticks(range(0,M,10))
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
        ax.set_xlim(0,M)
        ax.set_xticks(range(0,M,10))
        #ax.set_xticklabels(range(1,(M+1),10))
        ax.set_ylim(0,100)
        ax.xaxis.set_major_formatter(FormatStrFormatter('% 1.0f'))
        plt.tight_layout()
        c=c+1 

plt.figure()
for f in RSCs:
    plt.plot(V_Min_Av[f], linestyle=styles[f], label='Vmin with Rsc='+str(f), linewidth=1, color=cols[f])
plt.ylabel('Voltage (V)')
plt.plot([0,M],[216,216],color='Black',linestyle=":", linewidth=0.5, label='Statutory Min')
plt.xticks(ticks=range(0,M,10))
plt.xlabel('Number of EVs')
plt.xlim(0,M-1)
plt.grid(linewidth=0.2)
plt.legend()
