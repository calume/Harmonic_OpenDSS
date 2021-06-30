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
cf=2
cm='CC'
net_type=['rural','urban']

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

n_evs=50
n_buses=50
R=500
  ##--Number of Runs

##f_Rsc=pd.Series(dtype=float,index=RSCs,data=[1.6,0.62,0.21]) #for 185 mm - RURAL
##f_Rsc=pd.Series(dtype=float,index=RSCs,data=[0.77,1.7,0.72,0.305]) #for 185 mm - URBAN
  #for 185 mm - URBAN

####------Build Lines between B buses
RSCs=[33,66]
Summary={}
perH={}
V_Min_Av={}
allfails=np.array([])
AllH={}
for net in net_type:
    M=n_evs ##--Number of EVs
    B=n_buses ##--Number of Buses
    case=str(cm)+'_Unbalanced_'+str(M)+'EVs_'+str(R)+'Runs_'
    i=random.choice(list(rated_cc.keys()))
    Ch_Ratios={}
    Pass={}
    THD_Pass={}
    All_Pass={}    
    All_THDs={}
    Full_H_pass={}
    failers={}
    perH[net]={}
    Summary[net]={}
    VoltageMin={}
    V_Min_Av[net]={}
    seqz={}
    faults={}
    AllH[net]={}
    ##RSCs=['WPD_Zmax',15,33,66]   ##--- FOr Urban where WPD ZMax is much higher than corresponding RSC
    
      ###--- For Rural where WPD ZMax and RSC=15 are similar
    for f in RSCs:
        if f==66:
            M=M*cf
            B=B*cf
        if net=='urban':
            f_Rsc=pd.Series(dtype=float,index=RSCs,data=[0.78,0.327])   ##--- FOr Urban where WPD ZMax is much higher than corresponding RSC
        
        if net=='rural':
            f_Rsc=pd.Series(dtype=float,index=RSCs,data=[0.66,0.209])   
        print(net,f)
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
                dssText.Command="Compile Master_S.dss"
                ##------URBAN SETUP ----- For Min Source Impedence @ Secondary Tx of 0.022 +0.024j (WPD EV Emmissions report Table 2-7)
                if net=='urban':
                    dssText.Command ="Edit Transformer.TR1 Buses=[SourceBus 1] Conns=[Delta Wye] kVs=[11 0.415] kVAs=[500 500] XHL=0.01 ppm=0 tap=1.000"
                    dssText.Command ="Edit Reactor.R1 Bus1=1 Bus2=2 R=0.0212 X=0.0217 Phases=3 LCurve=L_Freq RCurve=R_Freq"
                    DSSTrans.First
                    DSSTrans.Wdg=1
                    DSSTrans.R=0.01
                    DSSTrans.Wdg=2
                    DSSTrans.R=0.01
                #--- Add Lines
                for L in range(1,B):
                    dssText.Command ="New Line.LINE"+str(L)+" Bus1="+str(L+1)+" Bus2="+str(L+2)+" phases=3 Linecode=D2 Length="+str(f_Rsc[f]/B)+" Units=km"
                n_d=n
                ##--- Add loads
                cv_factor=1
                if cm=='CC' or cm=='CV':
                    n_d=int(round((0.36+0.64/(n**0.5))*n,0))  ###--- Diversity factor for EV loads
                if cm=='CV':
                    cv_factor=0.5
                rb=[]
                for k in range(1,n_d+1):
                    i=random.choice(list(rated_cc.keys()))
                    b=random.choice(range(2,B+1))
                    rb.append(b-1)
                    p=random.choice(range(1,4))  
                    dssText.Command = "New Load.LOAD"+str(k)+" Phases=1 Status=1 Bus1="+str(b)+"."+str(p)+" kV=0.230 kW="+str(EV_power[i]*cv_factor)+" PF=1 spectrum="+str(i)
                
                dssText.Command="New monitor.M"+str(n)+" Line.LINE"+str(max(rb))+" Terminal=2"
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
                #os.remove('LVTest_Mon_m'+str(n)+'_1.csv')
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
        print(net,f,faults[f]['  1-Phase'][-1:].values)
        
        AllH[net][f]=pd.DataFrame(index=h_index, columns=range(1,M+1),dtype=float) 
        for n in range(1,M+1):
            AllH[net][f].loc[:,n]=Pass[f][n].sum(axis=1)
        allfails=np.append(allfails,list(AllH[net][f][AllH[net][f].sum(axis=1)/(M*R)<1].index))
        
        V_Min_Av[net][f]=VoltageMin[f].mean(axis=0)  ###
        
    allfails=np.unique(allfails)
    Summary[net]['All_Pass']={}
    Summary[net]['Pass']={}
    Summary[net]['THD_Pass']={}
    Summary[net]['All_THDs']={}
    for f in RSCs:
        Summary[net]['All_Pass'][f]=All_Pass[f]
        Summary[net]['Pass'][f]=Pass[f]
        Summary[net]['THD_Pass'][f]=THD_Pass[f]
        Summary[net]['All_THDs'][f]=All_THDs[f]
        
for net in net_type:
    for f in RSCs:
        perH[net][f]=AllH[net][f].loc[allfails]
    
pickle_out = open('results/Summary_'+case+'.pickle', "wb")
pickle.dump(Summary, pickle_out)
pickle_out.close()

pickle_out = open('results/AllHarmonics_'+case+'.pickle', "wb")
pickle.dump(perH, pickle_out)
pickle_out.close()

pickle_out = open('results/Vmin_'+case+'.pickle', "wb")
pickle.dump(V_Min_Av, pickle_out)
pickle_out.close()

# cf=1
# cm='CV'
# net_type=['rural','urban']
# n_evs=100
# n_buses=50
# R=500
# RSCs=[33,66]
# M=n_evs ##--Number of EVs
# B=n_buses ##--Number of Buses
# case=str(cm)+'_Unbalanced_'+str(M)+'EVs_'+str(R)+'Runs_'
   
# pick_in = open('results/Summary_'+case+'.pickle', "rb")
# Summary = pickle.load(pick_in)

# pick_in = open('results/AllHarmonics_'+case+'.pickle', "rb")
# perH = pickle.load(pick_in)

# pick_in = open('results/Vmin_'+case+'.pickle', "rb")
# V_Min_Av = pickle.load(pick_in)

styles=pd.Series(data=[':','-'],index=net_type)
cols=pd.Series(data=['tab:green','tab:red'],index=net_type)

for f in RSCs:
    M=n_evs ##--Number of EVs
    if f==66:
        M=M*cf
    fig, (ax1, ax2) = plt.subplots(2, sharex=False)
    ax1.set_title('RSC='+str(f))
    
    for net in net_type:
        ax1.plot(100-(Summary[net]['All_Pass'][f].sum(axis=0)/R*100),label=net,linestyle=styles[net], color=cols[net])
        ax2.plot(Summary[net]['All_THDs'][f].max(axis=0), label=net,linestyle=styles[net], color=cols[net])
    ax1.xaxis.set_major_formatter(FormatStrFormatter('% 1.0f'))
    ax2.xaxis.set_major_formatter(FormatStrFormatter('% 1.0f'))
    print('Max prob of THD Failure RSC='+str(f), (100-((Summary[net]['THD_Pass'][f]==True).sum(axis=0)/R*100).max()))
    
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

    
    if len(perH[net][f])>0:
        plt.figure('RSC='+str(f)+' Specific Harmonics 1',figsize=(5, 8))
        c=1
        for pl in perH[net][f].index[:4]:
            ax=plt.subplot(len(perH[net][f].index[:4]),1, c)
            ax.set_ylabel('% Failure')
            ax.text(.5,.8,'h='+str(pl),
                horizontalalignment='left',
                transform=ax.transAxes)
            for net in net_type:
                ax.plot(100-perH[net][f].loc[int(pl)][1:M]/R*100, label=net,linestyle=styles[net], color=cols[net])
            ax.legend()
            ax.xaxis.set_major_formatter(FormatStrFormatter('% 1.0f'))
            plt.grid(linewidth=0.2)
            ax.set_xlim(1,M)
            ax.set_xticks(range(0,M,10))
            ax.set_ylim(0,100)
            ax.xaxis.set_major_formatter(FormatStrFormatter('% 1.0f'))
            plt.tight_layout()
            c=c+1
    
    if len(perH[net][f])>5:
        plt.figure('RSC='+str(f)+' Specific Harmonics 2',figsize=(5, 8))
        c=1
        for pl in perH[net][f].index[5:]:
            ax=plt.subplot(len(perH[net][f].index[5:]),1, c)
            ax.set_ylabel('% Failure')
            ax.text(.5,.8,'h='+str(pl),
                horizontalalignment='left',
                transform=ax.transAxes)
            for net in net_type:
                ax.plot(100-perH[net][f].loc[int(pl)][1:M]/R*100, label=net,linestyle=styles[net], color=cols[net])
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
    
for f in RSCs:
    M=n_evs ##--Number of EVs
    if f==66:
        M=M*cf
    plt.figure('RSC='+str(f))
    for net in net_type:
        plt.plot(V_Min_Av[net][f], linestyle=styles[net], label='Vmin '+net, linewidth=1, color=cols[net])
    plt.ylabel('Voltage (V)')
    plt.plot([0,M],[216,216],color='Black',linestyle=":", linewidth=0.5, label='Statutory Min')
    plt.xticks(ticks=range(0,M,10))
    plt.xlabel('Number of EVs')
    plt.xlim(0,M-1)
    plt.grid(linewidth=0.2)
    plt.legend()


