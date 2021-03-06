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

study='compensate_derated'
study='remove_Zoe_BMW'
cms=['de']#,'CC','de']
counts={}
for cm in cms:   
    g55lims=pd.read_csv('g55limits.csv')
    counts[cm]={}
    if cm=='CV' or cm=='CC':
        rated_cc=pd.read_excel('rated_'+str(cm)+'_stats.xlsx', sheet_name=None)
        del rated_cc['BMW_3ph_'+str(cm)]
        del rated_cc['Zoe_3ph_'+str(cm)]
        del rated_cc['DC_'+str(cm)]
        EV_power=pd.Series(index=rated_cc.keys(),data=[6.6,6.6,7.2,7.2,7.2])
        if study=='remove_Zoe_BMW':
            del rated_cc['BMW_1ph_'+str(cm)]
            del rated_cc['Zoe_1ph_'+str(cm)]
            print(rated_cc.keys())
            EV_power=EV_power[list(rated_cc.keys())]
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
        EV_power=pd.Series(index=rated_cc.keys(),data=[1.32,2.64,3.96,5.28,1.32,2.64,3.96,5.28,1.35,2.7,4.05,5.4,1.35,2.7,4.05,5.4,1.35,2.7,4.05,5.4])
        if study=='compensate_derated':
            for kk in list(rated_cc.keys()):
                if kk[-2:]=='6A' or kk[-3:]=='12A' or kk[-3:]=='24A':
                    del rated_cc[kk]
            EV_power=EV_power[list(rated_cc.keys())]
            
        if study=='remove_Zoe_BMW':
            for kk in list(rated_cc.keys()):
                if kk[:3]=='zoe' or kk[:3]=='bmw':
                    del rated_cc[kk]
            EV_power=EV_power[list(rated_cc.keys())]       
    ###### Source Impedence is adjusted for Urban/Rural networks

    n_evs=100
    n_buses=100
    R=500
      ##--Number of Runs
    
    ####------Build Lines between B buses
    RSCs=[15,33,66]
    Summary={}
    perH={}
    V_Min_Av={}
    allfails=np.array([])
    AllH={}
    seqz={}
    case=str(cm)+'_Unbalanced_'+str(n_evs)+'EVs_'+str(R)+'Runs_'
    if study=='compensate_derated':
       case=str(cm)+'_Unbalanced_'+str(n_evs)+'EVs_'+str(R)+'Runs_Compensate_Derated' 
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
    V_Min_Av={}
    seqz={}
    faults={}
    AllH={}

    for f in RSCs:
        M=n_evs ##--Number of EVs
        B=n_buses ##--Number of Buses
        counts[cm][f]={}
        if f==15:
            M=int(M/2)
            B=int(B/2)
        f_Rsc=pd.Series(dtype=float,index=RSCs,data=[1.87,0.78,0.327])

        print(f)
        Ch_Ratios=pd.DataFrame()
        THD_Pass[f]=np.zeros(shape=(R,M),dtype=bool)
        All_Pass[f]=np.zeros(shape=(R,M),dtype=bool)
        All_THDs[f]=np.zeros(shape=(R,M),dtype=float)
        VoltageMin[f]=np.zeros(shape=(R,M),dtype=float)  
        Pass[f]={}
        for r in range(1,R+1):
            counts[cm][f][r]=np.zeros(shape=(M,3),dtype=int)
            print('Run',r)
            for n in range(1,M+1):
                if r==1:
                    Pass[f][n]=np.zeros(shape=(30,R),dtype=bool)
                Vh_ratios=pd.DataFrame()
                dssObj.ClearAll() 
                dssText.Command="Compile Master_S.dss"

                #--- Add Lines
                for L in range(1,B):
                    dssText.Command ="New Line.LINE"+str(L)+" Bus1="+str(L+1)+" Bus2="+str(L+2)+" phases=3 Linecode=D2 Length="+str(f_Rsc[f]/B)+" Units=km"
                n_d=n
                ##--- Add loads
                cv_factor=1
                if cm=='CC' or cm=='CV':
                    n_d=int(round((0.36+0.64/(n**0.5))*n,0))  ###--- Diversity factor for EV loads
                if cm=='de' and study=='compensate_derated':
                    n_d=int(round(((0.36+0.64/(n**0.5))*n)/0.58,0))
                rb=[]
                load_model=1
                if cm=='CC' or 'de':
                    load_model=5
                phases=[1,2,3]
                for k in range(1,n_d+1):
                    i=random.choice(list(rated_cc.keys()))
                    if cm=='CV' and (i=='Leaf_CV' or i=='Van_CV'):
                        cv_factor=0.5   
                    b=random.choice(range(2,B+1))
                    rb.append(b-1)
                    p=random.choice(phases)
                    counts[cm][f][r][n-1,p-1]=counts[cm][f][r][n-1,p-1]+1
                    if counts[cm][f][r][n-1,p-1]>((n_d/3)*1.1) and n_d>3:
                        phases.remove(p)
                        p=random.choice(phases) 
                    dssText.Command = "New Load.LOAD"+str(k)+" Model="+str(load_model)+" Phases=1 Status=1 Bus1="+str(b)+"."+str(p)+" kV=0.230 kW="+str(EV_power[i]*cv_factor)+" PF=1 spectrum="+str(i)
                dssText.Command="New monitor.M"+str(n)+" Line.LINE"+str(max(rb))+" Terminal=2"
                
                ###--- Solve Load Flow (and record Vmin)
                dssText.Command="Calcvoltagebases"
                dssText.Command="Solve"
                dssText.Command="Export Currents"
                dssText.Command="Export Voltages"
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
        print(f,faults[f]['  1-Phase'][-1:].values)
        
        AllH[f]=pd.DataFrame(index=h_index, columns=range(1,M+1),dtype=float) 
        for n in range(1,M+1):
            AllH[f].loc[:,n]=Pass[f][n].sum(axis=1)
        allfails=np.append(allfails,list(AllH[f][AllH[f].sum(axis=1)/(M*R)<1].index))
        
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
        
    for f in RSCs:
            perH[f]=AllH[f].loc[allfails]
        
    pickle_out = open('results/Summary_'+case+'.pickle', "wb")
    pickle.dump(Summary, pickle_out)
    pickle_out.close()
    
    pickle_out = open('results/PhaseCounts_'+case+'.pickle', "wb")
    pickle.dump(counts, pickle_out)
    pickle_out.close()
    
    pickle_out = open('results/AllHarmonics_'+case+'.pickle', "wb")
    pickle.dump(perH, pickle_out)
    pickle_out.close()
    
    pickle_out = open('results/Vmin_'+case+'.pickle', "wb")
    pickle.dump(V_Min_Av, pickle_out)
    pickle_out.close()


