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

cases=['Terminal_LoadY','End_LoadN','Terminal_LoadN']
cases=['End_LoadY']
for cse in cases:
    cms=['CC','de','CV']
    cms=['CC']
    for cm in cms:
        
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
        cf=2
        n_evs=50
        if cm=='CV':
            n_evs=100
            cf=1
        n_buses=50
        R=3
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
        seqz={}
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
            seqz[net]={}
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
                        if cse[:3]=='End':
                            dssText.Command="New monitor.M"+str(n)+" Line.LINE"+str(max(rb))+" Terminal=2"
                        if cse[:8]=='Terminal':
                            dssText.Command="New monitor.M"+str(n)+" Reactor.R1 Terminal=2"
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
                        if cse[-5:]=='LoadN':
                            dssText.Command="Solve Mode=harmonics"
                        if cse[-5:]=='LoadY':
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
                
                seqz[net][f]=pd.read_csv('LVTest_EXP_SEQZ.csv')
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
            
        # pickle_out = open('results/'+cse+'/Summary_'+case+'.pickle', "wb")
        # pickle.dump(Summary, pickle_out)
        # pickle_out.close()
        
        # pickle_out = open('results/'+cse+'/AllHarmonics_'+case+'.pickle', "wb")
        # pickle.dump(perH, pickle_out)
        # pickle_out.close()
        
        # pickle_out = open('results/'+cse+'/Vmin_'+case+'.pickle', "wb")
        # pickle.dump(V_Min_Av, pickle_out)
        # pickle_out.close()


