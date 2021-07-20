# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
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

####### OpenDSS Initialisation #########

dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")
dssText = dssObj.Text
DSSCircuit = dssObj.ActiveCircuit
DSSLoads=DSSCircuit.Loads;
DSSCktElement = DSSCircuit.ActiveCktElement
DSSLines=DSSCircuit.Lines;
dssObj.Start(0)
dssObj.AllowForms=False
cm='de'
phh='1ph'

if (cm=='CV' or cm=='CC') and phh=='1ph':
    rated_cc=pd.read_excel('rated_'+str(cm)+'_stats.xlsx', sheet_name=None)
    del rated_cc['BMW_3ph_'+str(cm)]
    del rated_cc['Zoe_3ph_'+str(cm)]
    del rated_cc['DC_'+str(cm)]
    EV_power=pd.Series(index=rated_cc.keys(),data=[6.6,6.6,7.2,7.2,7.2])

if (cm=='CV' or cm=='CC') and phh=='3ph':
    rcN={}
    rated_cc=pd.read_excel('rated_'+str(cm)+'_stats.xlsx', sheet_name=None)
    rcN['BMW_3ph_'+str(cm)]=rated_cc['BMW_3ph_'+str(cm)]
    rcN['Zoe_3ph_'+str(cm)]=rated_cc['Zoe_3ph_'+str(cm)]
    rated_cc=rcN
    EV_power=pd.Series(index=rated_cc.keys(),data=[11,22])
    
if cm=='de' and phh=='1ph':
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
 
if cm=='de' and phh=='3ph':
    rcN={}
    rated_cc=pd.read_excel('derated_stats.xlsx', sheet_name=None)
    rcN['bmw_3ph_6A']= rated_cc['bmw_3ph_6A']
    rcN['bmw_3ph_9A']= rated_cc['bmw_3ph_9A']
    rcN['bmw_3ph_12A']= rated_cc['bmw_3ph_12A']
    rcN['bmw_3ph_15A']= rated_cc['bmw_3ph_15A']
    rcN['zoe_3ph_6A']= rated_cc['zoe_3ph_6A']
    rcN['zoe_3ph_12A']= rated_cc['zoe_3ph_12A']
    rcN['zoe_3ph_18A']= rated_cc['zoe_3ph_18A']
    rcN['zoe_3ph_24A']= rated_cc['zoe_3ph_24A']
    rated_cc=rcN
    EV_power=pd.Series(index=rated_cc.keys(),data=[4.14,6.21,8.28,10.35,4.14,8.28,12.42,16.56])
 
    
    
g55lims=pd.read_csv('g55limits.csv')

####### OpenDSS Initialisation #########

dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")
dssText = dssObj.Text
DSSCircuit = dssObj.ActiveCircuit
DSSLoads=DSSCircuit.Loads
DSSTrans=DSSCircuit.Transformers

Vh_ratios=pd.DataFrame()
Ch_ratios=pd.DataFrame()
Pass=pd.DataFrame()
VoltageMin={}
VoltageSrc={}
RSCs=[15,33,66]
####---- Create Loads
p=75
ll={}
if phh=='1ph':
    cars=['van','leaf','bmw_1ph','zoe_1ph','kona']
    ratings=[6,12,18,24]
    ratings=[12]
if phh =='3ph':
    cars=['bmw_3ph','zoe_3ph']

f=33
seqz={}
faults={}
Pass={}
f_Rsc=pd.Series(dtype=float,index=RSCs,data=[1.87,0.78,0.327]) 
THD=pd.DataFrame(index=RSCs,columns=cars)
Vsummary=pd.DataFrame(index=cars,columns=RSCs)
for i in list(cars):
    if phh =='3ph':
        ratings={}
        ratings['bmw_3ph']=[6,9,12,15]
        ratings['zoe_3ph']=[6,12,18,24]
    ####---- Create Spectrum CSVs
    cc=0
    #i =list(rated_cc.keys())[3]
    vfig,vax=plt.subplots(figsize=(7, 3))
    plt.grid(linewidth=0.2)
    vax.text(.4,.9,str(i),
    horizontalalignment='left',
    transform=vax.transAxes)
    #cfig,cax=plt.subplots()
    #vax.set_title(i+' Voltage Harmonics')
    # cax.set_title(i+' Current Harmonics')
    vax.set_xlabel('h')
    vax.set_ylabel('V'r'$_h$(% of V'r'$_{fund}$'')')
    vax.set_xlim(1,29)
    pq=[-0.5,0,0.5]
    seqz[i]={}
    faults[i]={}
    if phh=='3ph':
        ratings=ratings[i]
    htch=pd.Series(index=RSCs,data=['','////','XXXX'])#,'\\\\'])
    coll=pd.Series(index=RSCs, data=['w','#eb3636','#90ee90'])#,'#add8e6'])
    for f in RSCs:
        if i==cars[0]:
            Pass[f]=pd.DataFrame()   
        r=12
        B=11##--Number of Buses
        g55lims=pd.read_csv('g55limits.csv')
        
        dssObj.ClearAll() 
        dssText.Command="Compile Master_S.dss"
           
        #--- Add Lines
        for L in range(1,B-1):
            dssText.Command ="New Line.LINE"+str(L)+" Bus1="+str(L+1)+" Bus2="+str(L+2)+" phases=3 Linecode=D2 Length="+str(f_Rsc[f]/(B-2))+" Units=km"
        #--- Add Loads
        if phh=='1ph':
            cp=1
            for k in range(1,B):
                for q in range(1,4):
                    dssText.Command = "New Load.LOAD"+str(cp)+" Model=5 Phases=1 Status=1 Bus1="+str(k+1)+"."+str(q)+" kV=0.230 kW="+str(EV_power[i+"_"+str(r)+"A"])+" PF=1 spectrum="+str(i+"_"+str(r)+"A")
                    cp=cp+1
        bmw_factor=B
        if i[:3]=='BMW':
            bmw_factor=bmw_factor*2-1
        oo=1
        if phh=='3ph':
            for cp in range(1,bmw_factor):
                dssText.Command = "New Load.LOAD"+str(cp)+" Phases=3 Status=1 Bus1="+str(oo)+" kV=0.4 kW="+str(EV_power[i+"_"+str(r)+"A"])+" PF=1 spectrum="+str(i+"_"+str(r)+"A")
                oo=oo+1
                if oo>4:
                    oo=2
            cp=cp+1
        
        #dssText.Command="New monitor.M1 Reactor.R1 Terminal=2"
        dssText.Command="New monitor.M1 Line.LINE"+str(B-2)+" Terminal=2"
        dssText.Command="Solve"
    
        bvs = list(DSSCircuit.AllBusVMag)
        Voltages = bvs[0::3], bvs[1::3], bvs[2::3]
        VoltArray = np.zeros((len(Voltages[0]), 3))
        iLoad=DSSLoads.First
        for z in range(0, 3):
            VoltArray[:, z] = np.array(Voltages[z], dtype=float)
        VoltageMin[i+'_'+str(r)+'A']=VoltArray[-1:].mean()
        VoltageSrc[i+'_'+str(r)+'A']=VoltArray[1].mean()
        
        Vsummary[f][i]=round(VoltageMin[i+'_'+str(r)+'A'],1)
            
        dssText.Command="Solve Mode=harmonics NeglectLoadY=Yes"
        dssText.Command="export monitors m1"
        res_Reactor=pd.read_csv('LVTest_Mon_m1_1.csv')
    
        Vh_ratios['h']=res_Reactor[' Harmonic']
        Vh_ratios['Lims']=g55lims['L']
        Vh_ratios['V']=res_Reactor[' V1']
        Vh_ratios['V_ratio']=res_Reactor[' V1']/res_Reactor[' V1'][0]*100
    
        
        Ch_ratios['h']=res_Reactor[' Harmonic']
        Ch_ratios['C_ratio']=res_Reactor[' I1']/res_Reactor[' I1'][0]*100
        
        Pass[f]['h']= res_Reactor[' Harmonic'][1:]
        Pass[f]['pass'+str(i)+str(p)]=Vh_ratios['V_ratio'][1:]<Vh_ratios['Lims'][1:]
        THD[i][f]=sum(Vh_ratios['V_ratio'][1:]**2)**0.5 
            
        x=Vh_ratios['h'][1:]
        y=Vh_ratios['V_ratio'][1:]
        lim=Vh_ratios['Lims'][1:]
        vax.bar(x+pq[cc],y, width=0.4,label='RSC='+str(f),edgecolor='black',hatch=htch[f],color=coll[f])   ###--- Plotting the Voltage harmonics
        vax.set_xticks(np.arange(1, 50, 2))
        
        vax.set_ylim(0,2)
        for n in x.index:  ###--- Plotting the G5 Limit
            if n<x.index[-1]:
                vax.plot([x[n]-0.5,x[n]+0.5],[lim[n],lim[n]],color='orange')
            if n==x.index[-1] and r==24:
                vax.plot([x[n]-0.5,x[n]+0.5],[lim[n],lim[n]],color='orange',label='G5/5 Limit')
        vax.legend()#framealpha=1,bbox_to_anchor=(0, 1.1), loc='upper left', ncol=2)
        
        # x=Ch_ratios['h'][1:]
        # y=Ch_ratios['C_ratio'][1:]
        # lim=g55lims['C'][1:]
        # cax.bar(x+pq[cc]/100-0.5,y, width=0.4,label=ne+' RSC'+str(f),hatch=htch[f],color=coll[ne])
        # cax.set_xticks(np.arange(1, 50, 2))
        # cax.legend()
        
        # cax.set_ylim(0,4)
        # for n in lim.index:
        #     if n <=lim.index[-1]:
        #         cax.plot([x[n]-0.5,x[n]+0.5],[lim[n],lim[n]],color='orange')
        #     if n >lim.index[-1]:
        #         cax.plot([x[n]-0.5,x[n]+0.5],[lim[n],lim[n]],color='orange',label='IEC 61000-3-12 Limit')
    
        dssText.Command="Solve Mode=Faultstudy"
        dssText.Command="export Faultstudy"
        dssText.Command="export seqz"
            
        seqz[i][r]=pd.read_csv('LVTest_EXP_SEQZ.csv')
        faults[i][r]=pd.read_csv('LVTest_EXP_FAULTS.csv')
        
        cc=cc+1
        print('Zterminal '+str(i+'_'+str(r)+'A'),round(seqz[i][r][' Z1'][1],4),'Zend',round(seqz[i][r][' Z1'][-1:].values[0],4))
        print('Fault End '+str(i+'_'+str(r)+'A'),faults[i][r]['  1-Phase'][-1:].values)
        print('Vsource '+str(i+'_'+str(r)+'A'),round(VoltageSrc[i+'_'+str(r)+'A'],2),'VEnd ',round(VoltageMin[i+'_'+str(r)+'A'],2))
    plt.savefig('figs/'+i+'_'+cm+'_'+str(cp-1)+'EVs_Balanced.png')
    plt.tight_layout()
    print(i)
    iline=DSSLines.First
    while iline>0:
        print(DSSLines.Name, 'Bus1=',DSSLines.Bus1,'Bus2=', DSSLines.Bus2)
        iline=DSSLines.Next
    iload=DSSLoads.First
    while iload>0:
        print(DSSLoads.Name, 'Bus=',DSSCktElement.BusNames, 'kW=',DSSLoads.kW, 'spectrum=', DSSLoads.Spectrum)
        iload=DSSLoads.Next
        
t=THD.transpose()
fails={}
evs={}
evs=pd.DataFrame(index=list(Pass.keys()),columns=Pass[15].columns[1:])
for i in Pass:
    idx=Pass[i].iloc[:,1:].sum(axis=1)<5
    fails[i]=Pass[i].loc[idx].transpose()
    fails[i].columns=fails[i].loc['h',:]
    for h in fails[i].index[1:]:
        evs[h][i]=fails[i].loc[h][fails[i].loc[h]==False].index.values
        
evs=evs.transpose()

