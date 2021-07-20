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
DSSLines=DSSCircuit.Lines;
dssObj.Start(0)
dssObj.AllowForms=False
DSSCktElement = DSSCircuit.ActiveCktElement

cm='CC'
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
g55lims['L'][0:10]=2
####### OpenDSS Initialisation #########

dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")
dssText = dssObj.Text
DSSCircuit = dssObj.ActiveCircuit
DSSLoads=DSSCircuit.Loads
DSSTrans=DSSCircuit.Transformers

Vh_ratios=pd.DataFrame()
Ch_ratios=pd.DataFrame()
Pass={}
VoltageMin={}
VoltageSrc={}
RSCs=[15,33,66]
master='Master_S'
####---- Create Loads
p=75
ll={}
Vsummary=pd.DataFrame(index=rated_cc.keys(),columns=RSCs)
THD=pd.DataFrame(index=RSCs,columns=list(rated_cc.keys()))
for i in list(rated_cc.keys()):
    print(i)
    ####---- Create Spectrum CSVs
    spectrum=pd.DataFrame()
    spectrum['h']=rated_cc[i]['Harmonic order']
    spectrum['mag']=rated_cc[i]['Ih mag '+str(p)+'% percentile']/rated_cc[i]['Ih mag '+str(p)+'% percentile'][0]*100
    spectrum['ang']=rated_cc[i]['Ih phase mean w.r.t. L1_Ih1']/spectrum['h']
    spectrum.to_csv('Spectrum'+i+'.csv', header=False, index=False)
    
    #i =list(rated_cc.keys())[3]
    vfig,vax=plt.subplots(figsize=(7, 3))
    plt.grid(linewidth=0.2)
    # vax.text(.4,.9,str(i),
    # horizontalalignment='left',
    # transform=vax.transAxes)
    #cfig,cax=plt.subplots()
    #vax.set_title(i+' Voltage Harmonics')
    # cax.set_title(i+' Current Harmonics')
    vax.set_xlabel('h')
    vax.set_ylabel('V'r'$_h$(% of V'r'$_{fund}$'')')
    vax.set_xlim(1,29)
    pq=[-.5,0,0.5]
    cc=0
    
    htch=pd.Series(index=RSCs,data=['////','','XXXX'])
    coll=pd.Series(index=RSCs, data=['w','#eb3636','yellow'])
    seqz={}
    faults={}
    currents={}
    voltages={}
    cvf=1
    if cm=='CV' and (i=='Leaf_CV' or i=='Van_CV'):
        cvf=0.5 
    seqz={}
    faults={}
    currents={}
    voltages={}
    B=5 ##--Number of Buses
    f_Rsc=pd.Series(dtype=float,index=RSCs,data=[1.87,0.78,0.327])   ##--- FOr Urban where WPD ZMax is much higher than corresponding RSC
    
    for f in RSCs:     
        if i==list(rated_cc.keys())[0]:
            Pass[f]=pd.DataFrame()           
        dssObj.ClearAll() 
        dssText.Command="Compile "+str(master)+".dss"
        #--- Add Lines
        for L in range(1,B-1):
            dssText.Command ="New Line.LINE"+str(L)+" Bus1="+str(L+1)+" Bus2="+str(L+2)+" phases=3 Linecode=D2 Length="+str(f_Rsc[f]/(B-2))+" Units=km"
        load_model=1
        if cm=='CC' or 'de':
            load_model=5
        #--- Add Loads
        if phh=='1ph':
            cp=1
            for k in range(1,B):
                for q in range(1,4):
                    dssText.Command = "New Load.LOAD"+str(cp)+" Phases=1 Model="+str(load_model)+" Status=1 Bus1="+str(k+1)+"."+str(q)+" kV=0.230 kW="+str(EV_power[i]*cvf)+" PF=1 spectrum="+str(i)
                    cp=cp+1
        bmw_factor=B
        if i[:3]=='BMW':
            bmw_factor=bmw_factor*2-1
        oo=2
        if phh=='3ph':
            for cp in range(1,bmw_factor):
                dssText.Command = "New Load.LOAD"+str(cp)+" Phases=3 Status=1 Bus1="+str(oo)+" kV=0.4 kW="+str(EV_power[i]*cvf)+" PF=1 spectrum="+str(i)
                oo=oo+1
                if oo>4:
                    oo=2
            cp=cp+1
        
        #dssText.Command="New monitor.M1 Reactor.R1 Terminal=2"
        dssText.Command="New monitor.M"+str(f)+" Line.Line3 Terminal=2"
        dssText.Command="Calcvoltagebases"
        dssText.Command="Solve"
        dssText.Command="export voltages"
        dssText.Command="export currents"
        dssText.Command="export powers"
        
        currents[f]=pd.read_csv('LVTest_EXP_CURRENTS.csv')
        voltages[f]=pd.read_csv('LVTest_EXP_VOLTAGES.csv')
        bvs = list(DSSCircuit.AllBusVMag)
        Voltages = bvs[0::3], bvs[1::3], bvs[2::3]
        VoltArray = np.zeros((len(Voltages[0]), 3))
        iLoad=DSSLoads.First
        for z in range(0, 3):
            VoltArray[:, z] = np.array(Voltages[z], dtype=float)
        VoltageMin[f]=VoltArray[-1:].mean()
        VoltageSrc[f]=VoltArray[2].mean()
        
        Vsummary[f][i]=round(VoltageMin[f],1)
            
        dssText.Command="Solve Mode=harmonics NeglectLoadY=Yes"
        dssText.Command="export monitors m"+str(f)
        res_Reactor=pd.read_csv('LVTest_Mon_m'+str(f)+'_1.csv')
    
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
        vax.bar(x+pq[cc],y, width=0.4,label='RSC'+str(f),edgecolor='black',hatch=htch[f],color=coll[f])   ###--- Plotting the Voltage harmonics
        vax.set_xticks(np.arange(1, 50, 2))
        
        vax.set_ylim(0,2)
        for n in x.index:  ###--- Plotting the G5 Limit
            if n<x.index[-1]:
                vax.plot([x[n]-0.5,x[n]+0.5],[lim[n],lim[n]],color='orange')
            if n==x.index[-1] and f==66:
                vax.plot([x[n]-0.5,x[n]+0.5],[lim[n],lim[n]],color='orange',label='G5/5 Limit')
        vax.legend()#framealpha=1,bbox_to_anchor=(0, 1), loc='upper left', ncol=2)
        plt.savefig('figs/'+i+'_'+str(cp-1)+'EVs_Balanced.png')
        # x=Ch_ratios['h'][1:]
        # y=Ch_ratios['C_ratio'][1:]
        # lim=g55lims['C'][1:]
        # cax.bar(x+pq[cc]/100-0.5,y, width=0.4,label=ne+' RSC'+str(f),hatch=htch[f],color=coll)
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
            
        seqz[f]=pd.read_csv('LVTest_EXP_SEQZ.csv')
        faults[f]=pd.read_csv('LVTest_EXP_FAULTS.csv')
        
        cc=cc+1
        print('Zterminal '+str(f),round(seqz[f][' Z1'][2],4),'Zend'+str(f),round(seqz[f][' Z1'][-1:].values[0],4),'Zdiff ',-round(seqz[f][' Z1'][2]-seqz[f][' Z1'][-1:].values[0],4))
        print('Fault End '+str(f),faults[f]['  1-Phase'][-1:].values)
        print('Vsource '+str(f),round(VoltageSrc[f],2),'VEnd '+str(f),round(VoltageMin[f],2))
        plt.tight_layout()
        
    print(i)
    iline=DSSLines.First
    while iline>0:
        print(DSSLines.Name, 'Bus1=',DSSLines.Bus1,'Bus2=', DSSLines.Bus2)
        iline=DSSLines.Next
    iload=DSSLoads.First
    while iload>0:
        print(DSSLoads.Name, 'Bus=',DSSCktElement.BusNames, 'kW=',DSSLoads.kW,'pf=',DSSLoads.PF,'spectrum=',DSSLoads.spectrum)
        iload=DSSLoads.Next

t=THD.transpose()
paa={}
fails={}
evs={}
evs=pd.DataFrame(index=list(Pass.keys()),columns=Pass[15].columns[1:])
for i in Pass:
    paa[i]=Pass[i].sum()
    idx=Pass[i].iloc[:,1:].sum(axis=1)<5
    fails[i]=Pass[i].loc[idx].transpose()
    fails[i].columns=fails[i].loc['h',:]
    for h in fails[i].index[1:]:
        evs[h][i]=fails[i].loc[h][fails[i].loc[h]==False].index.values
        
evs=evs.transpose()