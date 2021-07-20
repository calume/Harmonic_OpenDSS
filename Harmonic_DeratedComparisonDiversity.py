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

rated_cc=pd.read_excel('rated_CC_stats.xlsx', sheet_name=None)
del rated_cc['BMW_3ph_CC']
del rated_cc['Zoe_3ph_CC']
del rated_cc['DC_CC']
EV_power_cc=pd.Series(index=rated_cc.keys(),data=[6.6,6.6,7.2,7.2,7.2])
  
rated_de=pd.read_excel('derated_stats.xlsx', sheet_name=None)
del rated_de['bmw_3ph_6A']
del rated_de['bmw_3ph_9A']
del rated_de['bmw_3ph_12A']
del rated_de['bmw_3ph_15A']
del rated_de['zoe_3ph_6A']
del rated_de['zoe_3ph_12A']
del rated_de['zoe_3ph_18A']
del rated_de['zoe_3ph_24A']
EV_power_de=pd.Series(index=rated_de.keys(),data=[1.32,2.64,3.96,5.28,1.32,2.64,3.96,5.28,1.35,2.7,4.05,5.4,1.35,2.7,4.05,5.4,1.35,2.7,4.05,5.4])

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
Pass=pd.DataFrame()
VoltageMin={}
VoltageSrc={}
RSCs=[33,66]
master='Master_S'

pq=[-.6,-0.3,0,0.3,0.6]

cases=[]
for i in rated_cc.keys():
    cases.append(i.split('_')[0])
          

f_Rsc=pd.Series(dtype=float,index=RSCs,data=[0.78,0.327])
comps=['CC','De_6A','De_12A','De_18A','De_24A']
htch=pd.Series(index=comps,data=['/////','','XXXXX','','*'])
coll=pd.Series(index=comps, data=['w','#eb3636','w','yellow','g'])
thd=pd.Series(index=comps,dtype=(float))
passes=pd.DataFrame(columns=comps)
cc=0
vfig,vax=plt.subplots(figsize=(7,3))
plt.grid(linewidth=0.2)
# vax.text(.4,.9,str(i),
# horizontalalignment='left',
# transform=vax.transAxes)
vax.set_xlabel('h')
vax.set_ylabel('V'r'$_h$(% of V'r'$_{fund}$'')')
vax.set_xlim(1,29)
indexer=0
for c in comps:
    i=cases[indexer+3]
    dssObj.ClearAll() 
    dssText.Command="Compile "+str(master)+".dss"
    ## Add lines
    dssText.Command ="New Line.LINE1 Bus1=2 Bus2=3 phases=3 Linecode=D2 Length="+str(f_Rsc[33])+" Units=km"
    
    ###-- for bmw and zoe
    if i.lower()=='bmw' or i.lower()=='zoe':
        i=i+'_1ph'
    #--- Add Loads
    
    if c=='CC':
        dssText.Command = "New Load.LOAD1 Phases=1 Status=1 Bus1=3.1 kV=0.230 kW="+str(EV_power_cc[i+'_CC'])+" PF=1 spectrum="+str(i)+'_CC'
        dssText.Command = "New Load.LOAD2 Phases=1 Status=1 Bus1=3.1 kV=0.230 kW="+str(EV_power_cc[i+'_CC'])+" PF=1 spectrum="+str(i)+'_CC'
        
    if c.split('_')[0]=="De":
        amps=c.split('_')[1]
        n_de=int(round(2*EV_power_cc[i+'_CC']/EV_power_de[i.lower()+'_'+str(amps)],0))
        print('Full Power',EV_power_cc[i+'_CC'],'Derated power ', EV_power_de[i.lower()+'_'+str(amps)] , n_de, 'EVs', 'Total derated', round(n_de*EV_power_de[i.lower()+'_'+str(amps)],2))
        cf=(n_de*EV_power_de[i.lower()+'_'+str(amps)])/(2*EV_power_cc[i+'_CC'])
        for q in range(1,n_de+1):
            indexer=indexer+1
            if indexer>4:
                indexer=0
            i=cases[indexer]
            if i.lower()=='bmw' or i.lower()=='zoe':
                i=i+'_1ph'
            dssText.Command = "New Load.LOAD"+str(q)+" Phases=1 Status=1 Bus1=3.1 kV=0.230 kW="+str(EV_power_de[i.lower()+'_'+str(amps)]/cf)+" PF=1 spectrum="+str(i.lower()+'_'+str(amps))
        print('corrected',round(EV_power_de[i.lower()+'_'+str(amps)]/cf*n_de,2))
    dssText.Command="New monitor.M1 Line.Line1 Terminal=2"
    
    dssText.Command="Calcvoltagebases"
    dssText.Command="Solve"
    dssText.Command="export voltages"
    dssText.Command="export currents"
    dssText.Command="export powers"
    
    dssText.Command="Solve Mode=harmonics NeglectLoadY=Yes"
    dssText.Command="export monitors m1"
    
    res_Reactor=pd.read_csv('LVTest_Mon_m1_1.csv')
    
    res_bmw=pd.read_csv('LVTest_Mon_m2_1.csv')
    res_zoe=pd.read_csv('LVTest_Mon_m3_1.csv')
    res_kona=pd.read_csv('LVTest_Mon_m4_1.csv')
    
    Vh_ratios['h']=res_Reactor[' Harmonic']
    Vh_ratios['Lims']=g55lims['L']
    Vh_ratios['V']=res_Reactor[' V1']
    Vh_ratios['V_ratio']=res_Reactor[' V1']/res_Reactor[' V1'][0]*100
    
    thd[c]=sum(Vh_ratios['V_ratio'][1:]**2)**0.5
    passes[c] = Vh_ratios['V_ratio'][1:]<Vh_ratios['Lims'][1:]
    passes['h']=res_Reactor[' Harmonic']
    x=Vh_ratios['h'][1:]
    y=Vh_ratios['V_ratio'][1:]
    
    lim=Vh_ratios['Lims'][1:]
    vax.bar(x+pq[cc],y, width=0.4,label=c,edgecolor='black',hatch=htch[c],color=coll[c])   ###--- Plotting the Voltage harmonics
    cc=cc+1
    iline=DSSLines.First
    while iline>0:
        print(DSSLines.Name, 'Bus1=',DSSLines.Bus1,'Bus2=', DSSLines.Bus2,'Length= ',DSSLines.Length)
        iline=DSSLines.Next
    iload=DSSLoads.First
    while iload>0:
        print(DSSLoads.Name, 'Bus=',DSSCktElement.BusNames, 'kW=',DSSLoads.kW,'pf=',DSSLoads.PF,'spectrum=',DSSLoads.spectrum)
        iload=DSSLoads.Next

    
vax.set_xticks(np.arange(1, 50, 2))
vax.set_ylim(0,2)
for n in x.index:  ###--- Plotting the G5 Limit
    if n<x.index[-1]:
        vax.plot([x[n]-0.5,x[n]+0.5],[lim[n],lim[n]],color='orange')
    if n==x.index[-1]:
        vax.plot([x[n]-0.5,x[n]+0.5],[lim[n],lim[n]],color='orange',label='G5/5 Limit')
vax.legend()

dssText.Command="Solve Mode=Faultstudy"
dssText.Command="export Faultstudy"
dssText.Command="export seqz"
    
seqz=pd.read_csv('LVTest_EXP_SEQZ.csv')
faults=pd.read_csv('LVTest_EXP_FAULTS.csv')

cc=cc+1
print('Zterminal' ,round(seqz[' Z1'][1],4),'Zend',round(seqz[' Z1'][-1:].values[0],4))
print('Fault End' ,faults['  1-Phase'][-1:].values)
# ress=pd.DataFrame()
# ress['bmw']=res_bmw[' IAngle1']
# ress['kona']=res_kona[' IAngle1']

# def getDifference(b1, b2):
# 	r = (b2 - b1) % 360.0
# 	# Python modulus has same sign as divisor, which is positive here,
# 	# so no need to consider negative case
# 	if r >= 180.0:
# 		r -= 360.0
# 	return r

# ress['newdiff']=ress['bmw']-ress['kona']
# for w in ress['newdiff'].index:
#     ress['newdiff'][w]=getDifference(ress['bmw'][w], ress['kona'][w])

# ress['bmw_orig']=rated_cc['BMW_1ph_CC']['Ih phase mean w.r.t. L1_Ih1']
# ress['kona_orig']=rated_cc['Kona_CC']['Ih phase mean w.r.t. L1_Ih1']
# ress['olddiff']=ress['bmw_orig']-ress['kona_orig']

# for w in ress['olddiff'].index:
#     ress['olddiff'][w]=getDifference(ress['bmw_orig'][w], ress['kona_orig'][w])

    
plt.tight_layout()
plt.savefig('figs/'+i+'_Derated_MultipleCars_Diversity.png')

print(thd)