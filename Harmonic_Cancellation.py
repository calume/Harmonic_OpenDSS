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
net_type=['urban','rural']
master='Master_S'
####---- Create Loads
p=75
ll={}
Vsummary=pd.DataFrame(index=rated_cc.keys(),columns=['rural33', 'rural66', 'urban33', 'urban66'])
i=list(rated_cc.keys())[2]
print(i)
####---- Create Spectrum CSVs

vfig,vax=plt.subplots(figsize=(7, 3))
plt.grid(linewidth=0.2)
vax.set_xlabel('h')
vax.set_ylabel('V'r'$_h$(% of V'r'$_{fund}$'')')
vax.set_xlim(1,29)
pq=[-.6,-0.2,0.2,0.6]
cc=0

cases=list(rated_cc.keys())[2:]
#cases.append('Kona_CC')
cases.append('combined')

htch=pd.Series(index=cases,data=['////','','XXXXX',''])
coll=pd.Series(index=cases, data=['w','#eb3636','w','yellow'])
 
f_Rsc=pd.Series(dtype=float,index=RSCs,data=[0.66,0.209])

for i in cases:
    if i !='combined':
        spectrum=pd.DataFrame()
        spectrum['h']=rated_cc[i]['Harmonic order']
        spectrum['mag']=rated_cc[i]['Ih mag '+str(p)+'% percentile']/rated_cc[i]['Ih mag '+str(p)+'% percentile'][0]*100
        spectrum['ang']=rated_cc[i]['Ih phase mean w.r.t. L1_Ih1']
        spectrum.to_csv('Spectrum'+i+'.csv', header=False, index=False)
    
    dssObj.ClearAll() 
    dssText.Command="Compile "+str(master)+".dss"
    B=10
    ## Add lines
    for L in range(1,B-1):
        dssText.Command ="New Line.LINE"+str(L)+" Bus1="+str(L+1)+" Bus2="+str(L+2)+" phases=3 Linecode=D2 Length="+str(f_Rsc[33]/(B-2))+" Units=km"
    
    #--- Add Loads
    ca=0
    po=0
    for q in range(1,B):
        s=i
        if i == 'combined':
            s=cases[po]
        dssText.Command = "New Load.LOAD"+str(ca)+" Phases=1 Status=1 Bus1="+str(q+1)+".1 kV=0.230 kW="+str(EV_power[s])+" PF=1 spectrum="+str(s)
        # for ph in range(1,4):
        #     dssText.Command = "New Load.LOAD"+str(ca)+" Phases=1 Status=1 Bus1="+str(q+1)+"."+str(ph)+" kV=0.230 kW="+str(EV_power[s])+" PF=1 spectrum="+str(s)
        ca=ca+1
        po=po+1
        if po>2:
            po=0
        
    dssText.Command="New monitor.M1 Line.Line5 Terminal=2"
    # dssText.Command="New monitor.M2 Load.Load1 Terminal=1"
    # dssText.Command="New monitor.M3 Load.Load2 Terminal=1"
    # dssText.Command="New monitor.M4 Load.Load3 Terminal=1"
    
    dssText.Command="Calcvoltagebases"
    dssText.Command="Solve"
    dssText.Command="export voltages"
    dssText.Command="export currents"
    dssText.Command="export powers"
    
    dssText.Command="Solve Mode=harmonics NeglectLoadY=Yes"
    dssText.Command="export monitors m1"
    # dssText.Command="export monitors m2"
    # dssText.Command="export monitors m3"
    # dssText.Command="export monitors m4"
    
    res_Reactor=pd.read_csv('LVTest_Mon_m1_1.csv')
    
    res_bmw=pd.read_csv('LVTest_Mon_m2_1.csv')
    res_zoe=pd.read_csv('LVTest_Mon_m3_1.csv')
    res_kona=pd.read_csv('LVTest_Mon_m4_1.csv')
    
    Vh_ratios['h']=res_Reactor[' Harmonic']
    Vh_ratios['Lims']=g55lims['L']
    Vh_ratios['V'+str(i)+str(p)]=res_Reactor[' V1']
    Vh_ratios['V_ratio'+str(i)+str(p)]=res_Reactor[' V1']/res_Reactor[' V1'][0]*100
       
    x=Vh_ratios['h'][1:]
    y=Vh_ratios['V_ratio'+str(i)+str(p)][1:]
    
    lim=Vh_ratios['Lims'][1:]
    vax.bar(x+pq[cc],y, width=0.4,label=i,edgecolor='black',hatch=htch[i],color=coll[i])   ###--- Plotting the Voltage harmonics
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
    
