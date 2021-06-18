# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
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

####### OpenDSS Initialisation #########

dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")
dssText = dssObj.Text
DSSCircuit = dssObj.ActiveCircuit
DSSLoads=DSSCircuit.Loads;

cm='cc'
rated_cc=pd.read_excel('rated_'+cm+'_stats.xlsx', sheet_name=None)

if cm=='de':
    rated_cc=pd.read_excel('derated_stats.xlsx', sheet_name=None)

if cm=='cv':
    del rated_cc['BMW_3ph_CV']
    del rated_cc['Zoe_3ph_CV']
    del rated_cc['DC_CV']

if cm=='cc':
    del rated_cc['BMW_3ph_CC']
    del rated_cc['Zoe_3ph_CC']
    del rated_cc['DC_CC']
g55lims=pd.read_csv('g55limits.csv')

if cm=='de':
    del rated_cc['bmw_3ph_6A']
    del rated_cc['bmw_3ph_9A']
    del rated_cc['bmw_3ph_12A']
    del rated_cc['bmw_3ph_15A']
  
    del rated_cc['zoe_3ph_6A']
    del rated_cc['zoe_3ph_12A']
    del rated_cc['zoe_3ph_18A']
    del rated_cc['zoe_3ph_24A']
    
g55lims=pd.read_csv('g55limits.csv')

if cm=='cv' or cm=='cc':
    EV_power=pd.Series(index=rated_cc.keys(),data=[6.6,6.6,7.2,7.2,7.2])
if cm=='de':
    EV_power=pd.Series(index=rated_cc.keys(),data=[1.38,2.76,4.14,5.52,1.38,2.76,4.14,5.52,1.38,2.76,4.14,5.52,1.38,2.76,4.14,5.52,1.38,2.76,4.14,5.52])

####### OpenDSS Initialisation #########

dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")
dssText = dssObj.Text
DSSCircuit = dssObj.ActiveCircuit
DSSLoads=DSSCircuit.Loads;

Vh_ratios=pd.DataFrame()
Ch_ratios=pd.DataFrame()
Pass=pd.DataFrame()


for i in list(rated_cc.keys()):
    Loads=pd.read_csv('Loads.txt', delimiter=' ', names=['New','Load','Phases','Bus1','kV','kW','PF','spectrum'])
    Loads['spectrum'][0]='spectrum='+str(i)
    Loads['kW'][0]='kW='+str(EV_power[i])
    Loads.to_csv('Loads.txt', sep=" ", index=False, header=False)
    vfig,vax=plt.subplots()
    cfig,cax=plt.subplots()
    vax.set_title(i+' Voltage Harmonics')
    cax.set_title(i+' Current Harmonics')
    for p in [25,50,75]:
    ####---- Create Spectrum CSVs
        spectrum=pd.DataFrame()
        spectrum['h']=rated_cc[i]['Harmonic order']
        spectrum['mag']=rated_cc[i]['Ih mag '+str(p)+'% percentile']/rated_cc[i]['Ih mag '+str(p)+'% percentile'][0]*100
        spectrum['ang']=rated_cc[i]['Ih phase mean w.r.t. L1_Ih1']
        spectrum.to_csv('Spectrum'+i+'.csv', header=False, index=False)
        
        g55lims=pd.read_csv('g55limits.csv')
        
        dssObj.ClearAll() 
        dssText.Command="Compile E:/PNDC/TN-006Harmonic/Harmonic_OpenDSS/Master_S.dss"
        dssText.Command="Solve"
        dssText.Command="Solve Mode=harmonics NeglectLoadY=Yes"
        dssText.Command="export monitors m1"
        res_Reactor=pd.read_csv('LVTest_Mon_m1_1.csv')

        Vh_ratios['h']=res_Reactor[' Harmonic']
        Vh_ratios['Lims']=g55lims['L']
        Vh_ratios['V'+str(i)+str(p)]=res_Reactor[' V1']
        Vh_ratios['V_ratio'+str(i)+str(p)]=res_Reactor[' V1']/res_Reactor[' V1'][0]*100

        
        Ch_ratios['h']=res_Reactor[' Harmonic']
        Ch_ratios['C_ratio']=res_Reactor[' I1']/res_Reactor[' I1'][0]*100
        
        Pass['h']= res_Reactor[' Harmonic'][1:]
        Pass['pass'+str(i)+str(p)]=Vh_ratios['V_ratio'+str(i)+str(p)][1:]<Vh_ratios['Lims'][1:]
        
        x=Vh_ratios['h'][1:]
        y=Vh_ratios['V_ratio'+str(i)+str(p)][1:]
        lim=Vh_ratios['Lims'][1:]
        vax.bar(x+p/100-0.5,y, width=0.25,label='Q'+str(p))
        vax.set_xticks(np.arange(1, 50, 2))
        vax.legend()
        vax.set_ylim(0,2)
        for n in x.index:
            if n<=x.index[-1]:
                vax.plot([x[n]-0.5,x[n]+0.5],[lim[n],lim[n]],color='black')
            if n>x.index[-1]:
                vax.plot([x[n]-0.5,x[n]+0.5],[lim[n],lim[n]],color='black',label='G5/5 Limit')
        
        x=Ch_ratios['h'][1:]
        y=Ch_ratios['C_ratio'][1:]
        lim=g55lims['C'][1:]
        cax.bar(x+p/100-0.5,y, width=0.25,label='Q'+str(p))
        cax.set_xticks(np.arange(1, 50, 2))
        cax.legend()
        
        cax.set_ylim(0,4)
        for n in lim.index:
            if n <=lim.index[-1]:
                cax.plot([x[n]-0.5,x[n]+0.5],[lim[n],lim[n]],color='black')
            if n >lim.index[-1]:
                cax.plot([x[n]-0.5,x[n]+0.5],[lim[n],lim[n]],color='black',label='IEC 61000-3-12 Limit')