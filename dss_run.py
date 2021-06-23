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

dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")
dssText = dssObj.Text
DSSCircuit = dssObj.ActiveCircuit
DSSLoads=DSSCircuit.Loads
DSSLines=DSSCircuit.Lines
DSSElement=DSSCircuit.ActiveElement
dssObj.Start(0)
##### Load in the Harmonic Profiles ########
case='Test'
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
 
###### Source Impedence is adjusted for Urban/Rural networks

M=19 ##--Number of EVs
B=10 ##--Number of Buses
R=2##--Number of Runs

RSCs=['WPD_Zmax',15,33,66]   ##--- FOr Urban where WPD ZMax is much higher than corresponding RSC
##RSCs=[15,33,66]  ###--- For Rural where WPD ZMax and RSC=15 are similar

#f_Rsc=pd.Series(dtype=float,index=RSCs,data=[1.6,0.62,0.21]) #for 185 mm - RURAL
f_Rsc=pd.Series(dtype=float,index=RSCs,data=[0.77,1.7,0.72,0.305]) #for 185 mm - URBAN

####------Build Lines between B buses

i=random.choice(list(rated_cc.keys()))
admd_factor=1
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
for f in RSCs:
    print(f)
    Ch_Ratios[f]=pd.DataFrame()
    THD_Pass[f]=np.zeros(shape=(M,R),dtype=bool)
    All_Pass[f]=np.zeros(shape=(M,R),dtype=bool)
    All_THDs[f]=np.zeros(shape=(M,R),dtype=float)
    failers[f]=[]
    VoltageMin[f]=np.zeros(shape=(M,R),dtype=float)  
    Pass[f]={}
    for r in range(1,R+1):
        Pass[f][r]=np.zeros(shape=(30,M),dtype=bool)
        print('Run',r)
        for n in range(1,M+1):
            Vh_ratios=pd.DataFrame()
            dssObj.ClearAll() 
            dssText.Command="redirect Master_R.dss"
            ###--- Add Lines
            for L in range(1,B):
                dssText.Command ="New Line.LINE"+str(L)+" Bus1="+str(L+1)+" Bus2="+str(L+2)+" phases=3 Linecode=D2 Length="+str(f_Rsc[f]/B)+" Units=km"
            ###--- Add loads
            if n>=20:
                n=int(n/2)  
            for k in range(1,n+1):
                i=random.choice(list(rated_cc.keys()))
                b=random.choice(range(2,B+1))
                p=random.choice(range(1,4))
                dssText.Command = "New Load.LOAD"+str(k)+" Phases=1 Bus1="+str(b)+"."+str(p)+" kV=230 kW="+str(EV_power[i]*admd_factor)+" PF=1 spectrum="+str(i)
            ###--- Solve Load Flow (and record Vmin)
            dssText.Command="Solve"
            bvs = list(DSSCircuit.AllBusVMag)
            Voltages = bvs[0::3], bvs[1::3], bvs[2::3]
            VoltArray = np.zeros((len(Voltages[0]), 3))
            iLoad=DSSLoads.First
            while iLoad >0:
                print(DSSLoads.Name, DSSElement.BusNames, DSSLoads.kW,DSSLoads.spectrum)
                iLoad=DSSLoads.Next
            print(DSSCircuit.TotalPower)
            for i in range(0, 3):
                VoltArray[:, i] = np.array(Voltages[i], dtype=float)
            VoltageMin[f][(n-1),(r-1)]=VoltArray[-1:].min()

             ###--- Solve Harmonics
            dssText.Command="Solve Mode=harmonics NeglectLoadY=Yes"
            dssText.Command="export monitors m1"
            res_Reactor=pd.read_csv('LVTest_Mon_m1_1.csv')
            Vh_ratios['h']=res_Reactor[' Harmonic']
            Vh_ratios['V']=res_Reactor[' V1']
            Vh_ratios['V_ratio']=res_Reactor[' V1']/res_Reactor[' V1'][0]*100
            Vh_ratios['Lims']=g55lims['L']
            Ch_Ratios[f]['h']=res_Reactor[' Harmonic']
            Ch_Ratios[f][n]=res_Reactor[' I1']/res_Reactor[' I1'][0]*100
            Vh_ratios['V_ratio'][0] = sum(Vh_ratios['V_ratio'][1:]**2)**0.5
            Pass[f][r][:,(n-1)]=Vh_ratios['V_ratio']<Vh_ratios['Lims']
            All_THDs[f][(n-1),(r-1)]=Vh_ratios['V_ratio'].iloc[0]
        iLine=DSSLines.First
        while iLine >0:
            print(DSSLines.Name, DSSLines.Bus1, DSSLines.Bus2, DSSLines.R1, DSSLines.Length)
            iLine=DSSLines.Next
        All_Pass[f][:,r-1]=Pass[f][r].sum(axis=0)==30
        THD_Pass[f][:,r-1]=Pass[f][r][0,:]
        
#     dssText.Command="Solve Mode=Faultstudy"
#     dssText.Command="export Faultstudy"
#     dssText.Command="export seqz"
    
#     seqz[f]=pd.read_csv('LVTest_EXP_SEQZ.csv')
#     faults[f]=pd.read_csv('LVTest_EXP_FAULTS.csv')

    
#     V_Min_Av[f]=VoltageMin[f].mean(axis=1).values
#     V_Min_Av[f]=V_Min_Av.index=range(1,M+1)
    
#     for n in range(1,M+1):
#         failers[f].append((Full_H_pass[f][n].sum(axis=1)).index[Full_H_pass[f][n].sum(axis=1)<r].values)
#     failers[f]=np.unique(np.concatenate(failers[f]))
    
#     for n in range(1,M+1):
#         Full_H_pass[f][n]=Full_H_pass[f][n].loc[failers[f]]

#     perH[f]=pd.DataFrame(index=Full_H_pass[f][list(Full_H_pass[f].keys())[0]][1].index,columns=Full_H_pass[f][list(Full_H_pass[f].keys())[0]].keys())
#     perH[f]['h']=Pass[f]['h'][Full_H_pass[f][list(Full_H_pass[f].keys())[0]][1].index]
#     for j in Full_H_pass[f][list(Full_H_pass[f].keys())[0]].keys():
#         perH[f][j]=1-Full_H_pass[f][j].sum(axis=1)/Full_H_pass[f][j].count(axis=1)


#     Summary[f]={'All_Pass': All_Pass,'Pass': Pass, 'THD_Pass': THD_Pass,'All_THDs':All_THDs,'Full_H_pass':Full_H_pass}

# pickle_out = open('results/Summary_'+case+'.pickle', "wb")
# pickle.dump(Summary, pickle_out)
# pickle_out.close()

# pickle_out = open('results/AllHarmonics_'+case+'.pickle', "wb")
# pickle.dump(perH, pickle_out)
# pickle_out.close()

# pickle_out = open('results/Vmin_'+case+'.pickle', "wb")
# pickle.dump(V_Min_Av, pickle_out)
# pickle_out.close()


# # pick_in = open('results/Summary_'+case+'.pickle', "rb")
# # Unbalanced = pickle.load(pick_in)

# # pick_in = open('results/AllHarmonics_'+case+'.pickle', "rb")
# # perH = pickle.load(pick_in)

# # pick_in = open('results/Vmin_'+case+'.pickle', "rb")
# # V_Min_Av = pickle.load(pick_in)

# styles=pd.Series(data=[':','-.','-','--'],index=RSCs)

# fig, (ax1, ax2) = plt.subplots(2, sharex=False)
# for f in RSCs:
#     ax1.plot(100-(All_Pass[f].sum(axis=1)/All_Pass[f].count(axis=1))*100,label='RSC'+str(f),linestyle=styles[f])
#     ax2.plot(All_THDs[f].index,All_THDs[f].max(axis=1), label='RSC'+str(f),linestyle=styles[f])
#     ax1.xaxis.set_major_formatter(FormatStrFormatter('% 1.0f'))
#     ax2.xaxis.set_major_formatter(FormatStrFormatter('% 1.0f'))

# ax1.set_ylabel('% Probability of Failure')
# ax1.legend()
# ax1.grid(linewidth=0.2)
# ax1.set_xlim(1,M)
# ax1.set_ylim(0,100)
# ax2.set_ylabel('Maximum THD')
# ax2.legend()
# ax2.set_xlabel('Number of EVs')
# ax2.grid(linewidth=0.2)
# ax2.set_xlim(1,M)
# ax2.set_ylim(0,5)
# print('Max prob of THD Failure', (100-((THD_Pass[f]==True).sum(axis=1)/THD_Pass[f].count(axis=1))*100).max())
# plt.tight_layout()

# plt.figure('Specific Harmonics',figsize=(5, 8))

# allfails=[]
# for f in RSCs:
#     allfails.append(perH[f].index)
# allfails=np.unique(np.concatenate(allfails))

# c=1
# for pl in allfails:
#     ax=plt.subplot(len(allfails),1, c)
#     ax.set_ylabel('% Failure')
#     ax.text(.5,.8,'h='+str(int(perH[f]['h'][pl])),
#         horizontalalignment='left',
#         transform=ax.transAxes)
#     for f in RSCs:
#         if len(perH[f])>0:
#             if c<len(perH[f]):
#                 ax.plot(perH[f].loc[pl][1:M]*100,linestyle=styles[f])
#             if c==len(perH[f]):
#                 ax.plot(perH[f].loc[pl][1:M]*100, label='RSC='+str(f),linestyle=styles[f])
#     c=c+1
#     if c>len(allfails):
#         ax.legend()
#     ax.xaxis.set_major_formatter(FormatStrFormatter('% 1.0f'))
#     plt.grid(linewidth=0.2)
#     ax.set_xlim(1,M)
#     ax.set_ylim(0,100)
#     ax.xaxis.set_major_formatter(FormatStrFormatter('% 1.0f'))
#     plt.tight_layout()


# plt.figure()
# for f in RSCs:
#     plt.plot(V_Min_Av.index, V_Min_Av[f].values, linestyle=styles[f], label='Vmin with Rsc='+str(f), linewidth=1)
# plt.ylabel('Voltage (V)')
# plt.plot([1,M],[216,216],color='Black',linestyle=":", linewidth=0.5, label='Statutory Min')
# plt.xlabel('Number of EVs')
# plt.legend()
