# -*- coding: utf-8 -*-
"""
Created on Fri Jul  2 08:29:33 2021

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

cases=['Terminal_LoadY','End_LoadN','Terminal_LoadN']
cf=1
cm='CV'
net_type=['rural','urban']
n_evs=100
n_buses=50
R=500
RSCs=[33,66]
M=n_evs ##--Number of EVs
B=n_buses ##--Number of Buses
case=str(cm)+'_Unbalanced_'+str(M)+'EVs_'+str(R)+'Runs_'

cse='End_LoadY'
   
pick_in = open('results/'+cse+'/Summary_'+case+'.pickle', "rb")
Summary = pickle.load(pick_in)

pick_in = open('results/'+cse+'/AllHarmonics_'+case+'.pickle', "rb")
perH = pickle.load(pick_in)

pick_in = open('results/'+cse+'/Vmin_'+case+'.pickle', "rb")
V_Min_Av = pickle.load(pick_in)

styles=pd.Series(data=[':','-'],index=net_type)
cols=pd.Series(data=['tab:green','tab:red'],index=net_type)

for f in RSCs:
    M=n_evs ##--Number of EVs
    if f==66:
        M=M*cf
    fig, (ax1, ax2) = plt.subplots(2, sharex=False, figsize=(4.5, 4))
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
    ax1.set_xlabel('Number of EVs')
    ax2.set_ylabel('Maximum THD (%)')
    ax2.legend()
    ax2.set_xlabel('Number of EVs')
    ax2.grid(linewidth=0.2)
    ax2.set_xlim(1,M-1)
    ax2.set_ylim(0,5)
    ax2.set_xticks(range(0,M,10))
    #ax2.set_xticklabels(range(1,(M+1),10))
    plt.tight_layout()
    plt.savefig('figs/'+str(cm)+'_Failure_THD_'+str(f)+'.png')

    both = (perH[net_type[0]][f].sum(axis=1)+perH[net_type[1]][f].sum(axis=1))<(M*R*2*0.99)
    both=both[both]
    
    if len(both)>0:
        plt.figure('RSC='+str(f)+' Specific Harmonics 1',figsize=(4.5, 6.5))
        c=1
        for pl in both.index[:5]:
            ax=plt.subplot(len(both.index[:5]),1, c)
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
            ax.set_xlabel('Number of EVs')
            plt.tight_layout()
            c=c+1
        plt.savefig('figs/'+str(cm)+'_Specific_Harmonics1_'+str(f)+'.png')
        
    if len(both)>5:
        plt.figure('RSC='+str(f)+' Specific Harmonics 2',figsize=(4.5, 6.5))
        c=1
        for pl in both.index[5:]:
            ax=plt.subplot(len(both.index[5:]),1, c)
            ax.set_ylabel('% Failure')
            ax.text(.5,.8,'h='+str(pl),
                horizontalalignment='left',
                transform=ax.transAxes)
            for net in net_type:
                ax.plot(100-perH[net][f].loc[int(pl)][1:M]/R*100, label=net,linestyle=styles[net], color=cols[net])
            ax.legend()
            ax.xaxis.set_major_formatter(FormatStrFormatter('% 1.0f'))
            ax.set_xlabel('Number of EVs')
            plt.grid(linewidth=0.2)
            ax.set_xlim(0,M)
            ax.set_xticks(range(0,M,10))
            #ax.set_xticklabels(range(1,(M+1),10))
            ax.set_ylim(0,100)
            ax.xaxis.set_major_formatter(FormatStrFormatter('% 1.0f'))
            plt.tight_layout()
            c=c+1 
        plt.savefig('figs/'+str(cm)+'_Specific_Harmonics2_'+str(f)+'.png')
    
for f in RSCs:
    M=n_evs ##--Number of EVs
    if f==66:
        M=M*cf
    plt.figure('RSC='+str(f),figsize=(4, 3))
    for net in net_type:
        plt.plot(V_Min_Av[net][f], linestyle=styles[net], label='Vmin '+net, linewidth=1, color=cols[net])
    plt.ylabel('Voltage (V)')
    plt.plot([0,M],[216,216],color='Black',linestyle=":", linewidth=0.5, label='Statutory Min')
    plt.xticks(ticks=range(0,M,10))
    plt.xlabel('Number of EVs')
    plt.xlim(0,M-1)
    plt.grid(linewidth=0.2)
    plt.legend()
    plt.tight_layout()
    plt.savefig('figs/'+str(cm)+'_Vmin_'+str(f)+'.png')