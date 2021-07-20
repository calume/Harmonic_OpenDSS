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


n_evs=100
R=500
cases=['de']#,'de','CV']#,'de']
RSCs=[15,33,66]

Res_table=pd.DataFrame(index=RSCs,columns=['0Tol','5Tol','25Tol','Vmin','ThD'])
for cm in cases:
    M=n_evs ##--Number of EVs
    case=str(cm)+'_Unbalanced_'+str(M)+'EVs_'+str(R)+'Runs_'#Compensate_Derated'
    
      
    pick_in = open('results/Summary_'+case+'.pickle', "rb")
    Summary = pickle.load(pick_in)
    
    pick_in = open('results/AllHarmonics_'+case+'.pickle', "rb")
    perH = pickle.load(pick_in)
    
    pick_in = open('results/Vmin_'+case+'.pickle', "rb")
    V_Min_Av = pickle.load(pick_in)

    pick_in = open('results/PhaseCounts_'+case+'.pickle', "rb")
    phase_counts = pickle.load(pick_in)
    
    styles=pd.Series(data=[':','-','--'],index=RSCs)
    cols=pd.Series(data=['tab:green','tab:red','tab:blue'],index=RSCs)
    
    fig, (ax1, ax2) = plt.subplots(2, sharex=False, figsize=(4.5, 4.5))
    #ax1.set_title(cm)
    
    
    for f in RSCs:
        M=n_evs ##--Number of EVs
        if f==15 and cm!='CV':
            M=int(M/2)
        
        allps=100-(Summary['All_Pass'][f].sum(axis=0)/R*100)
        p_series=pd.Series(allps,index=range(1,len(allps)+1))
        thds=pd.Series(Summary['All_THDs'][f].max(axis=0),index=range(1,len(allps)+1))
        vmins=pd.Series(V_Min_Av[f],index=range(1,len(allps)+1))
        m,c=np.polyfit(vmins.index,vmins,1)
        mt,ct=np.polyfit(thds.index,thds,1)
        
        Res_table['0Tol'][f]=p_series[p_series>0].index[0]-1
        Res_table['5Tol'][f]=p_series[p_series>5].index[0]-1
        Res_table['25Tol'][f]=p_series[p_series>25].index[0]-1
        
        if vmins[-1:].values[0]<216:
            Res_table['Vmin'][f]=vmins[vmins<216].index[0]-1
        if vmins[-1:].values[0]>216:
            Res_table['Vmin'][f]=str(round((216-c)/m,0))+'*'
        
        if (thds>5).any()==True:
            Res_table['ThD'][f]=thds[thds>5].index[0]-1
        if (thds>5).any()==False:
            Res_table['ThD'][f]=str(round((5-ct)/mt,0))+'*'
            
        ax1.plot(allps,label='RSC='+str(f),linestyle=styles[f], color=cols[f])
        ax2.plot(Summary['All_THDs'][f].max(axis=0), label='RSC='+str(f),linestyle=styles[f], color=cols[f])
        ax1.xaxis.set_major_formatter(FormatStrFormatter('% 1.0f'))
        ax2.xaxis.set_major_formatter(FormatStrFormatter('% 1.0f'))
        print('Max prob of THD Failure RSC='+str(f), (100-((Summary['THD_Pass'][f]==True).sum(axis=0)/R*100).min()))
        
        ax1.set_ylabel('% Failure')
        #ax1.legend()
        ax1.grid(linewidth=0.2)
        ax1.set_xlim(1,M-1)
        ax1.set_xticks(range(0,M,10))
        #ax1.set_xticklabels(range(1,(M+1),10))
        ax1.set_ylim(0,100)
        ax1.set_xlabel('Number of Customers')
        ax2.set_ylabel('Maximum THD (%)')
        ax2.legend(framealpha=1,bbox_to_anchor=(0, 1.4), loc='upper left', ncol=3)
        ax2.set_xlabel('Number of Customers')
        ax2.grid(linewidth=0.2)
        ax2.set_xlim(1,M-1)
        ax2.set_ylim(0,6)
        ax2.set_xticks(range(0,M,10))
        #ax2.set_xticklabels(range(1,(M+1),10))
        plt.tight_layout()
    plt.savefig('figs/'+str(cm)+'_Failure_THD.png')
    
    
          
    both = (perH[33].sum(axis=1))<(M*R*0.9)
    both=both[both]
    for j in [0,3,6]:
        if len(both)>j:
            c=1
            plt.figure('RSC='+str(f)+' Specific Harmonics'+str(j),figsize=(4.5, 6.5)) 
            for pl in both.index[j:(j+3)]:
                ax=plt.subplot(len(both.index[j:(j+3)]),1, c)
                ax.set_ylabel('% Failure')
                ax.text(.5,.8,'h='+str(pl),
                    horizontalalignment='left',
                    transform=ax.transAxes)
                for f in RSCs:
                    ax.plot(100-perH[f].loc[int(pl)][1:M]/R*100, label='RSC='+str(f),linestyle=styles[f], color=cols[f])
                ax.legend(framealpha=1,bbox_to_anchor=(0, 1.3), loc='upper left', ncol=3)
                ax.xaxis.set_major_formatter(FormatStrFormatter('% 1.0f'))
                plt.grid(linewidth=0.2)
                ax.set_xlim(1,M)
                ax.set_xticks(range(0,M,10))
                ax.set_ylim(0,100)
                ax.xaxis.set_major_formatter(FormatStrFormatter('% 1.0f'))
                ax.set_xlabel('Number of Customers')
                plt.tight_layout()
                c=c+1
            plt.savefig('figs/'+str(cm)+'_Specific_Harmonics1_'+str(j)+'.png')
    
    plt.figure(cm,figsize=(4, 3))
    
    for f in RSCs:
        M=n_evs ##--Number of EVs
        plt.plot(V_Min_Av[f], linestyle=styles[f], label='RSC='+str(f), linewidth=1, color=cols[f])
        plt.ylabel('Voltage (V)')
    plt.plot([0,M],[216,216],color='Black',linestyle=":", linewidth=0.5, label='Statutory Min')
    plt.xticks(ticks=range(0,M,10))
    plt.xlabel('Number of Customers')
    plt.xlim(0,M-1)
    plt.grid(linewidth=0.2)
    plt.legend(framealpha=1,bbox_to_anchor=(1, 1.2), loc='upper right', ncol=2)
    plt.tight_layout()
    plt.savefig('figs/'+str(cm)+'_Vmin.png')