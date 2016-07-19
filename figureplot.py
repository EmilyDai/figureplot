# -*- coding: utf-8 -*-
"""
Created on Thu Dec 17 19:46:22 2015

@author: hp-hp
"""

import sys,xdrlib
import os
import xlrd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
from scipy.interpolate import interp1d
from matplotlib.ticker import MultipleLocator

path=os.getcwd()
LineColor=['','black','red','darkblue']
ll=['none','-','--','-.',':']
#myfont=matplotlib.font_manager.FontProperties(fname='times.ttf',size=12)
def GetDataFromExcel(file,by_index=0):
    data=xlrd.open_workbook(file)
#    print dir(data)
    table=data.sheets()[by_index]
    #print table
    nrows=table.nrows
    #print nrows,'nrows'
    ncols=table.ncols
    #print ncols,'nclos'
    databig=np.zeros((ncols,nrows-1))
    #print databig
    #list_x=[]
    for colnum in range(ncols):
        listx=[]
        #global databig
        #print 'colnum=',colnum
        col=table.col_values(colnum)
        for rownum in range(1,nrows):
            #print'rownum=',rownum
            
            listx.append(col[rownum])
            #print listx
        databig[colnum]=listx        
    return databig
    
def GetSheetNum(file):
    data=xlrd.open_workbook(file)
    return len(data.sheets())


     
def GetFlag(file,by_index):
    data=xlrd.open_workbook(file)
    flag=data.sheets()[by_index].name
    return flag
    
def GetLabel(file,by_index):
    linelabel=[]
    data=xlrd.open_workbook(file)
    table=data.sheets()[by_index]
    nrows=table.nrows
    ncols=table.ncols
    row0=table.row_values(0)
    for i in range(1,ncols):
#        aa='$'+row0[i]+'$'
        aa=row0[i]
        linelabel.append(aa)
#    print linelabel
    return linelabel
            
    
def GenerateFigure(x,labellist,flag,xmin,xmax,legendlocation):
    plt.figure(figsize=(8,4))
    xnum=x.shape[1]
    xnumnew=xnum*7
    ax=plt.gca()
###########################画图啦##################################################
    for i in range(1,x.shape[0]):
        f=interp1d(x[0], x[i], kind='cubic')
        xnew = np.linspace(x[0][0], x[0][xnum-1], num=xnumnew, endpoint=True)
        plt.plot(xnew,f(xnew),ll[i],color=LineColor[i],linewidth=2.5)         
        lg=plt.legend(labellist,shadow='true')
        for t in lg.get_texts():
#            print dir(t)
                        
            t.set_font_properties('Times New Roman')
            t.set_fontweight('bold')
            t.set_fontsize(12)

    font = {'family' : 'Times New Roman',  
        'color'  : 'black',  
        'weight' : 'bold',  
        'size'   : 12,  
        }

    for tick in ax.xaxis.get_major_ticks():
#        print dir(tick.label1)

        tick.label1.set_font_properties('Times New Roman')
        tick.label1.set_weight('bold')   
        tick.label1.set_fontsize(12)
        
    for tick in ax.yaxis.get_major_ticks():
        
        tick.label1.set_font_properties('Times New Roman')
        tick.label1.set_weight('bold')
        tick.label1.set_fontsize(12)
    wrapper=ax.spines
    wrapper['top'].set_linewidth(2)
    wrapper['bottom'].set_linewidth(2)
    wrapper['left'].set_linewidth(2)
    wrapper['right'].set_linewidth(2)
    
    for tickline in ax.xaxis.get_ticklines():
        tickline.set_markeredgewidth(1.5)
        tickline.set_markersize(4)
    for tickline in ax.yaxis.get_ticklines():
        tickline.set_markeredgewidth(1.5)
        tickline.set_markersize(4)
    
    if (flag.upper()=='VSWR'):
        plt.xlabel("Freq (Ghz)",fontdict=font)
        plt.ylabel("VSWR",fontdict=font)
        ax.get_legend()._loc=legendlocation

        if(xmin!='default' and xmax!='default'):
            plt.xlim(xmin,xmax)
        

    if (flag.upper()[:4]=='GAIN'):
        plt.xlim(-90,90)
        ax.xaxis.set_major_locator(MultipleLocator(30)) 
        plt.xlabel("Theta (deg)",fontdict=font)
        plt.ylabel("Gain (dB)",fontdict=font)
        ax.get_legend()._loc=8
    
    if (flag.upper()=='S11'):
        plt.xlabel("Freq (GHz)",fontdict=font)
        plt.ylabel("S11 (dB)",fontdict=font)
        ax.get_legend()._loc=legendlocation
        if(xmin!='default' and xmax!='default'):
            plt.xlim(xmin,xmax)
    
    if (flag.upper()[0:9]=='RCSVSFREQ'):
        plt.xlabel("Freq (Ghz)",fontdict=font)
        plt.ylabel("Monostatic RCS (dB)",fontdict=font)
        ax.get_legend()._loc=2
        if(xmin!='default' and xmax!='default'):
            plt.xlim(xmin,xmax)
#    plt.savefig('01.png',dpi=75)

def SaveFigure(filename,flag,dpivalue):
    figurename=''
    filenamenew=filename.split('.')[0]
    figurename=filenamenew+'_'+flag
#    print figurename
    plt.savefig(figurename,dpi=dpivalue)
#    print plt.savefig.__doc__

        
if __name__=='__main__':
    filename='jx.xlsx'
#    sheetnum=1
    sheetnum=GetSheetNum(filename)
    for sheetnum in range(sheetnum):
        flag=GetFlag(filename,sheetnum)
        x=GetDataFromExcel(filename,sheetnum)
        labelx=GetLabel(filename,sheetnum)
#    lablex=["$xx$","$yy$","$zz$","$aa$"]
        flagx=GetFlag(filename,sheetnum)
        GenerateFigure(x,labelx,flagx)
        SaveFigure(filename,flag)
#        print GetFlag('vswr.xlsx',1)

    
