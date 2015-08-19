import neo
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import xlsxwriter
import os
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.font_manager import FontProperties
import random
import brewer2mpl
import seaborn as sns  
rcdefsns = plt.rcParams.copy()
import sys
colors=brewer2mpl.get_map('Set1', 'Qualitative',9).mpl_colors

args=sys.argv
folder_name=args[1]
outputname=args[2]

for m,n,o in os.walk(folder_name):
    abfs=o
    break

fig_pdf=PdfPages('%s.pdf' %(outputname))
writer = pd.ExcelWriter('%s.xlsx' %(outputname), engine='xlsxwriter')    
for abf in sorted(abfs):
    print abf
    filein=neo.io.AxonIO(folder_name+abf)
    blocks=filein.read_block()
    allsweeps=pd.DataFrame(index=['Sweeps','Time','Amps_80','Amps_-80','T'])
    fig,(ax1,ax2) = plt.subplots(nrows=1, ncols=2)
    fig.suptitle(abf)

    for sweep in range(len(blocks.segments)):
        df=pd.DataFrame(data=zip(np.asarray(blocks.segments[sweep].analogsignals[0]),
                                 np.asarray(blocks.segments[sweep].analogsignals[1]),
                                 np.asarray(blocks.segments[sweep].analogsignals[2])),
                        columns=['Current','Voltage','T'])
        try:
            q=df[(df['Voltage']>80.) & (df['Voltage']<81.)].iloc[0]
            r=df[(df['Voltage']>-81.) & (df['Voltage']<-80.)].iloc[0]
            allsweeps[sweep]=[sweep,sweep*3,q['Current'],r['Current'],r['T']]
        except:
            pass
    allsweeps=allsweeps.T
    ax1.plot(allsweeps['Time'],allsweeps['Amps_80'],color='blue',label='+80mV')
    ax1.plot(allsweeps['Time'],allsweeps['Amps_-80'],color='red',label='-80mV')
    ax1.legend()
    ax1.set_xlabel('Time(s)')
    ax1.set_ylabel('Current (uA)')

    copy=allsweeps.copy()
    copy['T'][copy['T']>10.3]=np.nan
    index=copy['T'].first_valid_index()
    try:
        allsweeps['Normalized_current']=allsweeps['Amps_80']/max(allsweeps['Amps_80'])
        ax2.plot(allsweeps.ix[:index]['T'],allsweeps.ix[:index]['Normalized_current'],color='green')
        ax2.set_xlabel('Temperature (deg C)')
        ax2.set_ylabel('Normalized current')
        ax2.set_ylim([0,1])
        ax2.set_xlim([10,40])
        allsweeps[:index].to_excel(writer,sheet_name='%s'%(abf.replace('.abf','')))

    except:
        pass

    fig_pdf.savefig()
    plt.close(fig)

fig_pdf.close()
writer.close()
