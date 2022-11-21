# -*- coding: utf-8 -*-
"""
Created on Wed Nov  2 23:37:16 2022
@author: Leonardo Marín (leovimar@gmail.com)
"""

import powerfactory as pf
import pandas as pd
from Packages.Short_circuit.shc_functions import * 
from Packages.Dynamic_simulation.dynamic_simulation_functions import *
from timeit import default_timer as timer

app = pf.GetApplication()

app.EchoOff()

#==============================================================================
# Script execution
#==============================================================================     
shc  = app.GetFromStudyCase('ComShc') 
obj  = app.GetCalcRelevantObjects('BB_11.ElmTerm')[0]
vscs = app.GetCalcRelevantObjects('*.ElmVscmono') 

script = app.GetCurrentScript()

Rf = 0
Xf = 0

shc_gen_settings(shc, obj, Rf, Xf)

# Variable pre-allocation
method_col = []
model_col  = []
obj_col    = []
ip_col     = []
Ikss_col   = []
Iks_col    = []
Ik_col     = []
time_col   = []
                
for i in range(1, 6):
    if (i < 4):
        shc.iopt_mde = i
        app.PrintPlain('i = ' + str(i))
        if (i == 0) or (i == 1):
            for j in range(4):
                start = timer()
                app.PrintPlain('j = ' + str(j))
                method, model = shc_vsc_model_definition(vscs, i, j)
                shc.Execute()
                ip, Ikss, Iks, Ik = shc_get_results(obj, i)
                shc_data(method, model, obj, ip, Ikss, Iks, Ik, method_col, model_col, obj_col, ip_col, Ikss_col, Iks_col, Ik_col)
                end = timer()
                total_time = 1000*(end - start)
                time_col.append(total_time)
        elif (i == 2):
            for j in range(2):
                start = timer()
                app.PrintPlain('j = ' + str(j))
                method, model = shc_vsc_model_definition(vscs, i, j)
                shc.Execute()
                ip, Ikss, Iks, Ik = shc_get_results(obj, i)
                shc_data(method, model, obj, ip, Ikss, Iks, Ik, method_col, model_col, obj_col, ip_col, Ikss_col, Iks_col, Ik_col)
                end = timer()
                total_time = 1000*(end - start)
                time_col.append(total_time)       
        elif (i == 3):
            for j in range(5):
                start = timer()
                app.PrintPlain('j = ' + str(j))
                method, model = shc_vsc_model_definition(vscs, i, j)
                shc.Execute()
                ip, Ikss, Iks, Ik = shc_get_results(obj, i)
                shc_data(method, model, obj, ip, Ikss, Iks, Ik, method_col, model_col, obj_col, ip_col, Ikss_col, Iks_col, Ik_col)
                end = timer()
                total_time = 1000*(end - start)
                time_col.append(total_time)
    elif (i == 4):
        shc.iopt_mde       = 4
        shc.iopt_meth61363 = 0
        method = 'IEC 61363 Standard'
        model  = '-'
        try:
            start = timer()           
            shc.Execute()
            ip, Ikss, Iks, Ik = shc_get_results(obj, i)
            shc_data(method, model, obj, ip, Ikss, Iks, Ik, method_col, model_col, obj_col, ip_col, Ikss_col, Iks_col, Ik_col)
            end = timer()
            total_time = 1000*(end - start)
            time_col.append(total_time)
        except AttributeError: 
            app.PrintPlain('Method %s is not avaible for this system' %method)        
    else:
        start = timer()
        shc.iopt_mde       = 4
        shc.iopt_meth61363 = 1
        method = 'IEC 61363 EMT'
        model  = '-'
        shc.Execute()
        ip, Ikss, Iks, Ik = shc_get_results(obj, i)
        shc_data(method, model, obj, ip, Ikss, Iks, Ik, method_col, model_col, obj_col, ip_col, Ikss_col, Iks_col, Ik_col)
        end = timer()
        total_time = 1000*(end - start)
        time_col.append(total_time)

# Creating a data frame with the results
data_frame              = pd.DataFrame()
data_frame['SC method'] = method_col
data_frame['VSC model'] = model_col
#data_frame["Empty 1"]   = ""
# data_frame['SC at']     = obj_col
data_frame['ip']        = ip_col
data_frame['ip%']       = 100*data_frame['ip']/data_frame.iloc[-1]['ip']
data_frame["Empty 2"]   = ""
data_frame['Ikss']      = Ikss_col
data_frame['Ikss%']     = 100*data_frame['Ikss']/data_frame.iloc[-1]['Ikss']
data_frame["Empty 3"]   = ""
data_frame['Iks']       = Iks_col
#data_frame["Empty 4"]   = ""
data_frame['Ik']        = Ik_col
data_frame['Time']      = time_col

app.EchoOn()

com_inc = app.GetFromStudyCase('ComInc')
com_sim = app.GetFromStudyCase('ComSim')

clear_sim_events(app)

add_shc_event(app, obj, 0, 1, Rf, Xf) # fault_type: 0 = 3ph shc

time_step = 1e-3
time_stop = 10

setup_simulation(com_inc, com_sim, 'rms', 0.001, time_stop)

start = timer()
run_simulation(app, com_inc, com_sim)
end = timer()

total_time = 1000*(end - start)

app.PrintPlain(total_time)
Ik_sim = getattr(obj, 'm:Ishc') # (kA) Peak short-circuit current (instantaneous value)
app.PrintPlain(Ik_sim)

#method_col.append('RMS simulation')
#model_col.append('-')
#ip_col.append('-')
#Ikss_col.append('-')
#Iks_col.append('-')
#Ik_col.append('-')
#time_col.append(1000*total_time)

res_rms_sim = {'SC method': 'RMS simulation', 'VSC model': '-', 'ip': '-', 'ip%': '-', 'Empty 2': '', 'Ikss': '-', 'Ikss%': '-', 'Empty 3': '', 'Iks': '-', 'Ik': Ik_sim, 'Time': total_time}
  
data_frame = data_frame.append(res_rms_sim, ignore_index=True)


# Rounding results to a certain number of decimal places
results = data_frame.apply(pd.to_numeric, errors = 'coerce').round(2).fillna(data_frame)

# Exporting results as csv
results.to_csv('D:\Projects\Grical\Results.csv', encoding = 'utf-8', index = False)

# Showing results on screen
results_to_visualize = results.to_string(index = False)
app.PrintPlain(results_to_visualize)


