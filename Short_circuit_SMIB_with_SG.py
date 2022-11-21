# -*- coding: utf-8 -*-
"""
Created on Wed Nov  2 23:37:16 2022
@author: Leonardo Marín (leovimar@gmail.com)
"""

import powerfactory as pf
import pandas as pd

app = pf.GetApplication()

#==============================================================================
# Function to set general parameters of short-circuit calculation
#==============================================================================
def shc_calc_settings(shc, obj, Rf, Xf):
    shc.iopt_allbus = 0      # Fault location: 0 = user selection, 1 = busbars and junction nodes, 2 = all busbars
    shc.shcobj      = obj    # Fault location: bus or terminal object
    shc.iopt_shc    = '3psc' # Short-circuit type: 3psc = 3-phase, 2psc = 2-phase, (more options -> check in SC calculation object in PF) 
    shc.Rf          = Rf     # (Ohm) fault resistance
    shc.Xf          = Xf     # (Ohm) fault reactance
    shc.iopt_asc    = 0

#==============================================================================
# Function to set the short-circuit calculation method
#==============================================================================     
def shc_calc_execution(shc, shc_method):
    shc.iopt_mde = shc_method
    if (shc_method < 4):
        shc.iopt_mde = shc_method
    elif (shc_method == 4):
        shc.iopt_mde = shc_method
        shc.iopt_meth60363 = 0
    elif (shc_method == 5):
        shc.iopt_mde = 4
        shc.iopt_meth60363 = 1
    else:
        raise IndexError('SC method out of range')
        
#==============================================================================
# Function to set how VSCs are modeled for the short-circuit calculation
#==============================================================================                 
def shc_vsc_settigs(vscs, shc_method, vsc_model):
    for vsc in vscs:
        if (shc_method == 0) or (shc_method == 1): # VDE 0102 or IEC 60909
            if (vsc_model == 0):   # VDE/IEC: no short-circuit contribution
                vsc.iNoShcContr = 1
                return 'VDE/IEC', 'No SC contribution'
            elif (vsc_model == 1): # VDE/IEC: static converter-fed drive 
                vsc.iNoShcContr = 0
                vsc.iconfed     = 1
                return 'VDE/IEC', 'Static converter-fed drive'
            elif (vsc_model == 2): # VDE/IEC: full size converter
                vsc.iNoShcContr = 0
                vsc.iconfed     = 0
                vsc.psutype     = 'fsce'
                return 'VDE/IEC', 'Full size converter'
            elif (vsc_model == 3): # VDE/IEC: equivalent SM
                vsc.iNoShcContr = 0
                vsc.iconfed     = 0
                vsc.psutype     = 'syme'
                return 'VDE/IEC', 'Equivalent SM'
            else:
                raise IndexError('VSC model out of range for VDE/IEC short-circuit method')
        elif (shc_method == 2):    # ANSI
            if (vsc_model == 0):   # ANSI: no short-circuit contribution 
                vsc.iNoShcContr = 1
                return 'ANSI', 'No SC contribution'
            elif (vsc_model == 1): # ANSI: short-circuit contribution
                vsc.iNoShcContr = 0 
                return 'ANSI', 'SC contribution'
            else:
                raise IndexError('VSC model out of range for ANSI short-circuit method')
        elif (shc_method == 3):    # Complete    
            if (vsc_model == 0):   # Complete: equivalent SM
                vsc.iShcModel = 0
                return 'Complete', 'Equivalent SM'
            elif (vsc_model == 1): # Complete: dynamic voltage support
                vsc.iShcModel = 1
                return 'Complete', 'Dymamic voltage support'
            elif (vsc_model == 2): # Complete: constant V
                vsc.iShcModel = 2
                return 'Complete', 'Constant V'
            elif (vsc_model == 3): # Complete: constant I
                vsc.iShcModel = 3
                return 'Complete', 'Constant I'
            elif (vsc_model == 4): # Complete: full size converter
                vsc.iShcModel = 4
                return 'Complete', 'Full size converter'
            else:
                raise IndexError('VSC model out of range for Complete short-circuit method')
          
#==============================================================================
# Function to get the resultant variables of the the short-circuit calculation
#============================================================================== 
def shc_get_results(obj, shc_method):
    if (shc_method == 0) or (shc_method == 1): # VDE 0102 or IEC 60909
        ip   = getattr(obj, 'm:ip')      # (kA) Peak short-circuit current
        Ikss = getattr(obj, 'm:Ikss')    # (kA) Initial symmetrical short-circuit current
        Iks  = '-'                       # (kA) Not available
        Ik   = getattr(obj, 'm:Ik')      # (kA) Steady-state short-circuit current
    elif (shc_method == 2): # ANSI
        ip   = getattr(obj, 'm:Ipeak_m') # (kA) Peak short-circuit current (instantaneous value)
        Ikss = getattr(obj, 'm:Isym_m')  # (kA) Symmetrical momentary (first cycle) short-circuit current (RMS)
        Iks  = '-'                       # (kA) Not available
        Ik   = getattr(obj, 'm:Isym_30') # (kA) 30 cycle symmetrical current
    elif (shc_method == 3): # Complete
        ip   = getattr(obj, 'm:ip')      # (kA) Peak short-circuit current
        Ikss = getattr(obj, 'm:Ikss')    # (kA) Initial symmetrical short-circuit current
        Iks  = getattr(obj, 'm:Iks')     # (kA) Not available                  
        Ik   = '-'                       # (kA) Steady-state short-circuit current  
    else: # IEC 61363 Standard
        ip   = getattr(obj, 'm:ip')      # (kA) peak short-circuit current
        Ikss = getattr(obj, 'm:Ikss')    # (kA) initial symmetrical short-circuit current
        Iks  = '-'                       # (kA) Not available
        Ik   = '-'                       # (kA) Not available 
    return ip, Ikss, Iks, Ik

# ==============================================================================
# Function to generate a table with results of short-circuit calculation
# ============================================================================== 
def shc_data(method, model, obj, ip, Ikss, Iks, Ik):
    method_col.append(method)
    model_col.append(model)
    obj_col.append(getattr(obj, 'loc_name')) # Object name
    ip_col.append(ip)                        # (kA) Peak short-circuit current
    Ikss_col.append(Ikss)                    # (kA) Initial symmetrical short-circuit current
    Iks_col.append(Iks)                      # (kA) Transient short circuit current
    Ik_col.append(Ik)                        # (kA) steady-state short-circuit current
                
#==============================================================================
# Script execution
#==============================================================================     






shc  = app.GetFromStudyCase('ComShc')
obj  = app.GetCalcRelevantObjects('BB_03.ElmTerm')[0]
vscs = app.GetCalcRelevantObjects('*.ElmVscmono') 

Rf = 0
Xf = 0

shc_calc_settings(shc, obj, Rf, Xf)

# Variable pre-allocation
method_col = []
model_col  = []
obj_col    = []
ip_col     = []
Ikss_col   = []
Iks_col    = []
Ik_col     = []
                
for i in range(1,6):
    shc_calc_execution(shc, i)
    app.PrintPlain('i = ' + str(i))
    if (i == 0) or (i == 1):
        for j in range(4):
            app.PrintPlain('j = ' + str(j))
            method, model = shc_vsc_settigs(vscs, i, j)
            shc.Execute()
            ip, Ikss, Iks, Ik = shc_get_results(obj, i)
            shc_data(method, model, obj, ip, Ikss, Iks, Ik)
    elif (i == 2):
        for j in range(2):
            app.PrintPlain('j = ' + str(j))
            method, model = shc_vsc_settigs(vscs, i, j)
            shc.Execute()
            ip, Ikss, Iks, Ik = shc_get_results(obj, i)
            shc_data(method, model, obj, ip, Ikss, Iks, Ik)
    elif (i == 3):
        for j in range(5):
            app.PrintPlain('j = ' + str(j))
            method, model = shc_vsc_settigs(vscs, i, j)
            shc.Execute()
            ip, Ikss, Iks, Ik = shc_get_results(obj, i)
            shc_data(method, model, obj, ip, Ikss, Iks, Ik)
    elif (i == 4):
        try: 
            method = 'IEC 61363 Standard'
            model  = '-'
            shc.Execute()
            ip, Ikss, Iks, Ik = shc_get_results(obj, i)
            shc_data(method, model, obj, ip, Ikss, Iks, Ik)
        except: 'Method %s is not possible for the current system' %method
    else:
        method = 'IEC 61363 EMT'
        model  = '-'
        shc.Execute()
        ip, Ikss, Iks, Ik = shc_get_results(obj, i)
        shc_data(method, model, obj, ip, Ikss, Iks, Ik)
   
# Creating a data frame with the results
data_frame              = pd.DataFrame()
data_frame['SC method'] = method_col
data_frame['VSC model'] = model_col
# data_frame['SC at']     = obj_col
data_frame['ip']        = ip_col
data_frame['ip%']       = 100*data_frame['ip']/data_frame.iloc[-1]['ip']
data_frame['Ikss']      = Ikss_col
data_frame['Ikss%']     = 100*data_frame['Ikss']/data_frame.iloc[-1]['Ikss']
data_frame['Iks']       = Iks_col
data_frame['Ik']        = Ik_col

# Rounding results to a certain number of decimal places
results = data_frame.apply(pd.to_numeric, errors = 'coerce').round(2).fillna(data_frame)

# Exporting results as csv
results.to_csv('D:\Projects\Grical\Results.csv', encoding = 'utf-8', index = False)

# Showing results on screen
results_to_visualize = results.to_string(index = False)
app.PrintPlain(results_to_visualize)







                
    
