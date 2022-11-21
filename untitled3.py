# -*- coding: utf-8 -*-
"""
Created on Wed Nov  2 23:37:16 2022
@author: Leonardo Marín (leovimar@gmail.com)
"""
import powerfactory as pf
import pandas as pd

app = pf.GetApplication()

#==============================================================================
# Function definition
#==============================================================================

def shc_calc_settings(shc, obj, Rf, Xf):
    shc.iopt_allbus = 0      # Fault location: 0 = user selection, 1 = busbars and junction nodes, 2 = all busbars
    shc.shcobj      = obj    # Fault location: bus or terminal object
    shc.iopt_shc    = '3psc' # Short-circuit type: 3psc = 3-phase, 2psc = 2-phase, (more options -> check in SC calculation object in PF) 
    shc.Rf          = Rf     # (Ohm) fault resistance
    shc.Xf          = Xf     # (Ohm) fault reactance
    shc.iopt_asc    = 0

def shc_calc_execution(shc, shc_method):
    if (shc_method < 4):
        shc.iopt_mde = shc_method
        shc.Execute()
    elif (shc_method == 4):
        shc.iopt_mde = shc_method
        shc.iopt_meth60363 = 1
        shc.Execute()
    else:
        app.PrintPlain('Wrong short-circuit method')
      
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
    else:
        ip   = getattr(obj, 'm:ip')    # (kA) peak short-circuit current
        Ikss = getattr(obj, 'm:Ikss')  # (kA) initial symmetrical short-circuit current
        Iks  = '-'                     # (kA) Not available
        Ik   = '-'                     # (kA) Not available
    return ip, Ikss, Iks, Ik

# Function to put results together 
def shc_data(obj, ip, Ikss, Iks, Ik):
    obj_col.append(getattr(obj, 'loc_name')) # Object name
    ip_col.append(ip)                        # (kA) Peak short-circuit current
    Ikss_col.append(Ikss)                    # (kA) Initial symmetrical short-circuit current
    Iks_col.append(Iks)                      # (kA) Transient short circuit current
    Ik_col.append(Ik)                        # (kA) steady-state short-circuit current
    
    
# def shc_vsc_settigs(vsc, shc_method): # 0 = no shc
#     if (shc_method == 0) or (shc_method == 1):
#         for i in range(5):
#             if (i == 0):
#                 vsc.iNoShcContr = 1
#             elif (i == 1):
#                 vsc.iconfed = 1
#             else:
#                 vsc.psutype = i - 2
#     elif (shc_method == 2):
#         for i in range(3):
#             if (i == 0):
#                 vsc.iNoShcContr = 1
#             else:
#                 vsc.iNoShcContr = 0
#     elif (shc_method == 3):
#         for i in range(5):
#             vsc.iShcModeo = i         
                

def shc_vsc_settigs(vscs, shc_method): # 0 = no shc
    if (shc_method == 0) or (shc_method == 1):
        for i in range(5):
            if (i == 0):
                for vsc in vscs:
                    vsc.iNoShcContr = 1
                shc_calc_execution(shc, i)
                ip, Ikss, Iks, Ik = shc_get_results(obj, i)
                shc_data(obj, ip, Ikss, Iks, Ik)
            elif (i == 1):
                for vsc in vscs:
                    vscs.iconfed = 1
                shc_calc_execution(shc, i)
                ip, Ikss, Iks, Ik = shc_get_results(obj, i)
                shc_data(obj, ip, Ikss, Iks, Ik)
            else:
                for vsc in vscs:
                    vscs.psutype = i - 2
                shc_calc_execution(shc, i)
                ip, Ikss, Iks, Ik = shc_get_results(obj, i)
                shc_data(obj, ip, Ikss, Iks, Ik)
    elif (shc_method == 2):
        for i in range(3):
            if (i == 0):
                for vsc in vscs:
                    vscs.iNoShcContr = 1
                shc_calc_execution(shc, i)
                ip, Ikss, Iks, Ik = shc_get_results(obj, i)
                shc_data(obj, ip, Ikss, Iks, Ik)
            else:
                for vsc in vscs:
                    vscs.iNoShcContr = 0
                shc_calc_execution(shc, i)
                ip, Ikss, Iks, Ik = shc_get_results(obj, i)
                shc_data(obj, ip, Ikss, Iks, Ik)
    elif (shc_method == 3):
        for i in range(5):
                for vsc in vscs:
                    vscs.iShcMode = i
                shc_calc_execution(shc, i)
                ip, Ikss, Iks, Ik = shc_get_results(obj, i)
                shc_data(obj, ip, Ikss, Iks, Ik)
        
                
#==============================================================================
# Script execution
#==============================================================================     

shc = app.GetFromStudyCase('ComShc')  # Calling short-circuit command object (ComShc)
obj = app.GetCalcRelevantObjects('BB_03.ElmTerm')[0]
Rf  = 0
Xf  = 0

# Short-circuit methods
shc_methods = ['VDE 0102', 'IEC 60909', 'ANSI', 'Complete', 'IEC 61363 EMT']

# Variable pre-allocation
obj_col  = []
ip_col   = []
Ikss_col = []
Iks_col  = []
Ik_col   = []

vscs = app.GetCalcRelevantObjects('*.ElmVscmono')


shc_calc_settings(shc, obj, Rf, Xf)

for j in range(5):
    shc_vsc_settigs(vscs, j)



# for i in range(5):
#     shc_calc_execution(shc, i)
#     ip, Ikss, Iks, Ik = shc_get_results(obj, i)
#     shc_data(obj, ip, Ikss, Iks, Ik)

# # Creating a table from the results
# df              = pd.DataFrame()
# df['SC method'] = shc_methods
# df['SC at']     = obj_col
# df['ip']        = ip_col
# df['ip%']       = 100*df['ip']/df.iloc[-1]['ip']
# df['Ikss']      = Ikss_col
# df['Ikss%']     = 100*df['Ikss']/df.iloc[-1]['Ikss']
# df['Iks']       = Iks_col
# df['Ik']        = Ik_col

# # Round the results to a certain number of decimal places
# results = df.apply(pd.to_numeric, errors = 'coerce').round(2).fillna(df)

# app.PrintPlain('=============================================================')
# app.PrintPlain(results)



    

    

    





                
    
