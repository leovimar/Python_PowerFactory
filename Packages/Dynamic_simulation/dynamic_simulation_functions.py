# -*- coding: utf-8 -*-
"""
Created on Tue Nov  8 19:44:17 2022
@author: Leonardo Marín (leovimar@gmail.com)
"""

# import powerfactory as pf

# app = pf.GetApplication()

def setup_simulation(com_inc, com_sim, sim_type, time_step, time_stop): # sim_type: 'rms' or 'ins'
    com_inc.iopt_sim   =  sim_type
    com_inc.iopt_show  =  0
    com_inc.iopt_adapt =  0
    com_inc.start      = -0.1   
    com_sim.tstop      =  time_stop
    if (sim_type == 'rms'):
        com_inc.dtgrd = time_step
    else:
        com_inc.dtempt = time_step
    
def run_simulation(app, com_inc, com_sim):
    app.EchoOff()
    com_inc.Execute()
    app.EchoOn()
    com_sim.Execute()
  
def clear_sim_events(app):
    event_folder = app.GetFromStudyCase("Simulation Events/Fault.IntEvt")
    events       = event_folder.GetContents()
    for event in events:
        app.PrintPlain(event.loc_name)
        event.Delete()
  
def add_shc_event(app, obj, fault_type, time, Rf, Xf):
    event_folder   = app.GetFromStudyCase("Simulation Events/Fault.IntEvt")
    event          = event_folder.CreateObject("EvtShc", obj.loc_name)
    event.p_target = obj
    event.time     = time
    event.i_shc    = fault_type
    event.R_f      = Rf
    event.X_f      = Xf
    