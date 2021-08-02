import numpy as np
import matplotlib.pyplot as plt
import openpyxl as xl
from openpyxl.styles import PatternFill, Color
import pandas as pd
import sys, os
import random
from random import randrange
sys.path.append("C:\Program Files (x86)\DIgSILENT\PowerFactory 15.1\python")

import math
import csv
import powerfactory as pf

# app.ResetCalculation()
##########################################################################################
##########################################################################################

class PowerFactorySim():

    def __init__(self, project_name= 'Project', studycase_name= 'Study Case'):

        # start powerfactory
        self.app = pf.GetApplication() 

        self.app.ClearOutputWindow()
        # user
        user = self.app.GetCurrentUser()
        
        # Activate project
        self.app.ActivateProject(project_name)
        self.project = self.app.GetActiveProject()
        
        # Activate study case
        self.study_case = self.project.GetContents('Study Cases\\'+ studycase_name)[0][0]
        #print(self.stuy_case)
        self.study_case.Activate()


##################################################################################
##################################### LoadFlow Simulation #########################
##################################################################################

    def prepare_loadflow(self, ldf_model= 'balanced', voltage_dep_load= 0):

        models = {'balanced':0, 'unbalanced':1, 'dc':2}
        
        self.ldf = self.app.GetFromStudyCase('*.ComLdf')
        self.ldf.iopt_net = models[ldf_model]
        self.ldf.iopt_pq = voltage_dep_load        

    def run_loadflow(self):
        return bool(self.ldf.Execute())


    def get_bus_voltages(self):

        #voltages = {}
        bus_names = []
        buse_voltages = []
        bus_angles = []
        # get all buss in net
        buses = self.app.GetCalcRelevantObjects('*.ElmTerm')

        var = []
        for bus in buses:

            #voltages[bus.loc_name] = bus.GetAttribute('m:u')  # 'm.u' is magnitude of pu voltage.
            bus_names.append(bus.GetAttribute('loc_name'))
            bus_angles.append(bus.GetAttribute('m:phiu'))
            buse_voltages.append(bus.GetAttribute('m:u'))
            var.append(bus.GetAttribute('m:u'))
            var.append(bus.GetAttribute('m:phiu'))

        return  bus_names, buse_voltages, bus_angles, var #,voltages

##################################################################################
##################################### Dynamic Simulation #########################
##################################################################################

    def prepare_dynamic_sim (self, monitored_variables, sim_type= 'rms',
                             start_time =0.0, step_size= 0.01, end_time = 5):

        # get result file
        self.res = self.app.GetFromStudyCase('*.ElmRes')
        # select result var to monitor
        for elm_name, var_names in monitored_variables.items():
            
            # get all Elements that match elm_name
            elements = self.app.GetCalcRelevantObjects(elm_name)

            # select variables to monitor for each element
            for element in elements:
                self.res.AddVars(element, *var_names)

        # Retrive initial condition and time domain sim
        self.inc = self.app.GetFromStudyCase('ComInc')
        self.sim = self.app.GetFromStudyCase('ComSim')

        # set simulation type (RMS / ins(for EMT)). also set start/end/step_size time.
        self.inc.iopt_sim = sim_type
        self.inc.tstart = start_time
        self.inc.dtgrd = step_size
        self.inc.dtout = 0.1
        self.sim.tstop = end_time
        # execute initial cond.
        return (self.inc.Execute())


    def Run_dynamic_sim(self):
        return (self.sim.Execute())


    def get_dynamic_results(self ,elm_name=None, var_name=None, offset=0):

        """
        Simulation has been executed by run_dynamic_sim. This function
        gets results assuming the variables were called when setting up the simulation.
        Arguments:
             elm_name - element name (ie bus or line)
             var_name - name of variable
             offset - from what time step to truncate data, e.g if pre contingency data is not required
        returns:
            time stamp and required variables as lists
        """
        # read the results and time steps and store them as lists
        time = []
        var_values = []

        # load results from file
        self.app.ResLoadData(self.res)

        # get number of rows (time steps) in the results file
        n_rows = self.app.ResGetValueCount(self.res, 0)
        # get network element of interest
        element = self.app.GetCalcRelevantObjects(elm_name)[0]
        # find column in results file which holds the result of interest
        col_index = self.app.ResGetIndex(self.res, element, var_name)
            
        for i in range(0, n_rows - offset):
            time.append(self.app.ResGetData(self.res, i+offset, -1)[1])
            var_values.append(self.app.ResGetData(self.res, i+ offset, col_index)[1])
        return time, var_values



    def get_voltage_scan(self):

        """
        Function to return voltage scan results from previous RMS simulation
        """

        scan_folder = self.app.GetFromStudyCase('IntScn')
        fault = scan_folder.GetContents('*.ScnFrt')[0]
        num_of_violations = fault.GetNumberOfViolations()
        time_stamp = []
        for i in range(1, num_of_violations+1):
            time_stamp.append(fault.GetViolationTime(i))
        return num_of_violations, time_stamp

###################################################################################################          

    def set_all_loads_pq(self, p_load, q_load, scale_factor=None):

        """
        Function to set all loads in the system. If loads need to be scaled from 
        nominal system values set scale_factor.
        The arguments of this method are two dictionaries, one for active and one for reactive
        power. The keys of these dictionaries are load names.
        """

        loads = self.app.GetCalcRelevantObjects('*.ElmLod')

        if scale_factor is not None:
            for key in p_load:
                p_load[key] *= scale_factor
            for key in q_load:
                q_load[key] *= scale_factor
            for key in p_load:
                p_load[key] = int(p_load[key])
            for key in q_load:
                q_load[key] = int(q_load[key])

        for load in loads:
            load.plini = p_load[load.loc_name]
            load.qlini = q_load[load.loc_name]


    def get_all_loads_pq(self):

        """
        Function to return all system loads under current state
        """
        loads = self.app.GetCalcRelevantObjects('*.ElmLod')
        p_base = {}
        q_base = {}

        for load in loads:
            p_base[load.loc_name] = load.plini
            q_base[load.loc_name] = load.qlini

        return p_base, q_base
##########################################################################################

    def toggle_out_of_service(self, elm_name):
        #collect all elements that match elm_name
        elms = self.app.GetCalcRelevantObjects(elm_name)
        # if elment is out of service, switch to in service, else switch to out of service
        for elm in elms:
            elm.outserv = 1- elm.outserv


    def toggle_switches(self, elm_name):
        #collect all elements that match elm_name
        elms = self.app.GetCalcRelevantObjects(elm_name)
        #collect all switches
        sws = self.app.GetCalcRelevantObjects('*.StaSwitch')
        #find swiches corresponding to each elm and toggle them
        for elm in elms:
            cubs = elm.GetCubicle(0) + elm.GetCubicle(1)
            for sw in sws:
                if sw.fold_id in cubs:
                    sw.on_off = 1-  sw.on_off
                    
###########################################################################################

    def return_Motor_Buses(self):
        Motors = self.app.GetCalcRelevantObjects('*.ElmAsm')
        motor_buses = []
        for motor in Motors:
            if motor.outserv == 0:
                motor_buses.append(motor.GetNode(0))
        return motor_buses


    def return_inserv_Motors(self):
        Motors = self.app.GetCalcRelevantObjects('*.ElmAsm')
        motors = []
        for motor in Motors:
            if motor.outserv == 0:
                motors.append(motor)
        return motors

############################################################################################
####################################### Dynamic Scenarios ##################################
############################################################################################
   
    def enable_short_circuits(self):
        """
        since by default lines are unavailable, function to change all to available
        """
        lines = self.app.GetCalcRelevantObjects('*.ElmLne')
        for line in lines:
            if line.ishclne == 0:
                line.ishclne = 1
        print("All lines available for short circuit")


    def create_short_circuit(self, target_name,
                             time, fault_type,
                             duration=None,
                             fault_Res = 0,
                             Line_loc=False , name='sc'):

        """
        Creates short circuit event in the events folder
        if duration is specified, a clearing event will also be created
        Arguments:
            target_name - name of line
            time - start time of short circuit
            duration - duration before clearing
            fault_type - 0 for three phase, fault codes can be found in PF
        """

        # get element where the short circuit will be made
        target = self.app.GetCalcRelevantObjects(target_name)[0]
        if Line_loc:
            target.fshcloc = Line_loc
        # get events folder from active study case
        evt_folder = self.app.GetFromStudyCase('IntEvt')
        # create an empty event of type EvtShc
        evt_folder.CreateObject('EvtShc', name)
        # get the newly created event
        sc = evt_folder.GetContents(name + '.EvtShc')[0][0]
        # set time, target and type of short circuit (ie single or three phase)
        sc.p_target = target
        sc.time = time
        sc.i_shc = fault_type
        sc.R_f = fault_Res
        # set clearing event if required
        if duration is not None:
            # create an empty event for the clearing event
            evt_folder.CreateObject('EvtShc', name + '_clear')
            # get the new event
            scc = evt_folder.GetContents(name + '_clear' + '.EvtShc')[0][0]
            scc.time = time + duration
            scc.p_target = target
            scc.i_shc = 4


    def delete_short_circuit(self, name='sc'):

        """
        For running multiple events in a loop, old events need to be deleted
        and new ones re-initialised. This function deletes all events in
        events folder assuming they were created from the method in this class
        """
        # get the events folder
        evt_folder = self.app.GetFromStudyCase('IntEvt')
        # find short circuit events and clear if they exist
        sc = evt_folder.GetContents(name + '.EvtShc')[0]
        scc = evt_folder.GetContents(name + '_clear' + '.EvtShc')[0]
        if sc:
            sc[0].Delete()
        if scc:
            scc[0].Delete()

#################################################

    def create_Switch_Event(self, target_name,
                            time,
                            action = 'open',
                            name='se'): 


        # get element where the short circuit will be made
        target = self.app.GetCalcRelevantObjects(target_name)[0]
        
        # get events folder from active study case
        evt_folder = self.app.GetFromStudyCase('IntEvt')
        # create an empty event of type EvtShc
        evt_folder.CreateObject('EvtSwitch', name)
        # get the newly created event
        se = evt_folder.GetContents(name + '.EvtSwitch')[0][0]
        # set time, target and type of short circuit (ie single or three phase)
        se.p_target = target
        se.time = time
        
        if action == 'open':
            se.i_switch = 0
        elif  action == 'close':
            se.i_switch = 1
        
    
    def delete_Switch_Event(self, name='se'):

        # get the events folder
        evt_folder = self.app.GetFromStudyCase('IntEvt')
        # find switch events and clear if they exist
        se = evt_folder.GetContents(name + '.EvtSwitch')[0]
        
        se[0].Delete()
        



######################################################################################
####################################### simulation ###################################
######################################################################################

fault_loc = ['EN902-1.ElmLne', 'MY908-1.ElmLne', 'Line(24).ElmLne', 'Line Route.ElmLne'] 
fault_dur = [0.2, 0.2, 0.3, 0.5] 


Monitored_variables = {
                       '*.ElmAsm': ['n:u:bus1', 'm:phiui:bus1', 's:xspeed']
                       }

object1 = PowerFactorySim(project_name= 'Real(Korasan)500', studycase_name= 'Study Case')
#object1.app.Show()

Lines = object1.app.GetCalcRelevantObjects('*.ElmLne')

# enable lines short circuit
object1.enable_short_circuits()

Motor_Terminals = [
  '*\asadabad.ElmTerm','*\brdakn.ElmTerm', '*\beihagh.ElmTerm', '*\birjand.ElmTerm', '*\bushruye.ElmTerm',
   '*\dargaz.ElmTerm', '*\davarzan.ElmTerm', '*\dolatabad.ElmTerm', '*\fldkhsn.ElmTerm',
  '*\ghaen.ElmTerm', '*\gholaman.ElmTerm', '*\golbahar.ElmTerm', '*\Terminal_Golshahr.ElmTerm',
   '*\gonabad.ElmTerm', '*\hajiabad.ElmTerm', '*\istghazd.ElmTerm', '*\jlgrkh.ElmTerm', '*\khaf.ElmTerm',
  '*\khajerabi63.ElmTerm', '*\kohsangi63.ElmTerm', '*\mashhad.ElmTerm', '*\machnelec.ElmTerm',
   '*\mehrgan.ElmTerm', '*\nehbandan.ElmTerm','*\neishabur.ElmTerm', '*\nmyshgh.ElmTerm',
  '*\Terminal_Pardis.ElmTerm', '*\rashtkhr.ElmTerm','*\sabzvar.ElmTerm', '*\sahel.ElmTerm',
   '*\sahlabad.ElmTerm', '*\salehabad.ElmTerm', '*\sangan.ElmTerm', '*\sangbast.ElmTerm',
   '*\sarakhs.ElmTerm','*\sarayan.ElmTerm', '*\sarbishe.ElmTerm','*\sedeh.ElmTerm', '*\sim.ElmTerm',
   '*\solat.ElmTerm','*\tabas.ElmTerm', '*\toossG1.ElmTerm', '*\trbtjam(2).ElmTerm', '*\abousaed.ElmTerm',
  '*\ghazished.ElmTerm', '*\Torbatshed.ElmTerm','*\Kashmaeshed.ElmTerm', '*\Ataaeshed.ElmTerm',
   '*\Bjnoordshed.ElmTerm', '*\jajramshed.ElmTerm', '*\Dashtjovinshad.ElmTerm',
  '*\shirvanshed.ElmTerm', '*\Soltanabaadshed.ElmTerm', '*\feizabaadshed.ElmTerm', '*\Ferdos.ElmTerm',
   '*\ghuchanshed.ElmTerm','*\Tous63-2shed.ElmTerm', '*\tous63.ElmTerm',
  '*\farimanshed.ElmTerm','*\Terminal.ElmTerm', '*\tybaadshed.ElmTerm'
]


Motors = object1.return_inserv_Motors()



# run multiple rms simulation 
with open('result_rms.csv', mode='w', newline='') as csv_file:
    csvwriter = csv.writer(csv_file)
    Vars = ['t']
    for motor in Motors:
        Vars.append('u_'  + motor.GetAttribute('loc_name'))
        Vars.append('phi_'+ motor.GetAttribute('loc_name'))
        Vars.append('speed_'+ motor.GetAttribute('loc_name'))
    for motor in Motors:
        Vars.append('index_'+ motor.GetAttribute('loc_name'))
    csvwriter.writerow(Vars)

    for run in range(1):
        # define Event
        
        #fault_line = Lines[random.randint(0, len(Lines)-1)].GetAttribute('loc_name')+'.ElmLne'
        fault_line = Lines[22].GetAttribute('loc_name')+'.ElmLne'
        object1.create_short_circuit(target_name= fault_line,
                             time=1.3, fault_type= 0,
                             duration= 0.2 ,
                             fault_Res =  0,
                             Line_loc = 25, name='sc')

        
        object1.create_Switch_Event(target_name = Motors[40].GetAttribute('loc_name')+'.ElmAsm',
                            time = 1.7,
                            action = 'open',
                            name='ls1')
        
        '''
        object1.create_Switch_Event(target_name = Motors[40].GetAttribute('loc_name')+'.ElmAsm',
                            time = 1.9,
                            action = 'open',
                            name='ls2')
        
        object1.create_Switch_Event(target_name = Motors[59].GetAttribute('loc_name')+'.ElmAsm',
                            time = 2.1,
                            action = 'open',
                            name='ls3')
        
        object1.create_Switch_Event(target_name = Motors[4].GetAttribute('loc_name')+'.ElmAsm',
                            time = 2.3,
                            action = 'open',
                            name='ls4')
        
        object1.create_Switch_Event(target_name = Motors[1].GetAttribute('loc_name')+'.ElmAsm',
                            time = 2.5,
                            action = 'open',
                            name='ls5')
        
        object1.create_Switch_Event(target_name = Motors[14].GetAttribute('loc_name')+'.ElmAsm',
                            time = 2.7,
                            action = 'open',
                            name='ls6')
        
        
        object1.create_Switch_Event(target_name = Motors[27].GetAttribute('loc_name')+'.ElmAsm',
                            time = 2.9,
                            action = 'open',
                            name='ls7')
    
        object1.create_Switch_Event(target_name = Motors[17].GetAttribute('loc_name')+'.ElmAsm',
                                time = 3.1,
                                action = 'open',
                                name='ls8')

        object1.create_Switch_Event(target_name = Motors[32].GetAttribute('loc_name')+'.ElmAsm',
                                time = 3.3,
                                action = 'open',
                                name='ls9')
        
        
        object1.create_Switch_Event(target_name = Motors[38].GetAttribute('loc_name')+'.ElmAsm',
                                time = 3.5,
                                action = 'open',
                                name='ls10')
        '''    
            
        object1.prepare_dynamic_sim(
                                    monitored_variables = Monitored_variables, sim_type= 'rms',
                                    start_time =0.0, step_size= 0.02, end_time = 6)

        # Run rms sim
        object1.Run_dynamic_sim()

        
        # delete Event
        object1.delete_short_circuit(name='sc')
        
        
        object1.delete_Switch_Event(name='ls1')
        '''
        object1.delete_Switch_Event(name='ls2')
        object1.delete_Switch_Event(name='ls3')
        object1.delete_Switch_Event(name='ls4')
        object1.delete_Switch_Event(name='ls5')
        object1.delete_Switch_Event(name='ls6')
        object1.delete_Switch_Event(name='ls7')
        object1.delete_Switch_Event(name='ls8')
        object1.delete_Switch_Event(name='ls9')
        object1.delete_Switch_Event(name='ls10')        
        '''

        # get results
        buses = object1.app.GetCalcRelevantObjects('*.ElmTerm')
        bus_voltages =[]
        bus_angles = []
        Motor_speed = []

        for motor in Motors:
            t, voltage = object1.get_dynamic_results(motor.GetAttribute('loc_name')+ '.ElmAsm', 'n:u:bus1', offset=0)
            bus_voltages.append(voltage)
            _, angle   = object1.get_dynamic_results(motor.GetAttribute('loc_name')+ '.ElmAsm', 'm:phiui:bus1', offset=0)
            bus_angles.append(angle)
            _, speed   = object1.get_dynamic_results(motor.GetAttribute('loc_name')+ '.ElmAsm', 's:xspeed', offset=0)
            Motor_speed.append(speed)


        #object1.app.PrintPlain(bus_voltages)
        # write rms results to csv file
        var = [t]
        for i in range(len(Motors)):
            var.append(bus_voltages[i])
            var.append(bus_angles[i])
            var.append(Motor_speed[i])

        for i in range(len(Motors)):
            var.append((1- np.tan(np.array(np.deg2rad(bus_angles[i])))))

        #object1.app.PrintPlain(var)
        for row in zip(*var):
            csvwriter.writerow(row)