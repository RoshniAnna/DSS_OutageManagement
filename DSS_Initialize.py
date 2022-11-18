"""
In this file the sectionalizing and tie switch details are specified. 
The DER details are also specified (in particular the PV generators and BESS info). 
Also the path for the DSS file containing the circuit information is specified.
The final DSS circuit which will be used by the environment is created.
"""

import os
import networkx as nx
from  DSS_CircuitSetup import*

#------------------- User defined inputs to modify the standard test network ---------------------------------

sectional_swt=[{'no':1,'line':'632670'},
               {'no':2,'line':'671692'}]


tie_swt=[{'no':1,'from node':'646','from conn':'.2.3', 'to node':'684','to conn':'.1.3', 'length':2000,'code':'601'},
          {'no':2,'from node':'633','from conn':'.1.2.3','to node':'692','to conn':'.1.2.3','length':2000,'code':'601'}]

# Generic generators
#Jacobs, Nicholas, Shamina Hossain-McKenzie, and Adam Summers. "Modeling data flows with network calculus in cyber-physical systems: enabling feature analysis for anomaly detection applications." Information 12.6 (2021): 255.
generators=[{'no':1, 'bus':'645', 'numphase':1, 'phaseconn':'.2', 'size':20, 'kV':2.4, 'Gridforming':'No'},
            {'no':2, 'bus':'645', 'numphase':1, 'phaseconn':'.3', 'size':20, 'kV':2.4, 'Gridforming':'No'},
            {'no':3, 'bus':'634', 'numphase':3, 'phaseconn':'.1.2.3', 'size':1000, 'kV':0.48, 'Gridforming':'Yes'},
            {'no':4, 'bus':'684', 'numphase':1, 'phaseconn':'.1', 'size':50, 'kV':2.4, 'Gridforming':'No'},
            {'no':5, 'bus':'684', 'numphase':1, 'phaseconn':'.3', 'size':50, 'kV':2.4, 'Gridforming':'No'},
            {'no':6, 'bus':'680', 'numphase':3, 'phaseconn':'.1.2.3', 'size':1000, 'kV':4.16, 'Gridforming':'Yes'},
            {'no':7, 'bus':'675', 'numphase':3, 'phaseconn':'.1.2.3', 'size':500, 'kV':4.16,'Gridforming':'No'}]

substatn_id='sourcebus'
#------------ Define the network with additions of DER, BESS and switches -------------------------------------
def initialize():       
    FolderName=os.path.dirname(os.path.realpath("__file__"))
    DSSfile=r""+ FolderName+ "\IEEE13Nodeckt.dss"
    DSSCktobj=CktModSetup(DSSfile,sectional_swt,tie_swt,generators) # initially the sectionalizing switches close and tie switches open
    DSSCktobj.dssSolution.Solve() #solving snapshot power flows
    if DSSCktobj.dssSolution.Converged:
       conv_flag=1
    else:
       conv_flag=0    
    G_init=graph_struct(DSSCktobj)
    return DSSCktobj,G_init,conv_flag

DSSCktobj,G_init,conv_flag= initialize()   

#---------------------Create a dictionary with name of generator element and corresponding buses and also one with blackstart capability---------------------------

Generator_Buses={} # list of dictionary create for opendss element extraction
Generator_BlackStart={}
i= DSSCktobj.dssCircuit.Generators.First
while i>0:
      elemName =  DSSCktobj.dssCircuit.ActiveCktElement.Name
      bus_connectn = DSSCktobj.dssCircuit.ActiveCktElement.BusNames[0].split('.')[0]
      # phases = DSSCktobj.dssCircuit.ActiveCktElement.NumPhases
      Generator_Buses[elemName]=bus_connectn
      num=int(elemName[-1])-1
      if generators[num]['Gridforming'] == 'Yes':
         Generator_BlackStart[elemName]=1
      else:
         Generator_BlackStart[elemName]=0 
      i = DSSCktobj.dssCircuit.Generators.Next

#----------------- Create a dictionary with the name of loads and corresponding bus ---------------------
Load_Buses={}
i= DSSCktobj.dssCircuit.Loads.First
while i>0:
      elemName =  DSSCktobj.dssCircuit.ActiveCktElement.Name
      bus_connectn = DSSCktobj.dssCircuit.ActiveCktElement.BusNames[0].split('.')[0]
      # phases = DSSCktobj.dssCircuit.ActiveCktElement.NumPhases
      Load_Buses[elemName]=bus_connectn
      i = DSSCktobj.dssCircuit.Loads.Next
      
      
#------------------ List of the network buses and bus connections -------------------------------------      
node_list=list(G_init.nodes())
edge_list=list(G_init.edges()) #the fixed set of edges
nodes_conn=[]
for bus in node_list:
    nodes_conn.append(Bus(DSSCktobj,bus).nodes)
    

#-----------For each bus creating a generator element, load element, blackstart indicator------------
# gen_buses = np.array(list(Generator_Buses.values())) # all the generator buses
# gen_elems=  list(Generator_Buses.keys()) # all the generator elements
# load_buses = np.array(list(Load_Buses.values())) #all the load buses
# load_elems = list(Load_Buses.keys()) #all the laod elements
# Bus_Info=[]

# for n in node_list:   
#     blackstart_flag=0
#     gen_names=[gen_elems[x]  for x in  np.where(gen_buses == n)[0]]
#     load_names=[load_elems[y]  for y in  np.where(load_buses == n)[0]]
#     if (len(gen_names)!=0) and (gen_names[-1].split('.')[0]=='Storage'): #if there is storage then it is blackstart DER
#        blackstart_flag=1
#     Bus_Info.append({'Bus_id':n, 'Load_names':load_names, 'Generator_names':gen_names, 'Black_start':blackstart_flag})   
#     # Use this information to encode the graph

gen_buses = np.array(list(Generator_Buses.values())) # all the generator buses
gen_elems=  list(Generator_Buses.keys()) # all the generator elements
Gen_Info ={}
for n in node_list:
    blackstart_flag=0
    gen_names=[gen_elems[x]  for x in  np.where(gen_buses == n)[0]]
    if (len(gen_names)!=0):
        for g in gen_names:
            blackstart_flag= blackstart_flag + Generator_BlackStart[g] 
        Gen_Info[n]= {'Generators':gen_names, 'Blackstart':blackstart_flag}
    
     
     
'''
  
#----Parameters to evaluate min and max voltage at all nodes in all buses for time series simulation
Vmin=100
Vmax=0  
min_busmark=max_busmark=''  
min_phasemark=max_phasemark=0

  
#-------------- Defining new monitors at each PV element -----------------------------------
i = DSSCktobj.dssCircuit.Generators.First
while i>0:
      elem= DSSCktobj.dssCircuit.ActiveElement.Name          
      DSSCktobj.dssText.command='New Monitor.PV_pq_' + str(i) + ' element=' + elem + ' terminal=1 mode=1 ppolar=no'
      i = DSSCktobj.dssCircuit.Generators.Next


# -------------Solve time series      
DSSCktobj.dssText.command='Set mode=daily'
DSSCktobj.dssText.command='Set stepsize=1h' 
DSSCktobj.dssText.command='Set number=1' # this means that each solve command corresponds to one step
t=0
t_tot=24
 
# # Different switching combinations and check 
# action =[1,1,0,0]

# #-------------NEED TO IMPLEMENT ACTION HERE AFTER DEFINING MODE
# i=DSSCktobj.dssCircuit.SwtControls.First
# while (i>0):
#     # Swobj=DSSCktobj.dssCircuit.SwtControls.SwitchedObj
#     # DSSCktobj.dssCircuit.SetActiveElement(Swobj)
#     if action[i-1]==0:
#         # DSSCktobj.dssCircuit.ActiveCktElement.Open(1,0)
#         DSSCktobj.dssText.command=('Swtcontrol.' + DSSCktobj.dssCircuit.SwtControls.Name+ '.Action=o')
#         # DSSCktobj.dssText.command='open ' + Swobj +' term=1'       #switching the line open
#     else:
#         # DSSCktobj.dssCircuit.ActiveCktElement.Close(1,0)
#         DSSCktobj.dssText.command=('Swtcontrol.'+ DSSCktobj.dssCircuit.SwtControls.Name+ '.Action=c')
#         # DSSCktobj.dssText.command='close ' + Swobj +' term=1'      #switching the line close
#     i=DSSCktobj.dssCircuit.SwtControls.Next      
# DSSCktobj.dssSolution.Solve()


while t< t_tot:
 
#-----NEED TO IMPLEMENT SWITCH CHANGE FOR EACH TIME INSTANCE    
      i=DSSCktobj.dssCircuit.SwtControls.First
      while (i>0):
         Swobj=DSSCktobj.dssCircuit.SwtControls.SwitchedObj
         DSSCktobj.dssCircuit.SetActiveElement(Swobj)
         if action[i-1]==0:
             DSSCktobj.dssText.command='open ' + Swobj +' term=1'       #switching the line open
         else:
             DSSCktobj.dssText.command='close ' + Swobj +' term=1'      #switching the line close
         i=DSSCktobj.dssCircuit.SwtControls.Next  

      DSSCktobj.dssSolution.Solve()
      # ----Min and Max voltage with bus location and phase at which it occurs
      for i in range(len(node_list)):
          bus=node_list[i]
          V_bus=Bus(DSSCktobj,bus).Vmag
          for phase_co in nodes_conn[i]:
              if (V_bus[phase_co-1]<Vmin):
                  Vmin=V_bus[phase_co-1]
                  min_busmark=bus
                  min_phasemark=phase_co
              if (V_bus[phase_co-1]>Vmax):
                  Vmax=V_bus[phase_co-1]
                  max_busmark=bus
                  max_phasemark=phase_co

      print('The max voltage in pu is:{} and it occurs at bus {}.{}'.format(Vmax,max_busmark,max_phasemark))
      print('The min voltage in pu is:{} and it occurs at bus {}.{}'.format(Vmin,min_busmark,min_phasemark))
      
      # AllSwitches=[]
      # i=DSSCktobj.dssCircuit.SwtControls.First
      # while i>0:
      #     name=DSSCktobj.dssCircuit.SwtControls.Name
      #     line=DSSCktobj.dssCircuit.SwtControls.SwitchedObj
      #     br_obj=Branch(DSSCktobj,line)
      #     from_bus=br_obj.bus_fr
      #     to_bus=br_obj.bus_to
      #     DSSCktobj.dssCircuit.SetActiveElement(line)
      #     if(DSSCktobj.dssCircuit.ActiveCktElement.IsOpen(1,0)):
      #         sw_status=0
      #     else:
      #         sw_status=1
      #     AllSwitches.append({'switch name':name,'edge name':line, 'from bus':from_bus.split('.')[0], 'to bus':to_bus.split('.')[0], 'status':sw_status})
      #     i=DSSCktobj.dssCircuit.SwtControls.Next  
      # print(AllSwitches)
      t=t+1


# #----------- Export PV and Storage monitors (may increase overhead)(if you want to see excel data--uncomment)
# i=DSSCktobj.dssCircuit.Monitors.First
# while i>0:
#       name_mon= DSSCktobj.dssCircuit.Monitors.Name
#       DSSCktobj.dssText.command=('Export monitor ' + name_mon)   
#       i=DSSCktobj.dssCircuit.Monitors.Next    


# # using the COM interface with  AllVariableNames, AllVariableValues of elements but requires extraction at each time step
# DSSCktobj.dssCircuit.SetActiveClass('Storage')
# i = DSSCktobj.dssCircuit.ActiveClass.First
# while i>0:
#         elem= DSSCktobj.dssCircuit.ActiveClass.Name
#         Var_Names= DSSCktobj.dssCircuit.ActiveCktElement.AllVariableNames
#         Var_Vals= DSSCktobj.dssCircuit.ActiveCktElement.AllVariableValues
#         i = DSSCktobj.dssCircuit.ActiveClass.Next
# #- COM interface for extracting the storage SOC, kwh

'''  
#-----------End of simulation , now some code snippet for checking switches, voltages, currents
# Check switches status

AllSwitches=[]
i=DSSCktobj.dssCircuit.SwtControls.First
while i>0:
    name=DSSCktobj.dssCircuit.SwtControls.Name
    line=DSSCktobj.dssCircuit.SwtControls.SwitchedObj
    br_obj=Branch(DSSCktobj,line)
    from_bus=br_obj.bus_fr
    to_bus=br_obj.bus_to
    DSSCktobj.dssCircuit.SetActiveElement(line)
    # if(DSSCktobj.dssCircuit.ActiveCktElement.IsOpen(1,0)):
    #     sw_status=0
    # else:
    #     sw_status=1
    sw_status=DSSCktobj.dssCircuit.SwtControls.Action -1 
    AllSwitches.append({'switch name':name,'edge name':line, 'from bus':from_bus.split('.')[0], 'to bus':to_bus.split('.')[0], 'status':sw_status})
    i=DSSCktobj.dssCircuit.SwtControls.Next    


# # Check Node Voltages
V_nodes=[]
for n in list(G_init.nodes()):
    V=Bus(DSSCktobj, n).Vmag
    temp_conn=Bus(DSSCktobj, n).nodes
    V_nodes.append({'name':n, 'Connection': temp_conn, 'Voltage':V})


# Check Line Currents
I_nodes=[]
for e in list(G_init.edges(data=True)):
    branchname=e[2]['label'][0]
    DSSCktobj.dssCircuit.SetActiveElement(branchname)
    I=DSSCktobj.dssCircuit.ActiveCktElement.Powers
    # I=Branch(DSSCktobj, branchname).Cap
    I_nodes.append({'name':branchname, 'Current':I})       

