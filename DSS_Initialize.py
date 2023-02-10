"""
In this file the sectionalizing and tie switch details are specified. 
The DER details are also specified (in particular the PV generators and BESS info). 
Also the path for the DSS file containing the circuit information is specified.
The final DSS circuit which will be used by the environment is created.

---G_init is used to maintain order of nodes, edges and also the varying graph scenario
---G_base is used to simulate fault isolation
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

# G_init has both sectionalizing and tie switches  

#--------- Graph with normal operating topology (with only sectionalizing switches)--for outage simulation (fault isolation)
tie_edges=[]
i=DSSCktobj.dssCircuit.SwtControls.First
while i>0:
    name=DSSCktobj.dssCircuit.SwtControls.Name
    if name[:5]=='swtie':
       line=DSSCktobj.dssCircuit.SwtControls.SwitchedObj
       br_obj=Branch(DSSCktobj,line)
       from_bus=br_obj.bus_fr
       to_bus=br_obj.bus_to
       tie_edges.append((from_bus,to_bus))
    i=DSSCktobj.dssCircuit.SwtControls.Next 
G_base = G_init.copy()    
G_base.remove_edges_from(tie_edges)

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


#------ Assigning a black start indicator for generators--------------------------------
gen_buses = np.array(list(Generator_Buses.values())) # all the generator buses
gen_elems=  list(Generator_Buses.keys()) # all the generator elements
Gen_Info ={}
for n in node_list:
    blackstart_flag=0
    gen_names=[gen_elems[x]  for x in  np.where(gen_buses == n)[0]]
    if (len(gen_names)!=0):
        for g in gen_names:
            blackstart_flag = blackstart_flag + Generator_BlackStart[g] 
        Gen_Info[n]= {'Generators':gen_names, 'Blackstart':blackstart_flag}
    
    

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

SwitchLines=[(s['from bus'],s['to bus']) for s in AllSwitches]


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
