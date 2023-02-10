"""
In this file the functions to evaluate the state, reward are defined and also the action is implemented
"""

import win32com.client
import numpy as np
from  DSS_Initialize import*


# To get the details of switches and their status
def switchInfo(DSSCktobj):
    #Input:   DSS Circuit Object
    #Returns: A list of dictionaries which contains: 
    #         The name of the switch.
    #         The associated line of the switch (the edge label).
    #         The status of the switch in the DSS Circuit object
    
    AllSwitches=[]
    i=DSSCktobj.dssCircuit.SwtControls.First
    while i>0:
        name=DSSCktobj.dssCircuit.SwtControls.Name
        line=DSSCktobj.dssCircuit.SwtControls.SwitchedObj
        DSSCktobj.dssCircuit.SetActiveElement(line)
        sw_status=DSSCktobj.dssCircuit.SwtControls.Action -1
        AllSwitches.append({'switch name':name,'edge name':line, 'status':sw_status})
        
        i=DSSCktobj.dssCircuit.SwtControls.Next

    return AllSwitches     


def get_state(DSSCktobj,G, edgesout):
    #Input: DSS Circuit Object (COM interface for OpenDSS Circuit) and the equivalent Graph representation
    #Returns: Dictionary to indicate state which includes:
              # Total Energy Supplied, bus voltage (nodes), branch powerflow (edges), adjacency, voltage and convergence violations
    
    node_list=list(G_init.nodes())
    Adj_mat=nx.adjacency_matrix(G,nodelist=node_list)

    # Estimating the total energy supplied to the end users given the state encompassed in DSS Circuit
    DSSCktobj.dssTransformers.First
    KVA_base=DSSCktobj.dssTransformers.kva #To convert into per unit

    En_Supply=0
    for ld in list(DSSCktobj.dssLoads.AllNames): # For each load
        DSSCktobj.dssCircuit.SetActiveElement("Load." + ld) #set the load as the active element
        S=np.array(DSSCktobj.dssCircuit.ActiveCktElement.Powers) # Power vector (3 phase P and Q for each load)
        ctidx = 2 * np.array(range(0, min(int(S.size/ 2), 3)))
        P = S[ctidx] #active power in KW
        Q =S[ctidx + 1] #reactive power in KVar
        Power_Supp=sum(P) # total active power supplied at load
        if np.isnan(Power_Supp):
            Power_Supp=0    # Nodes which are isolated with loads but no generators return nan-- ignore that(consider as inactive)   
        En_Supply= En_Supply + Power_Supp
        

    # Extracting the pu node voltages at all buses
    
    Vmagpu=[]
    active_conn=[]
    for b in node_list:    
        V = Bus(DSSCktobj,b).Vmag
        active_conn.append(Bus(DSSCktobj,b).nodes)
        temp_flag = np.isnan(V) # Nodes which are isolated with loads but no generators return nan-- ignore that(consider as inactive)
        if np.any(temp_flag,where=True):
            V[temp_flag] = 0
            temp_conn=[n for n in active_conn[node_list.index(b)] if temp_flag[n-1]==False]
            active_conn[node_list.index(b)]=np.array(temp_conn) #only active node connections
        Vmagpu.append(V)
    
    # Extracting the pu average branch currents(also includes the open branches)           
    
    I_flow=[]
    for e in G_init.edges(data=True):
        branchname=e[2]['label']
        I=Branch(DSSCktobj, branchname).Cap
        I_flow.append(I)

    # The convergence test and violation penalty   
    if DSSCktobj.dssSolution.Converged:
       conv_flag=1
       Conv_const=0
    else:
       conv_flag=0
       Conv_const=100000# NonConvergence penalty   
 
    # The voltage violation
    V_viol=Volt_Constr(Vmagpu,active_conn)
     
    # To mask those switches which are out
    SwitchMasks=[]
    for x in SwitchLines:
        if x in edgesout:
           SwitchMasks.append(1)
        else:
           SwitchMasks.append(0)
            
    
    
    return {"EnergySupp":np.array([En_Supply]),"NodeFeat(BusVoltage)":np.array(Vmagpu), "EdgeFeat(Branchflow)":np.array(I_flow),"Adjacency":np.array(Adj_mat.todense()), "VoltageViolation":np.array([V_viol]), "ConvergenceViolation":Conv_const,"ActionMasking":np.array(SwitchMasks)}

    
def take_action(action,out_edges):
    #Input :object of type DSSObj.ActiveCircuit (COM interface for OpenDSS Circuit)
    #Input: action multi binary type. i.e., the status of each switch if it is 0 open and 1 close
    #Returns:the circuit object with action implemented (and slack assigned), also the graph scenario
    DSSCktObj,G_init,conv_flag= initialize() 
    G_sc=G_init.copy() # Copy to create graph scenario
    
   # -------------Implement Action on DSSCircuit Object
    i=DSSCktObj.dssCircuit.SwtControls.First
    while (i>0):
        if action[i-1]==0:
            # DSSCktobj.dssCircuit.ActiveCktElement.Open(1,0)
            DSSCktObj.dssText.command=('Swtcontrol.' + DSSCktObj.dssCircuit.SwtControls.Name+ '.Action=o')
            # DSSCktobj.dssText.command='open ' + Swobj +' term=1'       #switching the line open
        else:
            # DSSCktobj.dssCircuit.ActiveCktElement.Close(1,0)
            DSSCktObj.dssText.command=('Swtcontrol.'+ DSSCktObj.dssCircuit.SwtControls.Name+ '.Action=c')
            # DSSCktobj.dssText.command='close ' + Swobj +' term=1'      #switching the line close
        i=DSSCktObj.dssCircuit.SwtControls.Next 
        
    DSSCktObj.dssSolution.Solve()     

   #----Disable outage lines from DSSCircuit and also from Graph Scenario
    for o_e in out_edges:
        (u,v)=o_e
        branch_name= G_init.edges[o_e]['label'][0]
        if G_sc.has_edge(u,v):
            G_sc.remove_edge(u,v) # Remove the edge in graph domain
        # Remove the element from the DSSCktobj
        DSSCktObj.dssCircuit.SetActiveElement(branch_name)
        # DSSCktobj.dssCircuit.ActiveCktElement.Open(1,0)
        DSSCktObj.dssText.command=(branch_name + '.enabled="False"')
        # DSSCktobj.dssText.command='open ' + o_e  +' term=1' #open the terminal of line to remove it
    DSSCktObj.dssSolution.Solve()    
    
    #---------- Also remove the open switches from Graph Scenario
    i=DSSCktObj.dssCircuit.SwtControls.First
    while i>0:
          # name=DSSCktobj.dssCircuit.SwtControls.Name
          line=DSSCktObj.dssCircuit.SwtControls.SwitchedObj
          # DSSCktobj.dssCircuit.SetActiveElement(line)
          # if(DSSCktobj.dssCircuit.ActiveCktElement.IsOpen(1,0)):
          if DSSCktObj.dssCircuit.SwtControls.Action == 1: #Open is 1 in DSS
             b_obj=Branch(DSSCktObj, line)
             u=b_obj.bus_fr.split('.')[0]
             v=b_obj.bus_to.split('.')[0]
             if G_sc.has_edge(u,v):
                G_sc.remove_edge(u,v) # Remove the edge in graph domain
          i=DSSCktObj.dssCircuit.SwtControls.Next    
    
    # #----- Finding network components and find virtual slack ------#
    Components= list(nx.connected_components(G_sc)) #components formed due to outage
    Virtual_Slack=[] # for each component not connected to sourcebus...we will assign a slack 
    if len(Components) >1 : #Only if there exists a network component unconnected to sourcebus virtual slack is assigned
        for C in Components: #for each component        
            if substatn_id not in C: # for the component unconnected to sourcebus
                Slack_DER={'name':'','kVA':0}
                # Find the DER corresponding to slack bus (largest grid forming DER) in component     
                for gen_bus, gen_info in Gen_Info.items():
                    if gen_bus in C and gen_info['Blackstart']==1: #if generator is present and has gridforming capability
                       kva_val=0
                       for gen_name in gen_info['Generators']: #get total KVA at node
                           DSSCktObj.dssCircuit.SetActiveElement(gen_name)
                           kva_val= kva_val + float(DSSCktObj.dssCircuit.ActiveCktElement.Properties('kVA').Val)
                       if kva_val > Slack_DER ['kVA'] : # if multiple grid forming DERs, largest grid forming DER is slack
                          Slack_DER['kVA'] = kva_val
                          Slack_DER['name']= ('bus_'+ gen_bus)
                Virtual_Slack.append(Slack_DER)
    
    #---- Assign slack bus in DSSCkt                 
    for vs in Virtual_Slack:
        Vs_name=vs['name']
        if Vs_name != '':
            Vs_locatn=Vs_name.split('_')[1]
            Vs_MVA = Vs_MVAsc3 = vs['kVA']/1000 #MVA and MVAsc3 are set to be same
            # Vs_MVA = Vs_MVAsc3 = 3
            Vs_MVAsc1 = Vs_MVAsc3/3 # MVAsc1 approax 1/3 of MVAsc3
            
            DSSCktObj.dssCircuit.SetActiveBus(Vs_locatn)
            # DSSCktobj.dssBus.kVBase gives the per phase (phase to neutral) voltage
            Vs_kv =DSSCktObj.dssBus.kVBase * math.sqrt(3) # this has to be phase to phase
            DSSCktObj.dssText.command=('New Vsource.'+ Vs_name + ' bus1=' + Vs_locatn + ' basekV=' + str(Vs_kv) +' phases=3 Pu=1.00 angle=30 baseMVA=' + str(Vs_MVA) + ' MVAsc3=' + str(Vs_MVAsc3) +' MVAsc1=' + str(Vs_MVAsc1) + ' enabled =yes')
            # print(Vs_MVAsc3)
            # print(Vs_MVAsc1)
            # DSSCktobj.dssText.command = 'Formedit'+ ' Vsource.' +Vs_name   
            for gens in Gen_Info[Vs_locatn]['Generators']:
                DSSCktObj.dssText.command = (gens + '.enabled=no')             
    DSSCktObj.dssSolution.Solve()    
    
    return DSSCktObj,G_sc



# Constraint for voltage violation 
def Volt_Constr(Vmagpu,active_conn):
    #Input: The pu magnitude of node voltages at all buses, node activated or node phase of all buses
    Vmax=1.10
    Vmin=0.90
    
    V_ViolSum=0
    for i in range(len(active_conn)):
        for phase_co in active_conn[i]:
            if (Vmagpu[i][phase_co-1]<Vmin):            
                V_ViolSum = V_ViolSum + abs(Vmin-Vmagpu[i][phase_co-1])
            if (Vmagpu[i][phase_co-1]>Vmax):                      
                V_ViolSum = V_ViolSum + abs(Vmagpu[i][phase_co-1]-Vmax)
    return V_ViolSum  

         


def get_reward(observ_dict):
    #Input: A dictionary describing the state of the network
    # ----#Output: reward- minimize load outage, penalize non convergence and closing of outage lines
      
    reward= observ_dict['EnergySupp']-observ_dict['ConvergenceViolation']-observ_dict['VoltageViolation']
  
    return reward