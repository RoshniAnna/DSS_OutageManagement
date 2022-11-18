"""
In this file the functions to evaluate the state, reward are defined and also the action is implemented
"""
import win32com.client
import numpy as np
import networkx as nx
from  DSS_Initialize import*


# Get the nodes of the switches from this function---includes all the details of switches
def switchInfo(DSSCktobj):
    #Input the DSSCircuitobject and initial circuit graph
    #Returns: A list of dictionaries which contains 
    #         The name of the switch.
    #         The associated line which is the edge label
    #         The from bus and to bus of the line housing the switch
    #         The status of the switch in the DSSCircuit object passed as system state
    
    AllSwitches=[]
    i=DSSCktobj.dssCircuit.SwtControls.First
    while i>0:
        name=DSSCktobj.dssCircuit.SwtControls.Name
        line=DSSCktobj.dssCircuit.SwtControls.SwitchedObj
        # br_obj=Branch(DSSCktobj,line)
        # from_bus=br_obj.bus_fr
        # to_bus=br_obj.bus_to
        DSSCktobj.dssCircuit.SetActiveElement(line)
        # if(DSSCktobj.dssCircuit.ActiveCktElement.IsOpen(2,0)):
        #     sw_status=0
        # else:
        #     sw_status=1
        sw_status=DSSCktobj.dssCircuit.SwtControls.Action -1
        # AllSwitches.append({'switch name':name,'edge name':line, 'from bus':from_bus.split('.')[0], 'to bus':to_bus.split('.')[0], 'status':sw_status})
        AllSwitches.append({'switch name':name,'edge name':line, 'status':sw_status})
        
        i=DSSCktobj.dssCircuit.SwtControls.Next

    return AllSwitches     


    
def get_state(DSSCktobj,G):
    #Input: object of type DSSObj.ActiveCircuit (COM interface for OpenDSS Circuit)
    #Returns: dictionary of circuit loss, bus voltage, branch powerflow, radiality of network
    
    node_list=list(G_init.nodes())
    Adj_mat=nx.adjacency_matrix(G,nodelist=node_list)
    # Extracting pu loss for the DSS circuit object 
    DSSCktobj.dssTransformers.First
    KVA_base=DSSCktobj.dssTransformers.kva
    # P_loss=(DSSCktObj.dssCircuit.Losses[0])/(1000*KVA_base)
    # Q_loss=(DSSCktObj.dssCircuit.Losses[1])/(1000*KVA_base)

    ENS=0
    for ld in list(DSSCktobj.dssLoads.AllNames):
        DSSCktobj.dssCircuit.SetActiveElement("Load." + ld) #set the load as the active element
        S=np.array(DSSCktobj.dssCircuit.ActiveCktElement.Powers)
        ctidx = 2 * np.array(range(0, min(int(S.size/ 2), 3)))
        P = S[ctidx] #active power in KW
        Q =S[ctidx + 1] #angle in KVA
        Power_Supp=sum(P)
        if np.isnan(Power_Supp):
            Power_Supp=0.0
        # Demand = DSSCktobj.dssCircuit.ActiveCktElement.Properties('KW').Val[0]
        # ENS_ld= Demand -Power_Supp
        ENS_ld=Power_Supp
        ENS= ENS+ENS_ld


    # Extracting the pu node voltages at all buses
    Vmagpu=[]
    nodes_conn=[]
    for b in node_list:
        V=Bus(DSSCktobj,b).Vmag
        Vmagpu.append(V)
        nodes_conn.append(Bus(DSSCktobj,b).nodes)
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
       Conv_const=1000  # NonConvergence penalty   
 
    # The voltage violation
    V_viol=Volt_Constr(Vmagpu,nodes_conn)
        
    return {"ENS":ENS,"NodeFeat(BusVoltage)":np.array(Vmagpu), "EdgeFeat(branchflow)":np.array(I_flow),"Adjacency":np.array(Adj_mat.todense()), "VoltageViolation":V_viol, "Convergence":Conv_const}

    
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
def Volt_Constr(Vmagpu,nodes_conn):
    #Input: The pu magnitude of node voltages at all buses, node activated or node phase of all buses
    V_upper=1.10
    V_lower=0.90
    Vmin=100
    Vmax=0    
    for i in range(len(nodes_conn)):
        for phase_co in nodes_conn[i]:
            if (Vmagpu[i][phase_co-1]<Vmin):
                Vmin=Vmagpu[i][phase_co-1]
            if (Vmagpu[i][phase_co-1]>Vmax):
                Vmax=Vmagpu[i][phase_co-1]
    if (Vmax > V_upper) and (Vmin < V_lower):             
        V_viol=abs(Vmax-V_upper)+abs(V_lower-Vmin) # For the minimum and maximum voltage in the network(all nodes of all buses)
    else:
        V_viol=0
    return V_viol  

         
# Constraint for branch flow violation
def Flow_Constr(I_flow):
    I_upper=2 #pu upper limit for average branch current
    I_lower=-2 # pu lower limit for average branch current
    flow_viol=0
    for i in I_flow:
        if (i>I_upper) and (i<I_lower):
            flow_viol=flow_viol+abs(i-I_upper)+abs(I_lower-i) #sum of all branch current violations
        else:
            flow_viol=0
    return flow_viol

def get_reward(observ_dict):
    #Input: A dictionary describing the state of the network
    #Output: reward        
    reward= observ_dict['ENS']-observ_dict['Convergence']
    # TO DO: Voltage and Current violations multiplier

    return reward