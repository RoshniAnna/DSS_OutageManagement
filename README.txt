This encapsulates the environment for outage response to improve resilience.
The opendssdirect version is used here as the environment.
Both switching and load shedding/pickup is modeled.



Single step environment
Bounds for the generator loads. Based on this I need to modify the policy network.
I_mag becomes empty in some steps
dss._cffi_api_util.DSSException: (#8989) No active bus found! Activate one and retry.
EnergySupp, VoltageViolation, NodeFeat(BusVoltage), and EdgeFeat(Branchflow) are getting inf
some crazy value for
