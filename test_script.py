"""
Author: Steve Paul 
Date: 5/31/23 """

import numpy as np
import gym
from stable_baselines3 import PPO
import torch
from stable_baselines3 import A2C, PPO
from stable_baselines3.common.callbacks import CheckpointCallback, BaseCallback

from stable_baselines3.common.torch_layers import BaseFeaturesExtractor
from typing import Any, Callable, Dict, List, Optional, Tuple, Type, Union
import pickle
from stable_baselines3.common.utils import set_random_seed

from Configs.training_config import get_training_config
from stable_baselines3.common.vec_env import DummyVecEnv, SubprocVecEnv


if __name__ == '__main__':

    env_size = 34

    if env_size == 13:
        from Environments.DSSdirect_13bus_loadandswitching.DSS_OutCtrl_Env import DSS_OutCtrl_Env
        from Policies.Feature_Extractor import CustomGNN
        from Policies.CustomPolicies import ActorCriticGCAPSPolicy
    elif env_size == 34: # will add more conditions once the 13 bus is fixed
        from Environments.DSSdirect_34bus_loadandswitching.DSS_OutCtrl_Env import DSS_OutCtrl_Env
        from Policies.bus_34.Feature_Extractor import CustomGNN
        from Policies.bus_34.CustomPolicies import ActorCriticGCAPSPolicy

    env = DSS_OutCtrl_Env()
    model = PPO.load("/Users/stevepaul/Library/CloudStorage/Box-Box/Academics/Research/Graph Learning/ONR_Reconf_Journal/DSS_OutageManagement_2/Trained_Models/34_bus/MLP/MLP_1735000_steps.zip", env=env)

    dt = 0