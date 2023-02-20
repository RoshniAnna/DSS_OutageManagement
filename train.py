import numpy as np
import gym
from stable_baselines3 import PPO
# from stable_baselines.common import make_vec_env
from DSS_OutCtrl_Env import DSS_OutCtrl_Env
# import json
# import datetime as dt
import torch
# from stable_baselines3.common.utils import set_random_seed
from feedforwardPolicy import *
from stable_baselines3 import A2C, PPO
#from CustomPolicies import ActorCriticGCAPSPolicy
from DSS_Initialize import   *
from CustomPolicies import ActorCriticGCAPSPolicy

import pickle
from stable_baselines3.common.utils import set_random_seed

from stable_baselines3.common.vec_env import DummyVecEnv, SubprocVecEnv


class CustomGNN(BaseFeaturesExtractor):
    """
    :param observation_space: (gym.Space)
    :param features_dim: (int) Number of features extracted.
        This corresponds to the number of unit for the last layer.
    """

    def __init__(self,
                observation_space: gym.spaces.Box,
                 features_dim: int = 256,
                 n_layers=2,
                 n_dim=256,
                 n_p=1,
                 node_dim=3,
                 n_K=1,
                 ):
        super(CustomGNN, self).__init__(observation_space, features_dim)
        self.n_layers = n_layers
        self.n_dim = n_dim
        self.n_p = n_p
        self.n_K = n_K
        self.node_dim = node_dim
        self.init_embed = torch.nn.Linear(node_dim, n_dim * n_p)
        self.init_embed_depot = torch.nn.Linear(2, n_dim)

        self.W_L_1_G1 = torch.nn.Linear(n_dim * (n_K + 1) * n_p, n_dim)

        self.normalization_1 = torch.nn.BatchNorm1d(n_dim * n_p)

        self.W_F = torch.nn.Linear(n_dim * n_p, features_dim)
        self.full_context_nn = th.nn.Linear(19, 128)
        self.switch_encoder = torch.nn.Linear(16, 128)

        self.activ = torch.nn.Sigmoid()

    def forward(self, data):
        X = data['NodeFeat(BusVoltage)']
        num_samples, num_locations, _ = X.size()
        A = data["Adjacency"]
        # print(A.shape)
        D = torch.mul(torch.eye(num_locations).expand((num_samples, num_locations, num_locations)),
                      (A.sum(-1))[:, None].expand((num_samples, num_locations, num_locations)))

        # Layer 1

        # p = 3
        F0 = self.init_embed(X)
        # print(torch.isnan(X).to(torch.int32).sum())
        #print(X)
        # K = 3
        L = D - A
        
        g_L1_1 = self.W_L_1_G1(torch.cat((F0[:, :, :],
                                          torch.matmul(L, F0)[:, :, :]
                                          ),
                                         -1))

        F1 = g_L1_1  # torch.cat((g_L1_1), -1)
        # F1 = self.activ(F1)

        F_final = self.W_F(F1)

        h = F_final  # torch.cat((init_depot_embed, F_final), 1)
        # return (
        #     h,  # (batch_size, graph_size, embed_dim)
        #     h.mean(dim=1),  # average to get embedding of graph, (batch_size, embed_dim)
        # )
        switch_embeddings = self.switch_encoder(h.permute(0,2,1))
       
        context = self.full_context_nn(th.cat((data["EnergySupp"],data["VoltageViolation"], data["EdgeFeat(Branchflow)"]), -1))
        
        final = switch_embeddings.mean(dim=1)+context
        
        return final

def learning_rate_schedule(initial_value: float) -> Callable[[float], float]:

    def func(progress_remaining: float) -> float:

        # return max((progress_remaining**2) * initial_value, 0.0002)

        return  initial_value
    return func

def make_env(rank, seed=0):
    """
    Utility function for multiprocessed env.
    :param env_id: (str) the environment ID
    :param num_env: (int) the number of environments you wish to have in subprocesses
    :param seed: (int) the inital seed for RNG
    :param rank: (int) index of the subprocess
    """
    def _init():
        env = DSS_OutCtrl_Env()
        env.seed(seed + rank)
        return env
    set_random_seed(seed)
    return _init
if __name__ == '__main__':
# env = openDSSenv34()
    num_cpu = 1
    use_cuda = torch.cuda.is_available() 
    env = SubprocVecEnv([make_env(i) for i in range(num_cpu)])

    # n_envs = 4
    # env = make_vec_env(mTSPEnv, n_envs=n_envs, env_kwargs={"n_locations":21, "n_agents":5})

    rms_prop_eps = 1e-5
    policy_kwargs = dict(
        features_extractor_class=CustomGNN,
        features_extractor_kwargs=dict(features_dim=128,node_dim=3),
        #activation_fn=torch.nn.Sigmoid,
        net_arch=[dict(vf=[128,256])],
        device=torch.device("cuda:0" if use_cuda else "cpu")
        # optimizer_class = th.optim.RMSprop,
        # optimizer_kwargs = dict(alpha=0.89, eps=rms_prop_eps, weight_decay=0)
    )

    model = PPO(policy=ActorCriticGCAPSPolicy, env=env,tensorboard_log="logger/", policy_kwargs=policy_kwargs, verbose=1, n_steps=200, batch_size=100,
            gamma=1.00,
            learning_rate=learning_rate_schedule(0.00001),
                ent_coef=0.01
                ).learn(total_timesteps=80000, n_eval_episodes=1, log_interval=1, tb_log_name="s1")




    # model = A2C(ActorCriticGCAPSPolicy, env,tensorboard_log="logger/", policy_kwargs=policy_kwargs, verbose=1, n_steps=200,
    #         use_rms_prop=False,
    #         gamma=1.00,
    #         learning_rate=learning_rate_schedule(0.07),
    #             ).learn(total_timesteps=20000, n_eval_episodes=1, log_interval=1, tb_log_name="R1_new_env")
    # model = A2C('MultiInputPolicy', env, tensorboard_log="logger/", policy_kwargs=policy_kwargs, verbose=1,
    #             n_steps=1).learn(total_timesteps=20000, n_eval_episodes=1)
    # model.learn(total_timesteps=2000)


# observations = env.test_func()

    log_dir = "."
    model.save("s1")