from netpyne import specs, sim
import matplotlib
import pandas as pd

# Use Agg backend for matplotlib to avoid DISPLAY issues
matplotlib.use('Agg')

# Load circuit parameters from Excel
circuit_params = pd.read_excel('Circuit_param.xls', sheet_name=None, index_col=0)
cell_names = [i for i in circuit_params['conn_probs'].axes[0]]

netParams = specs.NetParams()

# Synaptic Mechanisms
netParams.synMechParams['AMPA'] = {'mod': 'Exp2Syn', 'tau1': 0.3, 'tau2': 3.0, 'e': 0}
netParams.synMechParams['NMDA'] = {'mod': 'Exp2Syn', 'tau1': 2.0, 'tau2': 65.0, 'e': 0}
netParams.synMechParams['GABA'] = {'mod': 'Exp2Syn', 'tau1': 1.0, 'tau2': 10.0, 'e': -80}

# Cell Parameters
cellParams = {}
for name in cell_names:
    cellParams[name] = {
        'secs': {
            'soma': {
                'geom': {'pt3d': [(0, 0, 0, 10), (0, 0, 10, 10)]},
                'mechs': {'pas': {'g': 0.0001, 'e': -65}}
            },
            'axon': {
                'geom': {'pt3d': [(0, 0, 0, 1), (100, 0, 0, 1)]},
                'mechs': {'hh': {'gnabar': 0.12, 'gkbar': 0.036, 'gl': 0.0003, 'el': -54.3}}
            }
        }
    }
netParams.cellParams.update(cellParams)

# Connectivity Parameters
connectivity_params = {
    'HL23PYR->HL23PYR': {
        'preConds': {'pop': 'HL23PYR'},
        'postConds': {'pop': 'HL23PYR'},
        'probability': circuit_params["conn_probs"].at['HL23PYR', 'HL23PYR'],
        'weight': circuit_params["syn_params"]['HL23PYRHL23PYR']['gmax'],
        'synMech': 'AMPA',
        'delay': 0.5,
        'plasticity': {
            'mech': 'TsodyksMarkram',
            'params': {
                'U': circuit_params["syn_params"]['HL23PYRHL23PYR']['Use'],
                'tau_rec': circuit_params["syn_params"]['HL23PYRHL23PYR']['Dep'],
                'tau_facil': circuit_params["syn_params"]['HL23PYRHL23PYR']['Fac'],
                'u0': circuit_params["syn_params"]['HL23PYRHL23PYR']['u0']
            }
        }
    },
    'HL23PYR->HL23SST': {
        'preConds': {'pop': 'HL23PYR'},
        'postConds': {'pop': 'HL23SST'},
        'probability': circuit_params["conn_probs"].at['HL23PYR', 'HL23SST'],
        'weight': circuit_params["syn_params"]['HL23PYRHL23SST']['gmax'],
        'synMech': 'AMPA',
        'delay': 0.5,
        'plasticity': {
            'mech': 'TsodyksMarkram',
            'params': {
                'U': circuit_params["syn_params"]['HL23PYRHL23SST']['Use'],
                'tau_rec': circuit_params["syn_params"]['HL23PYRHL23SST']['Dep'],
                'tau_facil': circuit_params["syn_params"]['HL23PYRHL23SST']['Fac'],
                'u0': circuit_params["syn_params"]['HL23PYRHL23SST']['u0']
            }
        }
    },
    'HL23PYR->HL23PV': {
        'preConds': {'pop': 'HL23PYR'},
        'postConds': {'pop': 'HL23PV'},
        'probability': circuit_params["conn_probs"].at['HL23PYR', 'HL23PV'],
        'weight': circuit_params["syn_params"]['HL23PYRHL23PV']['gmax'],
        'synMech': 'AMPA',
        'delay': 0.5,
        'plasticity': {
            'mech': 'TsodyksMarkram',
            'params': {
                'U': circuit_params["syn_params"]['HL23PYRHL23PV']['Use'],
                'tau_rec': circuit_params["syn_params"]['HL23PYRHL23PV']['Dep'],
                'tau_facil': circuit_params["syn_params"]['HL23PYRHL23PV']['Fac'],
                'u0': circuit_params["syn_params"]['HL23PYRHL23PV']['u0']
            }
        }
    },
    'HL23PYR->HL23VIP': {
        'preConds': {'pop': 'HL23PYR'},
        'postConds': {'pop': 'HL23VIP'},
        'probability': circuit_params["conn_probs"].at['HL23PYR', 'HL23VIP'],
        'weight': circuit_params["syn_params"]['HL23PYRHL23VIP']['gmax'],
        'synMech': 'AMPA',
        'delay': 0.5,
        'plasticity': {
            'mech': 'TsodyksMarkram',
            'params': {
                'U': circuit_params["syn_params"]['HL23PYRHL23VIP']['Use'],
                'tau_rec': circuit_params["syn_params"]['HL23PYRHL23VIP']['Dep'],
                'tau_facil': circuit_params["syn_params"]['HL23PYRHL23VIP']['Fac'],
                'u0': circuit_params["syn_params"]['HL23PYRHL23VIP']['u0']
            }
        }
    },
    'HL23PV->HL23PYR': {
        'preConds': {'pop': 'HL23PV'},
        'postConds': {'pop': 'HL23PYR'},
        'probability': circuit_params["conn_probs"].at['HL23PV', 'HL23PYR'],
        'weight': circuit_params["syn_params"]['HL23PVHL23PYR']['gmax'],
        'synMech': 'GABA',
        'delay': 0.5,
        'plasticity': {
            'mech': 'TsodyksMarkram',
            'params': {
                'U': circuit_params["syn_params"]['HL23PVHL23PYR']['Use'],
                'tau_rec': circuit_params["syn_params"]['HL23PVHL23PYR']['Dep'],
                'tau_facil': circuit_params["syn_params"]['HL23PVHL23PYR']['Fac'],
                'u0': circuit_params["syn_params"]['HL23PVHL23PYR']['u0']
            }
        }
    },
    'HL23PV->HL23SST': {
        'preConds': {'pop': 'HL23PV'},
        'postConds': {'pop': 'HL23SST'},
        'probability': circuit_params["conn_probs"].at['HL23PV', 'HL23SST'],
        'weight': circuit_params["syn_params"]['HL23PVHL23SST']['gmax'],
        'synMech': 'GABA',
        'delay': 0.5,
        'plasticity': {
            'mech': 'TsodyksMarkram',
            'params': {
                'U': circuit_params["syn_params"]['HL23PVHL23SST']['Use'],
                'tau_rec': circuit_params["syn_params"]['HL23PVHL23SST']['Dep'],
                'tau_facil': circuit_params["syn_params"]['HL23PVHL23SST']['Fac'],
                'u0': circuit_params["syn_params"]['HL23PVHL23SST']['u0']
            }
        }
    }
}
netParams.connParams.update(connectivity_params)

# Stimulation Parameters
stimParams = {}
stimuli = []
for stimulus in circuit_params['STIM_PARAM'].axes[0]:
    stim = {}
    for param_name in circuit_params['STIM_PARAM'].axes[1]:
        stim[param_name] = circuit_params['STIM_PARAM'].at[stimulus, param_name]
    new_param = circuit_params["syn_params"][stim['syn_params']].copy()
    new_param['gmax'] = stim['gmax']
    stim['new_param'] = new_param
    stimuli.append(stim)

for i, stim in enumerate(stimuli):
    stim_key = f'stim_{i}'
    stimParams[stim_key] = {
        'type': 'NetStim',
        'rate': stim['interval'],
        'noise': 0,
        'start': stim['start_time']
    }
    netParams.stimSourceParams[stim_key] = stimParams[stim_key]
    
    netParams.stimTargetParams[stim_key] = {
        'source': stim_key,
        'conds': {'pop': stim['cell_name']},
        'weight': stim['gmax'],
        'delay': stim['delay']
    }

# Create, simulate, and analyze network
simConfig = specs.SimConfig()
simConfig.createNEURONObj = False  # Do not require NEURON GUI
sim.createSimulateAnalyze(netParams, simConfig)
