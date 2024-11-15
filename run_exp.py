# This is for suggesting next experiment
from operator import index
import numpy as np
import pandas as pd
import openpyxl
import matplotlib.pyplot as plt

from summit import *
from optimize import MOBO
from utils import *

# Create experimental domain
domain = Domain()

# Set up the domain with design variables (continuous, discrete, categorical) and constraints
domain += ContinuousVariable(
    name="temp", description="temperature", bounds=[20.0, 50.0]
)

domain += ContinuousVariable(
    name="conc_koh", description="KOH_concentration", bounds=[0.10, 3.00]
)

domain += ContinuousVariable(
    name="conc_gly", description="glycerol_concentration", bounds=[0.05, 3.00]
)

domain += ContinuousVariable(
    name="flowrate", description="flowrate", bounds=[0.0, 500.0]
)

domain += ContinuousVariable(
    name="current", description="current_density", bounds=[20.0, 125.0]
)

# For adding categorical/discrete variables
# domain += CategoricalVariable(
#     name="solv",
#     description="Solvent",
#     levels=["THF", "EtOAc", "MeCN", "Toluene"],
# )


# Objectives
domain += ContinuousVariable(
    name="yld",
    description="yield of lactic acid",
    bounds=[0, 100],
    is_objective=True,
    maximize=True,
)

domain += ContinuousVariable(
    name="conv",
    description="glycerol conversion",
    bounds=[0, 100],
    is_objective=True,
    maximize=True,
)


domain += ContinuousVariable(
    name="feff",
    description="faradaic efficiency",
    bounds=[0, 100],
    is_objective=True,
    maximize=True,
)

#########################################################################################

# # Generate initial experiments
# initial_exp = 10  # Number of initial experiments
# strategy = MOBO(domain)
# result = strategy.suggest_experiments(initial_exp)
#
# # Save initial data to excel
# data_values = pd.DataFrame(result.values)
# col_names = result.columns.get_level_values('NAME')
# data_values.columns = col_names
#
# # Setting precision for specific columns
# data_values["temp"] = pd.to_numeric(data_values["temp"], errors="coerce").round(1)
# data_values["conc_koh"] = pd.to_numeric(data_values["conc_koh"], errors="coerce").round(1)
# data_values["conc_gly"] = pd.to_numeric(data_values["conc_gly"], errors="coerce").round(1)
# data_values["flowrate"] = pd.to_numeric(data_values["flowrate"], errors="coerce").round(1)
# data_values["current"] = pd.to_numeric(data_values["current"], errors="coerce").round(1)
#
# # Insert columns for objectives
# objs = ["yld", "conv", "feff"]
# for col_name in objs:
#     data_values.insert(data_values.columns.get_loc("strategy"), col_name, None)
#
# # Save to excel
# data_values.to_excel("test.xlsx", sheet_name="data")
# print(data_values)


########################################################################################

# Generate new experiments

# Load data from excel
current_exp = 10 # Number of current experiments
data = pd.read_excel('test.xlsx', sheet_name='data', header=None, usecols="B:J", skiprows=[0], nrows=current_exp,
                        names=["temp","conc_koh","conc_gly","flowrate","current","yld","conv","feff","strategy"])

dataset = DataSet.from_df(data)
strategy = MOBO(domain)
result = strategy.suggest_experiments(1, prev_res=dataset)

# Round the values
temp_new = result["temp"][0].round(1)
conc_koh_new = result["conc_koh"][0].round(1)
conc_gly_new = result["conc_gly"][0].round(1)
flowrate_new = result["flowrate"][0].round(1)
current_new = result["current"][0].round(1)
next_exp = [temp_new, conc_koh_new, conc_gly_new, flowrate_new, current_new, None, None, None, "MOBO"]
print(next_exp)

# Save to excel
data.loc[len(data)] = next_exp
print(data)
data.to_excel("test2.xlsx", sheet_name="data")
