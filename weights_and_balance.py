import numpy as np
from matplotlib import pyplot as plt
from openpyxl import load_workbook
import filepath1

# Excel Sheet Loading
cg_file_path = filepath1.cad_file_path
cg_file = load_workbook(filename=cg_file_path, data_only=True)
cg_params = cg_file['CG']  # Load into CG Sheet

# Init list holders
param_name = []
param_100_w = []
param_80_w = []
param_100_x = []
param_80_x = []
param_100_m = []
param_80_m = []

# Loop helpers
data_available = True
data_row_start = 3  # Row where data starts

# Loop through variable names and values, saving to list.
while data_available is True:
    # print(cad_params["A"+str(data_row_start)].value)
    if cg_params["A" + str(data_row_start)].value is not None:
        param_name.append(cg_params["A" + str(data_row_start)].value)
        param_100_w.append(cg_params["B" + str(data_row_start)].value)
        param_80_w.append(cg_params["C" + str(data_row_start)].value)
        param_100_x.append(cg_params["F" + str(data_row_start)].value)
        param_80_x.append(cg_params["G" + str(data_row_start)].value)
        param_100_m.append(cg_params["H" + str(data_row_start)].value)
        param_80_m.append(cg_params["I" + str(data_row_start)].value)
        data_row_start = data_row_start+1
    else:
        break

# Zip into a dictionary
w_100 = dict(zip(param_name, param_100_w))
m_100 = dict(zip(param_name, param_100_m))
x_100 = dict(zip(param_name, param_100_x))
w_80 = dict(zip(param_name, param_80_w))
m_80 = dict(zip(param_name, param_80_m))
x_80 = dict(zip(param_name, param_80_x))

# Add statics
big_weights = {"MTOW_w": cg_params["Q4"].value,
               "MTOW_m": cg_params["Q3"].value,
               "MTOW_x": cg_params["Q5"].value,
               "OWE_w": cg_params["Q23"].value,
               "OWE_m": cg_params["Q22"].value,
               "OWE_x": cg_params["Q24"].value,
               }
