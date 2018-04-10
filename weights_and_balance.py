import numpy as np
from matplotlib import pyplot as plt
from matplotlib.ticker import FuncFormatter
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

ax_lims=[big_weights["OWE_w"]*0.98, big_weights["MTOW_w"]*1.02]
c_bar = cg_params["M7"].value
print(c_bar)
h0 = cg_params["M8"].value


def mac_axes(mac_pos):
    lower_ax_point = (mac_pos-0.25) * ax_lims[0]
    upper_ax_point = (mac_pos-0.25) * ax_lims[1]
    print(lower_ax_point)
    return [lower_ax_point, upper_ax_point]


def to_tons(x, pos):
    'The two args are the value and tick position'
    return '%1.1fT' % (x*1e-3)


def to_mac(x, pos):
    return str(int(round(100*(x/ax_lims[0]+0.25)))) + '%'


formattery = FuncFormatter(to_tons)
formatterx = FuncFormatter(to_mac)


def plotit(mac_range=[0.11,0.51]):
    fig = plt.figure(figsize=(10, 7))
    ax = fig.add_subplot(1, 1, 1)
    ax.yaxis.set_major_formatter(formattery)
    ax.xaxis.set_major_formatter(formatterx)
    ax.xaxis.set_ticks_position('bottom')
    ax.yaxis.grid(which="major", color='k', linestyle='-', linewidth=1)

    # Plot MAC axes
    mac_spacing = 0.02
    mac_set = np.arange(mac_range[0], mac_range[1], mac_spacing)

    xtick_hold = []
    for mac_pos in mac_set:
        plt.plot(mac_axes(mac_pos), ax_lims, 'k')
        xtick_hold.append(mac_axes(mac_pos)[0])
        # plt.annotate(mac_pos, [mac_axes(mac_pos)[0], ax_lims[1]-(0.05*(ax_lims[1]-ax_lims[0]))])
    plt.xticks(xtick_hold)
    plt.ylabel("Mass")
    plt.xlabel("Moment (temp)")
    plt.xlim(mac_axes(mac_range[0]-0.01)[0], mac_axes(mac_range[1]+0.01)[0])
    plt.ylim(ax_lims[0], ax_lims[1])
    plt.show()

plotit()
print(to_mac(443,0))
