import numpy as np
from openpyxl import load_workbook

# Excel Sheet Loading
cad_file_path = '../../../../2. TechInt/Project Space/CAD/CADParametersNewAero.xlsx'
cad_file = load_workbook(filename=cad_file_path, data_only=True)
cad_params = cad_file['Sheet1']  # Load into Sheet1

# Init list holders
param_name = []
param_value = []

# Loop helpers
data_available = True
data_row_start = 7  # Row where data starts

# Loop through variable names and values, saving to list.
while data_available is True:
    # print(cad_params["A"+str(data_row_start)].value)
    if cad_params["A"+str(data_row_start)].value is not None:
        param_name.append(cad_params["A"+str(data_row_start)].value)
        param_value.append(cad_params["B"+str(data_row_start)].value)
        data_row_start = data_row_start+1
    else:
        break

# Zip into a dictionary
params = dict(zip(param_name, param_value))

# Utility Functions


def q(v):
    return 0.5 * 1.225 * (v**2)


def qS(v):
    return 0.5 * 1.225 * params['Sarea'] * (v**2)


def myceil(x, base):
    return (base * np.ceil(float(x)/base))


def myfloor(x, base):
    return (base * np.floor(float(x)/base))

# Constants/Values
dcyWBN_db = -0.43  # per rad
dcyV_db = -3.726  # per rad
dcyV_dd = 1.892  # per rad
c_bar = np.mean([params['Ct'], params['Cr']])
maxthrust = 44538  # N @ takoff
vto = 62.1  # m/s
cthrust = maxthrust / qS(vto)
dr = np.deg2rad(30)
Kr = 0.9
blade_count = 6
blade_diameter = params['EnginePropDiameter']
blade_centre = params['EngineWingPosOutboard']
engine_scale = 0.89
engine_intake_area = 0.16 * engine_scale
engine_diameter = np.sqrt(engine_intake_area/np.pi)*2
h0 = params['WingChordStart']/c_bar
mtow_pos = (14.436+0.364)/c_bar  # Approx P2B mtow pos from GA
print(engine_diameter)


def take_off_yaw():
    lvtp = ((params['TailRootRearPlane'] / c_bar) - h0) * c_bar
    lhs_top = (cthrust + c_drag_engine())*(blade_centre/c_bar)
    lhs_bottom = (dr * Kr * dcyV_dd * lvtp)/c_bar
    return lhs_top/lhs_bottom


def c_drag_engine():
    cd_wind = 0.1 * (engine_diameter**2)
    cd_prop = 0.00125 * blade_count * blade_diameter**2
    cd_engine_total = cd_prop + cd_wind
    return cd_engine_total


print(take_off_yaw())
