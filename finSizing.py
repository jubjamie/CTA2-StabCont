import matplotlib.pyplot as plt
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
maxthrust = 44459  # N @ takoff
vto = 62.4  # m/s
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


#print(take_off_yaw())

def plotit(r1, r2):
    # Create Range of h values
    step = 0.1
    r2 = r2 + step
    x_h = np.arange(r1, r2, 0.1)
    fig = plt.figure(figsize=(10, 7))
    ax = fig.add_subplot(1, 1, 1)
    ax.xaxis.set_ticks_position('bottom')
    y_tails = []
    y_heads = []
    plt.xlabel("h")
    plt.ylabel("ST/S")
    plt.plot([r1, r2], [take_off_yaw(), take_off_yaw()], label='Take Off Yaw')
    y_tails.append(take_off_yaw())
    y_heads.append(take_off_yaw())

    """This section constrains the graph correctly"""
    max_y = myceil(max(y_heads), step)
    #print(max_y)
    min_y = myfloor((min(y_tails)), step)
    #print(min_y)
    plt.ylim(min_y, max_y)
    plt.xlim(r1, r2-0.1)

    """Reference Plots
    This section plots reference points for the current aircraft configuration absed on data from the CAD param file"""
    ref_tail = min_y
    ref_head = ((max_y-ref_tail)*0.05) + ref_tail


    # Plot h0 position
    plt.plot([h0, h0], [ref_tail, ref_head])
    plt.annotate("Wing h0 Position", [h0, ref_head])
    # Plot MTOW CoG position (approx)
    plt.plot([mtow_pos, mtow_pos], [ref_tail, ref_head+(ref_head-ref_tail)])
    plt.annotate("MTOW C.o.G", [mtow_pos, ref_head+(ref_head-ref_tail)])
    # Plot Wing LE/TE position (approx)
    wingLE = (params['WingCentrePoint']-(params['Cr']/2))/c_bar
    wingTE = (params['WingCentrePoint'] + (params['Cr'] / 2)) / c_bar
    plt.plot([wingLE, wingLE], [ref_tail, ref_head])
    plt.annotate("Wing LE", [wingLE, ref_head])
    plt.plot([wingTE, wingTE], [ref_tail, ref_head])
    plt.annotate("Wing TE", [wingTE, ref_head])
    #Main Gear Position
    mainpos = params['MainGearPos']/c_bar
    plt.plot([mainpos, mainpos], [ref_tail, ref_head])
    plt.annotate("Main Gear Pos", [mainpos, ref_head])

    """Result Plots
    This section plots the results of the size finder. There are two options to call. One finds via a delta h range,
    the other via known fwd and aft positions"""
    # Known Delta h range

    plt.grid(True)
    plt.legend(loc='upper right', shadow=True)
    plt.savefig("finplot.png")
    plt.show()


plotit(5, 7)
