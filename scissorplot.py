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
#print(params)


# Utility Functions


def q(v):
    return 0.5 * 1.225 * (v**2)


def qS(v):
    return 0.5 * 1.225 * params['Sarea'] * (v**2)


# Key Aircraft Params
c_bar = np.mean([params['Ct'], params['Cr']])
maxthrust = 44538  # N @ takoff
vto = 62.1  # m/s
cthrust = maxthrust / qS(vto)
cm0 = -0.0663
cl_to = 2.549
clt = -1  # Tail CL
mtow = 35590  # kg
g = 9.81  # m/s/s
h0 = params['WingChordStart']/c_bar
print("h0 = " + str(h0))


def noseWheel():
    nosepos = params['NoseGearPos']
    mainpos = params['MainGearPos']
    condition_point = (0.975*(mainpos-nosepos)) + nosepos
    return condition_point/c_bar


def mainGearReaction():
    total_weight = mtow/qS(vto)
    return 0


def takeOffRotation(h):
    cl_moment = cl_to * (h-h0)
    print(cl_moment)
    ct_moment = cthrust * 0.25/c_bar  # Approx cg offset pos
    print(cthrust)
    main_gear_moment_distance = (params['MainGearPos']/c_bar) - h
    print(main_gear_moment_distance)
    weight_nondim = g*mtow/qS(vto)
    print(weight_nondim)
    tail_moment_arm = (params['TailRootRearPlane']/c_bar) - h
    print(tail_moment_arm)

    lhs_top = cm0 + cl_moment + ct_moment - (main_gear_moment_distance * (weight_nondim-cl_to))
    lhs_bottom = (clt * tail_moment_arm) - (main_gear_moment_distance * clt)

    return lhs_top/lhs_bottom


print(takeOffRotation(5.65))


def plotit(r1, r2):
    # Create Range of h values
    r2 = r2 + 0.1
    x_h = np.arange(r1, r2, 0.1)
    fig = plt.figure(figsize=(10, 7))
    ax = fig.add_subplot(1, 1, 1)
    ax.xaxis.set_ticks_position('bottom')
    plt.xlabel("h")
    plt.ylabel("ST/S")
    plt.plot(x_h, takeOffRotation(x_h))
    plt.plot([noseWheel(), noseWheel()], [min(takeOffRotation(x_h)), max(takeOffRotation(x_h))])
    plt.show()


plotit(5.4, 6.2)

