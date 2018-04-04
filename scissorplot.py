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


def myceil(x, base):
    return (base * np.ceil(float(x)/base))


def myfloor(x, base):
    return (base * np.floor(float(x)/base))


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
mtow_pos = (14.436+0.364)/c_bar  # Approx P2B mtow pos from GA


def noseWheel():
    nosepos = params['NoseGearPos']
    mainpos = params['MainGearPos']
    condition_point = (0.975*(mainpos-nosepos)) + nosepos
    return condition_point/c_bar


def mainGearReaction():
    total_weight = mtow/qS(vto)
    return 0


def takeOffRotation(h):
    cl_moment = cl_to * (h0-h)  # CL * distance between h and h0 (Centre of Lift)
    print("cl moment = " + str(cl_moment))
    ct_moment = cthrust * 0.5/c_bar  # Thrust Coeef * vertical distance to CoG (i.e. h)
    print("ct moment = " + str(cthrust))
    main_gear_moment_distance = (params['MainGearPos']/c_bar) - h  # Distance of main gear to h
    print("gear pos - h = " + str(main_gear_moment_distance))
    weight_nondim = g*mtow/qS(vto)  # The weight/qS
    print("W/qs = " + str(weight_nondim))
    tail_moment_arm = (params['TailRootRearPlane']/c_bar) - h  # The tail moment to h
    print("Tail moment arm distance = " + str(tail_moment_arm))

    lhs_top = cm0 + cl_moment - ct_moment - (main_gear_moment_distance * (weight_nondim-cl_to))
    #         Cm0 + cl(h-h0)  - ct(dist)  - reaction distance * the weight minus the cl lift
    # print("top = " + str(lhs_top))
    lhs_bottom = -((clt * tail_moment_arm) - (main_gear_moment_distance * clt))
    #            (tail lift * moment arm) - (distance * the rest of the vertically resolved bit)
    # print("bottom = " + str(lhs_bottom))

    return lhs_top/lhs_bottom


def kn(h):
    kn_limit = 0.05
    tail_moment_arm = (params['TailRootRearPlane'] / c_bar) - h0  # The tail moment to h
    a1 = 4.5
    a = 5.14
    d_e_alpha = 0.2

    lhs_top = kn_limit - h0 + h
    lhs_bottom = tail_moment_arm * (a1/a) * (1-(d_e_alpha))

    return lhs_top/lhs_bottom



print(takeOffRotation(5.65))


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
    plt.plot(x_h, takeOffRotation(x_h))
    y_tails.append(min(takeOffRotation(x_h)))
    y_heads.append(max(takeOffRotation(x_h)))

    plt.plot(x_h, kn(x_h))
    y_tails.append(min(kn(x_h)))
    y_heads.append(max(kn(x_h)))

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

    # Plot the noseWheel conditions
    plt.plot([noseWheel(), noseWheel()], [min_y, max_y])

    # Plot h0 position
    plt.plot([h0, h0], [ref_tail, ref_head])
    plt.annotate("Wing h0 Position", [h0, ref_head])
    # Plot MTOW CoG position (approx)
    plt.plot([mtow_pos, mtow_pos], [ref_tail, ref_head+0.05])
    plt.annotate("MTOW C.o.G", [mtow_pos, ref_head+0.05])
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
    plt.grid(True)
    plt.savefig("plot.png")
    plt.show()


plotit(5, 6.25)

