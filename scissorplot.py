import matplotlib.pyplot as plt
import numpy as np
from openpyxl import load_workbook
import filepath1

# Excel Sheet Loading
cad_file_path = filepath1.cad_file_path
# cad_file_path = '../../../../2. TechInt/Project Space/CAD/CADParametersNewAero.xlsx'
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
c_bar = cad_file['Interface']['B72'].value
maxthrust = cad_file['Interface']['F28'].value * 2  # N @ takoff
vto = 58.14  # m/s
cthrust = maxthrust / qS(vto)
cm0 = cad_file['Interface']['F29'].value
cl_to = cad_file['Interface']['F30'].value
cl_ld = cad_file['Interface']['F32'].value
cm0_ld = cad_file['Interface']['F33'].value
clt = -0.8  # Tail CL
mtow = cad_file['Interface']['F31'].value  # kg
g = 9.81  # m/s/s
h0 = cad_file['Interface']['B73'].value
print("h0 = " + str(h0))
mtow_pos_m = cad_file['Interface']['F16'].value
mtow_pos = mtow_pos_m/c_bar  # From Mass CG File
lvtp_cad_value = cad_file['Interface']["B95"].value
engine_vcg_dist = cad_file['Interface']['F15'].value
# params['MainGearPos'] = 15.93
a1 = 4.5  # From Horace
a = cad_file['Interface']['F34'].value  # From Horace
a_to = cad_file['Interface']['F35'].value  # From Horace
a_ld = cad_file['Interface']['F36'].value  # From Horace


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
    # print("cl moment = " + str(cl_moment))
    ct_moment = cthrust * engine_vcg_dist/c_bar  # Thrust Coeef * vertical distance to CoG (i.e. h)
    # print("ct moment = " + str(cthrust))
    main_gear_moment_distance = (params['MainGearPos']/c_bar) - h  # Distance of main gear to h
    # print("gear pos - h = " + str(main_gear_moment_distance))
    weight_nondim = g*mtow/qS(vto)  # The weight/qS
    # print("W/qs = " + str(weight_nondim))
    tail_moment_arm = (lvtp_cad_value/c_bar) - h  # The tail moment to h
    # print("Tail moment arm distance = " + str(tail_moment_arm))

    lhs_top = cm0 + cl_moment - ct_moment - (main_gear_moment_distance * (weight_nondim - cl_to))
    #         Cm0 + cl(h-h0)  - ct(dist)  - reaction distance * the weight minus the cl lift
    # print("top = " + str(lhs_top))
    lhs_bottom = (clt*(main_gear_moment_distance-tail_moment_arm))
    #            (tail lift * moment arm) - (distance * the rest of the vertically resolved bit)
    # print("bottom = " + str(lhs_bottom))

    return lhs_top/lhs_bottom


def landing(h):
    tail_moment_arm = (lvtp_cad_value / c_bar) - h0
    lhs_top = cm0_ld - (cl_ld*(h0-h))
    lhs_bottom = clt * tail_moment_arm
    return lhs_top/lhs_bottom


def kn(h):
    kn_limit = 0.05
    tail_moment_arm = (lvtp_cad_value / c_bar) - h0  # The tail moment to h

    d_e_alpha = 0.2

    lhs_top = kn_limit - h0 + h
    lhs_bottom = tail_moment_arm * (a1/a) * (1-(d_e_alpha))

    return lhs_top/lhs_bottom


def get_kn(fraction, h):
    tail_moment_arm = (lvtp_cad_value / c_bar) - h0  # The tail moment to h
    d_e_alpha = 0.2
    return h0-h+(fraction*tail_moment_arm*(a1/a)*(1-d_e_alpha))



def size_finder_delta(x_h, left_y, right_y, static_y, hrange):
    y_find = []
    x_find = []
    for y_ar in right_y:
        zr = np.polyfit(x_h, y_ar, 1)
        zl = np.polyfit(x_h, left_y, 1)
        x_find_i = ((zr[0]*hrange)+zr[1]-zl[1])/(zl[0]-zr[0])
        x_find.append(x_find_i)
        zl_poly = np.poly1d(zl)
        y_find.append(zl_poly(x_find_i))

    for y_ar in static_y:
        x_find_i = y_ar - hrange
        x_find.append(x_find_i)
        zl = np.polyfit(x_h, left_y, 1)
        zl_poly = np.poly1d(zl)
        y_find.append(zl_poly(x_find_i))

    y_res = max(y_find)
    x_res = x_find[y_find.index(y_res)]

    output = [x_res, x_res+hrange, y_res]

    return output


def size_finder_range(x_h, left_y, right_y, static_y, hf, ha):
    y_find = []
    x_find = []
    x_find_a = []
    # Find hf y point
    zl = np.polyfit(x_h, left_y, 1)
    zlf = np.poly1d(zl)
    hf_y = zlf(hf)

    for y_ar in right_y:
        zr = np.polyfit(x_h, y_ar, 1)
        zrf = np.poly1d(zr)
        ha_y = zrf(ha)
        y_find.append(ha_y)
        x_find.append(ha)

    for y_ar in static_y:
        y_find.append(zlf(hf))
        x_find.append(y_ar)

    # Find the highest right y
    if max(y_find) > hf_y:
        output = [(zlf - max(y_find)).roots[0], x_find[y_find.index(max(y_find))], max(y_find)]
    else:
        for y_ar in right_y:
            zr = np.polyfit(x_h, y_ar, 1)
            zrf = np.poly1d(zr)
            ha_x = (zrf - hf_y).roots[0]
            x_find_a.append(ha_x)

        for y_ar in static_y:
            x_find_a.append(y_ar)
        #print(x_find_a)
        output = [hf, min(x_find_a), hf_y]

    return output


def h_finder_from_mac(mac_pos):
    result = h0-0.25+mac_pos
    return result


def mac_finder_from_h(h):
    return (h-h0)+0.25


def v_bar(size):
    tail_moment_arm = (lvtp_cad_value / c_bar) - h0
    print("v_bar: " + str(size*tail_moment_arm))


def plotit(r1, r2, search_mac, ylims=None):
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
    plt.plot(x_h, takeOffRotation(x_h), label='Take Off')
    y_tails.append(min(takeOffRotation(x_h)))
    y_heads.append(max(takeOffRotation(x_h)))

    plt.plot(x_h, kn(x_h), label='Kn > 0.05')
    y_tails.append(min(kn(x_h)))
    y_heads.append(max(kn(x_h)))

    plt.plot(x_h, landing(x_h), label='Landing')
    y_tails.append(min(landing(x_h)))
    y_heads.append(max(landing(x_h)))

    """This section constrains the graph correctly"""
    if ylims is None:
        max_y = myceil(max(y_heads), step)
        #print(max_y)
        min_y = myfloor((min(y_tails)), step)
        #print(min_y)
    else:
        max_y = ylims[1]
        min_y = ylims[0]

    plt.ylim(min_y, max_y)
    plt.xlim(r1, r2-0.1)

    """Reference Plots
    This section plots reference points for the current aircraft configuration absed on data from the CAD param file"""
    ref_tail = min_y
    ref_head = ((max_y-ref_tail)*0.05) + ref_tail

    # Plot the noseWheel conditions
    plt.plot([noseWheel(), noseWheel()], [min_y, max_y], label='Nose-Wheel Reaction')

    # Plot h0 position
    plt.plot([h0, h0], [ref_tail, ref_head])
    plt.annotate("Wing h0 Position", [h0, ref_head])
    # Plot MTOW CoG position (approx)
    #plt.plot([mtow_pos, mtow_pos], [ref_tail, ref_head+(ref_head-ref_tail)])
    #plt.annotate("MTOW C.o.G", [mtow_pos, ref_head+(ref_head-ref_tail)])
    # Plot Wing LE/TE position (approx)
    wingLE = (params['WingCentrePoint']-(params['Cr']/2))/c_bar
    wingTE = (params['WingCentrePoint'] + (params['Cr'] / 2)) / c_bar
    #plt.plot([wingLE, wingLE], [ref_tail, ref_head])
    #plt.annotate("Wing LE", [wingLE, ref_head])
    #plt.plot([wingTE, wingTE], [ref_tail, ref_head])
    #plt.annotate("Wing TE", [wingTE, ref_head])
    #Main Gear Position
    mainpos = params['MainGearPos']/c_bar
    plt.plot([mainpos, mainpos], [ref_tail, ref_head])
    plt.annotate("Main Gear Pos", [mainpos, ref_head])

    """Result Plots
    This section plots the results of the size finder. There are two options to call. One finds via a delta h range,
    the other via known fwd and aft positions"""
    # Known Delta h range
    #size = size_finder_delta(x_h, takeOffRotation(x_h), [kn(x_h)], [noseWheel()], 0.4)
    # Compare landing and t/o case
    sizes = []
    sizes_results = []
    search_f = search_mac[0]
    search_a = search_mac[1]
    sizes_results.append(size_finder_range(x_h, takeOffRotation(x_h), [kn(x_h)], [noseWheel()], h_finder_from_mac(search_f),
                                           h_finder_from_mac(search_a)))
    sizes.append(sizes_results[-1][2])
    sizes_results.append(size_finder_range(x_h, landing(x_h), [kn(x_h)], [noseWheel()], h_finder_from_mac(search_f),
                                           h_finder_from_mac(search_a)))

    sizes.append(sizes_results[-1][2])
    size = sizes_results[sizes.index(max(sizes))]
    v_bar(size[2])
    plt.plot([size[0], size[1]], [size[2], size[2]], 'c')
    plt.annotate("SH/S = " + str(format(size[2], '.2f')) + " : Size = " + str(format(size[2]*params['Sarea'], '.2f')) + "m2", [size[0], size[2]+0.05])
    plt.annotate("h Range: " + str(format(size[0], '.2f')) + " <> " + str(format(size[1], '.2f')), [size[0], size[2]+0.1])
    plt.annotate("% Mac Range: " + str(format(100*mac_finder_from_h(size[0]), '.1f')) + "% <> " + str(format(100*mac_finder_from_h(size[1]), '.1f')) + "%", [size[0], size[2]+0.15])
    print("h Range: " + str(format(size[0], '.2f')) + " <> " + str(format(size[1], '.2f')))
    print("% Mac Range: " + str(format(mac_finder_from_h(size[0]), '.2f')) + " <> " + str(format(mac_finder_from_h(size[1]), '.2f')))
    print("SH/S = " + str(format(size[2], '.2f')) + " : Size = " + str(format(size[2]*params['Sarea'], '.2f')) + "m2")
    print("Min Kn = " + str(format(get_kn(size[2],size[1]), '.2f')))
    plt.grid(True)
    plt.legend(loc='upper right', shadow=True)
    plt.title("CTA2-100 HTP Scissor Plot", size=20)
    plt.savefig("tailplot.png")
    plt.show()


plotit(4.8, 6.25, [0.07, 0.45], ylims=[-0.4, 0.6])
