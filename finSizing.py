import matplotlib.pyplot as plt
import numpy as np
from openpyxl import load_workbook
import filepath1

# Excel Sheet Loading
cad_file_path = filepath1.cad_file_path
# cad_file_path = "B:\OneDrive\Group Business Design Project/2. TechInt\Project Space\CAD\CADParametersNewAero.xlsx"
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
dcnWBN_db = -0.462  # per rad
dcyV_db = -3.726  # per rad
dcyV_dr = 1.892  # per rad
c_bar = cad_file['Interface']['B66'].value
print(c_bar)
maxthrust = 44459  # N @ takoff
cruise_thrust = 9919  # N
vto = 62.4  # m/s
vc = 65.9
cthrust = maxthrust / qS(vto)
cthrust_airborne = cruise_thrust / qS(vc)
dr = np.deg2rad(30)
Kr = 0.9
blade_count = 6
blade_diameter = params['EnginePropDiameter']
blade_centre = params['EngineWingPosOutboard']
engine_scale = 0.89
engine_intake_area = 0.16 * engine_scale
engine_diameter = np.sqrt(engine_intake_area/np.pi)*2
h0 = cad_file['Interface']['B67'].value
h0_centre = params["WingChordStart"]/c_bar
print(h0)
max_bank_angle = np.deg2rad(5)
cl_vmca = 2.549
mtow_pos = cad_file['Interface']['F16'].value/c_bar  # From Mass CG File
lvtp_cad_value = cad_file['Interface']["B89"].value
fin_pos_cad_value = cad_file['Interface']["B80"].value


def c_drag_engine():
    cd_wind = 0.1 * (engine_diameter**2)
    cd_prop = 0.00125 * blade_count * blade_diameter**2
    cd_engine_total = cd_prop + cd_wind
    return 0.02


def take_off_yaw():
    lvtp = ((fin_pos_cad_value / c_bar) - h0_centre) * c_bar
    lhs_top = (cthrust + c_drag_engine())*(blade_centre/c_bar)
    lhs_bottom = (dr * Kr * dcyV_dr * lvtp) / c_bar
    return lhs_top/lhs_bottom


def airborne_combined(h):

    lvtp = ((fin_pos_cad_value / c_bar) - h0_centre) * c_bar
    print(lvtp)
    eq_a = dcnWBN_db + (dcyWBN_db * (h0_centre-h))  # Joe D
    print("a: " + str(eq_a))
    eq_b = dcyV_dr * Kr * dr  # Joe C
    print("b: " + str(eq_b))
    eq_c = (cthrust + c_drag_engine()) * blade_centre/c_bar  # Joe B
    print("c: " + str(eq_c))
    eq_d = dcyV_db * (lvtp/c_bar)  # Joe E
    print("d: " + str(eq_d))
    eq_e = (lvtp/c_bar) * eq_b  # Joe a
    print("e: " + str(eq_e))
    eq_f = (dcyV_db * eq_e) - (eq_d * eq_b)
    print("f: " + str(eq_f))
    eq_g = (eq_c * dcyV_db)-(eq_b * eq_a) - (eq_e * dcyWBN_db) + (eq_d * cl_vmca * max_bank_angle)
    print("g: " + str(eq_g))
    eq_h = (eq_c * dcyWBN_db)-(cl_vmca * eq_a * max_bank_angle)
    print("h: " + str(eq_h))
    return (eq_h/eq_g)


def oei_sideforce():
    beta = np.deg2rad(5)
    lhs_top = -(dcyWBN_db*beta) - (cl_vmca * max_bank_angle)
    lhs_bottom = (dcyV_db * beta) + (dcyV_dr * Kr * dr)
    return lhs_top/lhs_bottom


def crosswind_landing(h):
    vcw = 25*0.515
    vapp = 125*0.515
    beta = np.arctan(vcw/vapp)
    dr_cw = dr * 0.7
    Kr_cw = 0.98
    lvtp = ((fin_pos_cad_value / c_bar) - h0_centre) * c_bar

    lhs_top = beta * (dcnWBN_db + (dcyWBN_db * (h0_centre-h)))
    lhs_bottom = (dcyV_dr * Kr_cw * dr_cw * lvtp/c_bar) + (beta * dcyV_db * lvtp/c_bar)
    return lhs_top/lhs_bottom

def directional_stability(h):
    lvtp = ((fin_pos_cad_value / c_bar) - h0_centre) * c_bar
    cnb = 1.25
    lhs_top = cnb + (dcyWBN_db * (h0_centre-h)) - dcnWBN_db
    lhs_bottom = -dcyV_db * ((lvtp/c_bar) - (h0_centre-h))
    return lhs_top/lhs_bottom


def plotit(r1, r2):
    # Create Range of h values
    step = 0.1
    r2 = r2 + step
    x_h = np.arange(r1, r2, 0.1)
    fig = plt.figure(figsize=(12, 8))
    ax = fig.add_subplot(1, 1, 1)
    ax.xaxis.set_ticks_position('bottom')
    y_tails = []
    y_heads = []
    plt.xlabel("h")
    plt.ylabel("ST/S")
    plt.plot([r1, r2], [take_off_yaw(), take_off_yaw()], label='Take Off Yaw')
    y_tails.append(take_off_yaw())
    y_heads.append(take_off_yaw())

    plt.plot(x_h, airborne_combined(x_h), label='Airborne Case')
    y_tails.append(min(airborne_combined(x_h)))
    y_heads.append(max(airborne_combined(x_h)))

    plt.plot(x_h, crosswind_landing(x_h), label='Crosswind Landing')
    y_tails.append(min(crosswind_landing(x_h)))
    y_heads.append(max(crosswind_landing(x_h)))

    plt.plot(x_h, directional_stability(x_h), label='Directional Stability')
    y_tails.append(min(directional_stability(x_h)))
    y_heads.append(max(directional_stability(x_h)))


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


plotit(5.1, 5.8)
