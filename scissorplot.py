from matplotlib import pyplot
import numpy as np
import pandas as pd

# Excel Sheet Loading
cad_file_path = '../../../../2. TechInt/Project Space/CAD/CADParametersNewAero.xlsx'
cad_file = pd.ExcelFile(cad_file_path)
print(cad_file.sheet_names)

cad_params = pd.read_excel(cad_file_path, sheet_name=0, index_col=0, skiprows=5, usecols="A:B")
#print(cad_params.head(10))

print(cad_params.shape)
