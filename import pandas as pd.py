from openpyxl import load_workbook
import pandas as pd
import os
import matplotlib.pyplot as plt
import numpy as np
from tkinter import messagebox as mb

def normalize(cell_arr):
    intensity_data = [float(val) for val in cell_arr]
    max_intensity = max(intensity_data)
    normalized_data = [intensity / max_intensity for intensity in intensity_data]
    return normalized_data

# File folder location
file_names = os.listdir("Mg_Sheets")
xANDy = {}
offset = 0



for excel_names in file_names:
    # Data frame for excel sheets
    dataFrame = pd.read_csv("Mg_Sheets/" + excel_names)
    
    # Angle
    degree_data = dataFrame.iloc[26:1550, 0].values
    degree_data = [float(angle) for angle in degree_data]  # Convert degree strings to float
    
    # Intensity
    intensity_data = dataFrame.iloc[26:1550, 2].values
    intensity_data = normalize(intensity_data)
    
    # Hashmap contains folder name as key, contains 2d array for intensity angle
    xANDy[excel_names] = [degree_data, intensity_data]
    
    offset_data2 = [val + offset for val in intensity_data]
    plt.plot(degree_data, offset_data2, label=excel_names)
    
    offset += 0.5

# Labels for graph
plt.xlabel("Angle")
plt.ylabel("Normalized Intensity")
plt.legend()
plt.title("Mg Data")
plt.show()

res = mb.askquestion("Save Data?", "Do you want to save this to an excel file(s)")

if res == "yes":
    output_folder = "Output_CSVs"
    os.makedirs(output_folder, exist_ok=True)
    
    for key, value in xANDy.items():
        csv_name = os.path.splitext(key)[0] + "_normalized.csv"
        output_path = os.path.join(output_folder, csv_name)
        a = np.array(value[0])
        b = np.array(value[1])
        
        df = pd.DataFrame({key.replace('.csv', '') + "Degree": a, key.replace('.csv', '') + "Intensity": b})
        df.to_csv(output_path, index=False)
