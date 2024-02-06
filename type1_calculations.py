import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from PIL import Image

#MEAN
def calculate_mean(file_path, sheet_name='Sheet1', col_index=None):
    # Read the Excel file
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    # Validate col_index
    if col_index is None or col_index >= len(df.columns):
        raise ValueError("Column index is required and must be within the range of available columns.")
    
    # Extract column name using col_index
    column_name = df.columns[col_index]

    # Calculate the mean for the specified column
    mean_value = df.iloc[:, col_index].mean()
    
    # Return a dictionary with the column name and mean value
    return {'Mean': mean_value}

#STD DEVIATION
def calculate_std(file_path, sheet_name='Sheet1', col_index=None):
    # Read the Excel file
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    # Validate col_index
    if col_index is None or col_index >= len(df.columns):
        raise ValueError("Column index is required and must be within the range of available columns.")
    
    # Extract column name using col_index
    column_name = df.columns[col_index]

    # Calculate the standard deviation for the specified column
    std_value = df.iloc[:, col_index].std(ddof=1)  # ddof=1 for sample standard deviation
    
    # Return a dictionary with the column name and standard deviation value
    return {'Std Deviation': std_value}

#STUDY VARIATION
def calculate_stdy_var(file_path, sheet_name='Sheet1', col_index=None):
    # Read the Excel file
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    # Validate col_index
    if col_index is None or col_index >= len(df.columns):
        raise ValueError("Column index is required and must be within the range of available columns.")
    
    # Extract column name using col_index
    column_name = df.columns[col_index]

    # Calculate the study variation for the specified column
    study_variation = df.iloc[:, col_index].std(ddof=1) * 6
    
    # Return a dictionary with the column name and study variation value
    return {'Study Variable': study_variation}

#TOLERANCES
def calculate_tolerance(file_path, sheet_name='Sheet1', col_index=None):
    # Read the Excel file
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # Validate col_index
    if col_index is None or col_index >= len(df.columns):
        raise ValueError("Column index is required and must be within the range of available columns.")
    
    # Extract column name using col_index
    column_name = df.columns[col_index]

    # Calculate the mean and study variation for the specified column
    mean_value = calculate_mean(file_path, sheet_name, col_index)['Mean']
    std_dev_value_x_6 = calculate_stdy_var(file_path, sheet_name, col_index)['Study Variable']

    # Calculate USL and LSL for the specified column
    tolerance = {
        'USL': mean_value + std_dev_value_x_6,
        'LSL': mean_value - std_dev_value_x_6
    }

    # Return a dictionary with the column name and its tolerance values
    return {'Tolerance': tolerance}

#NUMBER OF MEASUREMENTS
def calculate_num_meas(file_path, sheet_name='Sheet1', col_index=None):
    # Read the Excel file
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    # Validate col_index
    if col_index is None or col_index >= len(df.columns):
        raise ValueError("Column index is required and must be within the range of available columns.")
    
    # Extract column name using col_index
    column_name = df.columns[col_index]

    # Count non-NA cells for the specified column
    num_meas = df.iloc[:, col_index].count()
    
    # Return a dictionary with the column name and number of measurements
    return {'# of Measurments': num_meas}

#T-STATISTICS
def calculate_t_stats(file_path, sheet_name='Sheet1', ref_value=None, col_index=None):
    # Read the Excel file
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # Validate col_index and ref_value
    if ref_value is None:
        raise ValueError("Reference value parameter is required.")
    if col_index is None or col_index >= len(df.columns):
        raise ValueError("Column index is required and must be within the range of available columns.")
    
    # Extract column name using col_index
    column_name = df.columns[col_index]

    # Calculate the number of measurements
    n = df[column_name].count()

    # Calculate the mean and standard deviation for the specified column
    mean_x = df[column_name].mean()
    s = df[column_name].std(ddof=1)
    
    # Calculate the t-score for the specified column
    t_stat = (np.sqrt(n) * abs(mean_x - ref_value)) / s if s != 0 else float('nan')

    # Return a dictionary with the column heading and the corresponding t-score
    return {'t-Statistics': t_stat}

#Cg
def calculate_Cg(file_path, sheet_name='Sheet1', col_index=None, K=20, tolerance=None):
    # Read the Excel file
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    # Validate col_index
    if col_index is None or col_index >= len(df.columns):
        raise ValueError("Column index is required and must be within the range of available columns.")
    if tolerance is None:
        raise ValueError("Tolerance is required.")

    # Use the previously defined functions to calculate mean, study variation, and tolerance
    SV = calculate_stdy_var(file_path, sheet_name, col_index)['Study Variable']
    
    # Calculate Cg for the specified column
    Cg = (K / 100) * ( tolerance/ SV)

    # Return a dictionary with the column name and its Cg value
    return {'Cg': Cg}

#Cgk
def calculate_Cgk(file_path, sheet_name='Sheet1', col_index=None, K=20, ref_value=None):
    # Read the Excel file
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    # Validate col_index and ref_value
    if col_index is None or col_index >= len(df.columns):
        raise ValueError("Column index is required and must be within the range of available columns.")
    if ref_value is None:
        raise ValueError("Reference value parameter is required but was not provided.")
    
    # Use the previously defined functions to calculate mean, study variation, and tolerance
    mean_value = calculate_mean(file_path, sheet_name, col_index)['Mean']
    tolerance = calculate_tolerance(file_path, sheet_name, col_index)['Tolerance']
    SV = calculate_stdy_var(file_path, sheet_name, col_index)['Study Variable']
    
    # Tolerance is the difference between USL and LSL
    tolerance_range = tolerance['USL'] - tolerance['LSL']
    
    # Calculate Cgk for the specified column
    bias = abs(mean_value - ref_value)
    Cgk = ((K / 200.0) * tolerance_range - bias) / (SV / 2)
    
    # Return a dictionary with the column name and its Cgk value
    column_name = df.columns[col_index]
    return {'Cgk': Cgk}

#%Var
def calculate_percent_var(file_path, sheet_name='Sheet1', col_index=None, K=20):
    # Read the Excel file
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    # Validate col_index
    if col_index is None or col_index >= len(df.columns):
        raise ValueError("Column index is required and must be within the range of available columns.")
    
    # Extract column name using col_index
    column_name = df.columns[col_index]

    # Calculate the study variation and tolerance for the specified column
    SV = calculate_stdy_var(file_path, sheet_name, col_index)['Study Variable']
    tolerance_values = calculate_tolerance(file_path, sheet_name, col_index)
    tolerance_range = tolerance_values['Tolerance']['USL'] - tolerance_values['Tolerance']['LSL']
    
    # Calculate % Var for the specified column
    percent_var = (SV / tolerance_range) * 100
    
    # Return a dictionary with the column name and its % Var and Cg value
    return {'% Var': percent_var}

#Graph generation
def generate_graph(col_index = None, ref_value=None, tol=None):
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # Validate col_index
    if col_index is None:
        raise ValueError("Column index is required.")
    # Validate tolerance
    if tol is None:
        raise ValueError("Tolerance  is required.")
    # Validate ref_value
    if ref_value is None:
        raise ValueError("Refrence Value is required.")
    if col_index < 0 or col_index >= len(df.columns):
        raise ValueError("Column index is out of range.")
    
    #calling all the functions
    means = calculate_mean(file_path,col_index=col_index)
    std = calculate_std(file_path,col_index=col_index)
    stdy_var = calculate_stdy_var(file_path,col_index=col_index)
    tolerance = calculate_tolerance(file_path,col_index=col_index)
    num_meas = calculate_num_meas(file_path,col_index=col_index)
    t_stat = calculate_t_stats(file_path,col_index=col_index,ref_value=ref_value)
    Cg = calculate_Cg(file_path,col_index=col_index, tolerance=tol)
    Cgk = calculate_Cgk(file_path,col_index=col_index,ref_value=255)
    var_percent = calculate_percent_var(file_path,col_index=col_index)

    results = {
        "Mean": means['Mean'],
        "Std Deviation": std['Std Deviation'],
        "Study Variation": stdy_var['Study Variable'],
        "USL": tolerance['Tolerance']['USL'],
        "LSL": tolerance['Tolerance']['LSL'],
        "# of Measurements": num_meas['# of Measurments'],
        "t-Statistics": t_stat['t-Statistics'],
        "Cg": Cg['Cg'],
        "Cgk": Cgk['Cgk'],
        "% Var": var_percent['% Var']
    }

    USL = tolerance['Tolerance']['USL']
    LSL = tolerance['Tolerance']['LSL']
    range_buffer_percentage = 0.30  # Adjust the buffer percentage as needed
    range_buffer = (USL - LSL) * range_buffer_percentage
    new_LSL = LSL - range_buffer
    new_USL = USL + range_buffer
    values = [f"{value:.2f}" if isinstance(value, float) else str(value) for value in results.values()]
    plt.figure(figsize=(12, 10))
    plt.subplot(211)  # Adjust this to create space for the table
    plt.plot(df.iloc[:,0], df.iloc[:,col_index], marker='o')
    plt.title(f'Type 1 Gage Study for {df.columns[col_index]}')
    plt.xlabel('Observations')
    plt.ylabel('Data')
    plt.ylim(new_LSL,new_USL)
    plt.axhline(y=LSL, color='r', linestyle='--', label='LSL')
    plt.axhline(y=USL, color='r', linestyle='--', label='USL')
    plt.axhline(y=ref_value, color='g', linestyle='-', label='USL', lw=0.5)
    plt.grid(True)
    plt.legend()

    # Prepare table data
    metrics = ["Mean", "Std Deviation", "Study Variation (6*Std Deviation)", "# of Measurements", "t-Statistics", "Cg", "Cgk", "% Var"]
    values = [means["Mean"], std["Std Deviation"], stdy_var["Study Variable"], num_meas["# of Measurments"], t_stat["t-Statistics"], Cg["Cg"], Cgk["Cgk"], var_percent["% Var"]]

    # Convert all numeric values to formatted strings for display
    formatted_values = [f"{value:.5f}" if isinstance(value, float) else str(value) for value in values]

    # Adding the table below the run chart
    table = plt.table(cellText=list(zip(metrics, formatted_values)),
            colLabels=['Metric', 'Value'],
            cellLoc='center', 
            loc='bottom', 
            bbox=[0.125, -1.2, 0.75, 1])  # Adjust bbox as needed to fit the table properly

    for (i, j), cell in table.get_celld().items():
        if i == 0:  # The first row, i.e., the column labels
            cell.set_fontsize(12)  # Optional: Adjust fontsize as needed
            cell.set_text_props(fontweight='bold')  # Making the text bold

    # Adjust layout
    plt.subplots_adjust(left=0.2, bottom=0.2, top=0.95)
    plt.savefig('./Images/type_1.png', format='png', dpi=300)
    plt.show()

    # Open the image file
    img = Image.open('./Images/type_1.png')
    left = 350
    top = 0
    right = 3510
    bottom = 2550
    cropped_img = img.crop((left, top, right, bottom))

    # Save the cropped image
    cropped_img.save(f'./Images/Graph_Col{col_index}.png')

file_path = 'Data.xlsx'
sheet_name = 'Sheet1'

generate_graph(col_index=4, ref_value=470, tol=20)
