
## Code Analysis

This code snippet is designed to perform statistical analysis on a dataset provided in an Excel file. It includes functions to calculate mean, standard deviation, study variation, tolerance, number of measurements, t-statistics, Cg, Cgk, and percent variation. Additionally, it contains a function to generate a graphical representation (a run chart) of the data along with a table showing various calculated metrics. Below, we break down the functionality, prerequisites, and usage instructions for the code.

### Functionality Overview

1. **Statistical Calculations:**
   - The functions `calculate_mean`, `calculate_std`, `calculate_stdy_var`, `calculate_tolerance`, `calculate_num_meas`, `calculate_t_stats`, `calculate_Cg`, `calculate_Cgk`, and `calculate_percent_var` are designed to perform specific statistical calculations on a column of data from an Excel spreadsheet. These functions require the file path to the spreadsheet, the sheet name, and the column index as inputs. Some functions also require additional parameters like a reference value or a constant \(K\).

2. **Graph Generation:**
   - The `generate_graph` function reads the dataset, uses the statistical functions to calculate various metrics, and then generates a run chart for a specified column from the dataset. It highlights the Upper Specification Limit (USL) and Lower Specification Limit (LSL) and includes a table below the chart with calculated metrics.

### Prerequisites

- **Python Libraries:** To run this code, you need to have Python installed along with the pandas, matplotlib, numpy, and PIL (Python Imaging Library) libraries. These can be installed via pip, Python's package installer, with the following commands:
  ```sh
  pip install pandas matplotlib numpy Pillow
  ```
- **Excel File:** The dataset should be in an Excel file (.xlsx) with the data organized such that each column can be analyzed independently based on the column index.

### Installation and Usage Instructions

1. **Install Required Libraries:** First, ensure that you have installed the necessary libraries mentioned above.

2. **Prepare the Dataset:** Make sure your dataset is in an Excel file and properly formatted for analysis. Each column should contain numerical data for the analysis, and the sheet name should match what is specified in the code or be adjusted in the code as needed.

3. **Run the Code:** 
   - You can run the entire script if it is saved as a `.py` file. Alternatively, you can run it in a Jupyter notebook or any Python interactive shell.
   - Modify the `file_path` variable to point to your Excel file. Adjust the `sheet_name` if your data is in a different sheet from the default 'Sheet1'.
   - The `generate_graph` function is already set up to run with a specified column index. If you want to analyze a different column, change the `col_index` parameter accordingly.

4. **View the Results:** After running the code, it will generate a PNG image file (`type_1.png`) showing the run chart and a table with the calculated metrics. A cropped version of this image (`type_1_cropped.png`) is also saved, which may be adjusted depending on the specific dimensions required for your analysis.

### Note on Code Customization

The code has predefined calculations and a graph generation process designed to work with a wide range of datasets. However, depending on your specific requirements, you may need to adjust the calculations (e.g., modifying the formula for `calculate_Cgk` or the definition of study variation) or the graph presentation (e.g., changing the limits or the appearance of the chart).

By following these instructions and understanding the functionality provided by each part of the code, you should be able to perform a detailed statistical analysis of your dataset and visualize key metrics through a run chart.
