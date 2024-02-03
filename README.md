
# Type 1 - Capabiltity Study

## Overview

The script is designed for statistical calculations and analysis on data from an Excel file. It includes functionalities for calculating mean, standard deviation, study variation, tolerances, number of measurements, t-statistics, capability indices (Cg, Cgk), and percentage variation.

## Requirements

- **Python Libraries:** pandas, matplotlib, numpy
- **Data Source:** An Excel file with a specific structure and data format.

## Inputs

- An Excel file path and sheet name specified within the script.
- Column indices may need to be provided for certain calculations.

## Functionalities

### Statistical Calculations

- **Mean Calculation:** Calculates the mean of a specified column.
- **Standard Deviation:** Calculates the standard deviation of a specified column.
- **Study Variation:** Defined as 6 times the standard deviation, indicating variability in a manufacturing context.
- **Tolerances:** Calculates upper and lower specification limits based on the data.
- **Number of Measurements:** Counts non-NA cells in a specified column.
- **T-Statistics:** Calculates the t-score for hypothesis testing.
- **Capability Indices (Cg, Cgk):** Measures the capability of a process to produce output within specification limits.
- **Percentage Variation:** Calculates the percentage of variation relative to the specified tolerances.

## Comments and Documentation

The script contains comments detailing each function's purpose and the methodology for calculations, making it easier to understand the logic behind each statistical calculation.

## Conclusion

This Python script offers a comprehensive toolkit for performing detailed statistical analysis on manufacturing data, with a focus on quality control metrics such as capability indices and study variation. Its modular design allows for easy adaptation and extension to different datasets or analytical requirements.
