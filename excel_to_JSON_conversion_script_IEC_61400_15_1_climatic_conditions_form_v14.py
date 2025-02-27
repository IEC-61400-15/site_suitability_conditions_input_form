"""Quick script to convert the Turbine columns to JSON inputs by wind direction bin.
This provides a simple way to convert all the turbine values to a list of lists in
the example notebook so the results can be copied and pasted into the template for each
sheet.

HOW TO RUN: python quick_convert.py <input>.csv <output>.json

NOTE: This is not a proper JSON output, just one that can be copied and pasted easily
"""

import sys

import numpy as np
import pandas as pd
import xlwings as xw


INDENT = "    "
INDENT2 = f"{INDENT}{INDENT}"
INDENT3 = f"{INDENT}{INDENT}{INDENT}"
INDENT4 = f"{INDENT}{INDENT}{INDENT}{INDENT}"

# NOTE: This is fixed to the original data sample, and must be updated for any changes
N_ws = 40
wd = np.arange(0, 360, 30)
N_samples = [40] * wd.size
cumsum_samples_lower = [sum(N_samples[:i]) for i in range(len(N_samples))]
cumsum_samples_upper = [
    sum(N_samples[:i + 1]) if i < len(N_samples) else sum(N_samples)
    for i in range(len(N_samples))
]

sheet_names_ranges = [
    ("Turbine Layout Summary", "A1:W10", [0], [9]),
    ("WS frequency","C3:R495", cumsum_samples_lower, cumsum_samples_upper),
    ("WS Weibull", "C3:O41", [0, 13, 26], [13, 26, 38]), # Needs an ALL Row (first element of the first two rows)
    ("Ambient Mean TI", "C3:O536", cumsum_samples_lower, cumsum_samples_upper),  # Needs an ALL Row
    ("SD TI", "C3:O536", cumsum_samples_lower, cumsum_samples_upper),  # Needs an ALL Row
    ("Extreme Ambient TI", "C3:O44", [0], [41]),  # Is only an ALL Row
    ("Temperature", "B3:G95", [0], [91]),  # Columns have to be updated
    ("Shear", "C3:O16", [0], [13]),  # Only an ALL row
    ("Inflow Angle", "B3:N16", [0], [13]),  # All, Max, values
    ("CcT", "B3:N6", [0], [3]),  # Only an ALL row
]


def convert_column_to_dict(data, lower, upper):
    """Converts the column of values into the proper list of list sizes."""
    return [data[i: j] for i, j in zip(lower, upper)]


def read_sheet_data(wb, sheet_name, data_range, lower, upper):
    data_dict = {}
    excel_sheet = wb.sheets[sheet_name]

    data = excel_sheet.range(data_range).options(pd.DataFrame, index=False).value.fillna(0)
    for i, col in enumerate(data.columns):
        if col in data_dict:
            data_dict[f"{col}{i}"] = convert_column_to_dict(data.iloc[:, i].tolist(), lower, upper)
        else:
            data_dict[col] = convert_column_to_dict(data.iloc[:, i].tolist(), lower, upper)
    
    return data_dict


if __name__ == "__main__":
    fn_read = sys.argv[1]
    fn_save = sys.argv[2]
    data = xw.Book(fn_read)
    data_dict = {}
    for name, data_range, lower, upper in sheet_names_ranges:
        data_dict[name] = read_sheet_data(data, name, data_range, lower, upper)

    with open(fn_save, "w") as f:
        f.write("{\n")
        for i, (sheet, values) in enumerate(data_dict.items()):
            f.write(f'{INDENT}"{sheet}": {{\n')
            for j, (col, col_values) in enumerate(values.items()):
                f.write(f'{INDENT2}"{col}": {{\n')
                f.write(f'{INDENT3}"{sheet}": [\n')
                for k, lst in enumerate(col_values):
                    try:
                        lst_str = ", ".join([f"{el:.18f}" for el in lst])
                    except ValueError:
                        # The summary data that is strings, not float data
                        lst_str = ", ".join([f'"{el}"' for el in lst])
                    if k == len(col_values) - 1:
                        f.write(f"{INDENT4}[{lst_str}]\n")
                    else:
                        f.write(f"{INDENT4}[{lst_str}],\n")
                f.write(f"{INDENT3}]\n")
                if j == len(values) - 1:
                    f.write(f"{INDENT2}}}\n")
                else:
                    f.write(f"{INDENT2}}},\n")
                
            if i == len(data_dict) - 1:
                f.write(f"{INDENT}}}\n")
            else:
                f.write(f"{INDENT}}},\n")
        f.write("}")
