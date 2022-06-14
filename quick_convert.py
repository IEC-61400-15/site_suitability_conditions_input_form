"""Quick script to convert the Turbine columns to JSON inputs by wind direction bin.
This provides a simple way to convert all the turbine values to a list of lists in
the example notebook so the results can be copied and pasted into the template for each
sheet.

HOW TO RUN: python quick_convert.py <input>.csv <output>.json

NOTE: This is not a proper JSON output, just one that can be copied and pasted easily
"""

import sys
from pprint import pprint

import pandas as pd


# NOTE: This is fixed to the original data sample, and must be updated for any changes
N_samples = (11, 21, 20, 18, 14, 18, 23, 23, 20, 21, 22, 21, 22)
cumsum_samples_lower = [sum(N_samples[:i]) for i in range(len(N_samples))]
cumsum_samples_upper = [
    sum(N_samples[:i + 1]) if i < len(N_samples) else sum(N_samples)
    for i in range(len(N_samples))
]


def convert_column_to_dict(data):
    """Converts the column of values into the proper list of list sizes."""
    return [data[i: j] for i, j in zip(cumsum_samples_lower, cumsum_samples_upper)]


if __name__ == "__main__":
    fn_read = sys.argv[1]
    fn_save = sys.argv[2]
    data = pd.read_csv(fn_read)
    data_dict = {}
    for col in data.columns:
        data_dict[col] = convert_column_to_dict(data[col].tolist())

    with open(fn_save, "w") as f:
        for k, v in data_dict.items():
            f.write(f'"{k}": {{\n')
            for i, lst in enumerate(v):
                # Consistently format the list to a string format
                lst_str = ", ".join([f"{el:.18f}" for el in lst])
                if i == 0:
                    f.write(f'    "WS Frequency all directions": [{lst_str}],\n')
                    f.write(f'    "WS Frequency": [\n')
                    continue
                if i == len(v) - 1:
                    f.write(f"        [{lst_str}]\n")
                else:
                    f.write(f"        [{lst_str}],\n")
            # Close the list of lists bracket
            f.write("    ]\n},\n")
