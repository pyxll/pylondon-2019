"""
Code for reading a dataset into a Pandas DataFrame
"""
from pyxll import xl_func
import pandas as pd
import os


@xl_func("str name, dict dtypes, str[] parse_dates, bool infer_datetime_format: object")
def read_csv_data(name, dtype=None, parse_dates=None, infer_datetime_format=False):
    """Reads a csv file into a pandas DataFrame"""

    # construct the file name and check it exists
    filename = os.path.join(os.path.dirname(__file__), "..", "data", name + ".csv")

    if not os.path.exists(filename):
        raise IOError(f"File '{filename}' does not exists")

    # read the csv file into a DataFrame
    df = pd.read_csv(filename,
                     dtype=dtype,
                     parse_dates=parse_dates,
                     infer_datetime_format=infer_datetime_format)

    return df
