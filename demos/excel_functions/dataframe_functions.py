"""
Collection of functions for manipulating DataFrames in Excel.
"""
from pyxll import xl_func, get_type_converter
import pandas as pd


@xl_func
def df_empty():
    return pd.DataFrame()


@xl_func("dataframe, int, str[], int: dataframe", auto_resize=True)
def df_head(df, n, columns=[], offset=0):
    """Return the first n rows of a DataFrame."""
    columns = [c for c in columns if c]
    if columns:
        df = df[columns]
    if offset:
        df = df.iloc[offset:]
    return df.head(n)


@xl_func("dataframe, int, str[], int: dataframe", auto_resize=True)
def df_tail(df, n, columns=[], offset=0):
    """Return the last n rows of a DataFrame."""
    columns = [c for c in columns if c]
    if columns:
        df = df[columns]
    if offset:
        df = df.iloc[offset:]
    return df.tail(n)


@xl_func("dataframe, str[], int: object")
def df_drop(df, columns, axis=1):
    """Drop columns from a dataframe"""
    columns = [c for c in columns if c is not None]
    return df.drop(columns, axis=axis)


@xl_func("dataframe df, float[] percentiles, string[] include, string[] exclude: dataframe<index=True>",
         auto_resize=True)
def df_describe(df, percentiles=None, include=None, exclude=None):
    """Describe a pandas DataFrame"""
    return df.describe(percentiles=percentiles,
                       include=include,
                       exclude=exclude)


@xl_func("dataframe, dict: object")
def df_eval(df, exprs):
    """Evaluate a string describing operations on DataFrame columns."""
    new_columns = {}
    for key, expr in exprs.items():
        new_columns[key] = df.eval(expr)

    new_df = pd.DataFrame(new_columns)
    return df.join(new_df)


@xl_func("dataframe, str, int: object")
def df_apply(df, func, axis=0):
    return df.apply(func, axis=axis)


@xl_func("dataframe, var, int: object")
def df_divide(df, x, axis=0):
    return df.divide(x, axis=axis)


@xl_func("dataframe, var: object")
def df_multiply(df, x, axis=0):
    return df.multiply(x, axis=axis)


@xl_func("dataframe, var, var: object")
def df_add(df, x, axis=0):
    return df.add(x, axis=0)


@xl_func("dataframe, var, var: object")
def df_subtract(df, x, axis=0):
    return df.subtract(x, axis=0)


@xl_func("dataframe, str[], str[], str[], var: object")
def df_pivot_table(df, index, columns=None, values=None, aggfunc="mean"):
    """Return reshaped DataFrame organized by given index / column values."""
    if isinstance(aggfunc, list):
        to_dict = get_type_converter("var", "dict")
        aggfunc = to_dict(aggfunc)

    df = df.reset_index()
    return df.pivot_table(index=index,
                          columns=columns,
                          values=values,
                          aggfunc=aggfunc)


@xl_func("dataframe: object")
def df_stack(df):
    return df.stack()
