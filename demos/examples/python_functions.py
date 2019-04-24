"""
Example code showing how to declare a Python function.
"""
from pyxll import xl_func


@xl_func
def simple_test(a, b):
    return a + b
