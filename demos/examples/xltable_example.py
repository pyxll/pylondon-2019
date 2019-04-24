from xltable import *
import pandas as pd

# create a dataframe with three columns where the last is the sum of the first two
df = pd.DataFrame({
        "A": list(range(0, 10)),
        "B": list(range(10, 20)),
        "C": Cell("A") + Cell("B")
})

header_style = CellStyle(bg_color=0x4472C4, text_color=0xFFFFFF, bold=True)

# create the named xlwriter Table instance
table1 = Table("table1", df, header_style=header_style)

# create the Workbook and Worksheet objects and add table to the sheet
sheet = Worksheet("Sheet1")
sheet.add_table(table1)

# Create a second table containing the totals
df2 = pd.DataFrame({
        "Total A": [Formula("SUM", Column("A", table="table1"))],
        "Total B": [Formula("SUM", Column("B", table="table1"))],
        "Total C": [Formula("SUM", Column("C", table="table1"))],
})

style = TableStyle(stripe_colors=(0xFFF2CC,))
table2 = Table("table2", df2, style=style, include_columns=False)

sheet.add_table(table2)

workbook = Workbook("xltable_example.xlsx")
workbook.add_sheet(sheet)

# write the workbook to the file (requires xlsxwriter)
workbook.to_xlsx()
