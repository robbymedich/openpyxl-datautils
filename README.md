# openpyxl-datautils
This package primarily is used to create (and overwrite) the CellRange object to fulfill two goals:
1. Match the functionality of Excel VBA’s Range object more closely
1. Embed conversions to and from Pandas DataFrames to allow for easier use in data analysis

Example:
```python
from openpyxl_datautils import load_workbook
from openpyxl_datautils import CellRange

with load_workbook('source_data.xlsx') as workbook:
    sheet_data = {}
    for worksheet in workbook.sheetnames:
        sheet_data[worksheet] = CellRange.from_string(workbook[worksheet], 'A1').create_df(expand_range=True)
```
