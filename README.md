# openpyxl-datautils
This package is primarily used to create (and overwrite) the CellRange object to fulfill two goals:
1. Match the functionality of Excel VBA’s Range object more closely.
1. Embed conversions to and from Pandas DataFrames to allow for easier use in data analysis.

**To better match Excel VBA's Range Object:**
* ```current_region()``` - Expands the range to the right, then down to include adjacent data
* ```values``` - Iterates over all rows in the range to return the values of each cell
* ```cells``` - Returns all cells in the range (useful to set styling of a cell)

**Pandas Integration:**
* ```create_df()``` - Create a Pandas DataFrame with the data contained in a given CellRange
    * For larger data sets this is about **30% faster** then pd.read_excel()
* ```write_df()``` - Write data from a Pandas DataFrame to the values of a CellRange
    * For larger data set this is about **20% faster** then df.to_excel()

### Installation
This package can be found on [PyPI](https://pypi.org/project/openpyxl-datautils/) and can can be installed with ```pip install openpyxl-datautils```

### Documentation
The [openpyxl-datautils wiki](https://github.com/robbymedich/openpyxl-datautils/wiki) provides documentation on the classes and functions contained in the package. Additionally, known issues can be found on GitHub's [Issues](https://github.com/robbymedich/openpyxl-datautils/issues) tab.

### Examples
* For basic examples of how to read and write data using the CellRange object see [basic_demo.ipynb](https://github.com/robbymedich/openpyxl-datautils/blob/main/examples/basic_demo.ipynb)
* Load data from all sheets in a workbook:
    ```python
    from openpyxl_datautils import load_workbook
    from openpyxl_datautils import CellRange

    with load_workbook('source_data.xlsx') as workbook:
        sheet_data = {}
        for worksheet in workbook.sheetnames:
            sheet_data[worksheet] = CellRange.from_string(workbook[worksheet], 'A1').create_df(expand_range=True)
    ```
