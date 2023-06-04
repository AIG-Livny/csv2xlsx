# csv2xlsx
This script is designed to converting CSV (comma separated values) table to XLS or XLSX Calc or Excell format.

## Features
- Case insensibility
- Absolute ot relative paths
- Support XLS and XLSX output formats
- Support special ESV (extended) format. 

ESV is extension for usual CSV files for creating multisheet books. Under the hood it is JSON with new line character support, so you can easly pull out CSV tables by hand. ESV was made for creating multisheets from one text file.

Example of ESV format in repository.

# Usage
You can select XLS or XLSX by extension output file. And use CSV or ESV as input
```
csv2xlsx.py input.csv output.xls
csv2xlsx.py input.csv output.xlsx
csv2xlsx.py input.esv output.xlsx
```


# Install
Required python version equal or greater than 3.6
```
pip install -r requirements.txt
```

