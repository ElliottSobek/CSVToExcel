[![Generic badge](https://img.shields.io/badge/development%20status-in%20development-red.svg "Development Status")](https://shields.io/)
    
# CSVToExcel

Converts one or multiple CSV files into one Excel Workbook

## Requirements

Requirements can be installed using pip/pip3.

XlsxWriter

### Example

pip install XlsxWriter

pip3 install XlsxWriter

### Options

* Print help menu (-h)
* Print program version (--version)
* Store strings as numeric values (-s)
* Force file writing even if input files are empty (-f)
* Suppress the startup banner

## Worksheet Naming

The passed in csv files are divided into their own worksheet(s) when written to the
excel file. The basename of each file is used as the worksheet name to help
organize data.

### Empty Files

When the force option is enabled, empty files are processed and added as new
worksheets with no data in the worksheet.

## Contribution

Anyone can contribute and work on any available issues.

Keep the same format as is used. After, simply make a pull request and have
patience.
