# Spreadsheet Reader
### Author: Peter Swanson
[![License: GPL v3](https://img.shields.io/badge/License-GPL%20v3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)
[![Python 2.7](https://img.shields.io/badge/Python-2.7-brightgreen.svg)](https://www.python.org/downloads/release/python-2714/)
[![openpyxl 2.5.6](https://img.shields.io/badge/openpyxl-2.5.6-brightgreen.svg)](https://pypi.org/project/openpyxl/)

## Background:
<b>This code allows a user to write large quantities of Python data to spreadsheet rows or columns.</b>

## Using the Code:
### Installing Dependencies:
Ensure the following are installed on the machine you are running the application on:
- Python 2.7 with Pip
- virtualenv for Python 2.7

Create a virtualenv and install the requirements from <i>requirements.txt</i> with pip:
```
$ virtualenv venv
$ source venv/bin/activate
(venv)$ pip install -r "requirements.txt"
``` 

### Reading Spreadsheets:
Spreadsheets can be opened by instantiating a <i>Spreadsheet</i> object with the name of the sheet.
The extension is optional.
```
>>> from Spreadsheet import Spreadsheet
>>> sheet = Spreadsheet("test_sheet.xlsx")
```
Cells can be read with the <i>Spreadsheet.read_row()</i> and <i>Spreadsheet.read_column()</i>
functions. Functions return a dictionary containing the read values. 
Row and column indices begin at 1.
```
>>> print sheet.read_row(row=1, start_col=1)
{u'Row 1': [u'Row', u'Text']}

>>> print sheet.read_column(col=1, header=True, start_row=1)
{u'Row': [u'Column', u'Text']}
```
If header=True in the <i>Spreadsheet.read_column()</i> method, The first value will be considered
the header and used as the column's dictionary key (see above example). 

### Writing and Appending to Spreadsheets:
Open spreadsheets can be written or appended to.

#### Writing to Spreadsheets:
Spreadsheets can be written to with the <i>Spreadsheet.write_row()</i> and <i>Spreadsheet.write_column()</i>
methods. The <i>content</i> parameter must be a list of strings. Each string in the list will be written to
a concurrent cell.
```
>>> sheet.write_row(row=1, content=['These', 'Are', 'Headers'], start_col=1, bold=True, italics=False) 
# Writes These, Are, and Headers in bold to the top three cells horizontally in bold.

>>> sheet.write_column(col=1, content=['Here', 'Are', 'Values'], start_row=2, bold=False, italics=True) 
# Writes Here, Are, and Values in italics to three columns under the 'These' heading.
```

#### Appending to Spreadsheets:
Spreadsheets can be appended to with the <i>Spreadsheet.append_row()</i> and <i>Spreadsheet.append_column()</i>
methods. The <i>content</i> parameter must be a list of strings. Each string in the list will be written to
a concurrent cell.
```
>>> sheet.append_row(row=1, content=['More', 'Up', 'Top'], bold=True, italics=False) 
# Writes More, Up, and Top in bold after 'Headers.'

>>> sheet.append_column(col=1, content=['More', 'In', 'Col'], bold=False, italics=True) 
# Writes More, In, and Col in italics to three columns under 'Values.'
```

#### Saving Spreadsheets:
Spreadsheets can be saved with the <i>Spreadsheet.save()</i> method. Sheets are saved automatically 
when they are written to.
```
>>> sheet.save()
# Saves the sheet
```

## Files:
- <i>Spreadsheet.py</i> - Class for reading and writing to spreadsheets.