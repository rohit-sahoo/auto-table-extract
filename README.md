[![PyPI version](https://badge.fury.io/py/auto-table-extract.svg)](https://badge.fury.io/py/auto-table-extract)
[![Python 3 | 3.6](https://img.shields.io/badge/python-3.6-blue.svg)](https://www.python.org/downloads/release/python-360/)
# Automated Table Extractor

A Python package to extract tables from PDF documents
This package helps to extract all the table contents from the PDF searchable and scanned pdf document and dumps it into an excel sheet.

## Main features
1) Creates an excel file having the extracted contents of the table from PDF
2) It can be used for both searchable(contents can be selected) as well as scanned documents
3) It works for any orientation of the document

## Requirements
1) pdfminer
2) scikit-learn
3) pandas
4) ocrmypdf

## Usage
```python
from auto_table_extract import auto_table_extract 
auto_table_extract("input_file") #input_file should be PDF
```

## License

**MIT License**

**Copyright (c) 2020 rohit-sahoo**

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
