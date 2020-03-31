[![PyPI version](https://badge.fury.io/py/auto-table-extract.svg)](https://badge.fury.io/py/auto-table-extract)
![PyPI - Python Version](https://img.shields.io/pypi/pyversions/auto-table-extract)
![GitHub](https://img.shields.io/github/license/rohit-sahoo/auto-table-extract)
![Libraries.io SourceRank](https://img.shields.io/librariesio/sourcerank/pypi/auto-table-extract)
![PyPI - Wheel](https://img.shields.io/pypi/wheel/auto-table-extract)
![PyPI - Downloads](https://img.shields.io/pypi/dm/auto-table-extract)
![Libraries.io dependency status for latest release](https://img.shields.io/librariesio/release/pypi/auto-table-extract)

# Automated Table Extractor
![Auto_Table_Extract](https://i.ibb.co/gF7nyvL/icon3.png)
<br />
`auto-table-extract` is a Python package to extract tables from PDF documents. <br />
This package helps to extract all the table contents (tables can be bordered, borderless or partially bordered) from the searchable and scanned pdf document and dumps it into an excel sheet.<br />

## What's new in version 2.0?
1) Added feature to extract tables from scanned documents.
2) Saves the file on desktop as "filename.xlsx" rather than output.xlsx, making it easier to track your files. 
3) Bug Fixes

## Main features
1) Can extract tables from bordered tables, partially bordered tables (missing column lines / missing row lines) and also from fully borderless tables.
2) Extracts only the text inside the table. It won't extract the paragraphs, matrix, bar-charts or any text outside table.
2) Creates an Excel file having the extracted contents of the table from PDF.
3) Can be used for both searchable(contents can be selected) as well as scanned documents.
4) Works for any orientation of the document.

## Requirements
1) pdfminer   
2) scikit-learn
3) pandas
4) ocrmypdf
5) Numpy
6) PyPDF2

## Installation
1) Install Tesseract OCR <br />
[Download Tesseract 32 bit](https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w32-setup-v5.0.0-alpha.20191030.exe)<br />
[Download Tesseract 64 bit](https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-v5.0.0-alpha.20191030.exe)<br />

2) Install ghostscript <br />
[Download ghostscript](https://www.ghostscript.com/download/gsdnld.html)


## Usage
```python
from auto_table_extract import auto_table_extract 
auto_table_extract("input_file") #input_file should be PDF
#The Excel file will be saved to your Desktop
```


## License

**Copyright 2020 Rohit-Sahoo**

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
