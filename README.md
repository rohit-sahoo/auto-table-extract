# Auto-Table-Extract: A System To Identify And Extract Tables From PDF To Excel 
Read the published Paper: [Research Paper on Auto-Table-Extract](https://www.ijstr.org/final-print/may2020/Auto-table-extract-A-System-To-Identify-And-Extract-Tables-From-Pdf-To-Excel.pdf)
## Introduction
The auto-table-extract system is capable of identifying tablular data within PDF documents and extracting all the tabular information into an excel  file. 

Table detection is the process of identifying tables from a document, extracting the cells contained in a table. 

The auto-table-extract system consists of three main modules: 1) Document conversion 2) Layout Analysis 3) Table detection and extraction. 

![Auto-Table-Extract System](https://drive.google.com/uc?id=1mgSpUkGmnbI12_fE83PZ3fyDfR6lsMsW)


**The two methods used for identification and extraction are:**

1) Table_with_Border ( For tables with fully recognizable borders)
The Table_with_Border method is used to determine the tables with the help of coordinates of text lines, characters, and text boxes provided by the PDFMiner. 

2) Table_without_Border (For partially bordered or borderless tables ). 
The Table_without_Border method uses the clustering method and coordinates of the text line to determine the table and extract its contents.


Further, a Pandas DataFrame consisting of extracted data is created, which is used to make the excel sheet containing the data. The output of the auto-table-extract system is an Excel document with the table’s information extracted from the PDF. 
