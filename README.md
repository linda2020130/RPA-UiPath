# RPA-UiPath
Collections of RPA works and learning notes.<br>

## Table of Contents
* [104](#104)
* [Invoice](#invoice)
* [Product Line](#PL)
* [Sales Order](#SO)
* [UiPath Notes](#notes)
* [VB Syntax](#syntax)

<h2 id="104">104</h2>

[🐳](/104)
* **Usage**: Generate a list of all current employees from internal website and send an email notice for any related change(e.g. new employees, resigned employees...).
* **Files**:
  * [104SSearch](/104/104Search.py): Python code using selenium and BeautifulSoup packages for employees data scraping from internal website.
  * [Find104Change](/104/Find104Change.bas): Excel VBA code for comparing current employee list to last one and generating table of differences.
<br>

<h2 id="invoice">Invoice</h2>

[🐳](/Invoice)
* [Extract_pdf_S](/Invoice/Extract_pdf_S.vb): Input a invoice pdf file from customer S, extract certain columns/rows of information, and output a table with extracted data.
* [PDF2Excel](/Invoice/PDF2Excel.py): Save tables in the pdf file to an excel file via Python.
* [SaveAttachment](/Invoice/SaveAttachment.vbs): Get attachments from msg files(email) and save to assigned folder via VBScript.
<br>

<h2 id="PL">Product Line</h2>

[🐳](/Product%20Line)
* **Usage**: Download current product line information and send an email notice for any related change(e.g. new product line, change PM...).
* **Files**:
  * [FindPLChange](/Product%20Line/FindPLChange.bas): Excel VBA code for comparing current product line information to last one and generating table of differences.
<br>

<h2 id="SO">Sales Order</h2>

[🐳](/Sales%20Order)
* **Usage**: Log in to customer's website, download *sales order pdf files* based on given time range, and generate/download *summary table* of current available sales orders.
* **Projects**:
  * [Extract SO Info_C](/Sales%20Order/Extract%20SO%20Info_C.vb): Extract data from customer C's SO pdf, and output a datatable.
  * [Extract SO Info_L](/Sales%20Order/Extract%20SO%20Info_L.vb): Extract data from customer C's SO pdf, and output a datatable.
  * [PDF2Text](/Sales%20Order/PDF2Text.py): Save strings in the pdf file to a text file via Python.
<br>

<h2 id="notes">UiPath Notes</h2>

[🐳](/UiPath_Notes.xlsx)
* **Sheets**:
  * **Lecture**: Notes of UiPath online lectures and examples.
  * **Project**: Notes of my RPA works including what activities were used, how to use, and some useful tips.
  * **Remark**: Illustrations of other sheets.
<br>

<h2 id="syntax">VB Syntax</h2>

[🐳](/VB%20syntax.vbs)
Quickly look up VB.NET common syntax.
