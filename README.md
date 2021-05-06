# RPA-UiPath
Collections of RPA works and learning notes.<br>

## [104](/104)
* **Usage**: Generate a list of all current employees from internal website and send an email notice for any related change(e.g. new employees, resigned employees...).
* **Files**:
  * [104SSearch](/104/104Search.py): Python code using selenium and BeautifulSoup packages for employees data scraping from internal website.
  * [Find104Change](/104/Find104Change.bas): Excel VBA code for comparing current employee list to last one and generating table of differences.
  * [Flowchart_Get 104 List](/104/Flowchart_Get%20104%20List.png): Flowchart of generating all current employees list from internal website.
    <details><summary>Flowchart</summary>
    
    ![Flowchart_Get 104 List](/104/Flowchart_Get%20104%20List.png)
    </details>
  * [Outline_104 Change](/104/Outline_104%20Change.png): Outline of sending an email notice for any 104 related change including invoke another process to generate current list of employees.
    <details><summary>Outline</summary>
    
    ![Outline_104 Change](/104/Outline_104%20Change.png)
    </details>
<br>

## [Invoice](/Invoice)
* [Extract_pdf_S](/Invoice/Extract_pdf_S.vbs): Input a invoice pdf file from customer S, extract certain columns/rows of information, and output a table with extracted data.
* [PDF2Excel](/Invoice/PDF2Excel.py): Save tables in the pdf files to excel files via Python.
* [SaveAttachment](/Invoice/SaveAttachment.vbs): Get attachments from msg files(email) and save to assigned folder via VBScript.
<br>

## [Product Line](/Product%20Line)
* **Usage**: Download current product line information and send an email notice for any related change(e.g. new product line, change PM...).
* **Files**:
  * [FindPLChange](/Product%20Line/FindPLChange.bas): Excel VBA code for comparing current product line information to last one and generating table of differences.
  * [Outline_PL Change](/Product%20Line/Outline_PL%20Change.png): Outline of downloading current product line information, invoking VBA code to generate table of differences, and sending an email notice for any product line related change.
    <details><summary>Outline</summary>
    
    ![Outline_PL Change](/Product%20Line/Outline_PL%20Change.png)
    </details>
<br>

## [Sales Order](/Sales%20Order)
* **Usage**: Log in to customer's website, download *sales order pdf files* based on given time range, and generate/download *summary table* of current available sales orders.
* **Projects**:
  * Customer A: 
    <details><summary>Outline</summary>
    
    ![Outline_A](/Sales%20Order/Outline_A.png)
    </details>
    
    * **Sales Order PDF**: Iterate through all sales order number to generate an url of the sales order pdf file, navigate to it, and hit print button.
    * **Summary Table**: Need to scrape data from multiple pages with uncertain total pages and no next-page button. --> Iterate through all possible number of page button and break if element is not found.
  * Customer B:
    <details><summary>Outline</summary>
    
    ![Outline_B](/Sales%20Order/Outline_B.png)
    </details>
    
    * **Sales Order PDF**: Only download pdf files with certain order status. Order status, order number, and id of print button need to be scraped from two tables on one website page. Check availability of the next-page button by element attribute before clicking.
    * **Summary Table**: Click button to download.
  * Customer M:
    <details><summary>Outline</summary>
    
    ![Outline_M](/Sales%20Order/Outline_M.png)
    </details>
    
    * **Sales Order PDF**: Iterate through each print button and get order number from element attribute. Click print button to download and name file by order number.
    * **Summary Table**: Click button to download.
  * Customer Q:
    <details><summary>Outline</summary>
    
    ![Outline_Q](/Sales%20Order/Outline_Q.png)
    </details>
    
    * **Sales Order PDF**: Get order number from downloaded summary table and generate urls based on order number. Navigate to each url to download the pdf file.
    * **Summary Table**: Click button to download.
  * [Extract SO Info_C](/Sales%20Order/Extract%20SO%20Info_C.vbs): Input a sales order from customer C, extract certain columns/rows of information, and output a table with extracted data.
  * [Extract SO Info_L](/Sales%20Order/Extract%20SO%20Info_L.vbs): Input a sales order from customer L, extract certain columns/rows of information, and output a table with extracted data.
<br>

## [UiPath Notes](/UiPath_Notes.xlsx)
* **Sheets**:
  * **Lecture**: Notes of UiPath online lectures and examples.
  * **Project**: Notes of my RPA works including what activities were used, how to use, and some useful tips.
  * **Remark**: Illustrations of other sheets.
<br>

## [VB Syntax](/VB%20syntax.vb)
Quickly look up VB.NET common syntax.
