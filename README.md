# RPA-UiPath
Collections of RPA works and learning notes.<br>

## [104](/104)
* **Usage**: Generate a list of all current employees from internal website and send an email notice for any related change(e.g. new employees, resigned employees...).
* **Files**:
  * [Find104Change](/104/Find104Change.vb): Excel VBA code for comparing current employee list to last one and generating table of differences.
  * [Flowchart_Get 104 List](/104/Flowchart_Get%20104%20List.png): Flowchart of generating all current employees list from internal website.
    <details><summary>Flowchart</summary>
    
    ![Flowchart_Get 104 List](/104/Flowchart_Get%20104%20List.png)
    </details>
  * [Outline_104 Change](/104/Outline_104%20Change.png): Outline of sending an email notice for any 104 related change including invoke another process to generate current list of employees.
    <details><summary>Outline</summary>
    
    ![Outline_104 Change](/104/Outline_104%20Change.png)
    </details>
<br>

## [Product Line](/Product%20Line)
* **Usage**: Download current product line information and send an email notice for any related change(e.g. new product line, change PM...).
* **Files**:
  * [FindPLChange](/Product%20Line/FindPLChange.vb): Excel VBA code for comparing current product line information to last one and generating table of differences.
  * [Outline_PL Change](/Product%20Line/Outline_PL%20Change.png): Outline of downloading current product line information, invoking VBA code to generate table of differences, and sending an email notice for any product line related change.
    <details><summary>Outline</summary>
    
    ![Outline_PL Change](/Product%20Line/Outline_PL%20Change.png)
    </details>
<br>

## [Sales Order](/Sales%20Order)
* **Usage**: Log in to customer's website, download sales orders pdf files based on given time range, and generate/download summary table of current available sales orders.
* **Projects**:
  * Customer A: (Tricky Parts)
    <details><summary>Outline</summary>
    
    ![Outline_A](/Sales%20Order/Outline_A.png)
    </details>
  * Customer B: (Tricky Parts)
    <details><summary>Outline</summary>
    
    ![Outline_B](/Sales%20Order/Outline_B.png)
    </details>
  * Customer M: (Tricky Parts)
    <details><summary>Outline</summary>
    
    ![Outline_M](/Sales%20Order/Outline_M.png)
    </details>
  * Customer Q: (Tricky Parts)
    <details><summary>Outline</summary>
    
    ![Outline_Q](/Sales%20Order/Outline_Q.png)
    </details>
  * [Extract SO Info_L](/Sales%20Order/Extract%20SO%20Info_L.vb): Input a sales order from customer L, extract certain columns/rows of information, and output a table with extracted data.
<br>

## [UiPath Notes](/UiPath_Notes.xlsx)
* **Sheets**:
  * **Lecture**: Notes of UiPath online lectures and examples.
  * **Project**: Notes of my RPA works including what activities were used, how to use, and some useful tips.
  * **Remark**: Illustrations of other sheets.



