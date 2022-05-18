' VB.NET常用語法整理



' ========= Directory ==========
' 取得資料夾內檔案數量
directory.getfiles(var_strPath,"*").Count
' 取得資料夾內xlsx檔案數量
directory.getfiles(var_strPath,"*.xlsx").Count
' 取得由資料夾內檔案名稱形成的array of string
directory.getfiles(var_strPath, "*")
' 取得父目錄
directory.getParent(var_strPath).FullName



' ========= Path ==========
Dim var_strPath As String = "C:\temp\test.xlsx"
' 取得檔案所在資料夾名稱
Path.GetDirectoryName(var_strPath)   ' 返回"C:\temp"
' 取得檔案副檔名
Path.GetExtension(var_strPath)   ' 返回".xlsx"
' 取得檔案名稱(含副檔名)
Path.GetFileName(var_strPath)   ' 返回"test.xlsx"
' 取得檔案名稱(不含副檔名)
Path.GetFileNameWithoutExtension(var_strPath)   ' 返回"test"



' ========= String =========
' 建立變數
Dim var_str As String
Dim strNumber As String = "123123"
' 尋找指定字元的位置
var_str.IndexOf("/")   ' 返回第一個出現"/"的index位置
' 尋找指定字元的位置
var_str.LastIndexOf("/")	' 返回最後一個出現"/"的index位置
' 尋找最後一個"\"在string中的位置
InStrRev(var_strPath,"\")
' 取路徑字串中的檔名(最後一個"\"後的字串)
Split(var_strPath,"\").Last.ToString
' 從左起擷取n個字串
Left(var_str, n)
' 從右起擷取n個字串
Right(var_str, n)
' 將指定字串("\")取代成另一個指定字串(""), 可取代為空字串來達到去除指定字串("\")的效果
var_strPath.Replace("\","")
' 去除字串變數頭尾的空白
var_str.Trim
' 取得字串長度
var_str.Length
' 去除var_str裡的最後一個字串
var_str.Remove(var_str.Length-1)
' 擷取指定位置字元形成子字串
var_str.Substring(0, n)   ' 從頭擷取n個字元
' 組字串, 亦可用_來分行
String.Format("Hi, my name is {0} and I'm {1}.", "Linda", 18)
String.Format(strVar, _
row(0).ToString.Trim, _
row(1).ToString.Trim
)
' 檢查字串內是否存在特定字串"\"
var_str.Contains("\")   ' return True or False)
' 輸出字串, 中間換行, 字串輸出完不換行
Console.write("Hi" + Environment.NewLine + "Bye")
Console.Write("Hi" + vbCrLf + "Bye")
' 輸出字串, 字串輸出完自動換行
Console.WriteLine("Hi")
' 字串轉小寫
var_str.ToLower
' 字串轉大寫
var_str.ToUpper
' 檢查字串是否為空或空白
String.IsNullOrEmpty(var_str)   ' return True or False
' 檢查字串是否為日期
IsDate(str_Date)   ' return True or False



' ========= Array =========
Dim splitSpace As String() = New String(){" "}
Dim arrItem As String() = var_str.Split(splitSpace, StringSplitOptions.RemoveEmptyEntries)
Dim items() As String() = var_str.Split(Chr(10))   ' ASCII Chr(10) = 換行
' 將字串依指定方式(分行, 空格)切割成array of string, 並剔除空字串
var_str.Split({Environment.NewLine,vbcrlf,vblf," ",vbtab,vbcr,vbNewLine},StringSplitOptions.RemoveEmptyEntries)
' 尋找指定字串("test")在array中的位置
Array.indexof(var_array, "test")   ' 返回index


' ========= List ===========
Dim listS As New List(Of String)
listS.Add("test")


' ========= Dictionary =========
' 取得Dictionary裡的值
Dim dictionary As New Dictionary(Of String, Integer)
dictionary.Add("apple", 5)
dictionary("apple")  ' return 5
dicSS = New Dictionary(Of String, String) From {{"A","aaa"},{"B","bbb"}}
dicSL = New Dictionary(Of String, List(Of String)) From {{"A", New List(Of String)({"a1","a2"})},{"B", New List(Of String)({"b1","b2"})}}
For each System.Collections.Generic.KeyValuePair


' ========= DataTable =========
' 建立變數
Dim dtTEST As New DataTable
Dim newCol1 As New Data.DataColumn("colName1", GetType(System.String))
Dim newCol2 As New Data.DataColumn("colName2", GetType(System.Decimal))
newCol1.DefaultValue = "0"
newCol2.DefaultValue = 0
Dim dtTEST() As datarow
' 建立一個與dtTEST相同欄位的空資料表
Dim dtTESTTemp As DataTable = dtTEST.Clone   
' 資料表內新增欄位
dtTEST.Columns.Add(newCol1)
dtTEST.Columns.Add(newCol2)
' 資料表內刪除欄位
dtTEST.Columns.Remove(strColName)	' 刪除名為strColName的欄位
dtTEST.Columns.Remove(0)	' 刪除第一欄位
' 資料表內新增資料列
dtTEST.ImportRow(dtRow)   ' maintains the rowstate of the imported row
dtTEST.Rows.Add(dtRow)   ' always sets the rowstate to added
dtTEST.Rows.Add(strNumber1, system.DBNull.Value)   ' 依照欄位順序與資料類型填入資料列內容
' 產生由資料列組成的array, 搭配For Each使用
dtTEST.Rows
' 產生由資料欄位組成的array
dtTEST.Columns
' 取得資料表的資料列數(不含表頭)
dtTEST.Rows.Count
' 取得資料表的資料欄數
dtTEST.Columns.Count
' 用欄位名稱取得欄位index
dtTEST.Columns.IndexOf(strColumnName)
' 取得資料表的第一個欄位內容
dtTEST.Columns(0).ToString
' 取得資料表的第一個欄位資料型態
dtTEST.Columns(0).DataType.ToString
' 將資料欄名稱"strColumnName"改為""TEST1"
dtTEST.Columns("strColumnName").ColumnName = "TEST1"
' 取得資料表內第1個資料列的第2個欄位內容
dtTEST.Rows(0)(1).ToString
' 取得資料表內第1個資料列的指定欄位名稱之內容
dtTEST.Rows(0)("strColumnName").ToString
' 取得資料型態
dtTEST.Rows(0)("strColumnName").GetType.ToString
' 更改資料列指定欄位名稱之內容為"TEST"
var_datarow("strColumnName") = "TEST"
var_datarow.Item("strColumnName") = "TEST"
' 從資料表中篩選出指定條件的資料列
' 指定條件若為字串, 需用引號''包起來
dtTEST2 = dtTEST.Select("strColumnName='TEST' and strNumber=1").CopyToDataTable
' 若要篩選的字串夾雜', 需使用兩次''來跳脫
strVar = test's
dtTEST2 = dtTEST.Select(String.Format("strColumnName='{0}' and strNumber=1",strVar.replace("'","''"))).CopyToDataTable
' 篩選出colName1等於colName2 或 colName1不等於colName2但colName1開頭為A或a(不分大小寫) 的資料
dtTEST3 = dtTEST.Select("[colName1]=[colName2] OR ([colName1]<>[colName2] AND [colName1] LIKE 'A%')").CopyToDataTable
' 篩選指定日期區間內資料
dtTEST2 = dtTEST.Select("[DATE] >= #2021/08/17# AND [DATE] <= #2021/08/20#").CopyToDataTable
dtTEST2 = dtTEST.Select(String.Format("[DATE] >= #{0}# AND [DATE] <= #{1}#",dateVariable1.ToString,dateVariable2.ToString)).CopyToDataTable
' 可指定資料列排列順序(按strNumber1欄位由小到大, 同strNumber1的再按strNumber2欄位由小到大)
dtTEST2 = dtTEST.Select("strColumnName='TEST'", "strNumber1 Asc, strNumber2 Asc").CopyToDataTable
' 複製資料表(資料欄列全部複製), 存在dtTEST2裡, 可搭配select使用
dtTEST2 = dtTEST.CopyToDataTable
' 複製資料表頭(僅資料欄位名稱, 資料型態等, 不複製資料列), 存在dtTEST3裡
dtTEST3 = dtTEST.Clone
' 清除資料列
dtTEST.Clear
' 依指定欄位在資料表中篩選出不重複的資料列, 返回由指定欄位組成的不重複資料列表(同SQL Distinct的效果)
dtTEST2 = dtTEST.DefaultView.ToTable(True, "colName1", "colName2")   ' 可放一個或多個欄位
' 對資料表依指定欄位排序, 預設為遞增, 可在欄位名稱後加上DESC改為遞減(其他未明列欄位仍存在)
dtTEST.DefaultView.Sort = "colName1 DESC, colName2"   ' colName1相同再依colName2排序
dtTEST = dtTEST.DefaultView.ToTable
' 檢查工作列表是否有被指派過
dtTEST is Nothing   ' return True or False
' 將dtTEST的欄位名稱組成array of string
dtTEST.Columns.Cast(Of DataColumn).Select(Function (x) x.Columnname).toArray
' 將dtTEST內依columnFilter欄位加總columnDecimal欄位(decimal欄位) -> 返回加總值
dtTEST.AsEnumerable.Where(Function(x) x(columnFilter).ToString.Equals(CurrentRow(columnFilter).ToString)).Sum(Function(x) Convert.ToDecimal(x(columnDecimal).ToString) ).ToString
' 將dtTEST內依columnFilter欄位加總columnInteger欄位(integer欄位) -> 返回最大值
dtTEST.AsEnumerable.Where(Function(x) x(columnFilter).ToString.Equals(CurrentRow(columnFilter).ToString)).Max(Function(x) Convert.ToInt32(x(columnInteger).ToString) ).ToString
' 將dtTEST內的columnInteger欄位加總 -> 返回加總值
dtTEST.AsEnumerable().Sum(Function(row) row.Field(Of Int32)(columnInteger))
' 將dtTEST內的columnDecimal欄位加總 -> 返回加總值
dtTEST.AsEnumerable().Sum(Function(row) Convert.ToDecimal(row(columnDecimal).ToString))
' 取dtTest內PO_DATE欄位最早日期 -> 返回日期變數
Convert.ToDateTime(dtTest.AsEnumerable().Min(Function(row) row("PO_DATE")))
' 取dtTest內PO_DATE欄位最晚日期 -> 返回日期變數
Convert.ToDateTime(dtTest.AsEnumerable().Max(Function(row) row("PO_DATE")))
' 取dtTest前20列
dtTop20 = dtTest.AsEnumerable.take(20).CopyToDataTable
' 取dtTest後10列
dtLast10 = dtTest.AsEnumerable.Reverse().take(10).CopyToDataTable



' ========= Excel工作表 =========
Dim app As New Excel.Application   ' app 是操作 Excel 的變數
Dim worksheet As Excel.Worksheet   ' Worksheet 代表的是 Excel 工作表
Dim workbook As Excel.Workbook   ' Workbook 代表的是一個 Excel 本體
' 取得由Excel工作表名稱組成的array of string
workbook.GetSheets



' ========= 資料型態轉換/日期相關 =========
' 字串轉日期
CDate(var_strDate)
' 日期依照指定日期格式轉字串
var_Date.ToString("yyyyMMdd")
' 當下時間轉字串
Date.Now.ToString("yyyyMMdd_HHmmss")
System.DataTime.Now.ToString("yyyyMMdd_HHmmss")
var_Date.ToString("dd-MMM-yy")		' var_Date為2021/12/31 => 返回 31-Dec-21
' 日期變數加天數(e.g. 加5天)
dateVariable.AddDays(5)
' 日期變數加年數(e.g. 加3年)
dateVariable.AddYears(3)
' 字串轉月份, True=縮寫, False=全名
MonthName(4, True)		' return "Apr"
MonthName(4)			' return "April"
' 返回星期幾
var_Date.DayOfWeek		' 返回0~6(星期日~星期六)
' 字串轉整數
CInt(var_str)
' 小數四捨五入轉整數
CInt(var_decimal)
' 字串轉double
CDbl(strNumber)
System.Convert.ToDouble(strNumber)
' 字串轉decimal
CDec(strNumber)
' 指定日期字串依日期格式(yyyyMMdd)轉日期
Datetime.ParseExact(strDate,“yyyyMMdd”,System.Globalization.CultureInfo.InvariantCulture)
' 時間差換算成天數(包含整數和小數, e.g. 3.25天)
dateDiff.TotalDays
' 字串第一字元轉字元
CChar(var_str(0))
' 字元轉ASCii編碼
Asc(var_char)
' Object轉Boolean
Convert.ToBoolean(objectVal)
' GetWeekOfYear可以回傳週數(int), 每年的第一個完整週(以星期天為一週的第一天)為第1週
' 參數1: 欲計算週數的日期 DateTime
' 參數2: 定義日曆週的列舉值 CalendarWeekRule
' FirstDay or 0: 指示一年的第一週開始於該年的第一天，並結束於被指定為該週第一天的前一天
' FirstFourDayWeek or 2: 指示一年的第一週有四天以上在被指定為該週的第一天之前
' FirstFullWeek or 1: 指示一年的第一週開始於，一年的第一天當天或之後被指定為一週第一天的那天
' 參數3: 代表一週第一天的列舉值 DayOfWeek
' 0~6, 星期天為0,星期一為1以此類推
' 若執行當下為2022/04/21, 則返回52
System.Globalization.CultureInfo.CreateSpecificCulture(“en-US”).Calendar.GetWeekOfYear(New Datetime(Now.Year,12,31),System.Globalization.CalendarWeekRule.FirstFullWeek,DayOfWeek.Sunday)		
' 檢查是否為DBNull
Convert.IsDBNull(objVariable)	' return True or False
' 可以是Null的資料型態
bolVar = New Nullable<Boolean>	' bolVar can be True or False or Null



' ========= Number ==========
' 建立變數
Dim qty As Decimal = Convert.ToDecimal(strNumber)   ' 將strNumber轉成數字存在qty變數裡
' 檢查是否為數字
strVar.IsNumeric   ' return True or False
' 相除並回傳整數結果
5 \ 4   ' return 1
' 相除回傳Double型別
5 / 4   ' return 1.25
' 小數後四捨五入取到第n位
Decimal.Round(0.12345, 3)	' return 0.124
' 無條件進位取整數
Math.Ceiling(1.01)  ' return 2
' 無條件捨去取整數
Math.Floor(1.99)  ' return 1



' ========== Loop =========
For Each row in dtTEST
	' Do Something
	If strNumber <> ""
		' Do Something
	ElseIf intNumber = 0
		' Do Something
	Else
		' Do Something
	End If
Next

For i As Integer = 0 To items.Length - 1
	If strNumber <> ""
		' Do Something
	Else If intNumber = 0
		' Do Something
		Exit For  ' 跳出for迴圈
	Else
		' Do Something
	End If
Next

' 產生數字list來跑迴圈(start, count)
Enumerable.Range(2, 3)   ' [2, 3, 4]



' ========== Except =========
' 中斷程式
System.Exception("Error!")
' 產出錯誤訊息
System.Exception.Message



' ========= Others =========
' 數字轉Excel欄位字母(1=>A, 2=>B...等) A=65, B=66, ...by ASCII編碼
If(var_int<=26, Convert.ToChar(var_int + 64).ToString, If(var_int mod 26 = 0,Convert.ToChar((var_int \ 26) + 63).ToString,Convert.ToChar((var_int \ 26) + 64).ToString) + If(var_int mod 26 = 0, "Z",Convert.ToChar((var_int mod 26) + 64).ToString))

