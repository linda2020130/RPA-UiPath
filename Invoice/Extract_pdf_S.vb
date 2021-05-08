Dim Invoice As String = ""
Dim InvoiceDate As String = ""
Dim PONumber As String = ""
Dim Total As Decimal = 0

Dim QtyShipped As String = ""
Dim ItemNumber As String = ""
Dim UnitPrice As String = ""
Dim TotalPrice As String = ""

Dim splitString As String() = New String(){" "}
Dim splitDollarSign As String() = New String(){"$"}
Dim file As String = Path.GetFileName(in_fileName)
Dim items() As String = in_text.Split(chr(10))

For i As Integer = 0 To items.Length - 1
  If(items(i).Contains("Invoice") And Invoice ="") 
    Invoice = items(i).Replace("Invoice","").Trim
  Else If(items(i).Contains("Date Invoiced"))
    InvoiceDate = CDate(items(i + 1).Trim).ToString("yyyyMMdd")
  Else If(items(i).Contains("Purchase Order No"))
    PONumber = items(i + 1).Trim
  Else If (items(i).Contains("Remit Wire Transfers"))
	Dim row As String() = items(i).Replace("$","").Trim.Split(splitString, StringSplitOptions.RemoveEmptyEntries)
	Total = CDec(row(row.Length-1).Trim)
	'system.Console.WriteLine(Total.ToString)
 
  Else If(items(i).Contains("Page 1 of 1"))
	Dim row As String() = items(i + 1).Split(splitString, StringSplitOptions.RemoveEmptyEntries)
	If(row.Length > 1)
		ItemNumber = row(0).Trim
	Else
		ItemNumber = items(i + 1).Trim + items(i + 2).Trim
	End If
  Else If(items(i).Contains("CUSTOMER P/N"))
    Dim row As String() = items(i + 1).Split(splitString, StringSplitOptions.RemoveEmptyEntries)
    QtyShipped = row(1).Trim
    UnitPrice = row(3).Trim
    TotalPrice = row(row.Length-1).Trim
	  
    in_dtDetail.Rows.Add(10, QtyShipped, ItemNumber, UnitPrice, TotalPrice)
 End If
Next

For Each row As DataRow In in_dtDetail.Rows
    in_dtMaster.Rows.Add(file, Invoice, InvoiceDate, PONumber, Total, row("Item"), row("Quantity"), row("PartNumber"), row("UnitPrice"), row("Amount"))
Next row 
in_dtDetail.Clear
