Dim splitSpace As String() = New String(){" "}
Dim splitLine As Char() = New Char(){chr(10)}

Dim BillTo As String
Dim VendorCode As String
Dim PO_NO As String
Dim PO_DATE As String
Dim PaymentTerm As String
Dim Incoterm As String
Dim ShipTo As String
Dim Plant As String

Dim items() As String = in_text.Split(splitLine , StringSplitOptions.RemoveEmptyEntries) 

For i As Integer = 0 To items.Length - 1
	'console.WriteLine(items(i))
	
	If(i=0)
		If(items(0).Contains("CPS2"))
			BillTo = items(1).ToString.Trim
			If BillTo = ""
				BillTo = items(2).ToString.Trim
			End If
		Else
			BillTo = items(0).ToString.Trim
		End If
	
	Else If(items(i).Contains("Vendor Code :"))
		Dim columns() As String = items(i).Split(splitSpace, StringSplitOptions.RemoveEmptyEntries)
		VendorCode = columns(3).Trim
		PO_NO = columns(7).Trim
		PO_DATE = columns(10).Trim
		
	Else If(items(i).Contains("Payment Term :"))
		Dim columns() As String = items(i).Split(splitSpace, StringSplitOptions.RemoveEmptyEntries)
		PaymentTerm = columns(4).Trim

	Else If(items(i).Contains("Incoterm :"))
		Dim columns() As String = items(i).Split(splitSpace, StringSplitOptions.RemoveEmptyEntries)
		Incoterm = columns(3).Trim
	
	Else If(items(i).Contains("Plant :"))
		Dim columns() As String = items(i).Split(splitSpace, StringSplitOptions.RemoveEmptyEntries)
		Plant = columns.Last.Trim
		Exit For
	
	'Else If(items(i).Contains("Ship to:"))
		'ShipTo = items(i).Substring( items(i).IndexOf("Ship to:") +8 ).Trim
		'Dim columns() As String = items(i).Split(splitSpace, StringSplitOptions.RemoveEmptyEntries)
		'ShipTo = columns.Last.Trim
		'Exit For	
	End If
Next

io_dtMaster.Rows.Add(BillTo, VendorCode, PO_NO, PO_DATE, PaymentTerm, Incoterm, Plant)
