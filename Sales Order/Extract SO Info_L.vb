Dim Contact As String = ""
Dim PO_NO As String = ""
Dim PO_DT As String = ""
Dim Plant As String = ""
Dim ShipTo As String = ""
Dim PayTo As String = ""
Dim Currency As String = ""
Dim PaymentTerm As String = ""

Dim boolDetail As Boolean = False
Dim PO_LIN As String = ""
Dim BUY_PN As String = ""
Dim VN_PN As String = ""
Dim PO_QTY As Decimal = 0
Dim PO_PRI As Decimal = 0
Dim SHIP_DT As String = ""

Dim splitSpace As String() = New String(){" "}
Dim splitDoubleSpace As String() = New String(){"  "}
Dim splitSlash As String() = New String(){"\"}
Dim splitSquare As String() = New String(){"■"}
Dim splitLine As Char() = New Char(){chr(10)}
Dim items() As String = in_text.Split(splitLine , StringSplitOptions.RemoveEmptyEntries) 


For i As Integer = 0 To items.Length - 1
	'console.WriteLine(items(i))
	
	If(items(i).Contains("Contact:") And Contact = "")
		Dim columns() As String = items(i).Split(splitSpace, StringSplitOptions.RemoveEmptyEntries)
		Contact = columns(1).Trim
	
	Else If(items(i).Contains("Purchase Order:") And PO_NO = "")
		Dim columns() As String = items(i).Split(splitDoubleSpace, StringSplitOptions.RemoveEmptyEntries) 
		PO_NO = columns(1).Trim
		PO_DT = columns(3).Trim
		Plant = columns(5).Trim
	
	Else If(items(i).Contains("Pay To:") And PayTo = "")
		Dim columns() As String = items(i + 1).Split(splitDoubleSpace, StringSplitOptions.RemoveEmptyEntries)
		PayTo = columns(0).Trim
		
	Else If(items(i).Contains("Currency:") And Currency = "")
		Dim columns() As String = items(i).Split(splitSpace, StringSplitOptions.RemoveEmptyEntries)
		Dim columns1() As String = items(i).Split(splitDoubleSpace, StringSplitOptions.RemoveEmptyEntries)
		Currency = columns(1).Trim
		PaymentTerm = columns1(2).Trim
		
	Else If(items(i).Contains("Ship To:") And ShipTo = "")
		For j As Integer = i + 1 To items.Length - 1
			If(items(j).Contains("Payment Term"))
				ShipTo = ShipTo.Trim
				Exit For
			Else
				Dim columns() As String = items(j).Split(splitDoubleSpace, StringSplitOptions.RemoveEmptyEntries)
				ShipTo += columns.Last.Trim + " "
			End If
		Next

	Else If(items(i).Contains("PR no") And boolDetail = False)
		boolDetail = True
		
	Else If(boolDetail And items(i).Trim.Length > 0)
		Dim columns() As String = items(i).Split(splitDoubleSpace, StringSplitOptions.RemoveEmptyEntries)
		If columns.Length >=8 And columns(0).Trim <> "Line"
			PO_LIN = columns(0).Trim
			BUY_PN = columns(1).Trim
			PO_PRI = CDec(columns(4).Trim)
			SHIP_DT = columns(5).Trim
			PO_QTY = CDec(columns(6).Trim)
			For j As Integer = i + 1 To items.Length - 1
				If(items(j).Contains("P/N:"))
					Dim columns1() As String = items(j).Split(splitDoubleSpace, StringSplitOptions.RemoveEmptyEntries)
					VN_PN = columns1(0).Substring(4).Trim
					Exit For
				End If
			Next
			
			io_dtDetail.Rows.add(PO_LIN, BUY_PN, PO_PRI, Convert.ToDateTime(SHIP_DT), PO_QTY, VN_PN)
			
		End If
	End If
Next

io_dtMaster.Rows.Add(Contact, PO_NO, Convert.ToDateTime(PO_DT), Plant, ShipTo, PayTo, Currency, PaymentTerm)
