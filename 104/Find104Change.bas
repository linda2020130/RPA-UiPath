Sub Find104Change()

    Dim curTotalRow, preTotalRow, chaTotalRow, chaTotalRow2 As Integer
    chaTotalRow2 = 1
    
    Application.ScreenUpdating = False
    ' 清空Change工作表
    With Worksheets("Change")
        If .Range("A2").Value <> "" Then
            chaTotalRow = .Range("A1").End(xlDown).Row
            .Rows("2:" & chaTotalRow).Delete
        End If
    End With

    ' 將新進員工貼到Change工作表
    With Worksheets("Current")
        curTotalRow = .Range("B1").End(xlDown).Row
        .Range("K2:K" & curTotalRow).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-10],Previous!C1,1,FALSE),""New"")"
        .Range("A1:K" & curTotalRow).AutoFilter Field:=11, Criteria1:="New"
        If .Range("A1:A" & curTotalRow).SpecialCells(xlCellTypeVisible).Count > 1 Then
            .Range("A2:C" & curTotalRow & ",F2:G" & curTotalRow & ",I2:J" & curTotalRow).SpecialCells(xlCellTypeVisible).Copy
            With Worksheets("Change")
                chaTotalRow = chaTotalRow2 + 1
                .Range("B" & chaTotalRow).PasteSpecial xlPasteValues
                chaTotalRow2 = .Range("B1").End(xlDown).Row
                .Range("A" & chaTotalRow & ":A" & chaTotalRow2).FormulaR1C1 = Sheets("Type").Range("A2").Value
            End With
        End If
        .AutoFilterMode = False
    End With
    
    ' 將離職員工貼到Change工作表
    With Worksheets("Previous")
        preTotalRow = .Range("B1").End(xlDown).Row
        .Range("K2:K" & preTotalRow).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-10],Current!C1,1,FALSE),""Left"")"
        .Range("A1:K" & preTotalRow).AutoFilter Field:=11, Criteria1:="Left"
        If .Range("A1:A" & preTotalRow).SpecialCells(xlCellTypeVisible).Count > 1 Then
            .Range("A2:C" & preTotalRow & ",F2:G" & preTotalRow & ",I2:J" & preTotalRow).SpecialCells(xlCellTypeVisible).Copy
            With Worksheets("Change")
                chaTotalRow = chaTotalRow2 + 1
                .Range("B" & chaTotalRow).PasteSpecial xlPasteValues
                chaTotalRow2 = .Range("B1").End(xlDown).Row
                .Range("A" & chaTotalRow & ":A" & chaTotalRow2).FormulaR1C1 = Sheets("Type").Range("A3").Value
            End With
        End If
        .AutoFilterMode = False
    End With
    
    ' 將中文名變更員工貼到Change工作表
    With Worksheets("Previous")
        .Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        .Range("A2:A" & preTotalRow).FormulaR1C1 = "=RC[1]&RC[2]"
    End With
    With Worksheets("Current")
        .Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        .Range("A2:A" & curTotalRow).FormulaR1C1 = "=RC[1]&RC[2]"
        .Range("M2:M" & curTotalRow).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-12],Previous!C1,1,FALSE),IF(RC[-1]=""New"",""New"",""Change""))"
        .Range("N2:N" & curTotalRow).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-2],Previous!C2:C3,2,FALSE),""New"")"
        .Range("A1:N" & curTotalRow).AutoFilter Field:=13, Criteria1:="Change"
        If .Range("B1:B" & curTotalRow).SpecialCells(xlCellTypeVisible).Count > 1 Then
            .Range("B2:D" & curTotalRow & ",G2:H" & curTotalRow & ",J2:K" & curTotalRow & ",N2:N" & curTotalRow).SpecialCells(xlCellTypeVisible).Copy
            With Worksheets("Change")
                chaTotalRow = chaTotalRow2 + 1
                .Range("B" & chaTotalRow).PasteSpecial xlPasteValues
                chaTotalRow2 = .Range("B1").End(xlDown).Row
                .Range("A" & chaTotalRow & ":A" & chaTotalRow2).FormulaR1C1 = Sheets("Type").Range("A4").Value
            End With
        End If
        .AutoFilterMode = False
    End With
    
    ' 將英文名變更員工貼到Change工作表
    With Worksheets("Previous")
        .Range("A2:A" & preTotalRow).FormulaR1C1 = "=RC[1]&RC[3]"
    End With
    With Worksheets("Current")
        .Range("A2:A" & curTotalRow).FormulaR1C1 = "=RC[1]&RC[3]"
        .Range("N2:N" & curTotalRow).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-2],Previous!C2:C4,3,FALSE),""New"")"
        .Range("A1:N" & curTotalRow).AutoFilter Field:=13, Criteria1:="Change"
        If .Range("B1:B" & curTotalRow).SpecialCells(xlCellTypeVisible).Count > 1 Then
            .Range("B2:D" & curTotalRow & ",G2:H" & curTotalRow & ",J2:K" & curTotalRow & ",N2:N" & curTotalRow).SpecialCells(xlCellTypeVisible).Copy
            With Worksheets("Change")
                chaTotalRow = chaTotalRow2 + 1
                .Range("B" & chaTotalRow).PasteSpecial xlPasteValues
                chaTotalRow2 = .Range("B1").End(xlDown).Row
                .Range("A" & chaTotalRow & ":A" & chaTotalRow2).FormulaR1C1 = Sheets("Type").Range("A5").Value
            End With
        End If
        .AutoFilterMode = False
    End With
    
    ' 將電話變更員工貼到Change工作表
    With Worksheets("Previous")
        .Range("A2:A" & preTotalRow).FormulaR1C1 = "=RC[1]&RC[6]"
    End With
    With Worksheets("Current")
        .Range("A2:A" & curTotalRow).FormulaR1C1 = "=RC[1]&RC[6]"
        .Range("N2:N" & curTotalRow).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-2],Previous!C2:C7,6,FALSE),""New"")"
        .Range("A1:N" & curTotalRow).AutoFilter Field:=13, Criteria1:="Change"
        If .Range("B1:B" & curTotalRow).SpecialCells(xlCellTypeVisible).Count > 1 Then
            .Range("B2:D" & curTotalRow & ",G2:H" & curTotalRow & ",J2:K" & curTotalRow & ",N2:N" & curTotalRow).SpecialCells(xlCellTypeVisible).Copy
            With Worksheets("Change")
                chaTotalRow = chaTotalRow2 + 1
                .Range("B" & chaTotalRow).PasteSpecial xlPasteValues
                chaTotalRow2 = .Range("B1").End(xlDown).Row
                .Range("A" & chaTotalRow & ":A" & chaTotalRow2).FormulaR1C1 = Sheets("Type").Range("A6").Value
            End With
        End If
        .AutoFilterMode = False
    End With
    
    ' 將Email變更員工貼到Change工作表
    With Worksheets("Previous")
        .Range("A2:A" & preTotalRow).FormulaR1C1 = "=RC[1]&RC[7]"
    End With
    With Worksheets("Current")
        .Range("A2:A" & curTotalRow).FormulaR1C1 = "=RC[1]&RC[7]"
        .Range("N2:N" & curTotalRow).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-2],Previous!C2:C8,7,FALSE),""New"")"
        .Range("A1:N" & curTotalRow).AutoFilter Field:=13, Criteria1:="Change"
        If .Range("B1:B" & curTotalRow).SpecialCells(xlCellTypeVisible).Count > 1 Then
            .Range("B2:D" & curTotalRow & ",G2:H" & curTotalRow & ",J2:K" & curTotalRow & ",N2:N" & curTotalRow).SpecialCells(xlCellTypeVisible).Copy
            With Worksheets("Change")
                chaTotalRow = chaTotalRow2 + 1
                .Range("B" & chaTotalRow).PasteSpecial xlPasteValues
                chaTotalRow2 = .Range("B1").End(xlDown).Row
                .Range("A" & chaTotalRow & ":A" & chaTotalRow2).FormulaR1C1 = Sheets("Type").Range("A7").Value
            End With
        End If
        .AutoFilterMode = False
    End With
    
    ' 將部門變更員工貼到Change工作表
    With Worksheets("Previous")
        .Range("A2:A" & preTotalRow).FormulaR1C1 = "=RC[1]&RC[9]"
    End With
    With Worksheets("Current")
        .Range("A2:A" & curTotalRow).FormulaR1C1 = "=RC[1]&RC[9]"
        .Range("N2:N" & curTotalRow).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-2],Previous!C2:C10,9,FALSE),""New"")"
        .Range("A1:N" & curTotalRow).AutoFilter Field:=13, Criteria1:="Change"
        If .Range("B1:B" & curTotalRow).SpecialCells(xlCellTypeVisible).Count > 1 Then
            .Range("B2:D" & curTotalRow & ",G2:H" & curTotalRow & ",J2:K" & curTotalRow & ",N2:N" & curTotalRow).SpecialCells(xlCellTypeVisible).Copy
            With Worksheets("Change")
                chaTotalRow = chaTotalRow2 + 1
                .Range("B" & chaTotalRow).PasteSpecial xlPasteValues
                chaTotalRow2 = .Range("B1").End(xlDown).Row
                .Range("A" & chaTotalRow & ":A" & chaTotalRow2).FormulaR1C1 = Sheets("Type").Range("A8").Value
            End With
        End If
        .AutoFilterMode = False
    End With
    
    ' 將職稱變更員工貼到Change工作表
    With Worksheets("Previous")
        .Range("A2:A" & preTotalRow).FormulaR1C1 = "=RC[1]&RC[10]"
    End With
    With Worksheets("Current")
        .Range("A2:A" & curTotalRow).FormulaR1C1 = "=RC[1]&RC[10]"
        .Range("N2:N" & curTotalRow).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-2],Previous!C2:C11,10,FALSE),""New"")"
        .Range("A1:N" & curTotalRow).AutoFilter Field:=13, Criteria1:="Change"
        If .Range("B1:B" & curTotalRow).SpecialCells(xlCellTypeVisible).Count > 1 Then
            .Range("B2:D" & curTotalRow & ",G2:H" & curTotalRow & ",J2:K" & curTotalRow & ",N2:N" & curTotalRow).SpecialCells(xlCellTypeVisible).Copy
            With Worksheets("Change")
                chaTotalRow = chaTotalRow2 + 1
                .Range("B" & chaTotalRow).PasteSpecial xlPasteValues
                chaTotalRow2 = .Range("B1").End(xlDown).Row
                .Range("A" & chaTotalRow & ":A" & chaTotalRow2).FormulaR1C1 = Sheets("Type").Range("A9").Value
            End With
        End If
        .AutoFilterMode = False
    End With
    Application.CutCopyMode = False
    
    With Worksheets("Current")
        .Rows("1:" & curTotalRow).Delete
    End With
    
    With Worksheets("Previous")
        .Rows("1:" & preTotalRow).Delete
    End With
    Application.ScreenUpdating = True
    
End Sub
