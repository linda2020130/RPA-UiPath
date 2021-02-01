Sub FindChange()

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

    ' Previous工作表加新欄位
    With Worksheets("Previous")
        preTotalRow = .Range("B1").End(xlDown).Row
        .Columns("E:E").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        .Range("E2:E" & preTotalRow).FormulaR1C1 = "=RC[-3]&RC[-1]&RC[6]"
    End With
    
    ' Current工作表加新欄位
    With Worksheets("Current")
        curTotalRow = .Range("B1").End(xlDown).Row
        .Columns("E:E").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        .Range("E2:E" & curTotalRow).FormulaR1C1 = "=RC[-3]&RC[-1]&RC[6]"
        
        ' 生成PM變更資料列
        .Range("T2:T" & curTotalRow).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-15],Previous!C5:C6,2,FALSE),""Not Exist"")"
        .Range("U2:U" & curTotalRow).FormulaR1C1 = "=IF(RC[-15]=RC[-1],""O"",""X"")"
        .Range("V2:V" & curTotalRow).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-17],Previous!C5:C9,5,FALSE),""Not Exist"")"
        .Range("A1:V" & curTotalRow).AutoFilter Field:=20, Criteria1:="<>Not Exist"
        .Range("A1:V" & curTotalRow).AutoFilter Field:=21, Criteria1:="X"
        If .Range("F1:F" & curTotalRow).SpecialCells(xlCellTypeVisible).Count > 1 Then
            .Range("B2:B" & curTotalRow & ",D2:D" & curTotalRow & ",F2:F" & curTotalRow & ",I2:I" & curTotalRow & ",T2:T" & curTotalRow & ",V2:V" & curTotalRow).SpecialCells(xlCellTypeVisible).Copy
            With Worksheets("Change")
                chaTotalRow = chaTotalRow2 + 1
                .Range("B" & chaTotalRow).PasteSpecial xlPasteValues
                chaTotalRow2 = .Range("B1").End(xlDown).Row
                .Range("A" & chaTotalRow & ":A" & chaTotalRow2).FormulaR1C1 = Sheets("Type").Range("A2").Value
            End With
        End If
        .AutoFilterMode = False
        
        ' 生成新增產品線之資料列
        .Range("T2:T" & curTotalRow).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-15],Previous!C5:C6,1,FALSE),""Not Exist"")"
        .Range("A1:V" & curTotalRow).AutoFilter Field:=20, Criteria1:="Not Exist"
        If .Range("F1:F" & curTotalRow).SpecialCells(xlCellTypeVisible).Count > 1 Then
            .Range("B2:B" & curTotalRow & ",D2:D" & curTotalRow & ",F2:F" & curTotalRow & ",I2:I" & curTotalRow).SpecialCells(xlCellTypeVisible).Copy
            With Worksheets("Change")
                chaTotalRow = chaTotalRow2 + 1
                .Range("B" & chaTotalRow).PasteSpecial xlPasteValues
                chaTotalRow2 = .Range("B1").End(xlDown).Row
                .Range("A" & chaTotalRow & ":A" & chaTotalRow2).FormulaR1C1 = Sheets("Type").Range("A3").Value
            End With
        End If
        .AutoFilterMode = False
    End With
    
    With Worksheets("Previous")
        ' 生成刪除產品線之資料列
        .Range("T2:T" & preTotalRow).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-15],Current!C5:C6,1,FALSE),""Not Exist"")"
        .Range("A1:T" & preTotalRow).AutoFilter Field:=20, Criteria1:="Not Exist"
        If .Range("F1:F" & preTotalRow).SpecialCells(xlCellTypeVisible).Count > 1 Then
            .Range("B2:B" & preTotalRow & ",D2:D" & preTotalRow & ",F2:F" & preTotalRow & ",I2:I" & preTotalRow).SpecialCells(xlCellTypeVisible).Copy
            With Worksheets("Change")
                chaTotalRow = chaTotalRow2 + 1
                .Range("B" & chaTotalRow).PasteSpecial xlPasteValues
                chaTotalRow2 = .Range("B1").End(xlDown).Row
                .Range("D" & chaTotalRow & ":E" & chaTotalRow2).Cut Destination:=.Range("F" & chaTotalRow & ":G" & chaTotalRow2)
                .Range("A" & chaTotalRow & ":A" & chaTotalRow2).FormulaR1C1 = Sheets("Type").Range("A4").Value
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
