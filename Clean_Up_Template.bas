Attribute VB_Name = "Clean_Up_Template"
Sub Clean_Up()

    Dim lastrow As Long
    Application.ScreenUpdating = False
    
''1.
'
    Workbooks.Open Filename:= _
        "Y:\Excel\MASTER FORMULARY DATA\MACROS\Multiple Forms.xlsx"

    Columns("A:b").Copy

    Windows("Different Strengths Test Doc.xlsx").Activate

    Sheets("sheet2").Select
    Range("a1").Select
    ActiveSheet.Paste
    Range("b99999").End(xlUp).Offset(1, 0).Select
    Selection = "End"

    Windows("Multiple Forms.xlsx").Activate

    Application.CutCopyMode = False
    ActiveWorkbook.Close False

    Sheets("sheet3").Select

    Sheets("sheet1").Select

    lastrow = [A20000].End(xlUp).Row
    Range("A2", Cells(lastrow, 10)).Copy
    Sheets("Sheet3").Select
    Range("A1").Select
    ActiveSheet.Paste

    lastrow = [A20000].End(xlUp).Row
    Range("B1", Cells(lastrow, 2)).Copy
    Range("K1").Select
    ActiveSheet.Paste

    Range("L1").Select
    ActiveSheet.Paste
    
    Sheets("Sheet2").Select
    Range("B20000").End(xlUp).Offset(1, 0).Select
    Selection = "End"
    Range("B1").Select
        
    Do Until Selection = "End"
    lookfor = Selection
    Sheets("Sheet3").Select
    Columns("L").Select
    
    
    Selection.Replace What:=lookfor, Replacement:="$$$", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False


    Sheets("Sheet2").Select
    Selection.Offset(1, 0).Select
    Loop

    Selection.ClearContents

    Sheets("Sheet3").Select
    
    lastrow = [B20000].End(xlUp).Row
    Range("M1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(COUNT(SEARCH(""$$$"",RC[-1])),""Keep"",""Clear"")"
    Range("M1").Select
    Selection.AutoFill Destination:=Range("M1", Cells(lastrow, 13)), Type:=xlFillDefault
    Range("M1", Cells(lastrow, 13)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Last = Cells(Rows.Count, "B").End(xlUp).Row
    For i = Last To 1 Step -1
        If (Cells(i, "M").Value) = "Clear" Then
        Cells(i, "M").ClearContents
        End If
    Next i
    
    Columns("K").Copy
    Range("L1").Select
    ActiveSheet.Paste
    
    Range("K20000").End(xlUp).Offset(1, 0).Select
    Selection = "End"
    Range("K1").Select
    
    Do Until Selection = "End"
    
    If Selection.Offset(0, 2) = "" Then
        
            Selection.Replace What:=" 1*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" 2*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" 3*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" 4*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" 5*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" 6*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" 7*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" 8*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" 9*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" 0*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" tab*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" subq*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" sub-q*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" pen*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" cap*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" oral*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" top*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
        Selection.Offset(1, 0).Select
        Else
        Selection.Offset(1, 0).Select
    End If
    Loop
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    lastrow = [A20000].End(xlUp).Row
    Range("B1", Cells(lastrow, 2)).Copy
    Range("L1").Select
    ActiveSheet.Paste

    Range("K200000").End(xlUp).Offset(1, 0).Select
    Selection = "End"

    Range("K1").Select
    
    Do Until Selection = "End"
    
    lookfor = Selection
    If Selection.Offset(0, 2) = "" Then
    
    Selection.Offset(0, 1).Select
    ActiveCell.Replace What:=lookfor, Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, -1).Select
    Else
    Selection.Offset(1, 0).Select
    End If
    Loop

    Selection.ClearContents

    lastrow = [K200000].End(xlUp).Row
    Range("K1", Cells(lastrow, 11)).Cut
    Range("B1").Select
    ActiveSheet.Paste
    Columns("K").Delete
    
    Last = Cells(Rows.Count, "B").End(xlUp).Row
    For i = Last To 1 Step -1
        If (Cells(i, "L").Value) = "Keep" Then
        Cells(i, "K").ClearContents
        End If
    Next i

    
    Columns("L").Delete
    
    Last = Cells(Rows.Count, "B").End(xlUp).Row
    For i = Last To 1 Step -1
        If (Cells(i, "B").Value) = "End" Then
        Cells(i, "A").EntireRow.Delete
        End If
    Next i
    
    lastrow = [A20000].End(xlUp).Row
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "=TRIM(SUBSTITUTE(RC[-1],CHAR(160),CHAR(32)))"
    Range("L1").Select
    Selection.AutoFill Destination:=Range("L1", Cells(lastrow, 12)), Type:=xlFillDefault
    Range("L1", Cells(lastrow, 12)).Select
    Selection.Copy
    Range("K1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("L").Delete

    Cells.FormatConditions.Delete
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("A:K").Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("Sheet3").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet3").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("E1:E3888"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet3").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Sheet3").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet3").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("B1:B3888"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet3").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Rows("1").Delete

''''''''''''''''''''''''COMBINE TIER INFO''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Last = Cells(Rows.Count, "A").End(xlUp).Row
    For i = Last To 1 Step -1
        If (Cells(i, "K").Value) <> "" Then
        Cells(i, "K").Value = Cells(i, "K") & " is " & Cells(i, "C") & " Tier " & Cells(i, "E")
        End If
    Next i
''''''''''''''''''''''''RESTRICTIONS - SOLO''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Last = Cells(Rows.Count, "A").End(xlUp).Row
    For i = Last To 1 Step -1
        If (Cells(i, "K").Value) <> "" And Cells(i, "G") = "Y" And Cells(i, "H") = "" And Cells(i, "I") = "" And Cells(i, "J") = "" Then
        Cells(i, "K") = Cells(i, "K") & " with a quantity limit"
        End If
        If (Cells(i, "K").Value) <> "" And Cells(i, "G") = "" And Cells(i, "H") = "Y" And Cells(i, "I") = "" And Cells(i, "J") = "" Then
        Cells(i, "K") = Cells(i, "K") & " with a prior authorization"
        End If
        If (Cells(i, "K").Value) <> "" And Cells(i, "G") = "" And Cells(i, "H") = "" And Cells(i, "I") = "Y" And Cells(i, "J") = "" Then
        Cells(i, "K") = Cells(i, "K") & " with a step therapy"
        End If
        If (Cells(i, "K").Value) <> "" And Cells(i, "G") = "" And Cells(i, "H") = "" And Cells(i, "I") = "" And Cells(i, "J") <> "" Then
        Cells(i, "K") = Cells(i, "K") & ". " & Cells(i, "J")
        End If
    Next i
''''''''''''''''''''''''RESTRICTIONS - DOUBLE''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Last = Cells(Rows.Count, "A").End(xlUp).Row
    For i = Last To 1 Step -1
        If (Cells(i, "K").Value) <> "" And Cells(i, "G") = "Y" And Cells(i, "H") = "Y" And Cells(i, "I") = "" And Cells(i, "J") = "" Then
        Cells(i, "K") = Cells(i, "K") & " with a quantity limit and prior authorization"
        End If
        If (Cells(i, "K").Value) <> "" And Cells(i, "G") = "Y" And Cells(i, "H") = "" And Cells(i, "I") = "Y" And Cells(i, "J") = "" Then
        Cells(i, "K") = Cells(i, "K") & " with a quantity limit and step therapy"
        End If
        If (Cells(i, "K").Value) <> "" And Cells(i, "G") = "Y" And Cells(i, "H") = "" And Cells(i, "I") = "" And Cells(i, "J") <> "" Then
        Cells(i, "K") = Cells(i, "K") & " with a quantity limit: " & Cells(i, "J")
        End If
    Next i
    Last = Cells(Rows.Count, "A").End(xlUp).Row
    For i = Last To 1 Step -1
        If (Cells(i, "K").Value) <> "" And Cells(i, "G") = "" And Cells(i, "H") = "Y" And Cells(i, "I") = "Y" And Cells(i, "J") = "" Then
        Cells(i, "K") = Cells(i, "K") & " with a prior authorization and step therapy"
        End If
        If (Cells(i, "K").Value) <> "" And Cells(i, "G") = "" And Cells(i, "H") = "Y" And Cells(i, "I") = "" And Cells(i, "J") <> "" Then
        Cells(i, "K") = Cells(i, "K") & " with a prior authorization: " & Cells(i, "J")
        End If
    Next i
    For i = Last To 1 Step -1
        If (Cells(i, "K").Value) <> "" And Cells(i, "G") = "" And Cells(i, "H") = "" And Cells(i, "I") = "Y" And Cells(i, "J") <> "" Then
        Cells(i, "K") = Cells(i, "K") & " with a step therapy: " & Cells(i, "J")
        End If
    Next i
''''''''''''''''''''''''RESTRICTIONS - TRIPLE''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = Last To 1 Step -1
        If (Cells(i, "K").Value) <> "" And Cells(i, "G") = "Y" And Cells(i, "H") = "Y" And Cells(i, "I") = "Y" And Cells(i, "J") = "" Then
        Cells(i, "K") = Cells(i, "K") & " with a quantity limit, prior authorization and step therapy"
        End If
        If (Cells(i, "K").Value) <> "" And Cells(i, "G") = "Y" And Cells(i, "H") = "" And Cells(i, "I") = "Y" And Cells(i, "J") <> "" Then
        Cells(i, "K") = Cells(i, "K") & " with a quantity limit and step therapy:" & Cells(i, "J")
        End If
        If (Cells(i, "K").Value) <> "" And Cells(i, "G") = "Y" And Cells(i, "H") = "Y" And Cells(i, "I") = "" And Cells(i, "J") <> "" Then
        Cells(i, "K") = Cells(i, "K") & " with a quantity limit and prior authorization: " & Cells(i, "J")
        End If
    Next i
    For i = Last To 1 Step -1
        If (Cells(i, "K").Value) <> "" And Cells(i, "G") = "" And Cells(i, "H") = "Y" And Cells(i, "I") = "Y" And Cells(i, "J") <> "" Then
        Cells(i, "K") = Cells(i, "K") & " with a prior authorization and step therapy: " & Cells(i, "J")
        End If
    Next i
''''''''''''''''''''''''RESTRICTIONS - QUAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = Last To 1 Step -1
        If (Cells(i, "K").Value) <> "" And Cells(i, "G") = "Y" And Cells(i, "H") = "Y" And Cells(i, "I") = "Y" And Cells(i, "J") <> "" Then
        Cells(i, "K") = Cells(i, "K") & " with a quantity limit, prior authorization and step therapy: " & Cells(i, "J")
        End If
    Next i









'''            If (Cells(i, "G").Value) <> "" Then
'''            Cells(i, "K") = Cells(i, "K") & " with a quantity limit."
'''            End If
'''            If (Cells(i, "H").Value) <> "" Then
'''            Cells(i, "K") = Cells(i, "K") & " with a prior authorization."
'''            End If
'''            If (Cells(i, "I").Value) <> "" Then
'''            Cells(i, "K") = Cells(i, "K") & " with a step therapy."
'''            End If

    lastrow = [B20000].End(xlUp).Row
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "=TRIM(SUBSTITUTE(RC[-11],CHAR(160),CHAR(32)))"
    Range("L1").Select
    Selection.AutoFill Destination:=Range("L1:V1"), Type:=xlFillDefault
    Range("L1:V1").Select
    Selection.AutoFill Destination:=Range("L1", Cells(lastrow, 22)), Type:=xlFillDefault
    Range("L1", Cells(lastrow, 22)).Select
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("L:V").Delete

'''''TEST'''''









    Last = Cells(Rows.Count, "A").End(xlUp).Row
    For i = Last To 2 Step -1
        If UCase(Cells(i, "B").Value) = UCase(Cells(i - 1, "B")) And Cells(i - 1, "K") = "" Then
        Cells(i - 1, "K") = Cells(i, "K")
        Cells(i, "A").EntireRow.Delete
        End If
    Next i

    Last = Cells(Rows.Count, "A").End(xlUp).Row
    For i = Last To 2 Step -1
        If UCase(Cells(i, "B").Value) = UCase(Cells(i - 1, "B")) Then
        Cells(i - 1, "K") = Cells(i - 1, "K") & "% " & Cells(i, "K")
        Cells(i, "A").EntireRow.Delete
        End If
    Next i


    Columns("K:k").Cut
    Range("L1").Select
    ActiveSheet.Paste

    Last = Cells(Rows.Count, "A").End(xlUp).Row
    For i = Last To 1 Step -1
        If (Cells(i, "L").Value) <> "" Then
        Cells(i, "K") = Cells(i, "c") & " Tier " & Cells(i, "E")
        End If
    Next i

    Columns("L:L").Select
    Selection.TextToColumns Destination:=Range("L1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="%", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True

    Range("A1").Select
    ActiveCell.SpecialCells(xlLastCell).Select
    Selection.Offset(0, 1).End(xlUp).Select
    Selection = "End"

    Range("A200000").End(xlUp).Offset(1, 10).Select
    Selection = "Stop"

    Range("K1").Select

    Do Until Selection = "Stop"
    If Selection = "" Then
    Selection.Offset(1, 0).Select
    Else

    lookfor = Selection
    Selection.Offset(0, 1).Select
    Selection.Replace What:=lookfor, Replacement:="@@@", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False


    Selection.Offset(1, -1).Select
    End If
    Loop
    Selection.ClearContents

    Columns("L").Select

    Selection.Replace What:="*@@@*", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    Range("M1").Select


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Range("A200000").End(xlUp).Offset(1, 10).Select
    Selection = "Stop"

    Range("A200000").End(xlUp).Offset(1, 12).Select
    Selection = "Stop"

    Range("K1").Select

    Do Until Selection = "End"

    Do Until Selection = "Stop"

    If Selection = "" Then
    Selection.Offset(1, 0).Select
    Else

    lookfor = Selection
    Selection.Offset(0, 2).Select

    Selection.Replace What:=lookfor, Replacement:="@@@", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

        Selection.Offset(1, -2).Select
    End If
    Loop



    Columns("M").Select

    Selection.Replace What:="*@@@*", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    Last = Cells(Rows.Count, "A").End(xlUp).Row
    For i = Last To 1 Step -1
        If (Cells(i, "M").Value) <> "" And (Cells(i, "L").Value) <> "" Then
            Cells(i, "L") = Cells(i, "L") & ". " & Cells(i, "M")
            End If
        If (Cells(i, "L").Value) = "" And (Cells(i, "M").Value) <> "" Then
            Cells(i, "L") = Cells(i, "M")
            Else
            Cells(i, "L") = Cells(i, "L")
        End If
    Next i

    
    Columns("M").Delete
    Range("M1").Select



    Range("A200000").End(xlUp).Offset(1, 12).Select
    Selection = "Stop"
    
    Range("m1").Select

    Loop

    Columns("m:m").Delete
    Range("k999999").End(xlUp).ClearContents

    
    Columns("k:K").Delete
    
    lastrow = [k99999].End(xlUp).Row
    Range("l1").Select
    ActiveCell.FormulaR1C1 = "=TRIM(SUBSTITUTE(RC[-1],CHAR(160),CHAR(32)))"
    Range("l1").Select
    Selection.AutoFill Destination:=Range("l1", Cells(lastrow, 12)), Type:=xlFillDefault
    Range("l1", Cells(lastrow, 12)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    Columns("K:k").Delete
    



    Last = Cells(Rows.Count, "A").End(xlUp).Row
    For i = Last To 1 Step -1
        If (Cells(i, "J").Value) <> "" Then
        Cells(i, "J").Value = Cells(i, "J") & ". " & Cells(i, "K")
        Else
        Cells(i, "J").Value = Cells(i, "K")
        End If
    Next i
    Columns("K").Delete



'''''END TEST''''''





    lastrow = [B20000].End(xlUp).Row
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "=TRIM(SUBSTITUTE(RC[-11],CHAR(160),CHAR(32)))"
    Range("L1").Select
    Selection.AutoFill Destination:=Range("L1:V1"), Type:=xlFillDefault
    Range("L1:V1").Select
    Selection.AutoFill Destination:=Range("L1", Cells(lastrow, 22)), Type:=xlFillDefault
    Range("L1", Cells(lastrow, 22)).Select
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("L:W").Delete

    Columns("A:J").Select
    Selection.NumberFormat = "General"
    Columns("K").Delete

''''''''''''''''''''''''''Import back into sheet 1''''''''''''''''''''''''''''''''''''''''

    Sheets("sheet1").Select
    lastrow = [b999999].End(xlUp).Row
    Range("b2", Cells(lastrow, 2)).ClearContents
    
    Last = Cells(Rows.Count, "A").End(xlUp).Row
    For i = Last To 2 Step -1
        If (Cells(i, "B").Value) = "" Then
        Cells(i, "A").EntireRow.Delete
        End If
    Next i
    
    Sheets("sheet3").Select
    
    lastrow = [a999999].End(xlUp).Row
    Range("a1", Cells(lastrow, 10)).Cut
    Sheets("sheet1").Select
    Range("a2").Select
    ActiveSheet.Paste
    
    Sheets("sheet3").Select
    Columns("a:z").Delete
    Sheets("sheet1").Select
    
    Sheets("Sheet2").Select
    Columns("A:B").Delete
    Range("A1").Select
    Sheets("Sheet1").Select
    
'    Application.CutCopyMode = False
'    ActiveWorkbook.Close True
    

End Sub
Sub Clean_Up_Multiple_Forms()

    Dim lastrow As Long
    
    Workbooks.Open Filename:= _
        "Y:\Excel\MASTER FORMULARY DATA\MACROS\Multiple Forms.xlsx"

    Columns("A:b").Copy

    Windows("Different Strengths Test Doc.xlsx").Activate

    Sheets("sheet2").Select
    Range("a1").Select
    ActiveSheet.Paste
    Range("b99999").End(xlUp).Offset(1, 0).Select
    Selection = "End"

    Windows("Multiple Forms.xlsx").Activate

    Application.CutCopyMode = False
    ActiveWorkbook.Close False

    Sheets("sheet3").Select

    Sheets("sheet1").Select

    lastrow = [A20000].End(xlUp).Row
    Range("A2", Cells(lastrow, 10)).Copy
    Sheets("Sheet3").Select
    Range("A1").Select
    ActiveSheet.Paste

    lastrow = [A20000].End(xlUp).Row
    Range("B1", Cells(lastrow, 2)).Copy
    Range("K1").Select
    ActiveSheet.Paste

    Range("L1").Select
    ActiveSheet.Paste
    
    Sheets("Sheet2").Select
    Range("B20000").End(xlUp).Offset(1, 0).Select
    Selection = "End"
    Range("B1").Select
        
    Do Until Selection = "End"
    lookfor = Selection
    Sheets("Sheet3").Select
    Columns("L").Select
    
    
    Selection.Replace What:=lookfor, Replacement:="$$$", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False


    Sheets("Sheet2").Select
    Selection.Offset(1, 0).Select
    Loop

    Selection.ClearContents

    Sheets("Sheet3").Select
    
    lastrow = [B20000].End(xlUp).Row
    Range("M1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(COUNT(SEARCH(""$$$"",RC[-1])),""Keep"",""Clear"")"
    Range("M1").Select
    Selection.AutoFill Destination:=Range("M1", Cells(lastrow, 13)), Type:=xlFillDefault
    Range("M1", Cells(lastrow, 13)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Last = Cells(Rows.Count, "B").End(xlUp).Row
    For i = Last To 1 Step -1
        If (Cells(i, "M").Value) = "Clear" Then
        Cells(i, "M").ClearContents
        End If
    Next i
    
    Columns("K").Copy
    Range("L1").Select
    ActiveSheet.Paste
    
    Range("K20000").End(xlUp).Offset(1, 0).Select
    Selection = "End"
    Range("K1").Select
    
    Do Until Selection = "End"
    
    If Selection.Offset(0, 2) = "" Then
        
            Selection.Replace What:=" 1*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" 2*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" 3*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" 4*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" 5*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" 6*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" 7*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" 8*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" 9*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" 0*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" tab*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" subq*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" sub-q*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" pen*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" cap*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" oral*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            Selection.Replace What:=" top*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
        Selection.Offset(1, 0).Select
        Else
        Selection.Offset(1, 0).Select
    End If
    Loop
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    lastrow = [A20000].End(xlUp).Row
    Range("B1", Cells(lastrow, 2)).Copy
    Range("L1").Select
    ActiveSheet.Paste

    Range("K200000").End(xlUp).Offset(1, 0).Select
    Selection = "End"

    Range("K1").Select
    
    Do Until Selection = "End"
    
    lookfor = Selection
    If Selection.Offset(0, 2) = "" Then
    
    Selection.Offset(0, 1).Select
    ActiveCell.Replace What:=lookfor, Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, -1).Select
    Else
    Selection.Offset(1, 0).Select
    End If
    Loop

    Selection.ClearContents

    lastrow = [K200000].End(xlUp).Row
    Range("K1", Cells(lastrow, 11)).Cut
    Range("B1").Select
    ActiveSheet.Paste
    Columns("K").Delete





End Sub
