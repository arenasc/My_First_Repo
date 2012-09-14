Attribute VB_Name = "Four_D"
Sub fourd_Clean_Up()
Attribute fourd_Clean_Up.VB_ProcData.VB_Invoke_Func = " \n14"

    Application.ScreenUpdating = False
    
    Workbooks.Open Filename:= _
        "C:\Excel\MASTER FORMULARY DATA\EXCEL EXTRACTIONS\4D Pharmacy Management Systems, Inc\305_3TO_4D Pharmacy Management Systems, Inc. EXTRACTION.xlsx"

'For column deletion

    Columns("D").Delete

'Clean Up Headers

   Cells.Replace What:="www.*", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="managed drug formulary*", Replacement:="", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="co-pay", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="dispensing", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

'Delete Blank Rows

    Columns("A:A").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete

    Last = Cells(Rows.Count, "A").End(xlUp).Row
    For i = Last To 1 Step -1
        If (Cells(i, "A").Value) = "Drug Name" Then
        Cells(i, "A").EntireRow.Delete
        End If
    Next i
    
'Restrictions

    lastrow = [A20000].End(xlUp).Row
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "=IF(COUNT(SEARCH(""QL"",RC3)),""QL"",""Clear"")"
    Range("D1").Select
    Selection.AutoFill Destination:=Range("D1", Cells(lastrow, 4)), Type:=xlFillDefault

    lastrow = [A20000].End(xlUp).Row
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=IF(COUNT(SEARCH(""PA"",RC3)),""PA"",""Clear"")"
    Range("E1").Select
    Selection.AutoFill Destination:=Range("E1", Cells(lastrow, 5)), Type:=xlFillDefault

    lastrow = [A20000].End(xlUp).Row
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "=IF(COUNT(SEARCH(""ST"",RC3)),""ST"",""Clear"")"
    Range("F1").Select
    Selection.AutoFill Destination:=Range("F1", Cells(lastrow, 6)), Type:=xlFillDefault

    lastrow = [A20000].End(xlUp).Row
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=IF(COUNT(SEARCH(""SP"",RC3)),""Specialty"",""Clear"")"
    Range("G1").Select
    Selection.AutoFill Destination:=Range("G1", Cells(lastrow, 7)), Type:=xlFillDefault
    
    lastrow = [A20000].End(xlUp).Row
    Range("D1", Cells(lastrow, 7)).Copy
    Range("D1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Last = Cells(Rows.Count, "D").End(xlUp).Row
    For i = Last To 1 Step -1
        If (Cells(i, "D").Value) = "Clear" Then
        Cells(i, "D").ClearContents
        End If
    Next i

    Last = Cells(Rows.Count, "E").End(xlUp).Row
    For i = Last To 1 Step -1
        If (Cells(i, "E").Value) = "Clear" Then
        Cells(i, "E").ClearContents
        End If
    Next i

    Last = Cells(Rows.Count, "F").End(xlUp).Row
    For i = Last To 1 Step -1
        If (Cells(i, "F").Value) = "Clear" Then
        Cells(i, "F").ClearContents
        End If
    Next i

    Last = Cells(Rows.Count, "G").End(xlUp).Row
    For i = Last To 1 Step -1
        If (Cells(i, "G").Value) = "Clear" Then
        Cells(i, "G").ClearContents
        End If
    Next i

    Columns("G").Delete
    
''2. Diabetics

    Range("A1").Select
    On Error GoTo NextD1
    Cells.Find(What:="*Accu?chek*", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 6) = "QL: 100 strips/month"
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert

    ActiveCell.End(xlDown).Select
    Range(Selection, Cells(Selection.Row, 8)).Copy
    Selection.End(xlUp).Offset(1, 0).Select
    Range(ActiveCell, Cells(ActiveCell.End(xlDown).Row, ActiveCell.Column)).Select
    ActiveSheet.Paste
    
    ActiveCell.Select
    Selection.Replace What:="*Accu?chek*", Replacement:="Accu-chek", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*Accu?chek*", Replacement:="Accu-Chek Active Test Strips", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*Accu?chek*", Replacement:="Accu-Chek Aviva Test Strips", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*Accu?chek*", Replacement:="Accu-Chek Compact Drum Strips", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select

NextD1:

    Range("A1").Select
    On Error GoTo NextD2a
    Cells.Find(What:="*apidra*", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    Rows(ActiveCell.Row).Insert

    ActiveCell.End(xlDown).Select
    Range(Selection, Cells(Selection.Row, 8)).Copy
    Selection.End(xlUp).Offset(1, 0).Select
    Range(ActiveCell, Cells(ActiveCell.End(xlDown).Row, ActiveCell.Column)).Select
    ActiveSheet.Paste
    
    ActiveCell.Select
    Selection.Replace What:="*apidra*", Replacement:="Apidra Solostar", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*apidra*", Replacement:="Apidra Vial", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select

NextD2a:


    Range("A1").Select
    On Error GoTo NextD2
    Cells.Find(What:="*bd*", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    Rows(ActiveCell.Row).Insert

    ActiveCell.End(xlDown).Select
    Range(Selection, Cells(Selection.Row, 8)).Copy
    Selection.End(xlUp).Offset(1, 0).Select
    Range(ActiveCell, Cells(ActiveCell.End(xlDown).Row, ActiveCell.Column)).Select
    ActiveSheet.Paste
    
    ActiveCell.Select
    Selection.Replace What:="*bd*", Replacement:="BD Ultra-Fine Needle", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select

NextD2:

    Range("A1").Select
    On Error GoTo NextD3
    Cells.Find(What:="*breeze*", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 6) = "QL: 100 strips/month"
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert

    ActiveCell.End(xlDown).Select
    Range(Selection, Cells(Selection.Row, 8)).Copy
    Selection.End(xlUp).Offset(1, 0).Select
    Range(ActiveCell, Cells(ActiveCell.End(xlDown).Row, ActiveCell.Column)).Select
    ActiveSheet.Paste
    
    ActiveCell.Select
    Selection.Replace What:="*breeze*", Replacement:="Breeze 2", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*breeze*", Replacement:="BREEZE 2 DISC TEST STRIP", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select

NextD3:

    Range("A1").Select
    On Error GoTo NextD4
    Cells.Find(What:="*contour*", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 6) = "QL: 100 strips/month"
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert

    ActiveCell.End(xlDown).Select
    Range(Selection, Cells(Selection.Row, 8)).Copy
    Selection.End(xlUp).Offset(1, 0).Select
    Range(ActiveCell, Cells(ActiveCell.End(xlDown).Row, ActiveCell.Column)).Select
    ActiveSheet.Paste
    
    ActiveCell.Select
    Selection.Replace What:="*contour*", Replacement:="Contour", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*contour*", Replacement:="CONTOUR TEST STRIPS", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select

NextD4:

    Range("A1").Select
    On Error GoTo NextD5
    Cells.Find(What:="*freestyle*", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 6) = "QL: 100 strips/month"
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert

    ActiveCell.End(xlDown).Select
    Range(Selection, Cells(Selection.Row, 8)).Copy
    Selection.End(xlUp).Offset(1, 0).Select
    Range(ActiveCell, Cells(ActiveCell.End(xlDown).Row, ActiveCell.Column)).Select
    ActiveSheet.Paste
    
    ActiveCell.Select
    Selection.Replace What:="*freestyle*", Replacement:="FreeStyle Freedom Lite", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*freestyle*", Replacement:="FreeStyle Lite", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*freestyle*", Replacement:="FreeStyle Lite", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*freestyle*", Replacement:="Freestyle Test Strips", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select

NextD5:

    Range("A1").Select
    On Error GoTo NextD7
    Cells.Find(What:="Humalog", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert

    ActiveCell.End(xlDown).Select
    Range(Selection, Cells(Selection.Row, 8)).Copy
    Selection.End(xlUp).Offset(1, 0).Select
    Range(ActiveCell, Cells(ActiveCell.End(xlDown).Row, ActiveCell.Column)).Select
    ActiveSheet.Paste
    
    ActiveCell.Select
    Selection.Replace What:="*Humalog*", Replacement:="Humalog Cartridge", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*Humalog*", Replacement:="Humalog Kwikpen", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*Humalog*", Replacement:="Humalog Mix 50/50 Kwikpen", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*Humalog*", Replacement:="Humalog Mix 50/50 Pen", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*Humalog*", Replacement:="Humalog Mix 50/50 Vial", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*Humalog*", Replacement:="Humalog Mix 75/25 Kwikpen", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*Humalog*", Replacement:="Humalog Mix 75/25 Pen", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*Humalog*", Replacement:="Humalog Mix 75/25 Vial", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*Humalog*", Replacement:="Humalog Pen", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*Humalog*", Replacement:="Humalog Vial", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select

NextD7:

    Range("A1").Select
    On Error GoTo NextD8
    Cells.Find(What:="Humulin", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert

    ActiveCell.End(xlDown).Select
    Range(Selection, Cells(Selection.Row, 8)).Copy
    Selection.End(xlUp).Offset(1, 0).Select
    Range(ActiveCell, Cells(ActiveCell.End(xlDown).Row, ActiveCell.Column)).Select
    ActiveSheet.Paste
    
    ActiveCell.Select
    Selection.Replace What:="*Humulin*", Replacement:="Humulin 50/50 Vial", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*Humulin*", Replacement:="Humulin 70/30 Cartridge", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*Humulin*", Replacement:="Humulin 70/30 Pen", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*Humulin*", Replacement:="Humulin 70/30 Vial", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*Humulin*", Replacement:="Humulin N Cartridge", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*Humulin*", Replacement:="Humulin N Pen", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*Humulin*", Replacement:="Humulin N Vial", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*Humulin*", Replacement:="Humulin R Cartridge", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*Humulin*", Replacement:="Humulin R Vial", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select

NextD8:

    Range("A1").Select
    On Error GoTo NextD9
    Cells.Find(What:="lantus", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert

    ActiveCell.End(xlDown).Select
    Range(Selection, Cells(Selection.Row, 8)).Copy
    Selection.End(xlUp).Offset(1, 0).Select
    Range(ActiveCell, Cells(ActiveCell.End(xlDown).Row, ActiveCell.Column)).Select
    ActiveSheet.Paste
    
    ActiveCell.Select
    Selection.Replace What:="*lantus*", Replacement:="Lantus OptiClik", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*lantus*", Replacement:="Lantus Solostar", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*lantus*", Replacement:="Lantus Vial", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select

NextD9:

    Range("A1").Select
    On Error GoTo NextD10
    Cells.Find(What:="Levemir", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert

    ActiveCell.End(xlDown).Select
    Range(Selection, Cells(Selection.Row, 8)).Copy
    Selection.End(xlUp).Offset(1, 0).Select
    Range(ActiveCell, Cells(ActiveCell.End(xlDown).Row, ActiveCell.Column)).Select
    ActiveSheet.Paste
    
    ActiveCell.Select
    Selection.Replace What:="*Levemir*", Replacement:="Levemir FlexPen", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*Levemir*", Replacement:="Levemir Vial", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select

NextD10:

    Range("A1").Select
    On Error GoTo NextD11
    Cells.Find(What:="Novolin", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert

    ActiveCell.End(xlDown).Select
    Range(Selection, Cells(Selection.Row, 8)).Copy
    Selection.End(xlUp).Offset(1, 0).Select
    Range(ActiveCell, Cells(ActiveCell.End(xlDown).Row, ActiveCell.Column)).Select
    ActiveSheet.Paste
    
    ActiveCell.Select
    Selection.Replace What:="*Novolin*", Replacement:="Novolin 70/30 Vial", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*Novolin*", Replacement:="Novolin N Vial", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*Novolin*", Replacement:="Novolin R Vial", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select

NextD11:

    Range("A1").Select
    On Error GoTo NextD12
    Cells.Find(What:="Novolog", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert

    ActiveCell.End(xlDown).Select
    Range(Selection, Cells(Selection.Row, 8)).Copy
    Selection.End(xlUp).Offset(1, 0).Select
    Range(ActiveCell, Cells(ActiveCell.End(xlDown).Row, ActiveCell.Column)).Select
    ActiveSheet.Paste
    
    ActiveCell.Select
    Selection.Replace What:="*Novolog*", Replacement:="Novolog Cartridge", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*Novolog*", Replacement:="NovoLog FlexPen", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*Novolog*", Replacement:="Novolog Mix 70/30 Cartridge", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*Novolog*", Replacement:="NovoLog Mix 70/30 FlexPen", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*Novolog*", Replacement:="Novolog Mix 70/30 Vial", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*Novolog*", Replacement:="Novolog Vial", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select

NextD12:

    Range("A1").Select
    On Error GoTo NextD15
    Cells.Find(What:="One Touch*", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 6) = "QL: 100 strips/month"
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert

    ActiveCell.End(xlDown).Select
    Range(Selection, Cells(Selection.Row, 8)).Copy
    Selection.End(xlUp).Offset(1, 0).Select
    Range(ActiveCell, Cells(ActiveCell.End(xlDown).Row, ActiveCell.Column)).Select
    ActiveSheet.Paste
    
    ActiveCell.Select
    Selection.Replace What:="*One Touch*", Replacement:="One Touch Test Strips", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*One Touch*", Replacement:="One Touch Ultra Test Strips", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*One Touch*", Replacement:="One Touch UltraMini Meter", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*One Touch*", Replacement:="One Touch UltraSoft Lancets", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*One Touch*", Replacement:="OneTouch", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select

NextD15:

    Range("A1").Select
    On Error GoTo NextD16
    Cells.Find(What:="true*", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 6) = "QL: 100 strips/month"
    Rows(ActiveCell.Row).Insert
    Rows(ActiveCell.Row).Insert

    ActiveCell.End(xlDown).Select
    Range(Selection, Cells(Selection.Row, 8)).Copy
    Selection.End(xlUp).Offset(1, 0).Select
    Range(ActiveCell, Cells(ActiveCell.End(xlDown).Row, ActiveCell.Column)).Select
    ActiveSheet.Paste
    
    ActiveCell.Select
    Selection.Replace What:="*true*", Replacement:="TRUETEST GLUCOSE TEST STRIPS", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select
    
    Selection.Replace What:="*true*", Replacement:="TRUETRACK GLUCOSE TEST STRIPS", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Offset(1, 0).Select

NextD16:

    Range("A20000").End(xlUp).Offset(1, 0).Select
    Selection = "NovoFine Autocover"
    Selection.Offset(1, 0).Select
    Selection = "NovoFine PenNeedles"
    Selection.Offset(1, 0).Select
    
    Last = Cells(Rows.Count, "A").End(xlUp).Row
    For i = Last To 1 Step -1
        If (Cells(i, "B").Value) = "" Then
        Cells(i, "B") = "3"
        Cells(i, "D") = "QL"
        End If
    Next i

    Last = Cells(Rows.Count, "A").End(xlUp).Row
    For i = Last To 1 Step -1
        If UCase(Cells(i, "A").Value) = "NUTROPIN" Then
        Cells(i, "A").Select
        Range(Selection, Cells(Selection.Row, 10)).Copy
        Range("A200000").End(xlUp).Offset(1, 0).Select
        ActiveSheet.Paste
        ActiveCell.Select
        Selection = "NUTROPIN AQ"
        End If
    Next i

    Call Four_D_Mini_Database

End Sub

Sub Four_D_Mini_Database()

    Windows("305_3TO_4D Pharmacy Management Systems, Inc. (Mini Database).xlsx").Activate

    Sheets("Sheet1").Select
    Columns("A:J").Delete

    Range("A1") = "Plan Name"
    Range("B1") = "Brand Name"
    Range("C1") = "Formulary Status"
    Range("D1") = "Benefit Design"
    Range("E1") = "Tier"
    Range("F1") = "Copay"
    Range("G1") = "Quantity Limit (QL)"
    Range("H1") = "Prior Authorization (PA)"
    Range("I1") = "Step Therapy (ST)"
    Range("J1") = "Comments"

    Range("A1:J1").Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$J$1"), , xlYes).Name = _
        "Table1"
    Range("Table1[#All]").Select
    ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleLight1"
    ActiveSheet.ListObjects("Table1").TableStyle = ""

    Range("A1:J1").Select
    Selection.Font.Bold = True

    Range("Table1[#Headers]").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Range("Table1[[#Headers],[Plan Name]:[Copay]]").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("Table1[[#Headers],[Quantity Limit (QL)]:[Step Therapy (ST)]]").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("Table1[[#Headers],[Comments]]").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

''''''''''''Clear OLD TABLE and SHEET 2 (Targeted drugs)'''''''''''''''
    Sheets("Sheet2").Select
    lastrow = [A20000].End(xlUp).Row
    Columns("A").Delete
    Sheets("Sheet1").Select

    Windows("305_3TO_4D Pharmacy Management Systems, Inc. EXTRACTION.xlsx").Activate

''''''''Move Data''''''''''

    lastrow = [A20000].End(xlUp).Row
    Range("A1", Cells(lastrow, 1)).Copy
    Windows("305_3TO_4D Pharmacy Management Systems, Inc. (Mini Database).xlsx").Activate
    Range("B2").Select
    ActiveSheet.Paste

    Windows("305_3TO_4D Pharmacy Management Systems, Inc. EXTRACTION.xlsx").Activate
    lastrow = [B200000].End(xlUp).Row
    Range("B1", Cells(lastrow, 2)).Copy
    Windows("305_3TO_4D Pharmacy Management Systems, Inc. (Mini Database).xlsx").Activate
    Range("E2").Select
    ActiveSheet.Paste

    Windows("305_3TO_4D Pharmacy Management Systems, Inc. EXTRACTION.xlsx").Activate
    lastrow = [A200000].End(xlUp).Row
    Range("D1", Cells(lastrow, 7)).Copy
    Windows("305_3TO_4D Pharmacy Management Systems, Inc. (Mini Database).xlsx").Activate
    Range("G2").Select
    ActiveSheet.Paste

'Labeling

    lastrow = [B200000].End(xlUp).Row
    Range("A2", Cells(lastrow, 1)).Select
    Selection = "4D Pharmacy Management Systems, Inc."

    lastrow = [B200000].End(xlUp).Row
    Range("D2", Cells(lastrow, 4)).Select
    Selection = "3 Tier Open"

'Formulary Status

    Last = Cells(Rows.Count, "E").End(xlUp).Row
    For i = Last To 1 Step -1
        If (Cells(i, "E").Value) = "1" Then
        Cells(i, "C").Value = "Preferred"
        End If
    Next i

    Last = Cells(Rows.Count, "E").End(xlUp).Row
    For i = Last To 1 Step -1
        If (Cells(i, "E").Value) = "2" Then
        Cells(i, "C").Value = "Preferred"
        End If
    Next i

    Last = Cells(Rows.Count, "E").End(xlUp).Row
    For i = Last To 1 Step -1
        If (Cells(i, "E").Value) = "3" Then
        Cells(i, "C").Value = "Non Preferred"
        End If
    Next i

'Restrictions

    Last = Cells(Rows.Count, "A").End(xlUp).Row
    For i = Last To 1 Step -1
        If (Cells(i, "G").Value) = "QL" Then
        Cells(i, "G").Value = "Y"
        End If
    Next i

    Last = Cells(Rows.Count, "A").End(xlUp).Row
    For i = Last To 1 Step -1
        If (Cells(i, "H").Value) = "PA" Then
        Cells(i, "H").Value = "Y"
        End If
    Next i

    Last = Cells(Rows.Count, "A").End(xlUp).Row
    For i = Last To 1 Step -1
        If (Cells(i, "I").Value) = "ST" Then
        Cells(i, "I").Value = "Y"
        End If
    Next i

    Last = Cells(Rows.Count, "B").End(xlUp).Row
    For i = Last To 1 Step -1
        If (Cells(i, "E").Value) = "" Then
        Cells(i, "C").Value = "Preferred"
        Cells(i, "E").Value = "1"
        End If
        If (Cells(i, "E").Value) = "X" Then
        Cells(i, "C").Value = "Not Covered"
        Cells(i, "E").Value = "NC"
        End If
    Next i

''Multiple Forms

    Range("A1").Select
    On Error GoTo Next1
    Cells.Replace What:="Dovonex", Replacement:="Dovonex Cream/Scalp Solution", _
        LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:= _
        False, ReplaceFormat:=False

Next1:

    Range("A1").Select
    On Error GoTo Next2
    Cells.Replace What:="Imitrex", Replacement:="Imitrex Tablets/Injection/Nasal Spray", _
        LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:= _
        False, ReplaceFormat:=False

Next2:

    Range("A1").Select
    On Error GoTo Next3
    Cells.Replace What:="Zomig", Replacement:="Zomig Tablets/Nasal Spray", _
        LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:= _
        False, ReplaceFormat:=False

Next3:

    Range("A1").Select
    On Error GoTo Next4
    Last = Cells(Rows.Count, "B").End(xlUp).Row
    For i = Last To 1 Step -1
        If UCase(Cells(i, "B").Value) = "VOLTAREN" And (Cells(i, "E").Value) = "2" Then
        Cells(i, "B") = "Voltaren Ophthalmic"
        End If
    Next i


Next4:


    Windows("305_3TO_4D Pharmacy Management Systems, Inc. EXTRACTION.xlsx").Activate
    Application.CutCopyMode = False
    ActiveWorkbook.Close False
    Windows("305_3TO_4D Pharmacy Management Systems, Inc. (Mini Database).xlsx").Activate
    Range("A1").Select
    ActiveWorkbook.Save
    
End Sub

