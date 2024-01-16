Attribute VB_Name = "Module2"
Sub formatsheets()
 lastrow = Cells.Find(What:="*", After:=[I1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
'set up the sheet
Range("I1").Value = "ticker"
Columns("I").AutoFit
Range("J1").Value = "yearly change"
Columns("j").AutoFit
Range("k1").Value = "percent change"
Columns("k").AutoFit
Range("L1").Value = "total stonk volume"
Columns("l").AutoFit
Range("Q1").Value = "ticker"
Columns("q").AutoFit
Range("r1").Value = "stonk value"
Columns("r").AutoFit
Range("n2").Value = "Greatest % increase"
Range("n3").Value = "Greatest % decrease"
Range("n4").Value = "Greatest Total Volume"
Range("R2", "R3").NumberFormat = "0.00%"
 Range("R4").Select
    Selection.NumberFormat = "#,##0"
Columns("L:L").Select
    Selection.NumberFormat = "#,##0"
Columns("K:K").NumberFormat = "0.00%"
    Columns("J:J").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(255, 0, 0)
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(74, 197, 57)
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False


End Sub

