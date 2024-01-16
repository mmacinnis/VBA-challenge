Attribute VB_Name = "Module1"
' Code borrowed from:
' The credit card exercise (wk2 segment 6)
' Multisheet: extendoffice.com
' Better safe than sorry
'



'setting up to run in multiple worksheets
Sub multisheet()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
  '      Sheets("Sheet1").Rows(2).Columns.AutoFit
        Call formatsheets
        Call stonkanalysis
    Next
    Application.ScreenUpdating = True
End Sub

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
Range("r1").Value = "value"
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

Sub stonkanalysis()
Dim ticker As String
Dim tickertotal As Double
tickertotal = 0
Dim volumetablerow As Integer
volumetablerow = 2
Dim stonkopen As Double
Dim stonkclose As Double



 lastrow = Cells.Find(What:="*", After:=[A1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row



stonkopen = Range("C2").Value


 ' Loop through all stonks
  For i = 2 To lastrow

    ' Check if we are still within the same stonk, if not do this
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
         ' Set the ticker symbol
      ticker = Cells(i, 1).Value

      ' Add to the total trading volume
      tickertotal = tickertotal + Cells(i, 7).Value

      ' Print the ticker symbol in the summary table
      Range("I" & volumetablerow).Value = ticker

      ' Print the trading volume in the table
      Range("L" & volumetablerow).Value = tickertotal

'get the year close for the stock
    stonkclose = Cells(i, 6)
    
    'calculate the change over the year and print it
     Range("J" & volumetablerow).Value = (stonkclose - stonkopen)
        
    'calculate percent change and print
    Range("K" & volumetablerow).Value = (Range("J" & volumetablerow).Value / stonkopen)
    
    'get open value for new stonk
    stonkopen = Cells(i + 1, 3).Value

      ' Add one to the summary table row
      volumetablerow = volumetablerow + 1
      
      ' Reset the volume total
      tickertotal = 0

      
    ' If the cell immediately following a row is the same ticker do this
    Else

      ' Add to the volume total
      tickertotal = tickertotal + Cells(i, 7).Value

    End If

  Next i
  
    bigvol = 0
    bigneg = 0
    bigpos = 0
    
  'another loop to look for the top values
  
    For j = 2 To Range("I2").End(xlDown).Row + 1

     If Cells(j, 12).Value > bigvol Then
     bigvol = Cells(j, 12).Value
     bigvoltick = Cells(j, 9).Value
    
        End If
    
    
    'biggest positive change
    If Cells(j, 11).Value > 0 And Cells(j, 11).Value > bigpos Then
    bigpos = Cells(j, 11).Value
    bigpostick = Cells(j, 9).Value
    
    'find the biggest negative change
    ElseIf Cells(j, 11).Value < 0 And Cells(j, 11).Value < bigneg Then
         bigneg = Cells(j, 11).Value
        bignegtick = Cells(j, 9).Value
    End If
               
    Next j
    'found the top value, print it
    Range("Q4").Value = bigvoltick
    Range("R4").Value = bigvol
    'print the percent changes
    Range("Q2").Value = bigpostick
    Range("R2").Value = bigpos
    Range("Q3").Value = bignegtick
    Range("R3").Value = bigneg
    
    End Sub

' ===NOTES===
' multisheet works - now to conditional formatting
' 1. make cells look better (add comma separators, widen column)
' 2. fix colors on condform column
'
'
'





