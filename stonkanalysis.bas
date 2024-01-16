Attribute VB_Name = "Module4"
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







