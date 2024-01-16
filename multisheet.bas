Attribute VB_Name = "Module3"
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
