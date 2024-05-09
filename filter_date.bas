Attribute VB_Name = "Module3"
Sub filter_date_range()
    Dim lastRow As Long
    Dim i As Long
    Dim ws As Worksheet
    Dim chooseColumn As String
    Dim dateColumn As Range
    Dim dateStartRange As String
    Dim dateEndRange As String
    Dim loopVariable As Variant
    
    
    Set ws = ActiveSheet
    
    chooseColumn = InputBox("Enter Column with date (e.g. F)")
    dateStartRange = InputBox("Enter Start date in format mm/dd/yyyy")
    dateEndRange = InputBox("Enter End date in format mm/dd/yyyy")
    
    Set dateColumn = ws.Columns(chooseColumn)
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
   
    For i = lastRow To 1 Step -1
        loopVariable = ws.Cells(i, dateColumn.Column).Value
        If Not loopVariable >= CDate(dateStartRange) And loopVariable <= CDate(dateEndRange) Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub
