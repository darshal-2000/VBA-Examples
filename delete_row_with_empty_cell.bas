Attribute VB_Name = "Module1"
Sub delete_row_with_empty_cell()
    Dim lastRow As Long
    Dim i As Long
    Dim checkRange As Range
    Dim ws As Worksheet
    Dim startColumn As String
    Dim endColumn As String
    
    Set ws = ActiveSheet
    
    startColumn = InputBox("Enter the start column (e.g., A):")
    endColumn = InputBox("Enter the end column (e.g., M):")
    
    Set checkRange = ws.Range(startColumn & ":" & endColumn)
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
   
    For i = lastRow To 1 Step -1
        If WorksheetFunction.CountBlank(Intersect(checkRange, ws.Rows(i))) > 0 Then
            Rows(i).Delete
        End If
    Next i
End Sub
