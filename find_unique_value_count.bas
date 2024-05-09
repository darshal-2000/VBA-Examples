Attribute VB_Name = "Module2"
Sub find_unique_values_count()
    Dim ws As Worksheet
    Dim chooseColumn As String
    Dim uniqueValues As Object
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    
    Set ws = ActiveSheet
    
    chooseColumn = InputBox("Enter Column for finding unique values (e.g. F)")
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Create a dictionary object to store unique values
    Set uniqueValues = CreateObject("Scripting.Dictionary")
    
    ' Loop through each cell in column A and add unique values to the dictionary
    For i = 1 To lastRow
        If Not uniqueValues.Exists(ws.Cells(i, chooseColumn).Value) Then
            uniqueValues.Add ws.Cells(i, chooseColumn).Value, ws.Cells(i, chooseColumn).Value
        End If
    Next i
    

    ws.Cells(1, "N").Value = "Count"
    ws.Cells(2, "N").Value = uniqueValues.Count
End Sub
