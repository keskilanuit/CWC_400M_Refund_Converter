Attribute VB_Name = "Module2"
Sub Retrieve_Data_from_400M()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    

    'Prompt user to select a file
    Set wb = Application.Workbooks.Open(Application.GetOpenFilename())
    
    'Assuming data is on the first sheet of the selected workbook
    Set ws = wb.Sheets(1)
    
    'Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "U").End(xlUp).row
    
    'Loop through all rows with data
    For i = lastRow To 2 Step -1 'Starting from the last row and working backwards to avoid issues with row deletion
        'Check if the cell in column AT has no value
               If Len(Trim(ws.Range("AT" & i).Value)) = 0 Then
            'Delete the entire row if it has no value in column AT
            ws.Rows(i).Delete
        End If
    Next i
        
            For i = lastRow To 2 Step -1
            If InStr(1, ws.Range("A" & i).Value, "Day", vbTextCompare) > 0 Then
            'Delete the entire row if it contains the text value "Day" in column A
            ws.Rows(i).Delete
        End If
    Next i
    
    Cells(1, 1).EntireRow.Delete
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "AT").End(xlUp).row
    
    With ws.Range("AT1:AV" & lastRow)
        .NumberFormat = "0"
        .Value = .Value
    End With
    
    
        'Close the selected workbook and save changes
    wb.Close SaveChanges:=True

    MsgBox "Finished Retrieving."
    
End Sub
