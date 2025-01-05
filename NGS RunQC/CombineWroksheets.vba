Sub CombineWroksheets()

    Dim ws As Worksheet
    Dim wsMaster As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Create a new worksheet for the combined data
    Set wsMaster = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsMaster.Name = "CombinedData"
    
    'Initialize the master worksheet
    wsMaster.Cells(1, 1).Value = "SheetName"
    wsMaster.Cells(1, 2).Value = "Data"
    
    'Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> wsMaster.Name Then
            lastRow = wsMaster.Cells(wsMaster.Rows.Count, "A").End(xlUp).Row + 1
            ws.UsedRange.Copy Destination:=wsMaster.Cells(lastRow, 1)
            'Insert the sheet name
            For i = lastRow To wsMaster.Cells(wsMaster.Rows.Count, "A").End(xlUp).Row
                wsMaster.Cells(i, 1).Value = ws.Name
            Next i
        End If
    Next ws
  
    
End Sub
