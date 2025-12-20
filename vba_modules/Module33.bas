' Code to delete the voltage-increasing region starting from OCV
' That is, delete data up to just before the voltage starts decreasing again

Sub cc()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim previousValue As Double
    Dim currentValue As Double
    Dim increasingRow As Long
    Dim minRow As Long
    
    ' Set Sheet1
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Find the last row with data in column C
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    ' Initialize values
    previousValue = ws.Cells(2, "C").Value ' Initialize with the value in row 2 of column C
    increasingRow = 0 ' Initialize (returns 0 if not found)
    
    ' Traverse column C data (from row 3 to the last row)
    For currentRow = 3 To lastRow
        currentValue = ws.Cells(currentRow, "C").Value
        
        ' If the value starts decreasing, store the row number and exit the loop
        If currentValue < previousValue Then
            increasingRow = currentRow
            minRow = increasingRow - 1
            Exit For
        End If
        
        ' Update previous value
        previousValue = currentValue
    Next currentRow
    
    For i = minRow To 2 Step -1
        ws.Cells(i, 3).Delete Shift:=xlUp ' Column C
        ws.Cells(i, 4).Delete Shift:=xlUp ' Column D
    Next i
    
    ' Result message (optional)
'    If increasingRow > 0 Then
'        MsgBox "The first decreasing value in column C is at row: " & increasingRow, vbInformation
'    Else
'        MsgBox "No decreasing value found in column C.", vbExclamation
'    End If

    ' Find the last row with data in column G
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    
    ' Initialize values
    previousValue = ws.Cells(2, "G").Value ' Initialize with the value in row 2 of column G
    increasingRow = 0 ' Initialize (returns 0 if not found)
    
    ' Traverse column G data (from row 3 to the last row)
    For currentRow = 3 To lastRow
        currentValue = ws.Cells(currentRow, "G").Value
        
        ' If the value starts decreasing, store the row number and exit the loop
        If currentValue < previousValue Then
            increasingRow = currentRow
            minRow = increasingRow - 1
            Exit For
        End If
        
        ' Update previous value
        previousValue = currentValue
    Next currentRow
    
    For i = minRow To 2 Step -1
        ws.Cells(i, 7).Delete Shift:=xlUp ' Column G
        ws.Cells(i, 8).Delete Shift:=xlUp ' Column H
    Next i
    
    ' Find the last row with data in column K
    lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
    
    ' Initialize values
    previousValue = ws.Cells(2, "K").Value ' Initialize with the value in row 2 of column K
    increasingRow = 0 ' Initialize (returns 0 if not found)
    
    ' Traverse column K data (from row 3 to the last row)
    For currentRow = 3 To lastRow
        currentValue = ws.Cells(currentRow, "K").Value
        
        ' If the value starts decreasing, store the row number and exit the loop
        If currentValue < previousValue Then
            increasingRow = currentRow
            minRow = increasingRow - 1
            Exit For
        End If
        
        ' Update previous value
        previousValue = currentValue
    Next currentRow
    
    For i = minRow To 2 Step -1
        ws.Cells(i, 11).Delete Shift:=xlUp ' Column K
        ws.Cells(i, 12).Delete Shift:=xlUp ' Column L
    Next i
    
    ' Find the last row with data in column O
    lastRow = ws.Cells(ws.Rows.Count, "O").End(xlUp).Row
    
    ' Initialize values
    previousValue = ws.Cells(2, "O").Value ' Initialize with the value in row 2 of column O
    increasingRow = 0 ' Initialize (returns 0 if not found)
    
    ' Traverse column O data (from row 3 to the last row)
    For currentRow = 3 To lastRow
        currentValue = ws.Cells(currentRow, "O").Value
        
        ' If the value starts decreasing, store the row number and exit the loop
        If currentValue < previousValue Then
            increasingRow = currentRow
            minRow = increasingRow - 1
            Exit For
        End If
        
        ' Update previous value
        previousValue = currentValue
    Next currentRow
    
    For i = minRow To 2 Step -1
        ws.Cells(i, 15).Delete Shift:=xlUp ' Column O
        ws.Cells(i, 16).Delete Shift:=xlUp ' Column P
    Next i
    
    ' Find the last row with data in column S
    lastRow = ws.Cells(ws.Rows.Count, "S").End(xlUp).Row
    
    ' Initialize values
    previousValue = ws.Cells(2, "S").Value ' Initialize with the value in row 2 of column S
    increasingRow = 0 ' Initialize (returns 0 if not found)
    
    ' Traverse column S data (from row 3 to the last row)
    For currentRow = 3 To lastRow
        currentValue = ws.Cells(currentRow, "S").Value
        
        ' If the value starts decreasing, store the row number and exit the loop
        If currentValue < previousValue Then
            increasingRow = currentRow
            minRow = increasingRow - 1
            Exit For
        End If
        
        ' Update previous value
        previousValue = currentValue
    Next currentRow
    
    For i = minRow To 2 Step -1
        ws.Cells(i, 19).Delete Shift:=xlUp ' Column S
        ws.Cells(i, 20).Delete Shift:=xlUp ' Column T
    Next i

End Sub
