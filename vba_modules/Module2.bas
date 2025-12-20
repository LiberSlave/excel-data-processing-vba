' Code to remove rows where current equals zero

Sub b()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim dValue As Double
    Dim firstNonZeroRow As Long
    Dim lastZeroRow As Long
    Dim i As Long
    Dim zeroRow As Long
    Dim rangeToCut As Range
    Dim a As Long
    
    ' Set Sheet1
    Set ws = ThisWorkbook.Sheets("Sheet1")
    a = 0
    
    ' Remove rows where current equals zero in column D
    If a = 0 Then
        ' Find the last row in column D
        lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
        
        ' Initialize
        firstNonZeroRow = 0 ' Initialize the row number of the first non-zero value
        
        ' Iterate through column D from row 2 to the last row
        For currentRow = 2 To lastRow
            dValue = ws.Cells(currentRow, "D").Value
            
            ' When a non-zero value is found, store the row number
            If dValue <> 0 Then
                firstNonZeroRow = currentRow
                Exit For ' Exit loop after finding the first non-zero value
            End If
        Next currentRow
        
        lastZeroRow = firstNonZeroRow - 1
        ws.Rows("2:" & lastZeroRow).Delete
        a = 1
        
        ' Result check message (optional)
        
'        If firstNonZeroRow > 0 Then
'            MsgBox "The first non-zero number in column D is in row: " & firstNonZeroRow, vbInformation
'        Else
'            MsgBox "No non-zero number found in column D.", vbExclamation
'        End If
        
        ' Delete rows where current equals zero
    End If
    
    ' Cut data from the first zero-current row to the last row in column D
    ' and paste it into columns G, H, and I
    If a = 1 Then
        ' Find the last row in column D
        lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
        
        ' Find the first zero value in column D starting from row 2
        zeroRow = 0 ' Initialize
        
        For currentRow = 2 To lastRow
            dValue = ws.Cells(currentRow, "D").Value
            
            If dValue = 0 Then
                zeroRow = currentRow ' Store the row number of the first zero value
                Exit For ' Exit loop after finding the first zero
            End If
        Next currentRow
        
        ' Select columns C, D, and E from zeroRow to the last row
        Set rangeToCut = ws.Range("C" & zeroRow & ":E" & lastRow)
        
        ' Copy the selected range and paste it into columns G, H, and I
        rangeToCut.Copy
        ws.Range("G" & 2).PasteSpecial Paste:=xlPasteValues ' Paste starting at row 2 in column G
        
        ' Clear the selected range (cut operation)
        rangeToCut.ClearContents
        
        a = 2
    End If
        
    ' Remove rows where current equals zero in column H
    If a = 2 Then
        ' Find the last row in column H
        lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
        
        ' Initialize
        firstNonZeroRow = 0 ' Initialize the row number of the first non-zero value
        
        ' Iterate through column H from row 2 to the last row
        For currentRow = 2 To lastRow
            dValue = ws.Cells(currentRow, "H").Value
            
            ' When a non-zero value is found, store the row number
            If dValue <> 0 Then
                firstNonZeroRow = currentRow
                Exit For ' Exit loop after finding the first non-zero value
            End If
        Next currentRow
        
        lastZeroRow = firstNonZeroRow - 1
        
        For i = lastZeroRow To 2 Step -1
            ws.Cells(i, 7).Delete Shift:=xlUp  ' Column G
            ws.Cells(i, 8).Delete Shift:=xlUp  ' Column H
            ws.Cells(i, 9).Delete Shift:=xlUp  ' Column I
        Next i
        
        a = 3
    End If
    
    ' Cut data from the first zero-current row in column H to the last row
    ' and paste it into columns K, L, and M
    If a = 3 Then
        ' Find the last row in column H
        lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
        
        ' Find the first zero value in column H starting from row 2
        zeroRow = 0 ' Initialize
        
        For currentRow = 2 To lastRow
            dValue = ws.Cells(currentRow, "H").Value
            
            If dValue = 0 Then
                zeroRow = currentRow ' Store the row number of the first zero value
                Exit For ' Exit loop after finding the first zero
            End If
        Next currentRow
        
        ' Select columns G, H, and I from zeroRow to the last row
        Set rangeToCut = ws.Range("G" & zeroRow & ":I" & lastRow)
        
        ' Copy the selected range and paste it into columns K, L, and M
        rangeToCut.Copy
        ws.Range("K" & 2).PasteSpecial Paste:=xlPasteValues ' Paste starting at row 2 in column K
        
        ' Clear the selected range (cut operation)
        rangeToCut.ClearContents
        
        a = 4
    End If
    
    ' Remove rows where current equals zero in column L
    If a = 4 Then
        ' Find the last row in column L
        lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
        
        ' Initialize
        firstNonZeroRow = 0 ' Initialize the row number of the first non-zero value
        
        ' Iterate through column L from row 2 to the last row
        For currentRow = 2 To lastRow
            dValue = ws.Cells(currentRow, "L").Value
            
            ' When a non-zero value is found, store the row number
            If dValue <> 0 Then
                firstNonZeroRow = currentRow
                Exit For ' Exit loop after finding the first non-zero value
            End If
        Next currentRow
        
        lastZeroRow = firstNonZeroRow - 1
        
        For i = lastZeroRow To 2 Step -1
            ws.Cells(i, 11).Delete Shift:=xlUp  ' Column K
            ws.Cells(i, 12).Delete Shift:=xlUp  ' Column L
            ws.Cells(i, 13).Delete Shift:=xlUp  ' Column M
        Next i
        
        a = 5
    End If
    
    ' Cut data from the first zero-current row in column L to the last row
    ' and paste it into columns O, P, and Q
    If a = 5 Then
        ' Find the last row in column L
        lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
        
        ' Find the first zero value in column L starting from row 2
        zeroRow = 0 ' Initialize
        
        For currentRow = 2 To lastRow
            dValue = ws.Cells(currentRow, "L").Value
            
            If dValue = 0 Then
                zeroRow = currentRow ' Store the row number of the first zero value
                Exit For ' Exit loop after finding the first zero
            End If
        Next currentRow
        
        ' Select columns K, L, and M from zeroRow to the last row
        Set rangeToCut = ws.Range("K" & zeroRow & ":M" & lastRow)
        
        ' Copy the selected range and paste it into columns O, P, and Q
        rangeToCut.Copy
        ws.Range("O" & 2).PasteSpecial Paste:=xlPasteValues ' Paste starting at row 2 in column O
        
        ' Clear the selected range (cut operation)
        rangeToCut.ClearContents
        
        a = 6
    End If
    
    ' Remove rows where current equals zero in column P
    If a = 6 Then
        ' Find the last row in column P
        lastRow = ws.Cells(ws.Rows.Count, "P").End(xlUp).Row
        
        ' Initialize
        firstNonZeroRow = 0 ' Initialize the row number of the first non-zero value
        
        ' Iterate through column P from row 2 to the last row
        For currentRow = 2 To lastRow
            dValue = ws.Cells(currentRow, "P").Value
            
            ' When a non-zero value is found, store the row number
            If dValue <> 0 Then
                firstNonZeroRow = currentRow
                Exit For ' Exit loop after finding the first non-zero value
            End If
        Next currentRow
        
        lastZeroRow = firstNonZeroRow - 1
        
        For i = lastZeroRow To 2 Step -1
            ws.Cells(i, 15).Delete Shift:=xlUp  ' Column O
            ws.Cells(i, 16).Delete Shift:=xlUp  ' Column P
            ws.Cells(i, 17).Delete Shift:=xlUp  ' Column Q
        Next i
        
        a = 7
    End If
    
    ' Cut data from the first zero-current row in column P to the last row
    ' and paste it into columns S, T, and U
    If a = 7 Then
        ' Find the last row in column P
        lastRow = ws.Cells(ws.Rows.Count, "P").End(xlUp).Row
        
        ' Find the first zero value in column P starting from row 2
        zeroRow = 0 ' Initialize
        
        For currentRow = 2 To lastRow
            dValue = ws.Cells(currentRow, "P").Value
            
            If dValue = 0 Then
                zeroRow = currentRow ' Store the row number of the first zero value
                Exit For ' Exit loop after finding the first zero
            End If
        Next currentRow
        
        ' Select columns O, P, and Q from zeroRow to the last row
        Set rangeToCut = ws.Range("O" & zeroRow & ":Q" & lastRow)
        
        ' Copy the selected range and paste it into columns S, T, and U
        rangeToCut.Copy
        ws.Range("S" & 2).PasteSpecial Paste:=xlPasteValues ' Paste starting at row 2 in column S
        
        ' Clear the selected range (cut operation)
        rangeToCut.ClearContents
        
        a = 8
    End If
    
    ' Remove rows where current equals zero in column T
    If a = 8 Then
        ' Find the last row in column T
        lastRow = ws.Cells(ws.Rows.Count, "T").End(xlUp).Row
        
        ' Initialize
        firstNonZeroRow = 0 ' Initialize the row number of the first non-zero value
        
        ' Iterate through column T from row 2 to the last row
        For currentRow = 2 To lastRow
            dValue = ws.Cells(currentRow, "T").Value
            
            ' When a non-zero value is found, store the row number
            If dValue <> 0 Then
                firstNonZeroRow = currentRow
                Exit For ' Exit loop after finding the first non-zero value
            End If
        Next currentRow
        
        lastZeroRow = firstNonZeroRow - 1
        
        For i = lastZeroRow To 2 Step -1
            ws.Cells(i, 19).Delete Shift:=xlUp  ' Column S
            ws.Cells(i, 20).Delete Shift:=xlUp  ' Column T
            ws.Cells(i, 21).Delete Shift:=xlUp  ' Column U
        Next i
        
        a = 9
    End If

End Sub
