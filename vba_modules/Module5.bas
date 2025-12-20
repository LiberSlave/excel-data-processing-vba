' Code to trim the remaining uncut step time at the bottom
' If this code runs too slowly due to lag, do it manually instead

Sub e()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cValue As Variant
    Dim dValue As Variant
    Dim eValue As Variant
    Dim gValue As Variant
    Dim hValue As Variant
    Dim iValue As Variant
    Dim kValue As Variant
    Dim lValue As Variant
    Dim mValue As Variant
    Dim oValue As Variant
    Dim pValue As Variant
    Dim qValue As Variant
    Dim sValue As Variant
    Dim tValue As Variant
    Dim uValue As Variant
    

    ' Set Sheet1
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Find the last row based on column E
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row

    ' Delete rows in reverse order (exclude row 1, start from row 2)
    For i = lastRow To 2 Step -1
        cValue = ws.Cells(i, "C").Value ' Value in column C
        dValue = ws.Cells(i, "D").Value ' Value in column D
        eValue = ws.Cells(i, "E").Value ' Value in column E

        ' Delete row if C and D are empty and only E has a value
        If IsEmpty(cValue) And IsEmpty(dValue) And Not IsEmpty(eValue) Then
            ws.Cells(i, 5).Delete Shift:=xlUp ' Column E (5th column)
        End If
    Next i
    
    
    ' Find the last row based on column I
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row

    ' Delete rows in reverse order (exclude row 1, start from row 2)
    For i = lastRow To 2 Step -1
        gValue = ws.Cells(i, "G").Value ' Value in column G
        hValue = ws.Cells(i, "H").Value ' Value in column H
        iValue = ws.Cells(i, "I").Value ' Value in column I

        ' Delete row if G and H are empty and only I has a value
        If IsEmpty(gValue) And IsEmpty(hValue) And Not IsEmpty(iValue) Then
            ws.Cells(i, 9).Delete Shift:=xlUp ' Column I (9th column)
        End If
    Next i
    
    
    
    ' Find the last row based on column M
    lastRow = ws.Cells(ws.Rows.Count, "M").End(xlUp).Row

    ' Delete rows in reverse order (exclude row 1, start from row 2)
    For i = lastRow To 2 Step -1
        kValue = ws.Cells(i, "K").Value ' Value in column K
        lValue = ws.Cells(i, "L").Value ' Value in column L
        mValue = ws.Cells(i, "M").Value ' Value in column M

        ' Delete row if K and L are empty and only M has a value
        If IsEmpty(kValue) And IsEmpty(lValue) And Not IsEmpty(mValue) Then
            ws.Cells(i, 13).Delete Shift:=xlUp ' Column M (13th column)
        End If
    Next i
    
    
    ' Find the last row based on column Q
    lastRow = ws.Cells(ws.Rows.Count, "Q").End(xlUp).Row

    ' Delete rows in reverse order (exclude row 1, start from row 2)
    For i = lastRow To 2 Step -1
        oValue = ws.Cells(i, "O").Value ' Value in column O
        pValue = ws.Cells(i, "P").Value ' Value in column P
        qValue = ws.Cells(i, "Q").Value ' Value in column Q

        ' Delete row if O and P are empty and only Q has a value
        If IsEmpty(oValue) And IsEmpty(pValue) And Not IsEmpty(qValue) Then
            ws.Cells(i, 17).Delete Shift:=xlUp ' Column Q (17th column)
        End If
    Next i
    
    
    
    ' Find the last row based on column U
    lastRow = ws.Cells(ws.Rows.Count, "U").End(xlUp).Row

    ' Delete rows in reverse order (exclude row 1, start from row 2)
    For i = lastRow To 2 Step -1
        sValue = ws.Cells(i, "S").Value ' Value in column S
        tValue = ws.Cells(i, "T").Value ' Value in column T
        uValue = ws.Cells(i, "U").Value ' Value in column U

        ' Delete row if S and T are empty and only U has a value
        If IsEmpty(sValue) And IsEmpty(tValue) And Not IsEmpty(uValue) Then
            ws.Cells(i, 21).Delete Shift:=xlUp ' Column U (21st column)
        End If
    Next i

    ' Completion message
    MsgBox "Step time has been deleted. Go to Module 6.", vbInformation
End Sub
