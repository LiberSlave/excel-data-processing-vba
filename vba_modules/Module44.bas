' Code to convert 2 cycles into 1 cycle.
' For cathode materials, the voltage profile is measured as
' decrease → increase → decrease → increase (2 cycles),
' and this code removes the first cycle (decrease → increase).
' The reason is that when 2-cycle raw data is used in kinetics analytics,
' a bug seems to occur where the plot is rendered as scattered points instead of a proper curve.
' Therefore, only 1 cycle is kept.
' Running this code is optional. You can skip it and move on to Module 5.
' (This code may cause lag.)

Sub dd()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, col As Long
    Dim startOfCycle As Long
    Dim endOfCycle As Long
    Dim foundStart As Boolean, foundEnd As Boolean
    Dim targetColumns As Variant
    Dim currentColumn As Variant
    
    ' Reference the current worksheet
    Set ws = ThisWorkbook.Sheets(1)
    
    ' Target columns to process (C, G, K, O, S)
    targetColumns = Array("C", "G", "K", "O", "S")
    
    ' Perform operation for each target column
    For Each currentColumn In targetColumns
        ' Find the last row
        lastRow = ws.Cells(ws.Rows.Count, currentColumn).End(xlUp).Row
        
        ' Initialize flags
        foundStart = False
        foundEnd = False
        
        ' Iterate through data
        For i = 2 To lastRow
            ' Find the start of voltage decrease
            If Not foundStart And ws.Cells(i, currentColumn).Value < ws.Cells(i - 1, currentColumn).Value Then
                startOfCycle = i - 1
                foundStart = True
            End If
            
            ' Find the point where voltage increases and then decreases again
            If foundStart And Not foundEnd Then
                If ws.Cells(i, currentColumn).Value > ws.Cells(i - 1, currentColumn).Value Then
                    ' During voltage increase
                    Do While ws.Cells(i, currentColumn).Value > ws.Cells(i - 1, currentColumn).Value And i < lastRow
                        i = i + 1
                    Loop
                    
                    ' Moment when voltage starts decreasing again
                    If ws.Cells(i, currentColumn).Value < ws.Cells(i - 1, currentColumn).Value Then
                        endOfCycle = i ' Store the row index where decrease resumes
                        foundEnd = True
                        Exit For
                    End If
                End If
            End If
        Next i
        
        ' Check and delete the first cycle data
        If foundStart And foundEnd Then
            For i = endOfCycle To 2 Step -1
                ws.Cells(i, currentColumn).Delete Shift:=xlUp
                ws.Cells(i, currentColumn).Offset(0, 1).Delete Shift:=xlUp ' Also delete the adjacent right column
            Next i
        Else
            MsgBox currentColumn & " column - unable to find a complete 1 cycle."
        End If
    Next currentColumn

End Sub
