' Code to swap the positions of voltage and current step time

Sub a()
    Dim ws As Worksheet
    Dim lastRowC As Long
    Dim lastRowB As Long
    Dim maxLastRow As Long

    ' Use the currently active worksheet
    Set ws = ActiveSheet

    ' Find the last row in columns C and B
    lastRowC = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    lastRowB = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Select the longer last row between columns C and B
    maxLastRow = Application.WorksheetFunction.Max(lastRowC, lastRowB)

    ' Copy column C to column E
    ws.Range("C1:C" & lastRowC).Copy Destination:=ws.Range("E1")

    ' Copy column B to column F
    ws.Range("B1:B" & lastRowB).Copy Destination:=ws.Range("F1")

    ' Clear the original data in columns C and B (optional)
    ' Uncomment if you want to clear the original columns
    ws.Range("C1:C" & lastRowC).ClearContents
    ws.Range("B1:B" & lastRowB).ClearContents

    ws.Columns("A").Delete

End Sub
