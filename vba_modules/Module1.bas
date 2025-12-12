Attribute VB_Name = "Module1"
' voltage와 current step time의 위치를 바꾸는 코드


Sub MoveColumns()
    Dim ws As Worksheet
    Dim lastRowC As Long
    Dim lastRowB As Long
    Dim maxLastRow As Long

    ' 현재 활성화된 워크시트를 사용
    Set ws = ActiveSheet

    ' C열과 B열의 마지막 행 찾기
    lastRowC = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    lastRowB = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    

    ' C열과 B열 중 마지막 행이 더 긴 쪽 선택
    maxLastRow = Application.WorksheetFunction.Max(lastRowC, lastRowB)

    ' C열을 E열로 복사
    ws.Range("C1:C" & lastRowC).Copy Destination:=ws.Range("E1")

    ' B열을 F열로 복사
    ws.Range("B1:B" & lastRowB).Copy Destination:=ws.Range("F1")

    ' C열과 B열의 원본 데이터 삭제 (선택 사항)
    ' Uncomment if you want to clear the original columns
    ws.Range("C1:C" & lastRowC).ClearContents
    ws.Range("B1:B" & lastRowB).ClearContents
    ws.Columns("A").Delete
   
    ' MsgBox "Columns moved successfully.", vbInformation
End Sub
