Attribute VB_Name = "Module3"
' ocv부터 0voltage 까지 voltage가 내려가는 구간 삭제하는 코드
' 즉, voltage가 줄어들다가 다시 늘어나는 직전부분까지 삭제



Sub FindFirstIncreasingRowInColumn()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim previousValue As Double
    Dim currentValue As Double
    Dim increasingRow As Long
    Dim minRow As Long
    
    ' Sheet1 설정
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' C열의 마지막 데이터가 있는 행 찾기
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    ' 초기값 설정
    previousValue = ws.Cells(2, "C").Value ' C열 2행의 값으로 초기화
    increasingRow = 0 ' 초기값 (찾지 못하면 0 반환)
    
    ' C열 데이터 순회 (3행부터 마지막 행까지)
    For currentRow = 3 To lastRow
        currentValue = ws.Cells(currentRow, "C").Value
        
        ' 값이 증가하면 해당 행 번호 저장 후 루프 종료
        If currentValue > previousValue Then
            increasingRow = currentRow
            minRow = increasingRow - 1
            Exit For
        End If
        
        ' 현재 값을 이전 값으로 업데이트
        previousValue = currentValue
    Next currentRow
    

    For i = minRow To 2 Step -1
        ws.Cells(i, 3).Delete Shift:=xlUp ' C열 (3번째 열)
        ws.Cells(i, 4).Delete Shift:=xlUp ' D열 (4번째 열)
    Next i
    
        ' 결과 메시지 출력
'    If increasingRow > 0 Then
'        MsgBox "The first increasing value in column C is at row: " & increasingRow, vbInformation
'    Else
'        MsgBox "No increasing value found in column C.", vbExclamation
'    End If

    ' G열의 마지막 데이터가 있는 행 찾기
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    
    ' 초기값 설정
    previousValue = ws.Cells(2, "G").Value ' C열 2행의 값으로 초기화
    increasingRow = 0 ' 초기값 (찾지 못하면 0 반환)
    
    ' G열 데이터 순회 (3행부터 마지막 행까지)
    For currentRow = 3 To lastRow
        currentValue = ws.Cells(currentRow, "G").Value
        
        ' 값이 증가하면 해당 행 번호 저장 후 루프 종료
        If currentValue > previousValue Then
            increasingRow = currentRow
            minRow = increasingRow - 1
            Exit For
        End If
        
        ' 현재 값을 이전 값으로 업데이트
        previousValue = currentValue
    Next currentRow
    

    For i = minRow To 2 Step -1
        ws.Cells(i, 7).Delete Shift:=xlUp ' G열 (7번째 열)
        ws.Cells(i, 8).Delete Shift:=xlUp ' H열 (8번째 열)
    Next i
    
    ' K열의 마지막 데이터가 있는 행 찾기
    lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
    
    ' 초기값 설정
    previousValue = ws.Cells(2, "K").Value ' K열 2행의 값으로 초기화
    increasingRow = 0 ' 초기값 (찾지 못하면 0 반환)
    
    ' K열 데이터 순회 (3행부터 마지막 행까지)
    For currentRow = 3 To lastRow
        currentValue = ws.Cells(currentRow, "K").Value
        
        ' 값이 증가하면 해당 행 번호 저장 후 루프 종료
        If currentValue > previousValue Then
            increasingRow = currentRow
            minRow = increasingRow - 1
            Exit For
        End If
        
        ' 현재 값을 이전 값으로 업데이트
        previousValue = currentValue
    Next currentRow
    

    For i = minRow To 2 Step -1
        ws.Cells(i, 11).Delete Shift:=xlUp ' K열 (11번째 열)
        ws.Cells(i, 12).Delete Shift:=xlUp ' L열 (12번째 열)
    Next i
    
    
    ' O열의 마지막 데이터가 있는 행 찾기
    lastRow = ws.Cells(ws.Rows.Count, "O").End(xlUp).Row
    
    ' 초기값 설정
    previousValue = ws.Cells(2, "O").Value ' O열 2행의 값으로 초기화
    increasingRow = 0 ' 초기값 (찾지 못하면 0 반환)
    
    ' O열 데이터 순회 (3행부터 마지막 행까지)
    For currentRow = 3 To lastRow
        currentValue = ws.Cells(currentRow, "O").Value
        
        ' 값이 증가하면 해당 행 번호 저장 후 루프 종료
        If currentValue > previousValue Then
            increasingRow = currentRow
            minRow = increasingRow - 1
            Exit For
        End If
        
        ' 현재 값을 이전 값으로 업데이트
        previousValue = currentValue
    Next currentRow
    

    For i = minRow To 2 Step -1
        ws.Cells(i, 15).Delete Shift:=xlUp ' O열 (15번째 열)
        ws.Cells(i, 16).Delete Shift:=xlUp ' P열 (16번째 열)
    Next i
    
    ' S열의 마지막 데이터가 있는 행 찾기
    lastRow = ws.Cells(ws.Rows.Count, "S").End(xlUp).Row
    
    ' 초기값 설정
    previousValue = ws.Cells(2, "S").Value ' S열 2행의 값으로 초기화
    increasingRow = 0 ' 초기값 (찾지 못하면 0 반환)
    
    ' S열 데이터 순회 (3행부터 마지막 행까지)
    For currentRow = 3 To lastRow
        currentValue = ws.Cells(currentRow, "S").Value
        
        ' 값이 증가하면 해당 행 번호 저장 후 루프 종료
        If currentValue > previousValue Then
            increasingRow = currentRow
            minRow = increasingRow - 1
            Exit For
        End If
        
        ' 현재 값을 이전 값으로 업데이트
        previousValue = currentValue
    Next currentRow
    

    For i = minRow To 2 Step -1
        ws.Cells(i, 19).Delete Shift:=xlUp ' O열 (19번째 열)
        ws.Cells(i, 20).Delete Shift:=xlUp ' P열 (20번째 열)
    Next i


 
    
End Sub

