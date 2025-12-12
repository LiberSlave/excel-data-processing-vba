Attribute VB_Name = "Module2"
' current가 0인 데이터 지우는 코드

Sub FindFirstNonZeroInD()
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
    
    ' Sheet1 설정
    Set ws = ThisWorkbook.Sheets("Sheet1")
    a = 0
    
    ' D열에서 current 가 0인것 지우기
    If a = 0 Then
        ' D열의 마지막 행 찾기
        lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
        
        ' 초기값 설정
        firstNonZeroRow = 0 ' 첫 번째 0이 아닌 숫자 행 번호 초기화
        
        ' D열을 2행부터 마지막 행까지 순회
        For currentRow = 2 To lastRow
            dValue = ws.Cells(currentRow, "D").Value
            
            ' 0이 아닌 숫자가 나오면 첫 번째 행 번호 반환
            If dValue <> 0 Then
                firstNonZeroRow = currentRow
                Exit For ' 첫 번째 0이 아닌 숫자가 나왔으면 루프 종료
            End If
        Next currentRow
        lastZeroRow = firstNonZeroRow - 1
        ws.Rows("2:" & lastZeroRow).Delete
        a = 1
        
        ' 결과 확인 메시지
        
'        If firstNonZeroRow > 0 Then
'            MsgBox "The first non-zero number in column D is in row: " & firstNonZeroRow, vbInformation
'        Else
'            MsgBox "No non-zero number found in column D.", vbExclamation
'        End If
'
        
        
        ' 커런트가 0인 행 삭제
    
        
    End If
    
    ' D열에서 current가 0이나오는 부분부터 마지막행까지 잘라내서 G,H,I 열에 붙여넣기
    If a = 1 Then
        ' D열의 마지막 행 찾기
        lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
        
        ' 2행부터 D열 확인하여 첫 번째 0을 찾기
        zeroRow = 0 ' 초기화
        
        For currentRow = 2 To lastRow
            dValue = ws.Cells(currentRow, "D").Value
            
            If dValue = 0 Then
                zeroRow = currentRow ' 첫 번째 0을 찾으면 행 번호를 저장
                Exit For ' 첫 번째 0을 찾으면 루프 종료
            End If
        Next currentRow
        ' C, D, E열에서 zeroRow부터 마지막 행까지 선택
        Set rangeToCut = ws.Range("C" & zeroRow & ":E" & lastRow)
        
        ' 선택한 범위를 잘라서 G, H, I열에 붙여넣기
        rangeToCut.Copy
        ws.Range("G" & 2).PasteSpecial Paste:=xlPasteValues ' G열 2행에 붙여넣기


        ' 선택한 범위 삭제 (잘라내기)
        rangeToCut.ClearContents
        
        a = 2
    End If
        
        
    ' MsgBox "The value of i is: " & i
    
    ' H열에서 current 가 0인것 지우기
    If a = 2 Then
        ' H열의 마지막 행 찾기
        lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
        
        ' 초기값 설정
        firstNonZeroRow = 0 ' 첫 번째 0이 아닌 숫자 행 번호 초기화
        
        ' H열을 2행부터 마지막 행까지 순회
        For currentRow = 2 To lastRow
            dValue = ws.Cells(currentRow, "H").Value
            
            ' 0이 아닌 숫자가 나오면 첫 번째 행 번호 반환
            If dValue <> 0 Then
                firstNonZeroRow = currentRow
                Exit For ' 첫 번째 0이 아닌 숫자가 나왔으면 루프 종료
            End If
        Next currentRow
        lastZeroRow = firstNonZeroRow - 1
        
        For i = lastZeroRow To 2 Step -1
            ws.Cells(i, 7).Delete Shift:=xlUp  ' G열 (7번째 열)
            ws.Cells(i, 8).Delete Shift:=xlUp  ' H열 (8번째 열)
            ws.Cells(i, 9).Delete Shift:=xlUp  ' I열 (9번째 열)
        Next i
        
        a = 3
    End If
    
    ' H열에서 current가 0이나오는 부분부터 마지막행까지 잘라내서 K,L,M 열에 붙여넣기
    If a = 3 Then
        ' H열의 마지막 행 찾기
        lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
        
        ' 2행부터 D열 확인하여 첫 번째 0을 찾기
        zeroRow = 0 ' 초기화
        
        For currentRow = 2 To lastRow
            dValue = ws.Cells(currentRow, "H").Value
            
            If dValue = 0 Then
                zeroRow = currentRow ' 첫 번째 0을 찾으면 행 번호를 저장
                Exit For ' 첫 번째 0을 찾으면 루프 종료
            End If
        Next currentRow
        ' G, H, I열에서 zeroRow부터 마지막 행까지 선택
        Set rangeToCut = ws.Range("G" & zeroRow & ":I" & lastRow)
        
        ' 선택한 범위를 잘라서 K, L, M열에 붙여넣기
        rangeToCut.Copy
        ws.Range("K" & 2).PasteSpecial Paste:=xlPasteValues ' K열 2행에 붙여넣기


        ' 선택한 범위 삭제 (잘라내기)
        rangeToCut.ClearContents
        
        a = 4
    End If
    
    ' L열에서 current 가 0인것 지우기
    If a = 4 Then
        ' L열의 마지막 행 찾기
        lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
        
        ' 초기값 설정
        firstNonZeroRow = 0 ' 첫 번째 0이 아닌 숫자 행 번호 초기화
        
        ' L열을 2행부터 마지막 행까지 순회
        For currentRow = 2 To lastRow
            dValue = ws.Cells(currentRow, "L").Value
            
            ' 0이 아닌 숫자가 나오면 첫 번째 행 번호 반환
            If dValue <> 0 Then
                firstNonZeroRow = currentRow
                Exit For ' 첫 번째 0이 아닌 숫자가 나왔으면 루프 종료
            End If
        Next currentRow
        lastZeroRow = firstNonZeroRow - 1
        
        For i = lastZeroRow To 2 Step -1
            ws.Cells(i, 11).Delete Shift:=xlUp  ' K열 (11번째 열)
            ws.Cells(i, 12).Delete Shift:=xlUp  ' L열 (12번째 열)
            ws.Cells(i, 13).Delete Shift:=xlUp  ' M열 (13번째 열)
        Next i
        
        a = 5
    End If
    
    ' L열에서 current가 0이나오는 부분부터 마지막행까지 잘라내서 O,P,Q 열에 붙여넣기
    If a = 5 Then
        ' L열의 마지막 행 찾기
        lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
        
        ' 2행부터 L열 확인하여 첫 번째 0을 찾기
        zeroRow = 0 ' 초기화
        
        For currentRow = 2 To lastRow
            dValue = ws.Cells(currentRow, "L").Value
            
            If dValue = 0 Then
                zeroRow = currentRow ' 첫 번째 0을 찾으면 행 번호를 저장
                Exit For ' 첫 번째 0을 찾으면 루프 종료
            End If
        Next currentRow
        ' K, L, M열에서 zeroRow부터 마지막 행까지 선택
        Set rangeToCut = ws.Range("K" & zeroRow & ":M" & lastRow)
        
        ' 선택한 범위를 잘라서 O, P, Q열에 붙여넣기
        rangeToCut.Copy
        ws.Range("O" & 2).PasteSpecial Paste:=xlPasteValues ' O열 2행에 붙여넣기


        ' 선택한 범위 삭제 (잘라내기)
        rangeToCut.ClearContents
        
        a = 6
    End If
    
    ' P열에서 current 가 0인것 지우기
    If a = 6 Then
        ' P열의 마지막 행 찾기
        lastRow = ws.Cells(ws.Rows.Count, "P").End(xlUp).Row
        
        ' 초기값 설정
        firstNonZeroRow = 0 ' 첫 번째 0이 아닌 숫자 행 번호 초기화
        
        ' P열을 2행부터 마지막 행까지 순회
        For currentRow = 2 To lastRow
            dValue = ws.Cells(currentRow, "P").Value
            
            ' 0이 아닌 숫자가 나오면 첫 번째 행 번호 반환
            If dValue <> 0 Then
                firstNonZeroRow = currentRow
                Exit For ' 첫 번째 0이 아닌 숫자가 나왔으면 루프 종료
            End If
        Next currentRow
        lastZeroRow = firstNonZeroRow - 1
        
        For i = lastZeroRow To 2 Step -1
            ws.Cells(i, 15).Delete Shift:=xlUp  ' O열 (15번째 열)
            ws.Cells(i, 16).Delete Shift:=xlUp  ' P열 (16번째 열)
            ws.Cells(i, 17).Delete Shift:=xlUp  ' Q열 (17번째 열)
        Next i
        
        a = 7
    End If
    
    ' P열에서 current가 0이나오는 부분부터 마지막행까지 잘라내서 S,T,U 열에 붙여넣기
    If a = 7 Then
        ' P열의 마지막 행 찾기
        lastRow = ws.Cells(ws.Rows.Count, "P").End(xlUp).Row
        
        ' 2행부터 P열 확인하여 첫 번째 0을 찾기
        zeroRow = 0 ' 초기화
        
        For currentRow = 2 To lastRow
            dValue = ws.Cells(currentRow, "P").Value
            
            If dValue = 0 Then
                zeroRow = currentRow ' 첫 번째 0을 찾으면 행 번호를 저장
                Exit For ' 첫 번째 0을 찾으면 루프 종료
            End If
        Next currentRow
        ' O, P, Q열에서 zeroRow부터 마지막 행까지 선택
        Set rangeToCut = ws.Range("O" & zeroRow & ":Q" & lastRow)
        
        ' 선택한 범위를 잘라서 S, T, U열에 붙여넣기
        rangeToCut.Copy
        ws.Range("S" & 2).PasteSpecial Paste:=xlPasteValues ' S열 2행에 붙여넣기


        ' 선택한 범위 삭제 (잘라내기)
        rangeToCut.ClearContents
        
        a = 8
    End If
    
    ' T열에서 current 가 0인것 지우기
    If a = 8 Then
        ' T열의 마지막 행 찾기
        lastRow = ws.Cells(ws.Rows.Count, "T").End(xlUp).Row
        
        ' 초기값 설정
        firstNonZeroRow = 0 ' 첫 번째 0이 아닌 숫자 행 번호 초기화
        
        ' T열을 2행부터 마지막 행까지 순회
        For currentRow = 2 To lastRow
            dValue = ws.Cells(currentRow, "T").Value
            
            ' 0이 아닌 숫자가 나오면 첫 번째 행 번호 반환
            If dValue <> 0 Then
                firstNonZeroRow = currentRow
                Exit For ' 첫 번째 0이 아닌 숫자가 나왔으면 루프 종료
            End If
        Next currentRow
        lastZeroRow = firstNonZeroRow - 1
        
        For i = lastZeroRow To 2 Step -1
            ws.Cells(i, 19).Delete Shift:=xlUp  ' S열 (19번째 열)
            ws.Cells(i, 20).Delete Shift:=xlUp  ' T열 (20번째 열)
            ws.Cells(i, 21).Delete Shift:=xlUp  ' U열 (21번째 열)
        Next i
        
        a = 9
    End If

        
        
    
    
     
     
    
End Sub


