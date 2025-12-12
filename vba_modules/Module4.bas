Attribute VB_Name = "Module4"
'2cycle을 1cycle로 변환시키는 코드입니다 voltage구간이 증가->감소->증가->감소로 2cycle이 측정되는데 이중 앞의 첫사이클(증가->감소)를 지웁니다.
'이유는 2cycle을 raw data로 kinetics analytics에 넣으면 플랏이 잘안되는(점으로 플랏됨)버그가 생기는거 같아서 1cycle만 넣기 위함입니다.


Sub FindCycleAndRemoveMultipleColumns()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, col As Long
    Dim startOfCycle As Long
    Dim endOfCycle As Long
    Dim foundStart As Boolean, foundEnd As Boolean
    Dim targetColumns As Variant
    Dim currentColumn As Variant
    
    ' 현재 워크시트 참조
    Set ws = ThisWorkbook.Sheets(1)
    
    ' 처리할 대상 열 (C, G, K, O, S)
    targetColumns = Array("C", "G", "K", "O", "S")
    
    ' 각 열에 대해 작업 수행
    For Each currentColumn In targetColumns
        ' 마지막 행 찾기
        lastRow = ws.Cells(ws.Rows.Count, currentColumn).End(xlUp).Row
        
        ' 초기화
        foundStart = False
        foundEnd = False
        
        ' 데이터 순회
        For i = 2 To lastRow
            ' 증가 시작점 찾기
            If Not foundStart And ws.Cells(i, currentColumn).Value > ws.Cells(i - 1, currentColumn).Value Then
                startOfCycle = i - 1
                foundStart = True
            End If
            
            ' 감소 후 다시 증가하는 점 찾기
            If foundStart And Not foundEnd Then
                If ws.Cells(i, currentColumn).Value < ws.Cells(i - 1, currentColumn).Value Then
                    ' 감소 중
                    Do While ws.Cells(i, currentColumn).Value < ws.Cells(i - 1, currentColumn).Value And i < lastRow
                        i = i + 1
                    Loop
                    
                    ' 다시 증가하는 순간
                    If ws.Cells(i, currentColumn).Value > ws.Cells(i - 1, currentColumn).Value Then
                        endOfCycle = i ' 다시 증가하는 셀 번호 저장
                        foundEnd = True
                        Exit For
                    End If
                End If
            End If
        Next i
        
        ' 1cycle 데이터 확인 및 삭제
        If foundStart And foundEnd Then
            MsgBox currentColumn & "열 - 다시 증가하는 순간의 셀 번호는: " & endOfCycle
            
            For i = endOfCycle To 2 Step -1
                ws.Cells(i, currentColumn).Delete Shift:=xlUp
                ws.Cells(i, currentColumn).Offset(0, 1).Delete Shift:=xlUp ' 해당 열의 오른쪽 열도 삭제
            Next i
        Else
            MsgBox currentColumn & "열 - 1cycle을 찾을 수 없습니다."
        End If
    Next currentColumn
End Sub

