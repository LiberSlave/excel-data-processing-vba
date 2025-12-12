Attribute VB_Name = "Module5"
' 밑에 step time 안잘린 부분 자르는 코드입니다

Sub DeleteRowsWithOnlyColumnEData()
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
    

    ' Sheet1 설정
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' 마지막 행 찾기 (E열 기준)
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row

    ' 역순으로 행 삭제 (1행은 제외하고 2행부터 시작)
    For i = lastRow To 2 Step -1
        cValue = ws.Cells(i, "C").Value ' C열 값
        dValue = ws.Cells(i, "D").Value ' D열 값
        eValue = ws.Cells(i, "E").Value ' E열 값

        ' C, D가 비어 있고 E에만 값이 있는 경우 행 삭제
        If IsEmpty(cValue) And IsEmpty(dValue) And Not IsEmpty(eValue) Then
            ws.Cells(i, 5).Delete Shift:=xlUp ' E열 (5번째 열)
        End If
    Next i
    
    
    ' 마지막 행 찾기 (I열 기준)
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row

    ' 역순으로 행 삭제 (1행은 제외하고 2행부터 시작)
    For i = lastRow To 2 Step -1
        gValue = ws.Cells(i, "G").Value ' G열 값
        hValue = ws.Cells(i, "H").Value ' H열 값
        iValue = ws.Cells(i, "I").Value ' I열 값

        ' G, H가 비어 있고 I에만 값이 있는 경우 행 삭제
        If IsEmpty(gValue) And IsEmpty(hValue) And Not IsEmpty(iValue) Then
            ws.Cells(i, 9).Delete Shift:=xlUp ' i열 9번째 열)
        End If
    Next i
    
    
    
    ' 마지막 행 찾기 (M열 기준)
    lastRow = ws.Cells(ws.Rows.Count, "M").End(xlUp).Row

    ' 역순으로 행 삭제 (1행은 제외하고 2행부터 시작)
    For i = lastRow To 2 Step -1
        kValue = ws.Cells(i, "K").Value ' K열 값
        lValue = ws.Cells(i, "L").Value ' L열 값
        mValue = ws.Cells(i, "M").Value ' M열 값

        ' K, L가 비어 있고 M에만 값이 있는 경우 행 삭제
        If IsEmpty(kValue) And IsEmpty(lValue) And Not IsEmpty(mValue) Then
            ws.Cells(i, 13).Delete Shift:=xlUp ' M열 (13번째 열)
        End If
    Next i
    
    
    ' 마지막 행 찾기 (Q열 기준)
    lastRow = ws.Cells(ws.Rows.Count, "Q").End(xlUp).Row

    ' 역순으로 행 삭제 (1행은 제외하고 2행부터 시작)
    For i = lastRow To 2 Step -1
        oValue = ws.Cells(i, "O").Value ' O열 값
        pValue = ws.Cells(i, "P").Value ' P열 값
        qValue = ws.Cells(i, "Q").Value ' Q열 값

        ' O, P가 비어 있고 Q에만 값이 있는 경우 행 삭제
        If IsEmpty(oValue) And IsEmpty(pValue) And Not IsEmpty(qValue) Then
            ws.Cells(i, 17).Delete Shift:=xlUp ' Q열 (17번째 열)
        End If
    Next i
    
    
    
    ' 마지막 행 찾기 (U열 기준)
    lastRow = ws.Cells(ws.Rows.Count, "U").End(xlUp).Row

    ' 역순으로 행 삭제 (1행은 제외하고 2행부터 시작)
    For i = lastRow To 2 Step -1
        sValue = ws.Cells(i, "S").Value ' S열 값
        tValue = ws.Cells(i, "T").Value ' T열 값
        uValue = ws.Cells(i, "U").Value ' U열 값

        ' S, T가 비어 있고 U에만 값이 있는 경우 행 삭제
        If IsEmpty(sValue) And IsEmpty(tValue) And Not IsEmpty(uValue) Then
            ws.Cells(i, 21).Delete Shift:=xlUp ' U열 (21번째 열)
        End If
    Next i

    ' 완료 메시지
    ' MsgBox "Rows with data only in column E have been deleted.", vbInformation
End Sub

