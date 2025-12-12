Attribute VB_Name = "Module6"
' 가공된 데이터를 새폴더에 메모장을 만들고 저장하는 코드입니다
' 새폴더 위치는 contribution data processing using vba.xlsm 파일이 저장되어있는 위치입니다


Sub ExportColumnsToTextFile()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim filePath As String
    Dim folderPath As String
    Dim txtFileName As String
    Dim txtFilePath As String
    Dim i As Long
    Dim textData As String

    ' Sheet1 설정
    Set ws = ThisWorkbook.Sheets("Sheet1")
    

    ' 현재 엑셀 파일 경로 가져오기
    filePath = ThisWorkbook.Path
    If filePath = "" Then
        MsgBox "이 파일은 저장되지 않았습니다. 저장 후 다시 시도해주세요.", vbExclamation
        Exit Sub
    End If

    ' 새 폴더 경로 설정
    folderPath = filePath & "\새폴더\"
    
    ' 폴더 생성
    On Error Resume Next
    MkDir folderPath
    On Error GoTo 0
    
    ' 마지막 행 찾기 (C열 기준)
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    ' 텍스트 파일 이름 및 경로 설정
    txtFileName = "0.1.txt"
    txtFilePath = folderPath & txtFileName

    ' C, D, E 열 데이터를 텍스트로 결합
    textData = ""
    For i = 2 To lastRow ' 헤더 제외 (2행부터 시작)
        textData = textData & ws.Cells(i, "C").Value & vbTab & _
                              ws.Cells(i, "D").Value & vbTab & _
                              ws.Cells(i, "E").Value & vbCrLf
    Next i

    ' 텍스트 파일에 데이터 쓰기
    Dim txtFile As Object
    Set txtFile = CreateObject("Scripting.FileSystemObject").CreateTextFile(txtFilePath, True)
    txtFile.Write textData
    txtFile.Close

    ' 완료 메시지
    ' MsgBox "C, D, E 열 데이터가 다음 위치에 저장되었습니다: " & vbCrLf & txtFilePath, vbInformation
    
    
    
    
    
    ' 마지막 행 찾기 (G열 기준)
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    
    ' 텍스트 파일 이름 및 경로 설정
    txtFileName = "0.3.txt"
    txtFilePath = folderPath & txtFileName

    ' G, H, I 열 데이터를 텍스트로 결합
    textData = ""
    For i = 2 To lastRow ' 헤더 제외 (2행부터 시작)
        textData = textData & ws.Cells(i, "G").Value & vbTab & _
                              ws.Cells(i, "H").Value & vbTab & _
                              ws.Cells(i, "I").Value & vbCrLf
    Next i

    ' 텍스트 파일에 데이터 쓰기
    Set txtFile = CreateObject("Scripting.FileSystemObject").CreateTextFile(txtFilePath, True)
    txtFile.Write textData
    txtFile.Close

    ' 완료 메시지
    ' MsgBox "G, H, I 열 데이터가 다음 위치에 저장되었습니다: " & vbCrLf & txtFilePath, vbInformation
    
    
    
        ' 마지막 행 찾기 (K열 기준)
    lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
    
    ' 텍스트 파일 이름 및 경로 설정
    txtFileName = "0.5.txt"
    txtFilePath = folderPath & txtFileName

    ' G, H, I 열 데이터를 텍스트로 결합
    textData = ""
    For i = 2 To lastRow ' 헤더 제외 (2행부터 시작)
        textData = textData & ws.Cells(i, "K").Value & vbTab & _
                              ws.Cells(i, "L").Value & vbTab & _
                              ws.Cells(i, "M").Value & vbCrLf
    Next i

    ' 텍스트 파일에 데이터 쓰기
    Set txtFile = CreateObject("Scripting.FileSystemObject").CreateTextFile(txtFilePath, True)
    txtFile.Write textData
    txtFile.Close
    
    
    
    ' 마지막 행 찾기 (O열 기준)
    lastRow = ws.Cells(ws.Rows.Count, "O").End(xlUp).Row
    
    ' 텍스트 파일 이름 및 경로 설정
    txtFileName = "0.8.txt"
    txtFilePath = folderPath & txtFileName

    ' O, P, Q 열 데이터를 텍스트로 결합
    textData = ""
    For i = 2 To lastRow ' 헤더 제외 (2행부터 시작)
        textData = textData & ws.Cells(i, "O").Value & vbTab & _
                              ws.Cells(i, "P").Value & vbTab & _
                              ws.Cells(i, "Q").Value & vbCrLf
    Next i

    ' 텍스트 파일에 데이터 쓰기
    Set txtFile = CreateObject("Scripting.FileSystemObject").CreateTextFile(txtFilePath, True)
    txtFile.Write textData
    txtFile.Close
    
    
    
    ' 마지막 행 찾기 (S열 기준)
    lastRow = ws.Cells(ws.Rows.Count, "S").End(xlUp).Row
    
    ' 텍스트 파일 이름 및 경로 설정
    txtFileName = "1.0.txt"
    txtFilePath = folderPath & txtFileName

    ' S, T, U 열 데이터를 텍스트로 결합
    textData = ""
    For i = 2 To lastRow ' 헤더 제외 (2행부터 시작)
        textData = textData & ws.Cells(i, "S").Value & vbTab & _
                              ws.Cells(i, "T").Value & vbTab & _
                              ws.Cells(i, "U").Value & vbCrLf
    Next i

    ' 텍스트 파일에 데이터 쓰기
    Set txtFile = CreateObject("Scripting.FileSystemObject").CreateTextFile(txtFilePath, True)
    txtFile.Write textData
    txtFile.Close


    ' 완료 메시지
    MsgBox "모든 데이터가 다음 위치에 저장되었습니다: " & vbCrLf & txtFilePath, vbInformation
    
    
    
    
    
    
    
    
    
End Sub


