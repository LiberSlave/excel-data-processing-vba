' Code to create text files in a new folder and save the processed data
' The new folder is created at the same location where
' "contribution data processing using vba.xlsm" is saved

Sub f()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim filePath As String
    Dim folderPath As String
    Dim txtFileName As String
    Dim txtFilePath As String
    Dim i As Long
    Dim textData As String

    ' Set Sheet1
    Set ws = ThisWorkbook.Sheets("Sheet1")
    

    ' Get the current Excel file path
    filePath = ThisWorkbook.Path
    If filePath = "" Then
        MsgBox "This file has not been saved. Please save the file and try again.", vbExclamation
        Exit Sub
    End If

    ' Set new folder path
    folderPath = filePath & "\newfolder\"
    
    ' Create the folder
    On Error Resume Next
    MkDir folderPath
    On Error GoTo 0
    
    ' Find the last row based on column C
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    ' Set text file name and path
    txtFileName = "0.1.txt"
    txtFilePath = folderPath & txtFileName

    ' Combine columns C, D, and E into text
    textData = ""
    For i = 2 To lastRow ' Exclude header (start from row 2)
        textData = textData & ws.Cells(i, "C").Value & vbTab & _
                              ws.Cells(i, "D").Value & vbTab & _
                              ws.Cells(i, "E").Value & vbCrLf
    Next i

    ' Write data to text file
    Dim txtFile As Object
    Set txtFile = CreateObject("Scripting.FileSystemObject").CreateTextFile(txtFilePath, True)
    txtFile.Write textData
    txtFile.Close

    ' Completion message (optional)
    ' MsgBox "Data from columns C, D, E has been saved at: " & vbCrLf & txtFilePath, vbInformation
    
    
    
    ' Find the last row based on column G
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    
    ' Set text file name and path
    txtFileName = "0.3.txt"
    txtFilePath = folderPath & txtFileName

    ' Combine columns G, H, and I into text
    textData = ""
    For i = 2 To lastRow ' Exclude header (start from row 2)
        textData = textData & ws.Cells(i, "G").Value & vbTab & _
                              ws.Cells(i, "H").Value & vbTab & _
                              ws.Cells(i, "I").Value & vbCrLf
    Next i

    ' Write data to text file
    Set txtFile = CreateObject("Scripting.FileSystemObject").CreateTextFile(txtFilePath, True)
    txtFile.Write textData
    txtFile.Close

    ' Completion message (optional)
    ' MsgBox "Data from columns G, H, I has been saved at: " & vbCrLf & txtFilePath, vbInformation
    
    
    
    ' Find the last row based on column K
    lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
    
    ' Set text file name and path
    txtFileName = "0.5.txt"
    txtFilePath = folderPath & txtFileName

    ' Combine columns K, L, and M into text
    textData = ""
    For i = 2 To lastRow ' Exclude header (start from row 2)
        textData = textData & ws.Cells(i, "K").Value & vbTab & _
                              ws.Cells(i, "L").Value & vbTab & _
                              ws.Cells(i, "M").Value & vbCrLf
    Next i

    ' Write data to text file
    Set txtFile = CreateObject("Scripting.FileSystemObject").CreateTextFile(txtFilePath, True)
    txtFile.Write textData
    txtFile.Close
    
    
    ' Find the last row based on column O
    lastRow = ws.Cells(ws.Rows.Count, "O").End(xlUp).Row
    
    ' Set text file name and path
    txtFileName = "0.8.txt"
    txtFilePath = folderPath & txtFileName

    ' Combine columns O, P, and Q into text
    textData = ""
    For i = 2 To lastRow ' Exclude header (start from row 2)
        textData = textData & ws.Cells(i, "O").Value & vbTab & _
                              ws.Cells(i, "P").Value & vbTab & _
                              ws.Cells(i, "Q").Value & vbCrLf
    Next i

    ' Write data to text file
    Set txtFile = CreateObject("Scripting.FileSystemObject").CreateTextFile(txtFilePath, True)
    txtFile.Write textData
    txtFile.Close
    
    
    ' Find the last row based on column S
    lastRow = ws.Cells(ws.Rows.Count, "S").End(xlUp).Row
    
    ' Set text file name and path
    txtFileName = "1.0.txt"
    txtFilePath = folderPath & txtFileName

    ' Combine columns S, T, and U into text
    textData = ""
    For i = 2 To lastRow ' Exclude header (start from row 2)
        textData = textData & ws.Cells(i, "S").Value & vbTab & _
                              ws.Cells(i, "T").Value & vbTab & _
                              ws.Cells(i, "U").Value & vbCrLf
    Next i

    ' Write data to text file
    Set txtFile = CreateObject("Scripting.FileSystemObject").CreateTextFile(txtFilePath, True)
    txtFile.Write textData
    txtFile.Close


    ' Completion message
    MsgBox "All data has been saved at:" & vbCrLf & txtFilePath, vbInformation

End Sub
