Attribute VB_Name = "Módulo1"
Sub exe2cells()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim TxtRng  As Range

    Set wb = ActiveWorkbook
    Set ws = wb.Sheets("prog")
    
    Dim byteArr() As Byte
    Dim fileInt As Integer: fileInt = FreeFile
    Open "main.exe" For Binary Access Read As #fileInt
    ReDim byteArr(0 To LOF(fileInt) - 1)
    Get #fileInt, , byteArr
    
    Dim index As Double
    
    Application.ScreenUpdating = False
    For index = LBound(byteArr) To UBound(byteArr)
        ws.Cells(index + 1, 1).Value = byteArr(index)
        'Debug.Print (byteArr(index))
    Next
    
    ws.Cells(1, 2).Value = UBound(byteArr) + 1
    
    For index = UBound(byteArr) + 1 To UBound(byteArr) + 100000
        ws.Cells(index + 1, 1).ClearContents
        'Debug.Print (byteArr(index))
    Next
    Application.ScreenUpdating = True
    
    
    Close #fileInt
End Sub
Sub cells2exe()
Attribute cells2exe.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim TxtRng  As Range

    Set wb = ActiveWorkbook
    Set ws = wb.Sheets("prog")
    
    Dim byteArr() As Byte
    
    Dim totalBytes As Double
    
    totalBytes = ws.Cells(1, 2).Value - 1
    
    ReDim byteArr(totalBytes)
    
    Dim index As Double
    For index = 0 To totalBytes
        byteArr(index) = CByte((ws.Cells(index + 1, 1).Value))
    Next
    
    ok_ = FileWriteBinary(byteArr, "main2.exe")
    
    Dim res As String
    
    res = ShellRun("main2.exe")
    Debug.Print (res)
    
End Sub
Function FileWriteBinary(vData() As Byte, sFileName As String, Optional bAppendToFile As Boolean = False) As Boolean
    Dim iFileNum As Integer, lWritePos As Long
    
    On Error GoTo ErrFailed
    If bAppendToFile = False Then
        If Len(Dir$(sFileName)) > 0 And Len(sFileName) > 0 Then
            'Delete the existing file
            VBA.Kill sFileName
        End If
    End If
    
    iFileNum = FreeFile
    Open sFileName For Binary Access Write As #iFileNum
    
    If bAppendToFile = False Then
        'Write to first byte
        lWritePos = 1
    Else
        'Write to last byte + 1
        lWritePos = LOF(iFileNum) + 1
    End If
    
    Put #iFileNum, lWritePos, vData
    Close iFileNum
    
    FileWriteBinary = True
    Exit Function

ErrFailed:
    FileWriteBinary = False
    Close iFileNum
    Debug.Print Err.Description
End Function

Public Function ShellRun(sCmd As String) As String

    'Run a shell command, returning the output as a string

    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")
    
    'run command
    Dim oExec As Object
    Dim oOutput As Object
    Set oExec = oShell.Exec(sCmd)
    Set oOutput = oExec.StdOut

    'handle the results as they are written to and read from the StdOut object
    Dim s As String
    Dim sLine As String
    While Not oOutput.AtEndOfStream
        sLine = oOutput.ReadLine
        If sLine <> "" Then s = s & sLine & vbCrLf
    Wend

    ShellRun = s

End Function
