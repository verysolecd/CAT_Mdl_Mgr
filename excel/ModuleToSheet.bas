Attribute VB_Name = "ModuleToSheet"
Option Explicit


Sub ExecToExcel()
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "请选择包含CATVBA模块的文件夹"
        .AllowMultiSelect = False
        If Not .Show = -1 Then
            MsgBox "操作已取消"
            Exit Sub
        End If

        folderPath = .SelectedItems(1)
    End With
    
    Dim arrFiles As Variant
    arrFiles = GetFilesWithExtensions( _
        folderPath, _
        Array(".bas", ".cls", ".frm") _
    )
    If UBound(arrFiles) < 0 Then Exit Sub

    ' 创建新工作簿
    Dim wb As Workbook
    Set wb = Workbooks.Add
    
    Call CreateTextFilesToNewWorkbook( _
        arrFiles, _
        wb _
    )

    arrFiles = GetFilesWithExtensions( _
        folderPath, _
        Array(".frx") _
    )
    If Not UBound(arrFiles) < 0 Then
        Call ExportFRXFilesToSheet( _
            arrFiles, _
            wb _
        )
    End If

    Dim SavePath As String
    SavePath = Application.GetSaveAsFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx")
    If SavePath <> "False" Then
        wb.SaveAs fileName:=SavePath, FileFormat:=xlOpenXMLWorkbook
    End If

    ' 关闭工作簿
    wb.Close SaveChanges:=False

End Sub


Private Function GetFilesWithExtensions( _
        ByVal folderPath As String, _
        ByVal extensions As Variant) As Variant

    Dim fileSystem As Object
    Dim folder As Object
    Dim file As Object
    Dim filePaths As Collection
    Dim ext As Variant
    Dim filePathArray() As String
    Dim i As Integer
    
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set folder = fileSystem.GetFolder(folderPath)
    Set filePaths = New Collection
    
    ' 遍历文件夹中的文件，查找具有指定扩展名的文件
    For Each file In folder.Files
        For Each ext In extensions
            If LCase(fileSystem.GetExtensionName(file.path)) = LCase(Replace(ext, ".", "")) Then
                filePaths.Add file.path
                Exit For
            End If
        Next ext
    Next file
    
    ' 将集合转换为数组
    ReDim filePathArray(1 To filePaths.Count)
    For i = 1 To filePaths.Count
        filePathArray(i) = filePaths(i)
    Next i
    
    GetFilesWithExtensions = filePathArray
End Function


Sub ExportFRXFilesToSheet(frxFilePaths As Variant, wb As Workbook)
    Dim ws As Worksheet
    Dim frxFilePath As Variant
    Dim frxFileName As String
    Dim frxContent As String
    
    ' 对每个frx文件创建新工作表
    For Each frxFilePath In frxFilePaths
        ' 创建新工作表
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        
        ' 获取文件名并设置到工作表
        frxFileName = Mid(frxFilePath, InStrRev(frxFilePath, "\") + 1)
        
        ' 将文件内容转换为十六进制并输出到工作表
        frxContent = ConvertFRXToHex(frxFilePath)
        ws.Cells(1, 1).Value = EncryptLine(frxFileName)
        ws.Cells(2, 1).Value = frxContent
    Next frxFilePath
    
    Set ws = Nothing
End Sub


Function ConvertFRXToHex(filePath As Variant) As String
    Dim fileNum As Integer
    Dim byteArray() As Byte
    Dim hexString As String
    Dim i As Integer
    
    fileNum = FreeFile
    Open filePath For Binary As #fileNum
    ReDim byteArray(LOF(fileNum) - 1)
    Get #fileNum, , byteArray
    Close #fileNum
    
    hexString = ""
    For i = LBound(byteArray) To UBound(byteArray)
        hexString = hexString & Right("0" & Hex(byteArray(i)), 2)
    Next i
    
    ConvertFRXToHex = hexString
End Function


Private Sub CreateTextFilesToNewWorkbook( _
        ByVal filePaths As Variant, _
        ByVal wb As Workbook)

    Dim ws As Worksheet
    Dim filePath As String
    Dim fileName As String
    Dim FileLine As String
    Dim fileNum As Integer
    Dim i As Integer
    Dim rowNum As Integer
    Dim EncryptedLine As String
    Dim EncryptedFileName As String
    
    ' 遍历选定的文件
    For i = LBound(filePaths) To UBound(filePaths)
        filePath = filePaths(i)
        
        ' 获取文件名并加密
        fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
        EncryptedFileName = EncryptLine(fileName)
        
        ' 获取默认工作表
        Set ws = wb.Sheets(wb.Sheets.Count)
        
        ' 在第1行第1列写入加密后的文件名
        ws.Cells(1, 1).Value = EncryptedFileName
        
        ' 读取文本文件内容，从第2行开始写入
        fileNum = FreeFile
        Open filePath For Input As #fileNum
        rowNum = 2
        Do While Not EOF(fileNum)
            Line Input #fileNum, FileLine
            EncryptedLine = EncryptLine(FileLine)
            ws.Cells(rowNum, 1).Value = EncryptedLine
            rowNum = rowNum + 1
        Loop
        Close #fileNum
        
        ' 添加新工作表
        If i < UBound(filePaths) Then
            wb.Sheets.Add After:=wb.Sheets(wb.Sheets.Count)
        End If
    Next i
    
    Set ws = Nothing
    Set wb = Nothing

End Sub


Private Function EncryptLine( _
        ByVal Line As String) As String

    Dim EncryptedLine As String
    Dim i As Integer
    Dim CharCode As Integer
    
    EncryptedLine = ""
    For i = 1 To Len(Line)
        CharCode = Asc(Mid(Line, i, 1))
        EncryptedLine = EncryptedLine & Chr(CharCode + 1)
    Next i
    
    EncryptLine = EncryptedLine

End Function

'*******************
Sub ExecToModule()
    Dim filePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "请选择Excel文件"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx"
        .AllowMultiSelect = False
        If Not .Show = -1 Then
            MsgBox "操作已取消"
        End If
        
        filePath = .SelectedItems(1)
    End With
    
    Dim dirPath As String
    dirPath = GetFolderPath(filePath)
    
    Call ExportSheetsToTextFilesWithDecryption( _
        filePath, _
        dirPath _
    )

End Sub


Private Function GetFolderPath( _
        ByVal filePath As String) As String

    Dim fileSystem As Object
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    
    ' 获取文件夹路径
    GetFolderPath = fileSystem.GetParentFolderName(filePath)
    
    Set fileSystem = Nothing
End Function


Private Sub ExportSheetsToTextFilesWithDecryption( _
        ByVal xlsxFilePath As String, _
        ByVal outputFolderPath As String)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim SavePath As String
    Dim DecryptedLine As String
    Dim fileNum As Integer
    Dim rowNum As Integer
    Dim EncryptedFileName As String
    Dim DecryptedFileName As String
    Dim fileExt As String
    Dim i As Long
    
    ' 打开指定的Excel文件
    Set wb = Workbooks.Open(xlsxFilePath)
    
    ' 处理每个工作表
    For Each ws In wb.Sheets
        ' 获取第1行包含的加密文件名
        EncryptedFileName = ws.Cells(1, 1).Value
        DecryptedFileName = DecryptLine(EncryptedFileName)
        
        ' 根据解密后的文件名确定文件扩展名
        fileExt = LCase(Right(DecryptedFileName, 3))
        
        ' 设置保存路径
        SavePath = outputFolderPath & "\" & DecryptedFileName
        
        If fileExt = "frx" Then
            ' 以二进制文件形式保存
            Call ExportSheetToFRXFiles(ws, outputFolderPath)
        Else
            ' 以文本文件形式保存
            fileNum = FreeFile
            Open SavePath For Output As #fileNum
            
            ' 从第2行开始读取内容并解密后写入文件
            rowNum = 2
            Do
                DecryptedLine = DecryptLine(ws.Cells(rowNum, 1).Value)
                Print #fileNum, DecryptedLine
                rowNum = rowNum + 1
            Loop Until ws.Cells(rowNum, 1).Value = "" And ws.Cells(rowNum + 1, 1).Value = ""
            
            Close #fileNum
        End If
    Next ws
    
    ' 关闭工作簿
    wb.Close SaveChanges:=False
    
    Set ws = Nothing
    Set wb = Nothing

End Sub


Sub ExportSheetToFRXFiles( _
        ByVal ws As Worksheet, _
        ByVal outputFolderPath As String)

    Dim fileSystem As Object
    Dim filePath As String
    Dim fileName As String
    Dim fileContent As String
    Dim fileNum As Integer
    Dim rowNum As Integer
    Dim byteArray() As Byte
    Dim hexString As String
    Dim i As Integer
    
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    
    rowNum = 1
    
    ' 处理每个FRX文件
    Do While ws.Cells(rowNum, 1).Value <> ""
        ' 获取文件名
        fileName = DecryptLine(ws.Cells(rowNum, 1).Value)
        filePath = outputFolderPath & "\" & fileName
        
        ' 将十六进制字符串转换为字节数组
        hexString = ws.Cells(rowNum + 1, 1).Value
        ReDim byteArray(Len(hexString) \ 2 - 1)
        For i = 0 To UBound(byteArray)
            byteArray(i) = CByte("&H" & Mid(hexString, 2 * i + 1, 2))
        Next i
        
        ' 将字节数组写入二进制文件
        fileNum = FreeFile
        Open filePath For Binary As #fileNum
        Put #fileNum, , byteArray
        Close #fileNum
        
        rowNum = rowNum + 2
    Loop
    
    Set fileSystem = Nothing

End Sub


Private Function DecryptLine( _
        ByVal Line As String) As String

    Dim DecryptedLine As String
    Dim i As Integer
    Dim CharCode As Integer
    
    DecryptedLine = ""
    For i = 1 To Len(Line)
        CharCode = Asc(Mid(Line, i, 1))
        DecryptedLine = DecryptedLine & Chr(CharCode - 1)
    Next i
    
    DecryptLine = DecryptedLine

End Function