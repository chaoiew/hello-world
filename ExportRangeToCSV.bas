Sub ExportA9ToBQ210ToCSV()
    Dim dataRange As Range
    Dim csvFilePath As String
    Dim fso As Object, txtFile As Object
    Dim rowNum As Long, colNum As Long
    Dim dataRow As String
    
    ' 设置要导出的单元格范围：A9到BQ210
    On Error Resume Next
    Set dataRange = ThisWorkbook.ActiveSheet.Range("A9:BQ210")
    On Error GoTo 0
    
    ' 检查范围是否有效
    If dataRange Is Nothing Then
        MsgBox "指定的单元格范围无效，请检查工作表是否存在", vbExclamation
        Exit Sub
    End If
    
    ' 检查范围内是否有数据
    If Application.CountA(dataRange) = 0 Then
        MsgBox "指定范围内没有数据", vbInformation
        Exit Sub
    End If
    
    ' 获取保存路径
    csvFilePath = Application.GetSaveAsFilename( _
        FileFilter:="CSV文件 (*.csv), *.csv", _
        Title:="保存CSV文件")
    
    ' 取消操作
    If csvFilePath = "False" Then Exit Sub
    
    ' 创建文件系统对象
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtFile = fso.CreateTextFile(csvFilePath, True, True) ' 使用UTF-8编码
    
    ' 遍历指定区域并写入CSV
    For rowNum = 1 To dataRange.Rows.Count
        dataRow = ""
        For colNum = 1 To dataRange.Columns.Count
            ' 处理包含逗号的内容，用双引号包裹
            If InStr(1, dataRange.Cells(rowNum, colNum).Value, ",") > 0 Then
                dataRow = dataRow & """" & Replace(dataRange.Cells(rowNum, colNum).Value, """", """""") & """"
            Else
                dataRow = dataRow & dataRange.Cells(rowNum, colNum).Value
            End If
            
            ' 添加逗号分隔（最后一列不加）
            If colNum < dataRange.Columns.Count Then
                dataRow = dataRow & ","
            End If
        Next colNum
        
        ' 写入一行数据
        txtFile.WriteLine dataRow
    Next rowNum
    
    ' 清理对象
    txtFile.Close
    Set txtFile = Nothing
    Set fso = Nothing
    
    MsgBox "数据已成功导出到：" & vbCrLf & csvFilePath, vbInformation
End Sub
