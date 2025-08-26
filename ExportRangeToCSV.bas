Sub BatchConvertToUTF8NoBOM_Compatible()
    Dim objShell As Object
    Dim strCommand As String
    Dim strFileExt As String
    Dim strFolderPath As String
    Dim blnRecursive As Boolean
    
    ' 配置转换参数
    strFileExt = "*.csv" ' 要转换的文件类型
    strFolderPath = "C:\Your\Target\Folder" ' 目标文件夹路径
    blnRecursive = True ' 是否递归处理子目录
    
    ' 创建Shell对象
    Set objShell = CreateObject("WScript.Shell")
    
    ' 构建兼容低版本PowerShell的命令
    ' 原理：先转为带BOM的UTF-8，再通过二进制方式移除前3个BOM字节
    If blnRecursive Then
        strCommand = "powershell -Command ""Get-ChildItem -Path '" & strFolderPath & "' -Filter '" & strFileExt & "' -Recurse | ForEach-Object { " & _
                    "$content = Get-Content $_.FullName; " & _
                    "Set-Content -Path $_.FullName -Value $content -Encoding UTF8; " & _
                    "$bytes = [System.IO.File]::ReadAllBytes($_.FullName); " & _
                    "[System.IO.File]::WriteAllBytes($_.FullName, $bytes[3..($bytes.Length-1)]) }"""
    Else
        strCommand = "powershell -Command ""Get-ChildItem -Path '" & strFolderPath & "' -Filter '" & strFileExt & "' | ForEach-Object { " & _
                    "$content = Get-Content $_.FullName; " & _
                    "Set-Content -Path $_.FullName -Value $content -Encoding UTF8; " & _
                    "$bytes = [System.IO.File]::ReadAllBytes($_.FullName); " & _
                    "[System.IO.File]::WriteAllBytes($_.FullName, $bytes[3..($bytes.Length-1)]) }"""
    End If
    
    ' 执行命令（隐藏窗口，等待完成）
    objShell.Run strCommand, 0, True
    
    ' 释放对象
    Set objShell = Nothing
    
    MsgBox "文件已批量转换为UTF-8无BOM格式！", vbInformation, "完成"
End Sub
