Sub BatchConvertToUTF8NoBOM()
    Dim objShell As Object
    Dim strCommand As String
    Dim strFileExt As String
    Dim strFolderPath As String
    Dim blnRecursive As Boolean
    
    ' 配置转换参数
    strFileExt = "*.csv" ' 要转换的文件类型，如 *.txt, *.html 等
    strFolderPath = "C:\Your\Target\Folder" ' 目标文件夹路径
    blnRecursive = True ' 是否递归处理子目录 (True/False)
    
    ' 创建Shell对象
    Set objShell = CreateObject("WScript.Shell")
    
    ' 构建PowerShell命令
    If blnRecursive Then
        ' 递归处理子目录
        strCommand = "powershell -Command ""Get-ChildItem -Path '" & strFolderPath & "' -Filter '" & strFileExt & "' -Recurse | ForEach-Object { $content = Get-Content $_.FullName; Set-Content -Path $_.FullName -Value $content -Encoding UTF8NoBOM }"""
    Else
        ' 仅处理当前目录
        strCommand = "powershell -Command ""Get-ChildItem -Path '" & strFolderPath & "' -Filter '" & strFileExt & "' | ForEach-Object { $content = Get-Content $_.FullName; Set-Content -Path $_.FullName -Value $content -Encoding UTF8NoBOM }"""
    End If
    
    ' 执行命令（隐藏窗口执行）
    objShell.Run strCommand, 0, True
    
    ' 释放对象
    Set objShell = Nothing
    
    MsgBox "文件编码转换完成！", vbInformation, "提示"
End Sub
