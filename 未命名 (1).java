Option Explicit

' 工作簿关闭前自动保存
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Application.ScreenUpdating = False
    
    ' 检查是否有未保存的更改
    If ThisWorkbook.Saved = False Then
        ThisWorkbook.Save
        Debug.Print "已自动保存: " & Now
    End If
    
    Application.ScreenUpdating = True
End Sub
