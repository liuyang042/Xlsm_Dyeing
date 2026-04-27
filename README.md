# TOPT Excel VBA .xlsm
## 检查有效期
场景一：程序执行前检查（宏入口处）
```vba
Sub MyProtectedMacro()
    If Not CheckExpiration() Then Exit Sub
    ' 此处放置你需要保护的业务代码
End Sub
```

场景二：点击工作表时检查（工作表事件）:代码（放在对应工作表的模块中，例如 Sheet1）
```vba
' Worksheet Activate Event with Protection
' Place this in the DATA sheet module
Private Sub Worksheet_Activate()
    ' Execute expiration verification first
    If Not CheckExpiration() Then
        ' If expired, deactivate this sheet
        Application.EnableEvents = False
        If ThisWorkbook.Sheets.Count > 1 Then
            ThisWorkbook.Sheets(1).Activate
        End If
        Application.EnableEvents = True
        Exit Sub
    End If
    ' Continue with normal activation if license valid
    Range("A2").Select
End Sub
```
场景三：打开工作簿时检查（工作簿事件）'(放在 ThisWorkbook 模块中)
在 ThisWorkbook 模块中使用 Workbook_Open 事件，打开文件时立即验证。如果过期，可以给出警告并强制关闭或限制功能。
```vba
Private Sub Workbook_Open()
    ' 打开工作簿时进行有效期检查
    If Not CheckExpiration() Then
        ' 过期提示后立即关闭工作簿（可根据需要调整）
        ThisWorkbook.Close SaveChanges:=False
    End If
End Sub
```
## 检查有效期代码
```vba
Public Function CheckExpiration() As Boolean
    Dim githubLink As String
    Dim http As Object
    Dim expireDate As Date
    Dim today As Date
    Dim httpStatus As Long                    ' 用来存放状态码

    githubLink = "https://raw.giteeusercontent.com/liuyang042/Xlsm_ALL/raw/main/expire.txt"

    ' 创建 WinHttp 对象
    On Error Resume Next
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    On Error GoTo 0
    If http Is Nothing Then
        MsgBox "Unable to create network component. Please contact author: Liu Yang Email: luckilyliuyang@163.com", vbCritical
        CheckExpiration = False
        Exit Function
    End If

    ' ===== 所有涉及网络的操作统一加错误保护，包括读取状态码 =====
    On Error Resume Next
    http.SetTimeouts 15000, 15000, 15000, 15000
    http.Open "GET", githubLink, False
    http.Send
    httpStatus = http.Status                 ' 即使 Send 失败，读取状态码的错误也会被忽略
    On Error GoTo 0
    ' 检查状态：如果请求或网络出错，状态码不会是 200
    If httpStatus <> 200 Then
        MsgBox "Failed to verify version. Please contact author: Liu Yang Email: luckilyliuyang@163.com", vbCritical
        CheckExpiration = False
        Exit Function
    End If

    ' 转换日期（Trim 足以处理常规空格）
    On Error Resume Next
    expireDate = DateValue(Trim$(http.ResponseText))
    On Error GoTo 0

    If expireDate = 0 Then
        MsgBox "Invalid expiration date format. Please contact author.", vbCritical
        CheckExpiration = False
        Exit Function
    End If

    today = Date
    If today > expireDate Then
        MsgBox "Version expired. Please contact author: Liu Yang Email: luckilyliuyang@163.com", vbCritical
        CheckExpiration = False
    Else
        CheckExpiration = True
    End If
End Function
```
