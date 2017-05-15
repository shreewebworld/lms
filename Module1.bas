Attribute VB_Name = "Module1"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Sub ShellEx(PathName As String)
'Sub used to open a non-excutable file
    If ShellExecute(&O0, "Open", PathName, vbNullString, vbNullString, 1) < 33 Then
        Handler Err
    End If

End Sub
