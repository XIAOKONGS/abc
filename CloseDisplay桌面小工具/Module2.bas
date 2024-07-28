Attribute VB_Name = "Module2"
Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "User32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetWindowText Lib "User32" Alias "GetWindowTextA" (ByVal Hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal Hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Const MaxControlUnit = 65535

Private Declare Function ShowWindow Lib "User32" (ByVal Hwnd As Long, ByVal nCmdShow As Long) As Long

Private Const SW_MAXIMIZE As Long = 3
Private Declare Function IsWindowVisible Lib "User32" (ByVal Hwnd As Long) As Long '判断窗口状态

Dim QQexternHwnd As Long

Public Sub icatch()
QQexternHwnd = FindWindow(vbNullString, "永恒传奇3")

            If QQexternHwnd > 0 Then
                SetForegroundWindow QQexternHwnd
            End If
            
End Sub


Public Sub ShowClose()
'显示或者隐藏窗口
    If IsWindowVisible(QQexternHwnd) Then
         '不显示窗口
         ShowWindow QQexternHwnd, 0
         Else
          '使窗口句柄最大
         ShowWindow QQexternHwnd, SW_MAXIMIZE
    End If
End Sub
