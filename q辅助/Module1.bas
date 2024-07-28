Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function GetForegroundWindow Lib "user32" () As Long


Public XiaoKongs_HWND As Long
Public yueHWND As Long

'第一次打开窗口时记住窗口句柄
Public Function Jizhu() As Long
Jizhu = GetForegroundWindow() '得到活动窗口的句柄
End Function

'检测窗口状态是否隐藏
'    If IsWindowVisible(131668) Then
'         Label2.Caption = "打开"
'         Else
'         Label2.Caption = "关闭"
'    End If

