Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function RegisterHotKey Lib "User32" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "User32" (ByVal hWnd As Long, ByVal id As Long) As Long

Private preWinProc As Long
Private Modifiers As Long
Private Const WM_HOTKEY = &H312
Private Const GWL_WNDPROC = (-4)
Private Const RHKmodifier = 0

Const RHK_END_ID As Long = 1    '注册end键
Const RHK_HOME_ID As Long = 2   '注册HOME键

Private Function wndproc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg = WM_HOTKEY Then
        Select Case (wParam)
        Case RHK_HOME_ID:
            MsgBox "000"
        Case RHK_END_ID:
            Shell "notepad", vbNormalFocus
        End Select
    End If
    '与form_load中的ADDRESSOF WNDPROC对应
    wndproc = CallWindowProc(preWinProc, hWnd, Msg, wParam, lParam)
End Function

Sub RegHotKeyA(FormHwnd As Long)
    preWinProc = GetWindowLong(FormHwnd, GWL_WNDPROC) '获得窗口过程的地址，或代表窗口过程的地址的句柄
    Call SetWindowLong(FormHwnd, GWL_WNDPROC, AddressOf wndproc)
    RegisterHotKey FormHwnd, RHK_END_ID, RHKmodifier, vbKeyEnd
    RegisterHotKey FormHwnd, RHK_HOME_ID, RHKmodifier, vbKeyHome
End Sub

Public Sub UnRegHotKey(FormHwnd As Long)
    SetWindowLong FormHwnd, GWL_WNDPROC, preWinProc
    Call UnregisterHotKey(FormHwnd, vbKeyEnd)
    Call UnregisterHotKey(FormHwnd, vbKeyHome)
End Sub
