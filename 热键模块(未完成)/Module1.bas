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

Const RHK_END_ID As Long = 1    'ע��end��
Const RHK_HOME_ID As Long = 2   'ע��HOME��

Private Function wndproc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg = WM_HOTKEY Then
        Select Case (wParam)
        Case RHK_HOME_ID:
            MsgBox "000"
        Case RHK_END_ID:
            Shell "notepad", vbNormalFocus
        End Select
    End If
    '��form_load�е�ADDRESSOF WNDPROC��Ӧ
    wndproc = CallWindowProc(preWinProc, hWnd, Msg, wParam, lParam)
End Function

Sub RegHotKeyA(FormHwnd As Long)
    preWinProc = GetWindowLong(FormHwnd, GWL_WNDPROC) '��ô��ڹ��̵ĵ�ַ��������ڹ��̵ĵ�ַ�ľ��
    Call SetWindowLong(FormHwnd, GWL_WNDPROC, AddressOf wndproc)
    RegisterHotKey FormHwnd, RHK_END_ID, RHKmodifier, vbKeyEnd
    RegisterHotKey FormHwnd, RHK_HOME_ID, RHKmodifier, vbKeyHome
End Sub

Public Sub UnRegHotKey(FormHwnd As Long)
    SetWindowLong FormHwnd, GWL_WNDPROC, preWinProc
    Call UnregisterHotKey(FormHwnd, vbKeyEnd)
    Call UnregisterHotKey(FormHwnd, vbKeyHome)
End Sub
