VERSION 5.00
Begin VB.Form CloseDisplayW 
   BorderStyle     =   0  'None
   Caption         =   "XIAOKONGS實驗室"
   ClientHeight    =   600
   ClientLeft      =   20355
   ClientTop       =   11115
   ClientWidth     =   1260
   Icon            =   "CloseDisplayW.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   600
   ScaleWidth      =   1260
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "XIAOKONGS"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "XIAOKONGS"
      ForeColor       =   &H00800000&
      Height          =   180
      Left            =   200
      TabIndex        =   1
      Top             =   120
      Width           =   810
   End
   Begin VB.Menu b 
      Caption         =   "XIAOKONGS"
      Visible         =   0   'False
      Begin VB.Menu xuanxiang7 
         Caption         =   "58"
      End
      Begin VB.Menu xuanxiang5 
         Caption         =   "京东"
      End
      Begin VB.Menu xuanxiang8 
         Caption         =   "-"
      End
      Begin VB.Menu xuanxiang2 
         Caption         =   "百度搜索"
      End
      Begin VB.Menu xuanxiang4 
         Caption         =   "打开 XIAOKONGS 實驗室"
      End
      Begin VB.Menu CloseComputer 
         Caption         =   "瞬间关闭系统"
      End
      Begin VB.Menu RestartComputer 
         Caption         =   "重启计算机"
      End
      Begin VB.Menu xuanxiang10 
         Caption         =   "系统优化"
      End
      Begin VB.Menu sb 
         Caption         =   "-"
      End
      Begin VB.Menu xuanxiang6 
         Caption         =   "建设银行网上银行"
      End
      Begin VB.Menu xuanxiang3 
         Caption         =   "关于桌面小工具"
      End
      Begin VB.Menu xuanxiang1 
         Caption         =   "退出"
      End
   End
End
Attribute VB_Name = "CloseDisplayW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This version only supports Windows 7
'经过特别优化的界面
'版权所有 XIAOKONGS 2017

Private Declare Function SendScreenMessage Lib "User32" _
   Alias "SendMessageA" _
  (ByVal Hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

Private Const MONITOR_ON = -1&
Private Const MONITOR_LOWPOWER = 1&
Private Const MONITOR_OFF = 2&
Private Const SC_MONITORPOWER = &HF170&
Private Const WM_SYSCOMMAND = &H112

Private Declare Function SetWindowPos Lib "User32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
'Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3
Private Const GW_CHILD = 5
Private Const GW_OWNER = 4
Private Const GW_MAX = 5
Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_SYSMENU = &H80000
Private Enum ESetWindowPosStyles
        SWP_SHOWWINDOW = &H40
        SWP_HIDEWINDOW = &H80
        SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
        SWP_NOACTIVATE = &H10
        SWP_NOCOPYBITS = &H100
        SWP_NOMOVE = &H2
        SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
        SWP_NOREDRAW = &H8
        SWP_NOREPOSITION = SWP_NOOWNERZORDER
        SWP_NOSIZE = &H1
        SWP_NOZORDER = &H4
        SWP_DRAWFRAME = SWP_FRAMECHANGED
        HWND_NOTOPMOST = -2
End Enum
Private Declare Function GetWindowRect Lib "User32" (ByVal Hwnd As Long, lpRect As RECT) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, _
ByVal lpoperation As String, ByVal lpfile As String, ByVal lpparameters As String, _
ByVal lpdirectory As String, ByVal nshowcmd As Long) As Long

'Private Declare Function RtlAdjustPrivilege Lib "NTDLL.DLL" (ByVal Privilege As Long, ByVal Enable As Boolean, ByVal Client As Boolean, WasEnabled As Long) As Long
Private Declare Function RtlAdjustPrivilege& Lib "ntdll" (ByVal Privileges As Long, Optional ByVal NewValue As Long = 1, Optional ByVal Thread As Long, Optional Value As Long)

Private Declare Function NtShutdownSystem Lib "NTDLL.DLL" (ByVal ShutdownAction As Long) As Long
'//前两句声明两个API函数，你可以在API函数查询器中查到这两个函数的功能和各个参数的意义
Private Const SE_SHUTDOWN_PRIVILEGE& = 19
Private Const shutdown& = 0
Private Const RESTART& = 1
Private Const HWND_TOPMOST = -1
Dim xa As Single, ya As Single
 
'显示、隐藏标题栏函数
Public Function ShowTitleBar(chenjl1031 As Form, ByVal bState As Boolean)
         Dim lStyle As Long
         Dim tR As RECT
         'Dim playscreen As Variant
         On Error Resume Next
         GetWindowRect chenjl1031.Hwnd, tR
         lStyle = GetWindowLong(chenjl1031.Hwnd, GWL_STYLE)
         If (bState) Then
            If chenjl1031.ControlBox Then
               lStyle = lStyle Or WS_SYSMENU
            End If
            If chenjl1031.MaxButton Then
               lStyle = lStyle Or WS_MAXIMIZEBOX
            End If
            If chenjl1031.MinButton Then
               lStyle = lStyle Or WS_MINIMIZEBOX
            End If
            If chenjl1031.Caption <> "" Then
               lStyle = lStyle Or WS_CAPTION
            End If
         Else
            lStyle = lStyle And Not WS_SYSMENU
            lStyle = lStyle And Not WS_MAXIMIZEBOX
            lStyle = lStyle And Not WS_MINIMIZEBOX
            lStyle = lStyle And Not WS_CAPTION
         End If
         SetWindowLong chenjl1031.Hwnd, GWL_STYLE, lStyle
'         SetWindowPos chenjl1031.hwnd, 0, tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, SWP_NOREPOSITION Or SWP_NOZORDER Or SWP_FRAMECHANGED
         chenjl1031.Refresh
End Function

'关闭 显示器
Public Function MonitorOff(Form As Form)
    
    Call SendScreenMessage(Form.Hwnd, WM_SYSCOMMAND, SC_MONITORPOWER, ByVal MONITOR_OFF)

End Function

'开启显示器
Public Function MonitorOn(Form As Form)
    
    Call SendScreenMessage(Form.Hwnd, WM_SYSCOMMAND, SC_MONITORPOWER, ByVal MONITOR_ON)

End Function

'关闭显示器电源 :)---深度睡眠
Public Function MonitorPowerDown(Form As Form)
    
    Call SendScreenMessage(Form.Hwnd, WM_SYSCOMMAND, SC_MONITORPOWER, ByVal MONITOR_LOWPOWER)
    
End Function

'查询显示器状态'需要引用 Microsoft WMI Scipting V1.2 Library
Public Function WMIVideoControllerInfo() As Long
    Dim WMIObjSet As SWbemObjectSet
    Dim obj As SWbemObject
    Dim St As String
    
    Set WMIObjSet = GetObject("winmgmts:{impersonationLevel=impersonate}"). _
                        InstancesOf("Win32_VideoController")
    
    On Local Error Resume Next
    
    
    For Each obj In WMIObjSet
        WMIVideoControllerInfo = obj.Availability
        
        Select Case WMIVideoControllerInfo
        Case 1
           St = "其他"
        Case 2
           St = "未知"
        Case 3
           St = "运行"
        Case 4
           St = "警告"
        Case 5
           St = "试验"
        Case 6
           St = "不适用"
        Case 7
           St = "关闭电源"
        Case 8
           St = "离线"
        Case 9
           St = "下班"
        Case 10
           St = "退化"
        Case 11
           St = "未安装"
        Case 12
           St = "安装错误"
        Case 13
           St = "省电-未知" '该装置被称为是在省电模式，但其确切身份不明。
        Case 14
           St = "省电-低功耗" '该装置是在省电状态，但仍然运作，可能会出现退化的表现。
        Case 15
           St = "省电-待命" '该设备不能正常运行，但可以使全部力量迅速
        Case 16
           St = "动力循环"
        Case 17
           St = "省电警告" '该装置是在预警状态，虽然也处于省电模式。
        End Select
    Next
End Function

Private Sub CloseComputer_Click()
Call CloseComputerBy
End Sub

Private Sub Command1_Click()
'beep
'MonitorOff Me
'MonitorPowerDown Me

ShowClose
'TestPostmessageDown

End Sub

Public Function ActivateWindow()
    Dim hWndTarget As Long

    ' 获取目标窗口的句柄
    hWndTarget = FindWindow(vbNullString, "TIM")
'     hWndTarget = FindWindow(抖音, vbNullString)
    Debug.Print hWndTarget

    If hWndTarget <> 0 Then
        ' 激活目标窗口到前台
        SetForegroundWindow hWndTarget
'        DoEvents
'        SendKeys "{DOWN}"
    Else
        MsgBox "无法找到目标窗口"
    End If
End Function

Private Sub Form_Load()

    Dim i, a As String
        If App.PrevInstance = True Then
            MsgBox "你已经启动了桌面小工具！", vbOKOnly + 48, "警告"
            End
        End If

    RtlAdjustPrivilege 20

    If Dir("c:\abc.txt") = "" Then
        Open "c:\abc.txt" For Output As #1
        Print #1, "Qiyiservice*sppsvc*iexplore*QyClient*QyFragment*QyPlayer*AndroidService*QyKernel*chrome*cloudmusic*QQprotect*AppleMobileDeviceService*mDNSResponder*rundll32*BaiduNetdisk*BaiduNetdiskHost*YunDetectService*"
'         表示新建文件写abc内容
        Close #1 '关闭文件
    End If

    i = WMIVideoControllerInfo
    ShowTitleBar CloseDisplayW, False
    SetWindowPos Me.Hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE    '置顶

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Me.Move Me.Left + X - xa, Me.Top + Y - ya
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xa = X
ya = Y
'利用PopupMenu方法
  If Button And vbRightButton Then
     CloseDisplayW.PopupMenu b, 0, X, Y '弹出菜单
  End If
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
 On Error Resume Next
 Dim i As Long
 For i = 1 To Data.Files.Count '逐个读取文件路径
        Debug.Print Data.Files(i)
    Next
End Sub

Private Sub Label1_Click()
'MsgBox "版权所有 XIAOKONGS 實驗室", vbOKOnly + 48, "XIAOKONGS 桌面小工具"
End Sub

Private Sub RestartComputer_Click()
RestartComputerBy
End Sub

Private Sub Timer1_Timer()
Call icatch
End Sub

Private Sub xuanxiang1_Click()
Unload Me
End Sub

Public Sub CloseComputerBy()
RtlAdjustPrivilege SE_SHUTDOWN_PRIVILEGE, 1, 0, 0
'//等同于RtlAdjustPrivilege（SE_SHUTDOWN_PRIVILEGE,1,0,0）,是对上面声明的第一个API函数的调用
NtShutdownSystem shutdown
'//同理，是对第二API函数的调用，参数为shutdown
End Sub

Public Sub RestartComputerBy()
RtlAdjustPrivilege SE_SHUTDOWN_PRIVILEGE, 1, 0, 0
'//等同于RtlAdjustPrivilege（SE_SHUTDOWN_PRIVILEGE,1,0,0）,是对上面声明的第一个API函数的调用
NtShutdownSystem RESTART
End Sub

Private Sub xuanxiang10_Click()
    Call beep
    Dim a$
    Open "c:\abc.txt" For Input As #1
    Do
    Input #1, a
    sss = sss & a & vbCrLf
    Loop Until EOF(1)
    Close #1
    Call RefreshStack
End Sub

'打开百度搜索
Private Sub xuanxiang2_Click()
    '119.29.135.68
    Dim lngReturn As Long
    lngReturn = ShellExecute(Me.Hwnd, "open", "http://www.baidu.com", "", "", 0)
End Sub

Private Sub xuanxiang3_Click()
MsgBox "版权所有 XIAOKONGS 實驗室", vbOKOnly + 48, "桌面小工具"
End Sub
'打开XIAOKONGS首页
Private Sub xuanxiang4_Click()
'119.29.135.68
Dim lngReturn As Long
lngReturn = ShellExecute(Me.Hwnd, "open", "http://14.103.51.243/inc/", "", "", 0)
End Sub
'打开京东
Private Sub xuanxiang5_Click()
'Dim ws
'Set ws = CreateObject("wscript.shell")
'ws.run "iexplore.exe www.baidu.com"
'https://www.jd.com
Dim lngReturn As Long
lngReturn = ShellExecute(Me.Hwnd, "open", "https://www.jd.com", "", "", 0)
End Sub

Private Sub xuanxiang6_Click()
'https://ibsbjstar.ccb.com.cn/CCBIS/V6/common/login.jsp?UDC_CUSTOMER_ID=&UDC_CUSTOMER_NAME=&UDC_COOKIE=5075ef964b8e1a03QTrz8E71o11A7f3Rzcx21550996929289szmD6asRy3auo9Ga5M2T9831212a51cbf058d6ca5f6bd2cc7e38&UDC_SESSION_ID=Ur8uOLZbJNxeYXb4f69b12070b8-20190224235422
Dim lngReturn As Long
lngReturn = ShellExecute(Me.Hwnd, "open", "https://ibsbjstar.ccb.com.cn/CCBIS/V6/common/login.jsp", "", "", 0)
End Sub

Private Sub xuanxiang7_Click()
Dim url
url = "https://nj.58.com/job.shtml?PGTID=0d100000-000a-cd3a-9b37-0e8c01f15ba1&ClickID=3"
Shell "cmd.exe /c start " & url, 0
End Sub
