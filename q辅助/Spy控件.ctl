VERSION 5.00
Begin VB.UserControl Spy�ؼ� 
   BackStyle       =   0  '͸��
   ClientHeight    =   4470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9105
   ScaleHeight     =   4470
   ScaleWidth      =   9105
   ToolboxBitmap   =   "Spy�ؼ�.ctx":0000
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   600
      Top             =   1080
   End
   Begin VB.Frame Frame2 
      Caption         =   "�϶�ͼ��"
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   975
      Begin VB.Image Picture1 
         Height          =   480
         Left            =   240
         Picture         =   "Spy�ؼ�.ctx":0312
         Top             =   240
         Width           =   480
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000000FF&
         BorderColor     =   &H8000000F&
         BorderWidth     =   5
         Height          =   495
         Left            =   120
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   615
      Left            =   3720
      Picture         =   "Spy�ؼ�.ctx":1EC4
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   11
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Image1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   600
      Picture         =   "Spy�ؼ�.ctx":3A76
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label12 
      Height          =   495
      Left            =   2280
      TabIndex        =   13
      Top             =   1920
      Width           =   6615
   End
   Begin VB.Label Label11 
      Caption         =   "����·��"
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label10 
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   1560
      Width           =   6615
   End
   Begin VB.Label Label9 
      Caption         =   "��������"
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label8 
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   1200
      Width           =   6735
   End
   Begin VB.Label Label7 
      Caption         =   "���̾��"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   840
      Width           =   6615
   End
   Begin VB.Label Label5 
      Caption         =   "�� ��PID"
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   480
      Width           =   6615
   End
   Begin VB.Label Label3 
      Caption         =   "���ھ��"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Width           =   6495
   End
   Begin VB.Label Label1 
      Caption         =   "����  ��"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.Image Picture2 
      Height          =   480
      Left            =   2280
      Picture         =   "Spy�ؼ�.ctx":5628
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "Spy�ؼ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'�ؼ���: spy�ؼ�
'��  Ȩ: www.51xue8xue8.com
'�� �� :ֱ��[��ȡ����pid,���̾��,��������,·��,������]
'ʹ��˵����:
'��һ��:��VB���̼������spy�ؼ�
'�ڶ���:��VB����Ϳ��Ե����������
'spy�ؼ�.����Pid
'spy�ؼ�.���̾��
'spy�ؼ�.���ھ��
'spy�ؼ�.��������
'spy�ؼ�.����·��


Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
 Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
 
 Private Declare Function GetCursorPos Lib "user32" (lpPoint As POSSITION) As Long
 Private Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
 Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Declare Function OpenProcess _
                Lib "Kernel32.dll" (ByVal dwDesiredAccess As Long, _
                                    ByVal bInheritHandle As Long, _
                                    ByVal dwProcessId As Long) As Long


 Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
 Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPriv As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long                'Used to adjust your program's security privileges, can't restore without it!
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As Any, ByVal lpName As String, lpLuid As LUID) As Long          'Returns a valid LUID which is important when making security changes in NT.
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
'Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Const TOKEN_QUERY As Long = &H8&
Private Const TOKEN_ADJUST_PRIVILEGES As Long = &H20&
Private Const SE_PRIVILEGE_ENABLED As Long = &H2

Private Type LUID
   lowpart As Long
   highpart As Long
End Type
Private Type LUID_AND_ATTRIBUTES
   pLuid As LUID
   Attributes As Long
End Type
Private Type TOKEN_PRIVILEGES
   PrivilegeCount As Long
   Privileges As LUID_AND_ATTRIBUTES
End Type
 
 
 
 Private Const HWND_TOPMOST = -1
Private Const SWP_SHOWWINDOW = &H40
 Private Const WM_GETTEXTLENGTH = &HE '&H ������---��16���Ʊ�ʶ��
 Private Const WM_GETTEXT = &HD '&H ������---��16���Ʊ�ʶ��

Private Type POSSITION
    x As Long
    y As Long
End Type
 
 Public ���̾�� As Long
 Public ����Pid As Long
 Public ���ھ�� As Long   'Ҫ��ȡ�ľ��
 Public �������� As String    'Ҫ��ȡ������
   Public ������ As String
   Public ����·�� As String
'��Դ����ԣ�www.51xue8xue8.com




Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
      On Error GoTo k
      'ע�ͣ�����ͼ��'
     Form1.MousePointer = 99
    Form1.MouseIcon = Picture2.Picture ' LoadResPicture(101, vbResIcon) ' LoadPicture(App.Path & "\ico.ico")
    Picture1.Picture = Nothing
k:
End Sub


Private Function WoNiu(ByVal hwd As Long) As String 'x��ʾ�����ڵ��������������������ƣ���y��ʾ�Ӵ��ڵ��������Ӵ����������ƣ�
    Dim a As Long
    Dim astr As String * 256 '��ʾֻ�ܴ洢256�ַ�,��Ȼ��Ҳ����д��1000��
    a = SendMessage(hwd, WM_GETTEXTLENGTH, 0&, vbNull)
          SendMessage hwd, WM_GETTEXT, a + 1, astr
 WoNiu = astr
End Function

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo m
    Form1.MousePointer = 0
    Picture1.Picture = Picture2.Picture
    Dim xy As POSSITION, Chuang As String * 20
   
    GetCursorPos xy
    ���ھ�� = WindowFromPointXY(xy.x, xy.y)
    
    GetWindowText ���ھ��, Chuang, 20 '���
    ������ = Chuang
    

    GetWindowThreadProcessId ���ھ��, ����Pid
   
     EnablePrivilege "SeDebugPrivilege" '���VB�ĵ���Ȩ��
    ���̾�� = OpenProcess(&H1F0FFF, 0, ����Pid)
 

 �������� = String(&H100, vbNullChar)   '����256���ȵ��ַ�����

 '��ȡ��������
 GetClassName ���ھ��, ByVal ��������, Len(��������)
 
         Label10.Caption = ��������  '��ʾ���������
         Label2.Caption = WoNiu(���ھ��) '������
         ������ = WoNiu(���ھ��)
         Label4.Caption = ���ھ�� '���ھ��
         Label6.Caption = ����Pid '����pid
         Label8.Caption = ���̾�� '���̾��
      
         Label12.Caption = GetProcessPathByProcessID(����Pid)   '����·��
         ����·�� = GetProcessPathByProcessID(����Pid)
'    Frame2.Caption = "�ѻ��" & ���ھ��
   Frame2.Caption = " �ѻ��"
m:
End Sub



 


'���ݽ��̺Ż�ȡ����·��������
Private Function GetProcessPathByProcessID(PID As Long) As String
    On Error GoTo Z
    Dim cbNeeded As Long
    Dim szBuf(1 To 250) As Long
    Dim Ret As Long
    Dim szPathName As String
    Dim nSize As Long
    Dim hProcess As Long
    hProcess = OpenProcess(&H400 Or &H10, 0, PID)
    If hProcess <> 0 Then
        Ret = EnumProcessModules(hProcess, szBuf(1), 250, cbNeeded)
        If Ret <> 0 Then
            szPathName = Space(260)
            nSize = 500
            Ret = GetModuleFileNameExA(hProcess, szBuf(1), szPathName, nSize)
            GetProcessPathByProcessID = Left(szPathName, Ret)
        End If
    End If
    Ret = CloseHandle(hProcess)
    If GetProcessPathByProcessID = "" Then
       GetProcessPathByProcessID = "SYSTEM"
    End If
    Exit Function
Z:
End Function






Private Function EnablePrivilege(seName As String) As Boolean
    Dim p_lngRtn As Long
    Dim p_lngToken As Long
    Dim p_lngBufferLen As Long
    Dim p_typLUID As LUID
    Dim p_typTokenPriv As TOKEN_PRIVILEGES
    Dim p_typPrevTokenPriv As TOKEN_PRIVILEGES
    p_lngRtn = OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, p_lngToken)
    If p_lngRtn = 0 Then
        Exit Function ' Failed
    ElseIf Err.LastDllError <> 0 Then
        Exit Function ' Failed
    End If
    p_lngRtn = LookupPrivilegeValue(0&, seName, p_typLUID)  'Used to look up privileges LUID.
    If p_lngRtn = 0 Then
        Exit Function ' Failed
    End If
    ' Set it up to adjust the program's security privilege.
    p_typTokenPriv.PrivilegeCount = 1
    p_typTokenPriv.Privileges.Attributes = SE_PRIVILEGE_ENABLED
    p_typTokenPriv.Privileges.pLuid = p_typLUID
    EnablePrivilege = (AdjustTokenPrivileges(p_lngToken, False, p_typTokenPriv, Len(p_typPrevTokenPriv), p_typPrevTokenPriv, p_lngBufferLen) <> 0)
End Function

Private Sub Timer1_Timer()

If ����Pid = 0 Then
      If Shape1.BorderColor = &H8000000F Then
       Shape1.BorderColor = vbWhite
      Else
       Shape1.BorderColor = &H8000000F
     End If
 Else
     Shape1.BorderColor = &H8000000F
End If
 

End Sub

