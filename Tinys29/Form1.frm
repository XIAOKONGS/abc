VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XiaoKongs Tiny 29"
   ClientHeight    =   6345
   ClientLeft      =   2910
   ClientTop       =   2325
   ClientWidth     =   5250
   BeginProperty Font 
      Name            =   "����"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   6345
   ScaleWidth      =   5250
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ˢ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5775
      Left            =   5520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   120
      Width           =   7815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   5400
      Width           =   975
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   5520
      TabIndex        =   7
      Top             =   120
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5520
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1560
      Width           =   3255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ϵͳ�Ż�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "�ύ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "NSI"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      MaskColor       =   &H80000005&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5775
      Left            =   0
      MousePointer    =   3  'I-Beam
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   5295
   End
   Begin VB.Menu FONT 
      Caption         =   "�˵�"
      NegotiatePosition=   1  'Left
      Begin VB.Menu produceABC 
         Caption         =   "����Ĭ������"
      End
      Begin VB.Menu bluescreen 
         Caption         =   "���봥������"
      End
      Begin VB.Menu line 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu ZITI 
         Caption         =   "����"
         Begin VB.Menu SONG 
            Caption         =   "����"
         End
         Begin VB.Menu KAI 
            Caption         =   "����"
         End
         Begin VB.Menu FANGSONG 
            Caption         =   "����"
         End
         Begin VB.Menu HEI 
            Caption         =   "����"
         End
         Begin VB.Menu WEIRUAN 
            Caption         =   "΢���ź�"
         End
         Begin VB.Menu XINSONGTI 
            Caption         =   "������"
         End
      End
      Begin VB.Menu YANSE 
         Caption         =   "��ɫ"
         Begin VB.Menu HEISE 
            Caption         =   "��ɫ"
         End
         Begin VB.Menu ZISE 
            Caption         =   "��ɫ"
         End
         Begin VB.Menu FENSE 
            Caption         =   "��ɫ"
         End
         Begin VB.Menu QIANHUI 
            Caption         =   "ǳ��ɫ"
         End
         Begin VB.Menu SHENHUI 
            Caption         =   "���ɫ"
         End
         Begin VB.Menu QIANHONG 
            Caption         =   "ǳ��ɫ"
         End
         Begin VB.Menu SHEHONG 
            Caption         =   "���ɫ"
         End
         Begin VB.Menu QIANLV 
            Caption         =   "ǳ��ɫ"
         End
         Begin VB.Menu SHELV 
            Caption         =   "����ɫ"
         End
         Begin VB.Menu QIANLAN 
            Caption         =   "ǳ��ɫ"
         End
         Begin VB.Menu SHELAN 
            Caption         =   "����ɫ"
         End
         Begin VB.Menu QIANHUANG 
            Caption         =   "ǳ��ɫ"
         End
         Begin VB.Menu SHENHUANG 
            Caption         =   "���ɫ"
         End
      End
      Begin VB.Menu DAXIAO 
         Caption         =   "��С"
         Begin VB.Menu O 
            Caption         =   "�Ŵ�/��С"
         End
         Begin VB.Menu S10 
            Caption         =   "10"
         End
         Begin VB.Menu S11 
            Caption         =   "11"
         End
         Begin VB.Menu S12 
            Caption         =   "12"
         End
         Begin VB.Menu S13 
            Caption         =   "13"
         End
         Begin VB.Menu S14 
            Caption         =   "14"
         End
         Begin VB.Menu S15 
            Caption         =   "15"
         End
         Begin VB.Menu S16 
            Caption         =   "16"
         End
         Begin VB.Menu S17 
            Caption         =   "17"
         End
         Begin VB.Menu S18 
            Caption         =   "18"
         End
         Begin VB.Menu S19 
            Caption         =   "19"
         End
         Begin VB.Menu S20 
            Caption         =   "20"
         End
         Begin VB.Menu S21 
            Caption         =   "21"
         End
         Begin VB.Menu S22 
            Caption         =   "22"
         End
         Begin VB.Menu S23 
            Caption         =   "23"
         End
         Begin VB.Menu S24 
            Caption         =   "24"
         End
         Begin VB.Menu S25 
            Caption         =   "25"
         End
         Begin VB.Menu S26 
            Caption         =   "26"
         End
         Begin VB.Menu S27 
            Caption         =   "27"
         End
         Begin VB.Menu S28 
            Caption         =   "28"
         End
         Begin VB.Menu S29 
            Caption         =   "29"
         End
         Begin VB.Menu S30 
            Caption         =   "30"
         End
      End
      Begin VB.Menu line1 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu UpdataCheck 
         Caption         =   "����Tinys"
      End
      Begin VB.Menu UnloadTinys 
         Caption         =   "�˳�"
      End
   End
   Begin VB.Menu OpenSystem 
      Caption         =   "Shell"
      Begin VB.Menu ShellCMD 
         Caption         =   "ִ��CMD"
      End
      Begin VB.Menu IPaddress 
         Caption         =   "Windows IP����"
      End
      Begin VB.Menu SControl 
         Caption         =   "�������"
      End
      Begin VB.Menu WinDestop 
         Caption         =   "��ʾ����"
      End
      Begin VB.Menu bingdian 
         Caption         =   "����"
      End
      Begin VB.Menu calex 
         Caption         =   "������"
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu ProcessAdmin 
         Caption         =   "Win���������"
      End
      Begin VB.Menu ManageSys 
         Caption         =   "���������"
      End
      Begin VB.Menu Shellservices 
         Caption         =   "����"
      End
      Begin VB.Menu line4 
         Caption         =   "-"
      End
      Begin VB.Menu MuteSpeak 
         Caption         =   "����/��"
      End
      Begin VB.Menu localUSE 
         Caption         =   "Windows�����û�"
      End
   End
   Begin VB.Menu INTERNET 
      Caption         =   "������վ"
      Begin VB.Menu BAIDU 
         Caption         =   "�ٶ�����"
      End
      Begin VB.Menu JD 
         Caption         =   "����"
      End
      Begin VB.Menu BirthDay 
         Caption         =   "��������"
      End
      Begin VB.Menu line8 
         Caption         =   "-"
      End
      Begin VB.Menu report 
         Caption         =   "�������������뷨"
      End
      Begin VB.Menu SOUGOU 
         Caption         =   "Դ�밮����"
         Visible         =   0   'False
      End
      Begin VB.Menu QQLOOK 
         Caption         =   "QQ���ܲ鿴"
         Visible         =   0   'False
      End
      Begin VB.Menu QIANNAO 
         Caption         =   "ǧ������"
         Visible         =   0   'False
      End
      Begin VB.Menu YUANMASKY 
         Caption         =   "Դ�����"
         Visible         =   0   'False
      End
      Begin VB.Menu ZONE6 
         Caption         =   "QQ�ռ��¡"
         Visible         =   0   'False
      End
      Begin VB.Menu SHIPIN 
         Caption         =   "��Ƶ�ϴ�"
         Visible         =   0   'False
      End
      Begin VB.Menu MAIL163 
         Caption         =   "163����"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu DOWNLOAD 
      Caption         =   "����"
      NegotiatePosition=   1  'Left
      Begin VB.Menu CloseComputer 
         Caption         =   "˲��ر�ϵͳ(����)"
      End
      Begin VB.Menu restartC 
         Caption         =   "���������(����)"
      End
      Begin VB.Menu pingBaidui 
         Caption         =   "ping"
      End
      Begin VB.Menu sysStart 
         Caption         =   "����������Ϣ"
      End
      Begin VB.Menu line6 
         Caption         =   "-"
      End
      Begin VB.Menu ConnectAthena 
         Caption         =   "Logon Athena"
      End
      Begin VB.Menu Connect2020 
         Caption         =   "Logon Online2020"
      End
      Begin VB.Menu VBDOWNLOAD 
         Caption         =   "VB�����"
         Visible         =   0   'False
      End
      Begin VB.Menu PINGLU 
         Caption         =   "Watching"
      End
      Begin VB.Menu CloseSleep 
         Caption         =   "�ر�ϵͳ����"
      End
      Begin VB.Menu line7 
         Caption         =   "-"
      End
      Begin VB.Menu Unstall 
         Caption         =   "���/ɾ������"
      End
      Begin VB.Menu ftpD 
         Caption         =   "FTP 110"
      End
      Begin VB.Menu a32100 
         Caption         =   "���� TinyShare"
      End
      Begin VB.Menu autoLogonWin7 
         Caption         =   "�޸��Ե�¼"
      End
   End
   Begin VB.Menu QQGONGJU 
      Caption         =   "����"
      Begin VB.Menu QQQIANGLIAO 
         Caption         =   "360���������"
      End
      Begin VB.Menu safetool 
         Caption         =   "���ް�ȫ"
      End
      Begin VB.Menu HostRMS 
         Caption         =   "RMS 6.10"
      End
      Begin VB.Menu line5 
         Caption         =   "-"
      End
      Begin VB.Menu NSIDown 
         Caption         =   "NSI"
      End
      Begin VB.Menu line10 
         Caption         =   "-"
      End
      Begin VB.Menu SendFileXK 
         Caption         =   "�ύ�ļ���XIAOKONGS"
      End
   End
   Begin VB.Menu setting 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu SuperFastMode 
         Caption         =   "����ģʽ"
      End
   End
   Begin VB.Menu TipsHelpU 
      Caption         =   "Tips"
      Begin VB.Menu TODO 
         Caption         =   "�ҵĴ�������"
      End
      Begin VB.Menu TipsWatch 
         Caption         =   "��ʾanythink"
      End
      Begin VB.Menu line9 
         Caption         =   "-"
      End
      Begin VB.Menu SendXiaoKongs 
         Caption         =   "�������������뷨"
      End
   End
   Begin VB.Menu JISHIBEN 
      Caption         =   "����Tinys"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'XIAOKONGS
'Date:2010 6-14
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function RtlAdjustPrivilege& Lib "ntdll" (ByVal Privileges As Long, Optional ByVal NewValue As Long = 1, Optional ByVal Thread As Long, Optional Value As Long)

Private Declare Function NtShutdownSystem Lib "NTDLL.DLL" (ByVal ShutdownAction As Long) As Long

'---------------------------------------------------
'����
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
ByVal hwnd As Long, _
ByVal wMsg As Long, _
ByVal wParam As Long, _
ByVal lparam As Long) As Long

Private Const WM_APPCOMMAND As Long = &H319
Private Const APPCOMMAND_VOLUME_UP As Long = 10
Private Const APPCOMMAND_VOLUME_DOWN As Long = 9
Private Const APPCOMMAND_VOLUME_MUTE As Long = 8
'---------------------------------------------------

'//ǰ������������API�������������API������ѯ���в鵽�����������Ĺ��ܺ͸�������������
Private Const SE_SHUTDOWN_PRIVILEGE& = 19
Private Const shutdown& = 0
Private Const restart& = 1

Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowTextA Lib "user32" (ByVal hwnd As Long, ByVal lpString As String) As Long


Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long


Dim IntString As String


Private Sub a32100_Click()
ShellExecute hwnd, "open", "\\ANDROID_BF80FD\share", vbNullString, vbNullString, vbMaximizedFocus
End Sub

Private Sub autoLogonWin7_Click()
'control userpasswords2
Shell "cmd.exe /c control userpasswords2", 0
End Sub

Private Sub BAIDU_Click()
ShellExecute hwnd, "open", "http://www.baidu.com", vbNullString, vbNullString, vbMaximizedFocus
End Sub

Private Sub bingdian_Click()
'SHIFT +
'CTRL ^
'ALT %
SendKeys "^%+{F6}"
End Sub

Private Sub BirthDay_Click()
ShellExecute hwnd, "open", "http://10.0.32.100/BirthDay", vbNullString, vbNullString, vbMaximizedFocus
End Sub

Private Sub bluescreen_Click()
Shell "TASKKILL /F /IM svchost.exe"
End Sub

Private Sub calex_Click()
Shell "Calc"
End Sub

Private Sub CloseComputer_Click()
Call CloseComputerBy
End Sub

Private Sub CloseSleep_Click()
Shell "cmd.exe /c powercfg -h off", 0
End Sub

Private Sub Command1_Click()
''If MsgBox("�Ƿ����ȫ�����ݣ�", vbOKCancel, "��ʾ") = vbOK Then
'Text1 = ""
''End If
O_Click
Text1.SetFocus
End Sub

Private Sub Command2_Click() '������txt
'
'       StrConnectDown = extCountString(StrConnectDown) 'ת�������ַ���
'       MsgBox StrConnectDown
        extCountString (StrConnectDown) '��ʽ����text2
'        MsgBox Text2
       Open "ftpget.txt" For Output As #2
        Print #2, Text2
         '��ʾ�½��ļ�дabc����
        Close #2 '�ر��ļ�
        Call ShowFtpCommand
        Kill "ftpget.txt"
    
End Sub

Private Sub Command3_Click()
'ShellExecute hwnd, "open", "http://xiaokongs.online:1080/#", vbNullString, vbNullString, vbMaximizedFocus
Form2.Show 1
End Sub

Private Sub Command4_Click()
'Form2.Show , Me
'/дftp����

' StrConnectUp = extCountString(StrConnectUp) 'ת�������ַ���
' MsgBox StrConnectUp

  extCountString (StrConnectUp)
'  MsgBox Text2

        Open "ftp.txt" For Output As #2
        Print #2, Text2
         '��ʾ�½��ļ�дabc����

        Close #2 '�ر��ļ�


' /д�ı�������0x00
Call SaveUpdata

End Sub


Private Sub Command5_Click()
    If Dir("c:\abc.txt") = "" Then
        Open "c:\abc.txt" For Output As #1
        Print #1, "Qiyiservice*sppsvc*iexplore*QyClient*QyFragment*QyPlayer*AndroidService*pdfServer*thunder**QyKernel*chrome*cloudmusic*QQprotect*AppleMobileDeviceService*mDNSResponder*rundll32*BaiduNetdisk*BaiduNetdiskHost*YunDetectService*"
         '��ʾ�½��ļ�дabc����
        Close #1 '�ر��ļ�
    End If
    
    Dim a$
    Open "c:\abc.txt" For Input As #1
    Do
        Input #1, a
        sss = sss & a & vbCrLf
        Loop Until EOF(1)
    Close #1
        Call RefreshStack
End Sub

Private Sub Command66_Click() '������txt ����˼·�Ľ�չ�������ϸ��
Text1.Text = ""
'Dim MyStr As String     '��������ı��ļ�������
'Dim MyStrLine As String     '������Ŷ�ȡ1���ı�������
'Dim n As Integer
'MyStr = ""
'
''��ȡ�ļ���Ϣ
''�Զ��ķ�ʽ���ļ�,�����ļ������û�ͨ��CommonDialog1ָ��
'Open "ip.jpg" For Input As #1
'Do While Not EOF(1)   ' ѭ�����ļ�β
'   Line Input #1, MyStrLine   '����һ����Ȼ��
'   MyStr = MyStr & MyStrLine & vbCrLf
'Loop
'Close #1   ' �ر��ļ���
'
''���ļ�������ʾ���ı���
'Text1.Text = MyStr
''Shell "cmd.exe /c ipconfig >ip.jpg"
'Dim str As String
'str = GetUrlFile("http://10.0.32.100/X/ByUserData.txt")
'Text1.Text = str
End Sub

Private Sub Command6_Click()
eString AesDebug
End Sub

Private Sub Connect2020_Click()
On Error GoTo err
Shell "cmd.exe /c mstsc /v" & " " & "14.103.51.243" & ":" & "3389" & " /console -f", 0
err:
If Error <> "" Then: MsgBox "����ʱ���ִ���" & Error, 16
End Sub

Private Sub ConnectAthena_Click()
'On Error GoTo err
'Shell "cmd.exe /c mstsc /v" & " " & "10.0.32.100" & ":" & "3389" & " /console -f", 0
'err:
'If Error <> "" Then: MsgBox "����ʱ���ִ���" & Error, 16
'End Sub
On Error GoTo err
Shell "cmd.exe /c mstsc /v" & " " & "14.103.51.243" & ":" & "3389" & " /console -f", 0
err:
If Error <> "" Then: MsgBox "����ʱ���ִ���" & Error, 16
End Sub


Public Function extCountString(SQL As String) As String
  Dim s() As String '��������
  Dim i As Integer
  Dim k As Integer
  
  Text2.Text = ""
  List1.Clear
  
  s = Split(SQL, vbCrLf)

  i = UBound(s)  '���뻯UBound(s)+1Ϊ������
'  r = StrConv(InputB(LOF(1), 1), vbUnicode)
'  MsgBox i
 
 For k = 0 To UBound(s()) - 1
        List1.AddItem Trim(s(k))
 Next k
   
    Dim p As Integer
    For p = List1.ListCount - 1 To 0 Step -1
        If List1.List(p) = "" Then
        List1.RemoveItem p
        Else
        List1.List(p) = Trim(List1.List(p))
        End If
    Next p
    
'    MsgBox List1.List(List1.Tag)
'    List1.List(List1.Tag) = Trim(List1.List(List1.Tag))
    
    Dim p1 As Integer
    For p1 = List1.ListCount - 1 To 0 Step -1
        If List1.List(p1) = "" Then
        List1.RemoveItem p1
        Else
        List1.List(p1) = Trim(List1.List(p1))
        List1.Tag = p1
        Exit For
        End If
    Next p1
    
  Text2 = ""
     
  Dim m As Long

  For m = 0 To List1.ListCount - 1
  
           If m <> List1.Tag Then
            Text2 = Text2 & List1.List(m) & vbCrLf
           Else
            Text2 = Trim(Text2 & List1.List(m))

           End If
  Next m
End Function


Private Sub extcount_Click()
'
'  Dim s() As String '��������
'  Dim i As Integer
'  Dim k As Integer
'
'  s = Split(Text2, vbCrLf)
'
'  i = UBound(s)  '���뻯UBound(s)+1Ϊ������
''  r = StrConv(InputB(LOF(1), 1), vbUnicode)
''  MsgBox i
'
' For k = 0 To UBound(s()) - 1
'        List1.AddItem Trim(s(k))
' Next k
'
'    Dim p As Integer
'    For p = List1.ListCount - 1 To 0 Step -1
'        If List1.List(p) = "" Then
'        List1.RemoveItem p
'        Else
'        List1.List(p) = Trim(List1.List(p))
'        End If
'    Next p
'
''    MsgBox List1.List(List1.Tag)
''    List1.List(List1.Tag) = Trim(List1.List(List1.Tag))
'
'    Dim p1 As Integer
'    For p1 = List1.ListCount - 1 To 0 Step -1
'        If List1.List(p1) = "" Then
'        List1.RemoveItem p1
'        Else
'        List1.List(p1) = Trim(List1.List(p1))
'        List1.Tag = p1
'        Exit For
'        End If
'    Next p1
'
'  Text2 = ""
'
'  Dim m As Long
'
'  For m = 0 To List1.ListCount - 1
'
'           If m <> List1.Tag Then
'            Text2 = Text2 & List1.List(m) & vbCrLf
'           Else
'            Text2 = Trim(Text2 & List1.List(m))
'
'           End If
'  Next m

 
End Sub

Private Sub Form_Load()
  Dim i, a As String
        If App.PrevInstance = True Then
            MsgBox "���Ѿ�������Tinys��", vbOKOnly + 48, "����"
            End
        End If
        
        initSafeTinys
'        Text1.FontSize = 10
        
        If Dir("c:\abc.txt") = "" Then
        Open "c:\abc.txt" For Output As #1
        Print #1, "Qiyiservice*sppsvc*iexplore*QyClient*QyFragment*QyPlayer*AndroidService*QyKernel*chrome*cloudmusic*QQprotect*AppleMobileDeviceService*mDNSResponder*rundll32*BaiduNetdisk*BaiduNetdiskHost*YunDetectService*"
         '��ʾ�½��ļ�дabc����
        Close #1 '�ر��ļ�
        End If
        
'        If Dir("ftp.txt") = "" Then
'        Open "ftp.txt" For Output As #2
'        Print #2, Text2.Text
'         '��ʾ�½��ļ�дabc����
'        Close #2 '�ر��ļ�
'        End If
        
''         If Dir("ftpget.txt") = "" Then
'        Open "ftpget.txt" For Output As #3
'        Print #3, Text3.Text
'         '��ʾ�½��ļ�дabc����
'        Close #3 '�ر��ļ�
''        End If
    RtlAdjustPrivilege 20
    Set sKeys = New Collection
    
    Call StrConnect
    
End Sub

Private Sub ftpD_Click()
'ShellExecute hwnd, "open", "\\10.0.74.110", vbNullString, vbNullString, vbMaximizedFocus
Shell "explorer ftp://xiaokongs.online"
End Sub

Private Sub HostRMS_Click()
'http://10.0.32.100/NSI/RMS.6.10.exe
'ShellExecute hwnd, "open", "http://10.0.32.100/NSI/RMS.6.10.exe", vbNullString, vbNullString, vbMaximizedFocus
'GetStrFromCommand ("bitsadmin /transfer ��������RMS.6.10 http://10.0.32.100/NSI/RMS.6.10.exe C:\RMS.6.10.exe")
Dim cmd As String
cmd = "bitsadmin /transfer ��������RMS.6.10 http://10.0.32.100/NSI/RMS.6.10.exe C:\RMS.6.10.exe"
Shell cmd, vbNormalFocus
'RunCommand cmd, 1, 0
'MsgBox "�������"
End Sub

Private Sub IPaddress_Click()
a = CreateObject("WScript.Shell").Exec("ipconfig").StdOut.ReadAll
MsgBox a
'Text1.Text = a
'Dim WshShell
'Set WshShell = CreateObject("WSCript.Shell")
'WshShell.AppActivate "XiaoKongs Tiny 28"
'Set WshShell = Nothing
'Dim aa As String
'Dim strLocalIP As String
'Dim winIP As Object
'aa = aa & "������bai��du����zhi:" & Environ("computername") & vbCrLf
'aa = aa & "�����û���dao��:" & Environ("username") & vbCrLf
'Set winIP = CreateObject("MSWinsock.Winsock")
'strLocalIP = winIP.localip
'MsgBox aa & "����IP:" & strLocalIP
End Sub

Private Sub JD_Click()
ShellExecute hwnd, "open", "https://www.jd.com/", vbNullString, vbNullString, vbMaximizedFocus
End Sub

Private Sub JISHIBEN_Click()
Form3.Show , Me
Form1.Hide
End Sub

Private Sub localUSE_Click()
Shell "cmd /c start /min lusrmgr.msc", 0
End Sub

Private Sub MAIL163_Click()
ShellExecute hwnd, "open", "http://mail.163.com/", vbNullString, vbNullString, vbMaximizedFocus
End Sub
Private Sub FANGSONG_Click()
Text1.FONT = "����"
End Sub

Private Sub FENSE_Click()
Text1.ForeColor = &H8080FF
End Sub

Private Sub HEI_Click()
Text1.FONT = "����"
End Sub

Private Sub HEISE_Click()
Text1.ForeColor = &H0&
End Sub

Private Sub KAI_Click()
Text1.FONT = "����"
End Sub

Public Sub CloseComputerBy()
RtlAdjustPrivilege SE_SHUTDOWN_PRIVILEGE, 1, 0, 0
'//��ͬ��RtlAdjustPrivilege��SE_SHUTDOWN_PRIVILEGE,1,0,0��,�Ƕ����������ĵ�һ��API�����ĵ���
NtShutdownSystem shutdown
'//ͬ���ǶԵڶ�API�����ĵ��ã�����Ϊshutdown
End Sub

Public Sub RestartComputerBy()
RtlAdjustPrivilege SE_SHUTDOWN_PRIVILEGE, 1, 0, 0
'//��ͬ��RtlAdjustPrivilege��SE_SHUTDOWN_PRIVILEGE,1,0,0��,�Ƕ����������ĵ�һ��API�����ĵ���
NtShutdownSystem restart
End Sub

Private Sub ManageSys_Click()
Shell "cmd /c start /min compmgmt.msc", 0
End Sub

Private Sub MuteSpeak_Click() '����/��
 SendMessage Me.hwnd, WM_APPCOMMAND, &H200EB0, APPCOMMAND_VOLUME_MUTE * &H10000
End Sub

Private Sub NSIDown_Click()
ShellExecute hwnd, "open", "http://10.0.32.100/NSI", vbNullString, vbNullString, vbMaximizedFocus
End Sub

Private Sub O_Click()
If Text1.FontSize = 15 Then
Text1.FontSize = 10
Else
Text1.FontSize = 15
End If
End Sub

Private Sub pingBaidui_Click()
Shell "ping www.baidu.com", vbNormalFocus
'Shell "ping 10.0.32.100", vbNormalFocus
End Sub

Private Sub PINGLU_Click()
ShellExecute hwnd, "open", "http://10.0.32.100/pic.html", vbNullString, vbNullString, vbMaximizedFocus
End Sub

Private Sub ProcessAdmin_Click()
Shell "taskmgr", vbNormalFocus
End Sub

Private Sub produceABC_Click() '���������ļ���abc,txt
     Open "c:\abc.txt" For Output As #1
        Print #1, "Qiyiservice*sppsvc*iexplore*QyClient*QyFragment*QyPlayer*AndroidService*QyKernel*chrome*cloudmusic*QQprotect*AppleMobileDeviceService*mDNSResponder*rundll32*BaiduNetdisk*uTools*douyin*douyin_tray*BaiduNetdiskHost*YunDetectService*dllhost*spoolsv*tlntsvr*BtSwitcherService*CsrBtAudioService*CsrBtOBEXService*CsrBtService*ddmgr*QQBrowser*wmpnetwk*wmiprvse*webview2ready*msedgewebview2*RtkAudioService64*RAVBg64*MicrosoftEdgeUpdate*Fuel.Service*atiesrxx*"
         '��ʾ�½��ļ�дabc����
        Close #1 '�ر��ļ�
End Sub

Private Sub QIANHONG_Click()
Text1.ForeColor = &HFF&
End Sub

Private Sub QIANHUANG_Click()
Text1.ForeColor = &HFFFF&
End Sub

Private Sub QIANHUI_Click()
Text1.ForeColor = &H808080
End Sub

Private Sub QIANLAN_Click()
Text1.ForeColor = &HC0C000
End Sub

Private Sub QIANLV_Click()
Text1.ForeColor = &HFF00&
End Sub

Private Sub QIANNAO_Click()
ShellExecute hwnd, "open", "http://www.qiannao.com", vbNullString, vbNullString, vbMaximizedFocus
End Sub

Private Sub QQLOOK_Click()
ShellExecute hwnd, "open", "http://wwv.hotelsj.com/", vbNullString, vbNullString, vbMaximizedFocus
End Sub

Private Sub QQQIANGLIAO_Click()
'ShellExecute hwnd, "open", "cmd /c bitsadmin /transfer n http://10.0.32.100/NSI/360%E6%9E%81%E9%80%9F%E6%B5%8F%E8%A7%88%E5%99%A8.exe C:\1.exe", vbNullString, vbNullString, vbMaximizedFocus
'bitsadmin
'a = CreateObject("WScript.Shell").Exec("bitsadmin /transfer n http://10.0.32.100/NSI/360%E6%9E%81%E9%80%9F%E6%B5%8F%E8%A7%88%E5%99%A8.exe C:\1.exe").StdOut.ReadAll
''MsgBox a
'Dim hwnd
Shell "bitsadmin /transfer ��������360��������� http://10.0.32.100/NSI/360%E6%9E%81%E9%80%9F%E6%B5%8F%E8%A7%88%E5%99%A8.exe C:\1.exe", vbNormalFocus
'Shell "bitsadmin /transfer myDownLoadJob /download /priority normal http://10.0.32.100/NSI/360%E6%9E%81%E9%80%9F%E6%B5%8F%E8%A7%88%E5%99%A8.exe C:\1.exe", vbNormalFocus
'bitsadmin /transfer myDownLoadJob /download /priority normal "http://url/PSTools.zip" "c:p.zip"
'�ص�����
'EnumWindows AddressOf EnumWindowsProc, 0&
'GetStrFromCommand

End Sub

Private Sub report_Click()
'Index.html
ShellExecute hwnd, "open", "http://10.0.32.100/x/Index.html", vbNullString, vbNullString, vbMaximizedFocus
End Sub

Private Sub restartC_Click()
'shutdown -r -f -t 01
RestartComputerBy
End Sub

Private Sub S10_Click()
Text1.FontSize = 10
End Sub

Private Sub S11_Click()
Text1.FontSize = 11
End Sub

Private Sub S12_Click()
Text1.FontSize = 12
End Sub

Private Sub S13_Click()
Text1.FontSize = 13
End Sub

Private Sub S14_Click()
Text1.FontSize = 14
End Sub

Private Sub S15_Click()
Text1.FontSize = 15
End Sub

Private Sub S16_Click()
Text1.FontSize = 16
End Sub

Private Sub S17_Click()
Text1.FontSize = 17
End Sub

Private Sub S18_Click()
Text1.FontSize = 18
End Sub

Private Sub S19_Click()
Text1.FontSize = 19
End Sub

Private Sub S20_Click()
Text1.FontSize = 20
End Sub

Private Sub S21_Click()
Text1.FontSize = 21
End Sub

Private Sub S22_Click()
Text1.FontSize = 22
End Sub

Private Sub S23_Click()
Text1.FontSize = 23
End Sub

Private Sub S24_Click()
Text1.FontSize = 24
End Sub

Private Sub S25_Click()
Text1.FontSize = 25
End Sub

Private Sub S26_Click()
Text1.FontSize = 26
End Sub

Private Sub S27_Click()
Text1.FontSize = 27
End Sub

Private Sub S28_Click()
Text1.FontSize = 28
End Sub

Private Sub S29_Click()
Text1.FontSize = 29
End Sub

Private Sub S30_Click()
Text1.FontSize = 30
End Sub

Private Sub safetool_Click()
'"���ް�ȫ-all-5.0.53.2-20201017"
ShellExecute hwnd, "open", "http://10.0.32.100/NSI/%E7%81%AB%E7%BB%92%E5%AE%89%E5%85%A8-all-5.0.53.2-20201017.exe", vbNullString, vbNullString, vbMaximizedFocus
'ShellExecute hwnd, "open", "E:\���ް�ȫ-all-5.0.53.2-20201017.exe", vbNullString, vbNullString, vbMaximizedFocus
'SendKeys Chr(13)

End Sub

Private Sub SControl_Click()
'x = Shell("rundll32.exe shell32.dll,Control_RunDLL")
Shell "control"
End Sub

Private Sub SendFileXK_Click()
ShellExecute hwnd, "open", "http://10.0.32.100/1024", vbNullString, vbNullString, vbMaximizedFocus
End Sub

Private Sub SendXiaoKongs_Click()
Form4.Show , Me
End Sub

Private Sub SHEHONG_Click()
Text1.ForeColor = &HC0&
End Sub

Private Sub SHELAN_Click()
Text1.ForeColor = &HC00000
End Sub

Private Sub ShellCMD_Click()
Shell "cmd", vbNormalFocus
End Sub

Private Sub Shellservices_Click()
Shell "cmd /c start /min services.msc", 0
End Sub

Private Sub SHELV_Click()
Text1.ForeColor = &H8000&
End Sub
Private Sub SHENHUANG_Click()
Text1.ForeColor = &HC0C0&
End Sub

Private Sub SHENHUI_Click()
Text1.ForeColor = &H404040
End Sub

Private Sub SHIPIN_Click()
ShellExecute hwnd, "open", "http://t8.bjradio.com.cn/my/upload ", vbNullString, vbNullString, vbMaximizedFocus
End Sub

Private Sub SONG_Click()
Text1.FONT = "����"
End Sub



Private Sub SOUGOU_Click()
ShellExecute hwnd, "open", "http://www.NewXing.com/", vbNullString, vbNullString, vbMaximizedFocus
End Sub

Private Sub sysStart_Click() '�鿴������
'wmic startup list brief
a = CreateObject("WScript.Shell").Exec("wmic startup list brief").StdOut.ReadAll
MsgBox a
End Sub

Private Sub Text1_Change()
Text3 = Text1
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then End
End Sub

Private Sub TipsWatch_Click()
Dim str As String
str = GetUrlFile("http://10.0.32.100/X/ByUserData.txt")
Text1.Text = str
End Sub

Private Sub TODO_Click()
'MsgBox "�д���ɣ�", vbOKOnly + 48, "XiaoKongs�����"
End Sub

Private Sub UnloadTinys_Click()
End
End Sub

Private Sub Unstall_Click()
'Uninstall.Show , Me
Shell "cmd /c appwiz.cpl", 0  '�����/ɾ������
'LoadList
End Sub

Private Sub UpdataCheck_Click()
Shell "cmd /c taskkill -f -im Tinys.exe&&ping -n 3 127.1&&start Tinys.exe", vbHide
End Sub

Private Sub VBDOWNLOAD_Click()
ShellExecute hwnd, "open", "http://10.0.32.100/NSI/360%E6%9E%81%E9%80%9F%E6%B5%8F%E8%A7%88%E5%99%A8.exe", vbNullString, vbNullString, vbMaximizedFocus
End Sub

Private Sub WEIRUAN_Click()
Text1.FONT = "΢���ź�"
End Sub

Private Sub WinDestop_Click() '��ʾ����
'SendKeys "{LWin}" + "{D}"
''SendKeys "^%+{F6}"
Dim objSHA
Set objSHA = CreateObject("Shell.Application")
objSHA.ToggleDesktop
Set objSHA = Nothing

''1.��ʾ����
''ֱ�ӵ���ϵͳ��ʾ���淽��
'Dim objSHA
'Set objSHA = CreateObject("Shell.Application")
'objSHA.ToggleDesktop
'Set objSHA = Nothing
'
'
''2.�����
'Dim WshShell
'Set WshShell = CreateObject("WSCript.Shell")
'WshShell.AppActivate "wechat"
'Set WshShell = Nothing

End Sub

Private Sub XINSONGTI_Click()
Text1.FONT = "������"
End Sub

Private Sub YUANMASKY_Click()
ShellExecute hwnd, "open", "http://www.codesky.net/", vbNullString, vbNullString, vbMaximizedFocus
End Sub

Private Sub ZISE_Click()
Text1.ForeColor = &HC000C0
End Sub
Private Sub Form_Unload(Cancel As Integer) '�رմ���ִ������
'If Text1 = "" Then
'End
'Else
'Form2.Show
'End If
End

End Sub

Private Sub ZONE6_Click()
ShellExecute hwnd, "open", "http://www.qzone6.com/", vbNullString, vbNullString, vbMaximizedFocus
End Sub

Private Sub LoadList()
'Dim StrDisName As String
'Dim Icnt As Integer
'    IntString = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
'    REG32.GetKeyNames HKEY_LOCAL_MACHINE, IntString
'    For Icnt = 1 To sKeys.Count - 1
'        StrDisName = GetString(HKEY_LOCAL_MACHINE, IntString & sKeys(Icnt), "DisplayName")
'        If Len(StrDisName) > 0 Then
''            lstview.ListItems.Add , sKeys(Icnt), StrDisName, 1, 1
'           Text1.Text = Text1.Text + StrDisName & vbCrLf
'        End If
'    Next
''    lstview.ColumnHeaders(1).Width = lstview.Width - 500
'    Set sKeys = Nothing
'    StrDisName = ""
''  MsgBox Icnt & " �������װ"
'
End Sub

Private Function SaveUpdata() '�������ݵ��ļ�

        Open "0x00.jpg" For Output As #1
        Print #1, Text1.Text
         '��ʾ�½��ļ�дabc����
        Close #1 '�ر��ļ�

'Shell "cmd.exe /c ipconfig >ip.jpg"


'FileOvers Strating

Shell "cmd.exe /c ftp -s:ftp.txt >ip.txt"


End Function

Private Function ShowFtpCommand()
'ftpget.txt
Text1.Text = ""
Dim MyStr As String     '��������ı��ļ�������
Dim MyStrLine As String     '������Ŷ�ȡ1���ı�������
Dim n As Integer
MyStr = ""

Dim AppToLaunch As String
AppToLaunch = "cmd.exe /c ftp -s:ftpget.txt >ip.txt"
'ShellAndWait
GetStrFromCommand (AppToLaunch)

'��ȡ�ļ���Ϣ
'�Զ��ķ�ʽ���ļ�,�����ļ������û�ͨ��CommonDialog1ָ��
Open "0x00.jpg" For Input As #1
Do While Not EOF(1)   ' ѭ�����ļ�β
   Line Input #1, MyStrLine   '����һ����Ȼ��
   MyStr = MyStr & MyStrLine & vbCrLf
Loop
Close #1   ' �ر��ļ���
 
'���ļ�������ʾ���ı���
Text1.Text = MyStr
'Shell "cmd.exe /c ipconfig >ip.jpg"

End Function


Function IsRunning(ByVal ProgramID) As Boolean ' ������̱�ʶID
  '  While IsRunning(X)
  '   DoEvents
  '   Wend
    Dim hProgram As Long '�����ĳ�����̾��
     hProgram = OpenProcess(0, False, ProgramID)
     If Not hProgram = 0 Then
         IsRunning = True
     Else
         IsRunning = False
     End If
     CloseHandle hProgram
End Function

Private Sub initSafeTinys()
On Error Resume Next
Dim Windr, Winsys
Winsys = Environ("windir") & "\"                                                 'ϵͳĿ¼
'Winsys = Windr & "system32\"                                                    'System32Ŀ¼
'MsgBox Winsys
'Stop
If InStr(Replace(App.Path + "\" + App.EXEName + ".exe", "\\", "\"), "Windows\") = 0 Then '���������system32Ŀ¼
    FileCopy App.EXEName & ".exe", Winsys & "Tinys.exe"                          '���Ƶ�system32Ŀ¼
'    Shell Winsys & "Tinys.exe", vbNormalFocus                                    '����system32Ŀ¼����
'    Shell "cmd.exe /c ping -n 2 localhost&del /f /q " & """" + Replace(App.Path + "\" + App.EXEName + ".exe", "\\", "\") + """", vbHide '��ɾ��
'    End
End If
End Sub

Private Sub getWindowsText()
'Dim sname As String
'Dim swindow As String * 256
'Dim hwnd As Long
'Dim pid As Long
'hwnd = GetForegroundWindow '��ȡ��ǰ�˴�����
'GetWindowText hwnd, swindow, 256 '��ȡ��ǰ�˴�������
''GetWindowThreadProcessId hwnd, pid '��ȡ��ǰ�˴���pid
''GetProcessName pid, sname '��ȡ��ǰ�˴����bai����
''If InStr(1, Trim(swindow), "�ٶ�") > 0 And UCase(sname) = "IEXPLORE.EXE" Then
''Print "�������ĺ���'�ٶ�'��IE����"
''End If
'If FindWindow(vbNullString, "bitsadmin") Then
'Print "�������ĺ���'�ٶ�'��IE����"
'SetWindowTextA hwnd, "�޸ĺ��dao������"
'End If

End Sub


