VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000006&
   BorderStyle     =   0  'None
   Caption         =   "Connecting"
   ClientHeight    =   2565
   ClientLeft      =   5325
   ClientTop       =   5445
   ClientWidth     =   4290
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command3 
      Caption         =   "��ʼ����"
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "д��"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      Height          =   615
      Left            =   2160
      Picture         =   "Form2.frx":0CCA
      ScaleHeight     =   555
      ScaleWidth      =   1035
      TabIndex        =   4
      Top             =   4200
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   720
      Picture         =   "Form2.frx":12A0
      ScaleHeight     =   435
      ScaleWidth      =   1275
      TabIndex        =   3
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Text            =   "192.168.3.200 1080"
      Top             =   810
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      Caption         =   "ϵͳ�Ż�"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3240
      TabIndex        =   5
      Top             =   480
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      Caption         =   " ���ӷ�����IP��ַ��"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1710
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Download by http://www.NewXing.com


Option Explicit

Dim WithEvents objFtpClient As FTP
Attribute objFtpClient.VB_VarHelpID = -1

'Private Sub Command1_Click()
'MsgBox objFtpClient.UpFile("c:\1.jpg", "/SampleTest.jpg")
'End Sub
Dim nport As Integer
Dim nIP As String

Dim Filepath As String


Public Function ToSaveFile()

'If Dir(Filepath) = "" Then

      Open Filepath For Output As #1
        Print #1, Text1.Text
         '��ʾ�½��ļ�дText1.Text����
        Close #1 '�ر��ļ�
'        End If
  
End Function


Private Function ToReadFile()

If Dir(Filepath) = "" Then
MsgBox "δ�ҵ������ļ�"
Else
Dim a$
    Open Filepath For Input As #1
    Do
        Input #1, a
        sss = sss & a & vbCrLf
        Loop Until EOF(1)
    Close #1
    Text1.Text = sss
'        Call RefreshStack
    End If
    
End Function



Private Sub Form_Activate()
Label1.Caption = "" 'NULL
Command3_Click
End Sub

Private Sub Form_Load()

    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
    
    Filepath = App.Path & "\" & "tinys.ini"
    Call ToReadFile '��ȡ�����ļ��еĵ�ַ
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    objFtpClient.Logout
'    Set objFtpClient = Nothing
End Sub

Private Sub Label2_Click()
  sss = ""
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

Private Sub Label3_Click()
    
    Label1.Caption = "Checking..."
    DoEvents
    Sleep (500)
    
    Dim i As Long, s() As String
    Dim pos As Integer
    
    pos = Len(Text1) - InStr(Text1.Text, " ") 'pos nport����
    
    nport = Right(Text1.Text, Len(Text1) - InStr(Text1.Text, " "))
    nIP = Left(Text1.Text, Len(Text1) - pos - 1)
    
'    MsgBox nIP & ":" & nport

    Set objFtpClient = New FTP
    If objFtpClient.Login(nIP, nport) Then
    Label1.Caption = "�ѳɹ�������������"
    Else
    Label1.Caption = "û���ҵ�������"
    End If
    
     objFtpClient.Logout
    Set objFtpClient = Nothing

'    objFtpClient.EnumFile "/", True

End Sub

'Private Sub objFtpClient_EnumFileProc(FileName As String, Attr As VbFileAttribute, Size As Long, Create As String, Modify As String, Cancel As Boolean)
'    If (Attr Or vbDirectory) = Attr Then
'        Debug.Print Format(Modify, "yyyy-mm-dd hh:nn:ss"), "<DIR>", , FileName
''        XIAOKONGS��׼�������Ǻ���ӵ��������˵���/��
'        If Left(FileName, 1) = "/" Then FileName = Right(FileName, Len(FileName) - 1)
'        List1.AddItem "+" & FileName
'    Else
'        Debug.Print Format(Modify, "yyyy-mm-dd hh:nn:ss"), , Size, FileName
'        List1.AddItem FileName
'    End If
'End Sub


Private Sub Command1_Click()
If Text1 = "" Then
Label1 = "�ļ�������Ϊ�գ��������ļ���"
Else
'Open Text1 + ".TXT" For Output As #1 '�����д·����ֻд�ļ�����������ֱ������
'Print #1, Form1.Text1
'Close #1
'MsgBox "�Ѿ����浽��Ŀ¼��,�ټ�!", vbOKOnly, "����ɹ�"
'End



    Dim i As Long, s() As String
    Dim pos As Integer
    
    pos = Len(Text1) - InStr(Text1.Text, " ") 'pos nport����
    
    nport = Right(Text1.Text, Len(Text1) - InStr(Text1.Text, " "))
    nIP = Left(Text1.Text, Len(Text1) - pos - 1)
    
'    MsgBox nIP & ":" & nport
'-----------------------------------------------
'���������ַ���
Form3.Text1.Text = Text1.Text  '���Ŀ���ַ
Call StrConnect 'ehco
'-----------------------------------------------
'Dim i As Long, s() As String
'    Set objFtpClient = New FTP
''    MsgBox objFtpClient.Login(Right(Text1.Text, 4))
'
''    objFtpClient.EnumFile "/", True
'-----------------------------------------------
End If
Call ToSaveFile '�������ļ�
Me.Hide
End Sub

Private Sub Command2_Click()
'End
Unload Me

End Sub

Private Sub Command3_Click()
Label3_Click
End Sub

Private Sub image1_Click() '�رհ�ť
Form1.Show
Unload Form2
End Sub

'/////������ȫѡ���ļ���(Text1)///////////////////////

'/////�����ǹرհ�ťЧ��,����ƶ��¼�,����ͼƬ�����л�///////////////////////

'Private Sub image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Image1 = Picture1
'End Sub
'Private Sub image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Image1 = Picture2
'End Sub

