VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1605
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   1605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "����ˮ��"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "XIAOKONGS"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "XIAOKONGS�����"
      ForeColor       =   &H00400000&
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1350
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal Scan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Const KEYEVENTF_KEYUP = &H2 '�ͷŰ�������
Dim BolIsMove As Boolean, MousX As Long, MousY As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long '�жϴ���״̬
'--------------����͸��+
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
' �����������б�������λ���κ�������ڵ�ǰ��
Private Const SWP_NOSIZE& = &H1
' ���ִ��ڴ�С
Private Const SWP_NOMOVE& = &H2
' ���ִ���λ��

Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Const SW_HIDE = 0

Private Sub Command2_Click()
On Error Resume Next
    If IsWindowVisible(Command2.Tag) Then
         ShowWindow Command2.Tag, 0
         Else
         ShowWindow Command2.Tag, 1
    End If
End Sub

Private Sub Form_Load()

'ǿ�Ʊ�������ǰ
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

End Sub
'--------------����͸��-

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then BolIsMove = True
MousX = x
MousY = y
''����PopupMenu����
'  If Button And vbRightButton Then
'     PopupMenu wj    '�����˵�
'  End If
End Sub
 
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim CurrX As Long, CurrY As Long
If BolIsMove Then
 CurrX = Me.Left - MousX + x
 CurrY = Me.Top - MousY + y
 Me.Move CurrX, CurrY
End If
End Sub
 
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
BolIsMove = False
End Sub

'�س�������
Private Sub Command1_Click()
On Error Resume Next
    If IsWindowVisible(Command1.Tag) Then
         ShowWindow Command1.Tag, 0
         Else
         ShowWindow Command1.Tag, 1
    End If
End Sub

Private Sub Label1_Click()
End
End Sub


