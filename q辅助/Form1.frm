VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ô��ھ��"
   ClientHeight    =   2265
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   3090
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   3090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin Vb����1.Spy�ؼ� Spy�ؼ�1 
      Height          =   975
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
   End
   Begin VB.CommandButton Command9 
      Caption         =   "���"
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "����ˮ��"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton command7 
      Caption         =   "XIAOKONGS"
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "���ھ��"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "������"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "��������"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����·��"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��ȡ���̾��"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ȡ����PIDֵ"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "˵�����϶�ͼ�굽QQ�����ͷ�"
      Height          =   180
      Left            =   360
      TabIndex        =   9
      Top             =   1320
      Width           =   2340
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "��ʾ����������϶�����ͼ�굽Ŀ�괰���ϼ��ɻ�ȡ����Ӧ�������ˣ������Ļ�����www.51xue8xue8.com �����������������ۡ�"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�ؼ���: spy�ؼ�
'��  Ȩ: www.51xue8xue8.com
'�� �� :ֱ��[��ȡ����pid,���̾��,��������,·��,������]
'ʹ��˵����:
'��һ��:��VB���̼������spy�ؼ�
'�ڶ���:��VB����Ϳ��Ե����������
'spy�ؼ�1.����Pid
'spy�ؼ�1.���̾��
'spy�ؼ�1.���ھ��
'spy�ؼ�1.��������
'spy�ؼ�1.����·��


Private Sub Command1_Click()
 MsgBox Spy�ؼ�1.����Pid

End Sub

Private Sub Command2_Click()
 MsgBox Spy�ؼ�1.���̾��
End Sub

Private Sub Command3_Click()
 MsgBox Spy�ؼ�1.����·��
End Sub


Private Sub Command4_Click()
MsgBox Spy�ؼ�1.��������
End Sub

Private Sub Command5_Click()
MsgBox Spy�ؼ�1.������
End Sub

Private Sub Command6_Click()
MsgBox Spy�ؼ�1.���ھ��
End Sub

Private Sub command7_Click()
XiaoKongs_HWND = Spy�ؼ�1.���ھ��
Form2.Command1.Tag = XiaoKongs_HWND
End Sub

Private Sub Command8_Click()
yueHWND = Spy�ؼ�1.���ھ��
Form2.Command2.Tag = yueHWND
End Sub

Private Sub Command9_Click()
Form2.Show
Unload Me
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
    MsgBox "���Ѿ������˱������", vbOKOnly + 48, "����"
    End
End If
Load Form2

End Sub

