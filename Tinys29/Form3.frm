VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "关于 Tinys"
   ClientHeight    =   2400
   ClientLeft      =   9720
   ClientTop       =   5565
   ClientWidth     =   4230
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4230
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   120
      TabIndex        =   5
      Text            =   "192.168.3.200 1080"
      Top             =   3000
      Width           =   4095
   End
   Begin VB.TextBox Text3 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "Form3.frx":0CCA
      Top             =   5160
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form3.frx":0CEA
      Top             =   3720
      Width           =   4095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FFFF&
      Height          =   255
      Left            =   360
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "XIAOKONGS by 10.28"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   2640
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Form3
Form1.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Show
End Sub

Private Sub Picture1_Click()
frm_Help.Show
End Sub

