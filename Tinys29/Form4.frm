VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "向 XiaoKongs 发送反馈"
   ClientHeight    =   3480
   ClientLeft      =   5985
   ClientTop       =   4740
   ClientWidth     =   5655
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5655
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "发送"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404040&
      Caption         =   "无论对方是否为好友,输入QQ即可聊天! 只有QQ2008能使用此功能!"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   735
      Left            =   720
      TabIndex        =   3
      Top             =   3840
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "对方QQ: "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   4200
      Width           =   1575
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Command2_Click()
'Unload Form4
'End Sub
''////下面是所谓的QQ强聊代码,需要建立一个Microsoft Internet Controls控件，若打开源码出错，请重新建立此空间控件
'Private Sub Command1_Click()
'If Len(Text1) >= 12 Or Len(Text1) < 4 Then '通过判断输入字符长度来判断输入正误
'MsgBox "您输入的QQ号有误，请重新输入！", vbOKOnly, "错误提示"
'Else
'WebBrowser1.Navigate "Tencent://Message/?Menu=YES&Exe=&Uin=" & Text1.Text
'WebBrowser1.Stop
'End If
'End Sub
'
'Private Sub Form_Load()
'
'End Sub

Private Sub Command1_Click()
Dim str As String
If Text1.Text = "" Then Exit Sub
str = "http://10.0.32.100/x/SendXiaoKongs.asp?content=" & Text1.Text
str = SendMSG(str)
Text1.Text = ""
Me.Hide
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
'MsgBox GetUrlFile("http://10.0.32.100/NSI/%E9%9A%8F%E7%9D%80%E6%80%9D%E8%B7%AF%E7%9A%84%E8%BF%9B%E5%B1%95%E4%B8%8D%E6%96%AD%E6%B7%BB%E5%8A%A0%E7%BB%86%E8%8A%82.txt")
End Sub
