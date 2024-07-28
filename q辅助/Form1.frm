VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "获得窗口句柄"
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
   StartUpPosition =   3  '窗口缺省
   Begin Vb工程1.Spy控件 Spy控件1 
      Height          =   975
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
   End
   Begin VB.CommandButton Command9 
      Caption         =   "完成"
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "镜花水月"
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
      Caption         =   "窗口句柄"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "窗口名"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "窗口类名"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "进程路径"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "获取进程句柄"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "获取进程PID值"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "说明：拖动图标到QQ界面释放"
      Height          =   180
      Left            =   360
      TabIndex        =   9
      Top             =   1320
      Width           =   2340
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "提示：请用鼠标拖动以上图标到目标窗口上即可获取到对应的数据了，不懂的话请来www.51xue8xue8.com 交流交流，讨论讨论。"
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
'控件名: spy控件
'版  权: www.51xue8xue8.com
'功 能 :直观[获取进程pid,进程句柄,窗口类名,路径,窗口名]
'使用说明书:
'第一步:在VB工程加载这个spy控件
'第二步:在VB里面就可以调用以下语句
'spy控件1.进程Pid
'spy控件1.进程句柄
'spy控件1.窗口句柄
'spy控件1.窗口类名
'spy控件1.进程路径


Private Sub Command1_Click()
 MsgBox Spy控件1.进程Pid

End Sub

Private Sub Command2_Click()
 MsgBox Spy控件1.进程句柄
End Sub

Private Sub Command3_Click()
 MsgBox Spy控件1.进程路径
End Sub


Private Sub Command4_Click()
MsgBox Spy控件1.窗口类名
End Sub

Private Sub Command5_Click()
MsgBox Spy控件1.窗口名
End Sub

Private Sub Command6_Click()
MsgBox Spy控件1.窗口句柄
End Sub

Private Sub command7_Click()
XiaoKongs_HWND = Spy控件1.窗口句柄
Form2.Command1.Tag = XiaoKongs_HWND
End Sub

Private Sub Command8_Click()
yueHWND = Spy控件1.窗口句柄
Form2.Command2.Tag = yueHWND
End Sub

Private Sub Command9_Click()
Form2.Show
Unload Me
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
    MsgBox "你已经启动了本插件！", vbOKOnly + 48, "警告"
    End
End If
Load Form2

End Sub

