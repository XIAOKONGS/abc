VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "ONLINE 3389"
   ClientHeight    =   1860
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   1860
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
On Error GoTo err
Shell "cmd.exe /c mstsc /v" & " " & "14.103.51.243" & ":" & "3389" & " /console -f", 0
Me.Hide
DoEvents
End
err:
If Error <> "" Then: MsgBox "连接时出现错误：" & Error, 16
End Sub
