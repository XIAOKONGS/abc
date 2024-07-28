VERSION 5.00
Begin VB.Form frm_Help 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "帮助"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   10.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   8190
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   5520
      Width           =   255
   End
   Begin VB.ListBox List1 
      Height          =   300
      Left            =   360
      TabIndex        =   2
      Top             =   5520
      Width           =   615
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6720
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   8175
   End
End
Attribute VB_Name = "frm_Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
        Unload Me
        Unload Form3
End Sub

Private Sub Form_Load()
        Me.Icon = Form1.Icon
        Readtxt
        Command3_Click
        
    CommandGoText_Click
    
End Sub


'Private Sub ReadGOING()
'
'Dim p As String
'Dim sss1 As String
'
'p = App.Path & "\" & "ip.txt"
'
'If Dir(p) = "" Then
'    Text1.Text = "未找到配置文件"
'Else
'        Dim b$
'        Open p For Input As #1
'        Do
'            Input #1, b
'            sss1 = sss1 & b & vbCrLf
'            Loop Until EOF(1)
'        Close #1
'        Text1.Text = sss1
'End If
'
'End Sub



Public Sub Readtxt()

Dim Filepath As String


Filepath = App.Path & "\" & "ip.txt"
If Dir(Filepath) = "" Then
MsgBox "未找到配置文件"
Else
Dim a$
    Open Filepath For Input As #1
    Do
        Input #1, a
        List1.AddItem a
        sss = sss & a & vbCrLf
        Loop Until EOF(1)
    Close #1
    Text1.Text = sss
'        Call RefreshStack
    End If
End Sub

Private Sub Command1_Click()
MsgBox read
End Sub

Private Sub Command2_Click() '分析文本框


  Dim s() As String '定义数组
  Dim i As Integer
  
  s = Split(Text1, vbCrLf)

  i = UBound(s)  '理想化UBound(s)+1为虚拟量
'  r = StrConv(InputB(LOF(1), 1), vbUnicode)
 MsgBox i
'
End Sub



Private Sub Command4_Click()

        List1.Refresh
        List1.RemoveItem List1.ListCount - 1
        
        Dim c As Long
        For c = 0 To List1.ListCount

                If List1.List(c) = "===============================================" Then
                    List1.RemoveItem c
                Else
'                List1.AddItem a
                End If
      
            Next c
            
End Sub


Public Function addString()

    If List1.List(List1.ListCount) = "===============================================" Then
    Else
    List1.List(List1.ListCount) = "==============================================="
    End If
    
        If List1.List(List1.ListCount - 1) = "===============================================" Then
         List1.RemoveItem List1.ListCount - 1
        End If
    
    
End Function


Private Sub CommandGoText_Click()

  Dim i As Long
  Dim m As Long
  Dim W As String

  For i = 0 To List1.ListCount - 1

'           MsgBox List1.List(i)
           Text3 = Text3 & List1.List(i) & vbCrLf
  Next i
  
'  "                                 "
  
  Text3.Text = "命令已执行完毕:" & vbCrLf & "===============================================" & vbCrLf & Text3.Text & "===============================================" & "                                 " & Now
           
End Sub



Public Function read() As String

  Dim r
  Dim Filepath As String
  Filepath = App.Path & "\" & "tinys.ini"
  Dim s() As String
  
  Open Filepath For Binary As #1
  s = Split(Input$(LOF(1), 1), vbCrLf)

  read = UBound(s) + 1 & "<行数"
'  r = StrConv(InputB(LOF(1), 1), vbUnicode)
  Close #1
'
'  Text1 = r
End Function

Private Sub Command3_Click()
  Dim i As Long
  Dim m As Long

  For i = 0 To List1.ListCount
'  MsgBox List1.List(i) & i
                If List1.List(i) = "" Then
                    List1.List(i) = "==============================================="
                Else
'                List1.AddItem a
                End If
      
            Next i
            
            
'     For m = 0 To List1.ListCount
'
'                If List1.List(m) = "==============" Then
''                    List1.RemoveItem m
'                Else
''                List1.AddItem a
'                End If
'
'            Next m
            
     Command4_Click
     
       For i = 0 To List1.ListCount
'  MsgBox List1.List(i) & i
                If List1.List(i) = "" Then
                    List1.List(i) = "==============================================="
                Else
'                List1.AddItem a
                End If
      
            Next i
            
            
'     For m = 0 To List1.ListCount
'
'                If List1.List(m) = "==============" Then
''                    List1.RemoveItem m
'                Else
''                List1.AddItem a
'                End If
'
'            Next m
            
     Command4_Click
     
     Call addString
     

End Sub


Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then cmdClose_Click
End Sub
