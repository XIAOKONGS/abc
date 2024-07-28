VERSION 5.00
Begin VB.Form TestC 
   Caption         =   "Form6"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8025
   LinkTopic       =   "Form6"
   ScaleHeight     =   7305
   ScaleWidth      =   8025
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command8 
      Caption         =   "格式化所有"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton command9 
      Caption         =   "分析末尾"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "去掉末尾空行"
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   5400
      Width           =   7695
   End
   Begin VB.CommandButton CommandGoText 
      Caption         =   "导入文本框"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "转换"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "分析文本框"
      Height          =   255
      Left            =   6360
      TabIndex        =   3
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "分析文本行数"
      Height          =   255
      Left            =   6360
      TabIndex        =   2
      Top             =   4560
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   4380
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Height          =   4335
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "TestC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Readtxt()

Dim Filepath As String


Filepath = App.Path & "\" & "tinys.ini"
If Dir(Filepath) = "" Then
MsgBox "未找到配置文件"
Else
Dim a$
    Open Filepath For Input As #1
    Do
        Input #1, a
        
'        If Trim(a) <> "" Then List1.AddItem a
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

Private Sub Command10_Click()

    Dim i As Integer
    For i = List1.ListCount - 1 To 0 Step -1
        If List1.List(i) = "" Then
        List1.RemoveItem i
        Else
        List1.List(i) = Trim(List1.List(i))
        List1.Tag = i
        Exit For
        End If
    Next i
    
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

                If List1.List(c) = "====================================================" Then
                    List1.RemoveItem c
                Else
'                List1.AddItem a
                End If
      
            Next c
            
End Sub


Public Function addString()

    If List1.List(List1.ListCount) = "====================================================" Then
    Else
    List1.List(List1.ListCount) = "===================================================="
    End If
    
        If List1.List(List1.ListCount - 1) = "====================================================" Then
         List1.RemoveItem List1.ListCount - 1
        End If
    
    
End Function


Private Sub Command8_Click()
 Dim i As Integer
    For i = List1.ListCount - 1 To 0 Step -1
        If List1.List(i) = "" Then
        List1.RemoveItem i
        Else
        List1.List(i) = Trim(List1.List(i))
        List1.Tag = i
        End If
    Next i
End Sub

Private Sub command9_Click()

    Dim i As Integer
    For i = List1.ListCount - 1 To 0 Step -1
        If List1.List(i) = "" Then
        List1.RemoveItem i
        Else
        List1.List(i) = Trim("[文件末尾>>]:" & List1.List(i))
        List1.Tag = i
        Exit For
        End If
    Next i
    
  
End Sub

Private Sub CommandGoText_Click()

  Dim i As Long
  Dim m As Long
  Dim W As String

  For i = 0 To List1.ListCount - 1
  
           If i <> List1.Tag Then
            Text2 = Text2 & List1.List(i) & vbCrLf
           Else
            Text2 = Trim(Text2 & List1.List(i))
           End If
  Next i
           
End Sub

Private Sub Form_Load()
Readtxt
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
                    List1.List(i) = "===================================================="
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
                    List1.List(i) = "===================================================="
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

