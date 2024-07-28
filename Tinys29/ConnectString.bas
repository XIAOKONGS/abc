Attribute VB_Name = "ConnectString"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public StrConnectUp As String
Public StrConnectDown As String
Public AesDebug As String

Public Sub StrConnect()

   Form3.Text1.Text = Form2.Text1.Text
   StrConnectUp = "open " & Form3.Text1.Text + vbCrLf + Form3.Text2.Text
   StrConnectDown = "open " & Form3.Text1.Text + vbCrLf + Form3.Text3.Text
   
   AesDebug = "open " & Form3.Text1.Text + vbCrLf + Form3.Text2.Text
   

End Sub

Public Function eString(SQL As String) As String 'todo待测试
  Dim s() As String '定义数组
  Dim i As Integer
  Dim k As Integer
  
  Dim e As String
  
  Form1.List1.Clear
  
  s = Split(SQL, vbCrLf)

  i = UBound(s)  '理想化UBound(s)+1为虚拟量
 
 For k = 0 To UBound(s()) - 1
        Form1.List1.AddItem Trim(s(k))
 Next k
   
    Dim p As Integer
    For p = Form1.List1.ListCount - 1 To 0 Step -1
        If Form1.List1.List(p) = "" Then
        Form1.List1.RemoveItem p
        Else
        Form1.List1.List(p) = Trim(Form1.List1.List(p))
        End If
    Next p
    
    Dim p1 As Integer
    For p1 = Form1.List1.ListCount - 1 To 0 Step -1
        If Form1.List1.List(p1) = "" Then
        Form1.List1.RemoveItem p1
        Else
        Form1.List1.List(p1) = Trim(Form1.List1.List(p1))
        Form1.List1.Tag = p1
        Exit For
        End If
    Next p1
    

     
  Dim m As Long

  For m = 0 To Form1.List1.ListCount - 1
  
           If m <> Form1.List1.Tag Then
            e = e & Form1.List1.List(m) & vbCrLf
           Else
            e = Trim(e & Form1.List1.List(m))

           End If
  Next m
  
  MsgBox e
  
End Function
