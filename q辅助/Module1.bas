Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function GetForegroundWindow Lib "user32" () As Long


Public XiaoKongs_HWND As Long
Public yueHWND As Long

'��һ�δ򿪴���ʱ��ס���ھ��
Public Function Jizhu() As Long
Jizhu = GetForegroundWindow() '�õ�����ڵľ��
End Function

'��ⴰ��״̬�Ƿ�����
'    If IsWindowVisible(131668) Then
'         Label2.Caption = "��"
'         Else
'         Label2.Caption = "�ر�"
'    End If

