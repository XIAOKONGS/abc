Attribute VB_Name = "process"

'firewall++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Type MIB_TCPROW_OWNER_PID ''���ǵ�TCP_TABLE_CLASS��TCP_TABLE_OWNER_PID_ALL,��GetExtendedTcpTable �õ�,
    dwState As Long
    dwLocalAddr As Long
    dwLocalPort As Long
    dwRemoteAddr As Long
    dwRemotePort As Long
    dwOwningPid As Long
End Type
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal Handle As Long) As Long
Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, _
                                                        ByVal bInheritHandle As Long, _
                                                        ByVal dwProcId As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, _
                                                        ByVal hModule As Long, _
                                                        ByVal ModuleName As String, _
                                                        ByVal nSize As Long) As Long
Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, _
                                                        ByRef lphModule As Long, _
                                                        ByVal cb As Long, _
                                                        ByRef cbNeeded As Long) As Long
                                                        
                                                        
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
Private Const MAX_PATH = 256
Private Const AF_INET6 = 23
Private Const AF_INET = 2
Public Enum TCP_TABLE_CLASS
  TCP_TABLE_BASIC_LISTENER
  TCP_TABLE_BASIC_CONNECTIONS
  TCP_TABLE_BASIC_ALL
  TCP_TABLE_OWNER_PID_LISTENER
  TCP_TABLE_OWNER_PID_CONNECTIONS
  TCP_TABLE_OWNER_PID_ALL
  TCP_TABLE_OWNER_MODULE_LISTENER
  TCP_TABLE_OWNER_MODULE_CONNECTIONS
  TCP_TABLE_OWNER_MODULE_ALL
End Enum
Private Declare Function htons Lib "ws2_32.dll" (ByVal dwLong As Long) As Long
Public Declare Function GetExtendedTcpTable Lib "IPHLPAPI.DLL" (pTcpTableEx As Any, lSize As Long, ByVal bOrder As Long, ByVal Flags As Long, ByVal TableClass As TCP_TABLE_CLASS, ByVal bReserved As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private pTablePtr() As Byte
Public nRows As Long
Private pDataRef As Long
Public sss As String

Public Declare Function SetTcpEntry Lib "IPHLPAPI.DLL" (ByRef pTcpTable _
As MIB_TCPROW_OWNER_PID) As Long

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long

'Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long 'API�ж�����Ϊ�ջ�û�г�ʼ��

Private Const SND_ASYNC = &H1 '�첽���ţ�����Ͷ�ռ����

Private Const SND_NODEFAULT = &H2 '��ʹ��ȱʡ����

Private Const SND_MEMORY = &H4 'ָ��һ���ڴ��ļ�

Private Const SND_FILENAME = &H20000 'ָ��һ��ʵ���ļ�

Private Const SND_LOOP = &H8 'ѭ������

Private Const SND_ALIAS_START = 0 '��������

Dim b() As Byte
'firewall++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'*functiondevider*
Public Sub RefreshStack()
On Error Resume Next
Dim i As Long
Dim tcpTable As MIB_TCPROW_OWNER_PID
    pDataRef = 0
    
Dim k As Integer
Dim spfwdatp1() As String
Dim astrgfir As String
spfwdatp1() = Split(sss, "*")
For i = 0 To nRows ' read 24 bytes at a time
'    astrgfir = ""
'    CopyMemory tcpTable, pTablePtr(0 + pDataRef + 4), LenB(tcpTable)
'
'        If tcpTable.dwRemoteAddr <> 0 Or GetPort(tcpTable.dwRemotePort) <> 0 Or GetPort(tcpTable.dwLocalPort) <> 0 Then
'            Debug.Print "״̬:"; c_state(tcpTable.dwState); ",";
'            Debug.Print "����IP:"; GetIPAddress(tcpTable.dwLocalAddr); ",";
'            Debug.Print "����PORT:"; GetPort(tcpTable.dwLocalPort); ",";
'            Debug.Print "Զ��IP:"; tcpTable.dwRemoteAddr; ",";
'            Debug.Print "Զ��PORT:"; GetPort(tcpTable.dwRemotePort); ",";
'            Debug.Print "����ID:"; tcpTable.dwOwningPid; ",";
'            Debug.Print "������:"; getPidPathName(tcpTable.dwOwningPid)
'            astrgfir = GetIPAddress(tcpTable.dwRemoteAddr) & getPidPathName(tcpTable.dwOwningPid)
            
            For k = 0 To UBound(spfwdatp1()) - 1
'                If InStr(LCase(astrgfir), LCase(spfwdatp1(k))) > 0 Then
'                   tcpTable.dwState = 12
'                   SetTcpEntry tcpTable
                    CloseProcess LCase(spfwdatp1(k))
'                End If
            Next k
'        End If
'        pDataRef = pDataRef + LenB(tcpTable)
        DoEvents
Next i
MsgBox "����ϵͳ�Ż���ϣ�", vbOKOnly + 48, "XIAOKONGS�����"
End Sub

Public Sub CloseProcess(process As String)
On Error Resume Next
Dim s
s = process & ".exe"
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

's = "mir3.exe"
Set colProcessList = objWMIService.ExecQuery _
("Select * from Win32_Process Where Name='" & s & "'")
For Each objProcess In colProcessList
objProcess.Terminate '��������
Next

Set objProcess = Nothing
Set colProcessList = Nothing
Set objWMIService = Nothing
End Sub

Public Sub beep()
    b = LoadResData(101, "WAVE")
'    IniArray = SafeArrayGetDim(B)
    sndPlaySound b(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY 'Or SND_LOOP
End Sub


