Attribute VB_Name = "ShellCode"
Option Explicit
Option Base 0
'Powered by barenx
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, _
 ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, _
 ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDirectory As String, _
 lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Private Type STARTUPINFO
 cb As Long
 lpReserved As String
 lpDesktop As String
 lpTitle As String
 dwX As Long
 dwY As Long
 dwXSize As Long
 dwYSize As Long
 dwXCountChars As Long
 dwYCountChars As Long
 dwFillAttribute As Long
 dwFlags As Long
 wShowWindow As Integer
 cbReserved2 As Integer
 lpReserved2 As Long
 hStdInput As Long
 hStdOutput As Long
 hStdError As Long
End Type
Private Type PROCESS_INFORMATION
 hProcess As Long
 hThread As Long
 dwProcessId As Long
 dwThreadId As Long
End Type
Private Type SECURITY_ATTRIBUTES
 nLength As Long
 lpSecurityDescriptor As Long
 bInheritHandle As Long
End Type
Private Const NORMAL_PRIORITY_CLASS As Long = &H20&
Private Const STARTF_USESTDHANDLES As Long = &H100&
Private Const STARTF_USESHOWWINDOW As Long = &H1&
Private Const SW_HIDE As Long = 0&
Private Const SW_NORMAL As Long = 1
Private Const INFINITE As Long = &HFFFF&
     
Public Function GetStrFromCommand(CommandLine As String) As String
On Error Resume Next
 Dim si As STARTUPINFO 'used to send info the CreateProcess
 Dim pi As PROCESS_INFORMATION 'used to receive info about the created process
 Dim retval As Long 'return value
 Dim hRead As Long 'the handle to the read end of the pipe
 Dim hWrite As Long 'the handle to the write end of the pipe
 Dim sBuffer(0 To 63) As Byte 'the buffer to store data as we read it from the pipe
 Dim lgSize As Long 'returned number of bytes read by readfile
 Dim sa As SECURITY_ATTRIBUTES
 Dim strResult As String 'returned results of the command line
 
' With sa
' .nLength = Len(sa)
' .bInheritHandle = 1& 'inherit, needed for this to work
' .lpSecurityDescriptor = 0&
' End With
 
 sa.nLength = Len(sa)
 sa.bInheritHandle = 1& 'inherit, needed for this to work
 sa.lpSecurityDescriptor = 0&
 
 retval = CreatePipe(hRead, hWrite, sa, 0&)
 
' With si
' .cb = Len(si)
' .cb = Len(si)
' .dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW 'tell it to use (not ignore) the values below
' .wShowWindow = SW_HIDE
' .hStdOutput = hWrite 'pass the write end of the pipe as the processes standard output
' End With
 
 si.cb = Len(si)
 si.cb = Len(si)
 si.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW 'tell it to use (not ignore) the values below
 si.wShowWindow = SW_HIDE
 si.hStdOutput = hWrite 'pass the write end of the pipe as the processes standard output
 
 
 retval = CreateProcess(vbNullString, CommandLine & vbNullChar, sa, sa, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, vbNullString, si, pi)
 WaitForSingleObject pi.hProcess, INFINITE
' Do While ReadFile(hRead, sBuffer(0), 64, lgSize, ByVal 0&)
' strResult = strResult & StrConv(sBuffer(), vbUnicode)
' Erase sBuffer()
' If lgSize <> 64 Then Exit Do
' Loop
 CloseHandle pi.hProcess
 CloseHandle pi.hThread
 CloseHandle hRead
 CloseHandle hWrite
 GetStrFromCommand = Replace(strResult, vbNullChar, "")
End Function

Public Function RunCommand(ByVal CommandLine As String, ByVal WaitForIt As Boolean, Optional ByVal ShowWindow As Boolean = False) As Long
On Error Resume Next
Dim si As STARTUPINFO 'used to send info the CreateProcess
Dim pi As PROCESS_INFORMATION 'used to receive info about the created process
Dim retval As Long 'return value
'Dim hRead As Long 'the handle to the read end of the pipe
'Dim hWrite As Long 'the handle to the write end of the pipe
'Dim lgSize As Long 'returned number of bytes read by readfile
Dim sa As SECURITY_ATTRIBUTES
'Dim strResult As String 'returned results of the command line
With sa
.nLength = Len(sa)
.bInheritHandle = 1& 'inherit, needed for this to work
.lpSecurityDescriptor = 0&
End With
If ShowWindow Then
  With si
    .cb = Len(si)
    .dwFlags = STARTF_USESHOWWINDOW  'tell it to use (not ignore) the values below
    .wShowWindow = SW_NORMAL
    '.hStdOutput = hWrite 'pass the write end of the pipe as the processes standard output
  End With
Else
  With si
    .cb = Len(si)
    .dwFlags = STARTF_USESHOWWINDOW  'tell it to use (not ignore) the values below
    .wShowWindow = SW_HIDE
    '.hStdOutput = hWrite 'pass the write end of the pipe as the processes standard output
  End With
End If
retval = CreateProcess(vbNullString, CommandLine & vbNullChar, sa, sa, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, vbNullString, si, pi)
If WaitForIt Then
  WaitForSingleObject pi.hProcess, INFINITE
End If
RunCommand = pi.dwProcessId
CloseHandle pi.hProcess
CloseHandle pi.hThread
End Function

