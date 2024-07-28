VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'* ******************************************************* *
'*    模块名称：FTP.cls
'*    模块功能：使用wininet API进行FTP操作
'*    作者：lyserver
'*    联系方式：http://blog.csdn.net/lyserver
'* ******************************************************* *
Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime(1) As Long
        ftLastAccessTime(1) As Long
        ftLastWriteTime(1) As Long
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * 260
        cAlternate As String * 14
End Type
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String) As Boolean
Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, lpdwCurrentDirectory As Long) As Boolean
Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean
Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hFtpSession As Long, ByVal lpszLocalFile As String, ByVal lpszRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Long
Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal lpszExisting As String, ByVal lpszNew As String) As Boolean
Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long
Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (lpdwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
Private Const INTERNET_FLAG_PASSIVE = &H8000000
Private Const INTERNET_FLAG_ASYNC = &H10000000
Private Const INTERNET_FLAG_EXISTING_CONNECT = &H20000000
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const INTERNET_SERVICE_FTP = 1
Private Const INTERNET_SERVICE_HTTP = 3
Private Const INTERNET_CONNECTION_MODEM = &H1
Private Const INTERNET_CONNECTION_LAN = &H2
Private Const INTERNET_CONNECTION_PROXY = &H4
Private Const FTP_TRANSFER_TYPE_ASCII = 1
Private Const FTP_TRANSFER_TYPE_BINARY = 2
Private Const ERROR_NO_MORE_FILES = 18&
Private Const ERROR_INTERNET_EXTENDED_ERROR = 12003
Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type
Private Declare Function FileTimeToSystemTime Lib "kernel32" (ByRef lpFileTime As Long, lpSystemTime As SYSTEMTIME) As Long

Public Event EnumFileProc(FileName As String, Attr As VbFileAttribute, Size As Long, Create As String, Modify As String, Cancel As Boolean)

Dim m_hInternet As Long, m_hConnect As Long, m_Cancel As Boolean

Private Sub Class_Initialize()
    m_hInternet = InternetOpen("FTP Appliction", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, INTERNET_FLAG_NO_CACHE_WRITE)
End Sub

Private Sub Class_Terminate()
    If m_hConnect <> 0 Then InternetCloseHandle m_hConnect
    InternetCloseHandle m_hInternet
End Sub

Public Function Login(Server As String, Optional Port As Integer = 1080, Optional UserName As String = "9999", Optional Password = "9999") As Boolean
    If m_hInternet = 0 Then Exit Function
    If m_hConnect <> 0 Then Logout
    m_hConnect = InternetConnect(m_hInternet, Server, Port, UserName, Password, INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE Or INTERNET_FLAG_EXISTING_CONNECT, 0)
    Login = (m_hConnect <> 0)
End Function

Public Function Logout() As Boolean
    If m_hConnect <> 0 Then
        InternetCloseHandle m_hConnect
        m_hConnect = 0
        Logout = True
    End If
End Function

Public Function GetDirectory() As String
    Dim strPath As String, nLen As Long
    
    If m_hConnect = 0 Then Exit Function
    nLen = 260
    strPath = String(nLen, vbNullChar)
    FtpGetCurrentDirectory m_hConnect, strPath, nLen
    GetDirectory = Left(strPath, InStr(strPath, vbNullChar) - 1)
End Function

Public Function SetDirectory(ByVal FtpPath As String) As Boolean
    If m_hConnect = 0 Then Exit Function
    FtpSetCurrentDirectory m_hConnect, FtpPath
    SetDirectory = (GetLastError <> ERROR_INTERNET_EXTENDED_ERROR)
End Function

Public Function CreateDirectory(ByVal FtpPath As String) As Boolean
    FtpCreateDirectory m_hConnect, FtpPath
    CreateDirectory = (GetLastError <> ERROR_INTERNET_EXTENDED_ERROR)
End Function

Public Function DeleteDirectory(ByVal FtpPath As String) As Boolean
    FtpRemoveDirectory m_hConnect, FtpPath
    DeleteDirectory = (GetLastError <> ERROR_INTERNET_EXTENDED_ERROR)
End Function

Public Sub EnumFile(ByVal strPath As String, Optional ByVal LookInSubPath As Boolean)
    Static strOldPath As String
    Dim wfd As WIN32_FIND_DATA
    Dim hFind As Long, i As Long
    Dim strFile As String, strSubPath As Variant
    Dim CreateTime As String, ModifyTime As String, SysTm As SYSTEMTIME

    strPath = Replace(strPath, "/", "/")
    If Right(strPath, 1) <> "/" Then strPath = strPath & "/"
    hFind = FtpFindFirstFile(m_hConnect, strPath, wfd, INTERNET_FLAG_RELOAD, 0)
    Do While GetLastError() <> ERROR_NO_MORE_FILES
        strFile = strPath & Left(wfd.cFileName, InStr(wfd.cFileName, vbNullChar) - 1)
        FileTimeToSystemTime wfd.ftCreationTime(0), SysTm
        CreateTime = DateSerial(SysTm.wYear, SysTm.wMonth, SysTm.wDay) & " " & TimeSerial(SysTm.wHour, SysTm.wMinute, SysTm.wSecond)
        FileTimeToSystemTime wfd.ftLastWriteTime(0), SysTm
        ModifyTime = DateSerial(SysTm.wYear, SysTm.wMonth, SysTm.wDay) & " " & TimeSerial(SysTm.wHour, SysTm.wMinute, SysTm.wSecond)
        If (wfd.dwFileAttributes Or vbDirectory) = wfd.dwFileAttributes Then
            RaiseEvent EnumFileProc(strFile, wfd.dwFileAttributes, -1, CreateTime, ModifyTime, m_Cancel)
            If LookInSubPath Then strSubPath = strSubPath & strFile & vbNullChar
        Else
            RaiseEvent EnumFileProc(strFile, wfd.dwFileAttributes, wfd.nFileSizeLow, CreateTime, ModifyTime, m_Cancel)
        End If
        If InternetFindNextFile(hFind, wfd) = 0 Or m_Cancel Then Exit Do
    Loop
    InternetCloseHandle hFind
    
    If LookInSubPath = False Or Len(strPath) = 0 Then Exit Sub
    strSubPath = Split(strSubPath, vbNullChar)
    For i = 0 To UBound(strSubPath) - 1
        EnumFile strSubPath(i), True
    Next
End Sub

Public Function Rename(ByVal FtpOldName As String, ByVal FtpNewName As String) As Boolean
    FtpRenameFile m_hConnect, FtpOldName, FtpNewName
    Rename = (GetLastError <> ERROR_INTERNET_EXTENDED_ERROR)
End Function

Public Function DeleteFile(ByVal FtpFile As String) As Boolean
    FtpDeleteFile m_hConnect, FtpFile
    DeleteFile = (GetLastError <> ERROR_INTERNET_EXTENDED_ERROR)
End Function

Public Function UpFile(ByVal LocalFile As String, Optional ByVal FtpFile As String) As Boolean
    If m_hConnect = 0 Then Exit Function
    If Len(Dir(LocalFile)) = 0 Or Left(Dir(LocalFile), 1) = "." Then Exit Function
    If Len(FtpFile) = 0 Then
        If InStr(LocalFile, "/") = 0 Then
            FtpFile = LocalFile
        Else
            FtpFile = StrReverse(LocalFile)
            FtpFile = StrReverse(Left(FtpFile, InStr(FtpFile, "/") - 1))
        End If
    End If
    FtpPutFile m_hConnect, LocalFile, FtpFile, 1, 0
    UpFile = (GetLastError <> ERROR_INTERNET_EXTENDED_ERROR)
End Function

'Public Function DownFile(ByVal FtpFile As String, Optional ByVal LocalFile As String) As Boolean
''    If m_hConnect = 0 Then Exit Function
''    If Len(LocalFile) = 0 Then
''        If InStr(FtpFile, "/") = 0 Then
''            LocalFile = FtpFile
''        Else
''            LocalFile = StrReverse(FtpFile)
''            LocalFile = StrReverse(Left(LocalFile, InStr(LocalFile, "/") - 1))
''        End If
''    End If
''    FtpGetFile m_hConnect, FtpFile, LocalFile, False, FILE_ATTRIBUTE_ARCHIVE, FTP_TRANSFER_TYPE_BINARY, 0
''    DownFile = (GetLastError <> ERROR_INTERNET_EXTENDED_ERROR)
'End Function


