Attribute VB_Name = "mod"
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sUrl As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer

Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, _
    ByVal lparam As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, _
    ByVal lngWMsg As Long, ByVal lngWParam As Long, ByVal lngLparam As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
    (ByVal hwnd As Long, ByVal lpClassName As String, _
    ByVal nMaxCount As Long) As Long

Private Declare Function SetWindowTextA Lib "user32" (ByVal hwnd As Long, ByVal lpString As String) As Long

Public Const scUserAgent = "XiaoKongs Internet Online 1.0"
Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Public Const INTERNET_OPEN_TYPE_DIRECT = 1
Public Const INTERNET_OPEN_TYPE_PROXY = 3
Public Const INTERNET_FLAG_RELOAD = &H80000000
Const WM_CLOSE = &H10

Function EnumWindowsProc(ByVal hwnd As Long, ByVal lparam As Long) As Boolean
    Dim s As String, W As String
    Dim H As Long
    s = String(256, 0)
    W = String(256, 0)
    Call GetWindowText(hwnd, s, 256)
    Call GetClassName(hwnd, W, 256)
    W = Left(W, InStr(W, Chr(0)) - 1)
    s = Left(s, InStr(s, Chr(0)) - 1)
    If Len(s) > 0 Then
'        Form1.Text1.Text = Form1.Text1.Text & S & "-->" & hwnd & "-->" & W & vbCrLf
    End If
'    不让IE运行的代码
    If InStr(s, "bitsadmin") <> 0 Then
        H = FindWindow(vbNullString, s)
'        PostMessage H, WM_CLOSE, 0&, 0&
        SetWindowTextA H, "修改后的dao标自题"
    End If
    EnumWindowsProc = True
End Function

Public Function GetUrlFile(stUrl As String) As String
    Dim lgInternet As Long, lgSession As Long
    Dim stBuf As String * 1024
    Dim inRes As Integer
    Dim lgRet As Long
    Dim stTotal As String
    stTotal = vbNullString
    lgSession = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    DoEvents
    If lgSession Then
        lgInternet = InternetOpenUrl(lgSession, stUrl, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
'              lgInternet = InternetOpenUrl(lgSession, stUrl, vbNullString, 0, INTERNET_FLAG_NO_CACHE_WRITE, 0)
        If lgInternet Then
            Do
                inRes = InternetReadFile(lgInternet, stBuf, 1024, lgRet)
                stTotal = stTotal & StrConv(LeftB$(StrConv(stBuf, vbFromUnicode), lgRet), vbUnicode)
            Loop While (lgRet <> 0)
        End If
        inRes = InternetCloseHandle(lgInternet)
    End If
    GetUrlFile = stTotal
End Function

Public Function SendMSG(stUrl As String) As String
    Dim lgInternet As Long, lgSession As Long
    Dim stBuf As String * 1024
    Dim inRes As Integer
    Dim lgRet As Long
    Dim stTotal As String
    stTotal = vbNullString
    lgSession = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    DoEvents
    If lgSession Then
        lgInternet = InternetOpenUrl(lgSession, stUrl, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
'              lgInternet = InternetOpenUrl(lgSession, stUrl, vbNullString, 0, INTERNET_FLAG_NO_CACHE_WRITE, 0)
        If lgInternet Then
            Do
                inRes = InternetReadFile(lgInternet, stBuf, 1024, lgRet)
                stTotal = stTotal & StrConv(LeftB$(StrConv(stBuf, vbFromUnicode), lgRet), vbUnicode)
            Loop While (lgRet <> 0)
        End If
        inRes = InternetCloseHandle(lgInternet)
    End If
    SendMSG = stTotal
End Function

