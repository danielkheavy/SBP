Attribute VB_Name = "Module4"
Option Explicit

Public Declare Function InternetOpen _
               Lib "wininet.dll" _
               Alias "InternetOpenA" (ByVal sAgent As String, _
                                      ByVal lAccessType As Long, _
                                      ByVal sProxyName As String, _
                                      ByVal sProxyBypass As String, _
                                      ByVal lFlags As Long) As Long

Public Declare Function InternetOpenUrl _
               Lib "wininet.dll" _
               Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, _
                                         ByVal sURL As String, _
                                         ByVal sHeaders As String, _
                                         ByVal lHeadersLength As Long, _
                                         ByVal lFlags As Long, _
                                         ByVal lContext As Long) As Long

Public Declare Function InternetReadFile _
               Lib "wininet.dll" (ByVal hFile As Long, _
                                  ByVal sBuffer As String, _
                                  ByVal lNumBytesToRead As Long, _
                                  lNumberOfBytesRead As Long) As Integer

Public Declare Function InternetCloseHandle _
               Lib "wininet.dll" (ByVal hInet As Long) As Integer

Public Const IF_FROM_CACHE = &H1000000

Public Const IF_MAKE_PERSISTENT = &H2000000

Public Const IF_NO_CACHE_WRITE = &H4000000
       
Private Const BUFFER_LEN = 256

Public Function GetUrlSource(sURL As String) As String

    Dim sBuffer   As String * BUFFER_LEN, iResult As Integer, sData As String

    Dim hInternet As Long, hSession As Long, lreturn As Long

    'get the handle of the current internet connection
    hSession = InternetOpen("vb wininet", 1, vbNullString, vbNullString, 0)

    'get the handle of the url
    If hSession Then hInternet = InternetOpenUrl(hSession, sURL, vbNullString, 0, IF_NO_CACHE_WRITE, 0)

    'if we have the handle, then start reading the web page
    If hInternet Then
        'get the first chunk & buffer it.
        iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lreturn)
        sData = sBuffer

        'if there's more data then keep reading it into the buffer
        Do While lreturn <> 0
            iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lreturn)
            sData = sData + Mid(sBuffer, 1, lreturn)
        Loop

    End If
   
    'close the URL
    iResult = InternetCloseHandle(hInternet)

    GetUrlSource = sData

End Function

Function busca_ip() As String

    'permite ver ip publica de este momento
    Dim RawSource      As String

    Dim SourceEndFound As Boolean, UserIP As String

    Dim cC             As Single

    'Command1.Enabled = False
    RawSource = GetUrlSource("http://checkip.dyndns.org/")
    cC = 77

    Do Until SourceEndFound = True

        If Mid(RawSource, cC, 1) <> "<" Then
            UserIP = UserIP & Mid(RawSource, cC, 1)
        Else
            SourceEndFound = True

        End If

        cC = cC + 1
    Loop
    busca_ip = UserIP

    'Command1.Enabled = True
End Function
