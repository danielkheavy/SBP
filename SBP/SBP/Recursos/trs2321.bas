Attribute VB_Name = "Module3"
Option Explicit

Global ComNum     As Long

Global bRead(255) As Byte

Type COMSTAT

    fCtsHold As Long
    fDsrHold As Long
    fRlsdHold As Long
    fXoffHold As Long
    fXoffSent As Long
    fEof As Long
    fTxim As Long
    fReserved As Long
    cbInQue As Long
    cbOutQue As Long

End Type

Type COMMTIMEOUTS

    ReadIntervalTimeout As Long
    ReadTotalTimeoutMultiplier As Long
    ReadTotalTimeoutConstant As Long
    WriteTotalTimeoutMultiplier As Long
    WriteTotalTimeoutConstant As Long

End Type

Type DCB

    DCBlength As Long
    BaudRate As Long
    fBinary As Long
    fParity As Long
    fOutxCtsFlow As Long
    fOutxDsrFlow As Long
    fDtrControl As Long
    fDsrSensitivity As Long
    fTXContinueOnXoff As Long
    fOutX As Long
    fInX As Long
    fErrorChar As Long
    fNull As Long
    fRtsControl As Long
    fAbortOnError As Long
    fDummy2 As Long
    wReserved As Integer
    XonLim As Integer
    XoffLim As Integer
    ByteSize As Byte
    Parity As Byte
    StopBits As Byte
    XonChar As Byte
    XoffChar As Byte
    ErrorChar As Byte
    EofChar As Byte
    EvtChar As Byte

End Type

Type OVERLAPPED

    Internal As Long
    InternalHigh As Long
    Offset As Long
    OffsetHigh As Long
    hEvent As Long

End Type

Type SECURITY_ATTRIBUTES

    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long

End Type

Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function GetLastError Lib "kernel32" () As Long
Declare Function ReadFile _
        Lib "kernel32" (ByVal hFile As Long, _
                        lpBuffer As Any, _
                        ByVal nNumberOfBytesToRead As Long, _
                        lpNumberOfBytesRead As Long, _
                        lpOverlapped As Long) As Long
Declare Function WriteFile _
        Lib "kernel32" (ByVal hFile As Long, _
                        lpBuffer As Any, _
                        ByVal nNumberOfBytesToWrite As Long, _
                        lpNumberOfBytesWritten As Long, _
                        lpOverlapped As Long) As Long
Declare Function SetCommTimeouts _
        Lib "kernel32" (ByVal hFile As Long, _
                        lpCommTimeouts As COMMTIMEOUTS) As Long
Declare Function GetCommTimeouts _
        Lib "kernel32" (ByVal hFile As Long, _
                        lpCommTimeouts As COMMTIMEOUTS) As Long
Declare Function BuildCommDCB _
        Lib "kernel32" _
        Alias "BuildCommDCBA" (ByVal lpDef As String, _
                               lpDCB As DCB) As Long
Declare Function SetCommState _
        Lib "kernel32" (ByVal hCommDev As Long, _
                        lpDCB As DCB) As Long
Declare Function CreateFile _
        Lib "kernel32" _
        Alias "CreateFileA" (ByVal lpFileName As String, _
                             ByVal dwDesiredAccess As Long, _
                             ByVal dwShareMode As Long, _
                             ByVal lpSecurityAttributes As Long, _
                             ByVal dwCreationDisposition As Long, _
                             ByVal dwFlagsAndAttributes As Long, _
                             ByVal hTemplateFile As Long) As Long
Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long

Function fin_com()
    fin_com = CloseHandle(ComNum)

End Function

Function FlushComm()
    FlushFileBuffers (ComNum)

End Function

Function Init_Com(ComNumber As String, Comsettings As String) As Boolean

    On Error GoTo handelinitcom

    Dim ComSetup As DCB, answer, Stat As COMSTAT, RetBytes As Long

    Dim retval As Long

    Dim CtimeOut As COMMTIMEOUTS, BarDCB As DCB

    ' Open the communications port for read/write (&HC0000000).
    ' Must specify existing file (3).
    ComNum = CreateFile(ComNumber, &HC0000000, 0, 0&, &H3, 0, 0)

    If ComNum = -1 Then
        MsgBox "Com Port " & ComNumber & " not available. Use Serial settings (on the main menu) to setup your ports.", 48
        Init_Com = False
        Exit Function

    End If

    'Setup Time Outs for com port
    CtimeOut.ReadIntervalTimeout = 20
    CtimeOut.ReadTotalTimeoutConstant = 1
    CtimeOut.ReadTotalTimeoutMultiplier = 1
    CtimeOut.WriteTotalTimeoutConstant = 10
    CtimeOut.WriteTotalTimeoutMultiplier = 1
    retval = SetCommTimeouts(ComNum, CtimeOut)

    If retval = -1 Then
        retval = GetLastError()
        MsgBox "Unable to set timeouts for port " & ComNumber & " Error: " & retval
        retval = CloseHandle(ComNum)
        Init_Com = False
        Exit Function

    End If

    retval = BuildCommDCB(Comsettings, BarDCB)

    If retval = -1 Then
        retval = GetLastError()
        MsgBox "Unable to build Comm DCB " & Comsettings & " Error: " & retval
        retval = CloseHandle(ComNum)
        Init_Com = False
        Exit Function

    End If

    retval = SetCommState(ComNum, BarDCB)

    If retval = -1 Then
        retval = GetLastError()
        MsgBox "Unable to set Comm DCB " & Comsettings & " Error: " & retval
        retval = CloseHandle(ComNum)
        Init_Com = False
        Exit Function

    End If
    
    Init_Com = True
handelinitcom:
    Exit Function

End Function

Function ReadCommPure() As String

    On Error GoTo handelpurecom

    Dim RetBytes   As Long, I As Integer, ReadStr As String, retval As Long

    Dim CheckTotal As Integer, CheckDigitLC As Integer

    retval = ReadFile(ComNum, bRead(0), 255, RetBytes, 0)
    ReadStr = ""

    If (RetBytes > 0) Then

        For I = 0 To RetBytes - 1
            ReadStr = ReadStr & Chr(bRead(I))
        Next I

    Else
        FlushComm

    End If

    'Return the string read from serial port
    ReadCommPure = ReadStr
handelpurecom:
    Exit Function

End Function

Function WriteCOM32(COMString As String) As Integer

    On Error GoTo handelwritelpt

    Dim RetBytes As Long, LenVal As Long

    Dim retval   As Long
    
    If Len(COMString) > 255 Then
        WriteCOM32 Left$(COMString, 255)
        WriteCOM32 Right$(COMString, Len(COMString) - 255)
        Exit Function

    End If
    
    For LenVal = 0 To Len(COMString) - 1
        bRead(LenVal) = Asc(Mid$(COMString, LenVal + 1, 1))
    Next LenVal

    '    bRead(LenVal) = 0
    retval = WriteFile(ComNum, bRead(0), Len(COMString), RetBytes, 0)
    'espera_segundo 2
    '    FlushComm
    WriteCOM32 = RetBytes
    
handelwritelpt:
    Exit Function

End Function

