VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RAWPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' A very simple printer object used to send data directly
' to the printer.
'
' Date       Notes
' 1-12-97    Written by TPA  tarheit@wcoil.com
'
'
'
' Basic usage:
'    Dim p As New RAWPrinter
'    p.NewDoc("My Document")
'    p.PrintText "This is a test."
'    p.PrintFile("C:\myfile.txt") 'Send file directly to the printer.
'    p.EndDoc
'

'
' Printer errors used with RAWPrinter, and future
' printer objects and related functions.
'

Public Enum PrinterErrors

    vbPE_CantOpenPrinter = 2000   ' Can't Open the printer device.
    vbPE_CantStartJob             ' Can't Start the print job.
    vbPE_CantStartPage            ' Can't start printing a page.
    vbPE_UnSentBytes              ' Some bytes were not successfully sent to the printer.
    vbPE_KillDocFailed            ' Could not cancel the print job.
    vbPE_CantChangeName           ' Can't change document name.
    vbPE_FailedWrite              ' Failed write to printer.
    vbPE_ReadFileError            ' Could not read from file.
    vbPE_CantEndPage              ' Call to end page failed.
    vbPE_CantEndDoc               ' Call to close doc failed.
    vbPE_CantChangeDevice         ' Can't change device while printing.
    vbPE_CantCreateDC             ' Can't create a device context.

End Enum

Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (hpvDest As Any, _
                                       hpvSource As Any, _
                                       ByVal cbCopy As Long)
            
Private Declare Function OpenPrinter _
                Lib "winspool.drv" _
                Alias "OpenPrinterA" (ByVal pPrinterName As String, _
                                      phPrinter As Long, _
                                      pDefault As Any) As Long
   
Private Declare Function StartDocPrinter _
                Lib "winspool.drv" _
                Alias "StartDocPrinterA" (ByVal hPrinter As Long, _
                                          ByVal Level As Long, _
                                          pDocInfo As Any) As Long
   
Private Declare Function StartPagePrinter _
                Lib "winspool.drv" (ByVal hPrinter As Long) As Long
   
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
   
Private Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
   
Private Declare Function EndPagePrinter _
                Lib "winspool.drv" (ByVal hPrinter As Long) As Long
   
Private Type DOC_INFO_1

    pDocName As String
    pOutputFile As String
    pDatatype As String

End Type

Private Declare Function SetJob _
                Lib "winspool.drv" _
                Alias "SetJobA" (ByVal hPrinter As Long, _
                                 ByVal JobId As Long, _
                                 ByVal Level As Long, _
                                 pJob As Any, _
                                 ByVal Command As Long) As Long

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

Private Type JOB_INFO_1

    JobId As Long
    pPrinterName As String
    pMachineName As String
    pUserName As String
    pDocument As String
    pDatatype As String
    pStatus As String
    Status As Long
    Priority As Long
    Position As Long
    TotalPages As Long
    PagesPrinted As Long
    Submitted As SYSTEMTIME

End Type

Private Const JOB_POSITION_UNSPECIFIED = 0

Private Declare Function GetJob _
                Lib "winspool.drv" _
                Alias "GetJobA" (ByVal hPrinter As Long, _
                                 ByVal JobId As Long, _
                                 ByVal Level As Long, _
                                 pJob As Any, _
                                 ByVal cdBuf As Long, _
                                 pcbNeeded As Long) As Long

Private Const MAX_PRIORITY = 99

Private Const MIN_PRIORITY = 1

Private Const DEF_PRIORITY = 1

Private Declare Function WritePrinter _
                Lib "winspool.drv" (ByVal hPrinter As Long, _
                                    pBuf As Any, _
                                    ByVal cdBuf As Long, _
                                    pcWritten As Long) As Long

Private Const JOB_CONTROL_PAUSE = 1

Private Const JOB_CONTROL_RESUME = 2

Private Const JOB_CONTROL_CANCEL = 3

Private Const JOB_CONTROL_RESTART = 4

Private Const JOB_CONTROL_DELETE = 5

Private lPrinter      As Long       ' Printer handle

Private lBytesWritten As Long  ' Number of bytes written

Private lBytesSent    As Long     ' Number of bytes that should have been written.

Private lJob          As Long           ' Print job handle

Private sDocName      As String     ' Name of the document

Private sDeviceName   As String  ' Device name.

Private bJobStarted   As Boolean ' Have we started a print job.

Public Sub NewPage()

    If Not bJobStarted Then
        NewDoc
    Else

        'end last page
        If EndPagePrinter(lPrinter) <= 0 Then
            Call EndDocPrinter(lPrinter)
            Call ClosePrinter(lPrinter)
            bJobStarted = False
            Err.Raise vbPE_CantEndPage, "RAWPrinter", "Can't end page."
            Exit Sub

        End If
      
        If StartPagePrinter(lPrinter) <= 0 Then
            Call EndDocPrinter(lPrinter)
            Call ClosePrinter(lPrinter)
            bJobStarted = False
            Err.Raise vbPE_CantStartPage, "RAWPrinter", "Can't start page."
            Exit Sub

        End If

    End If

End Sub

Public Sub NewDoc(Optional DocName As String = "", _
                  Optional FileName As String = vbNullString)

    Dim di As DOC_INFO_1
   
    If bJobStarted Then
        EndDoc

    End If
      
    If OpenPrinter(sDeviceName, lPrinter, ByVal 0&) <= 0 Then
        Err.Raise vbPE_CantOpenPrinter, "RAWPrinter", "Can't Open Printer Device"
        Exit Sub

    End If
   
    If DocName <> "" Then
        sDocName = DocName

    End If
   
    di.pDocName = sDocName & vbNullChar

    If FileName = vbNullString Then
        di.pOutputFile = FileName
    Else
        di.pOutputFile = FileName & vbNullChar

    End If

    di.pDatatype = "RAW" & vbNullChar
   
    lJob = StartDocPrinter(lPrinter, 1, di)
   
    If lJob <= 0 Then
        Call ClosePrinter(lPrinter)
        Err.Raise vbPE_CantStartJob, "RAWPrinter", "Can't start print job."
        Exit Sub

    End If
   
    If StartPagePrinter(lPrinter) <= 0 Then
        Call EndDocPrinter(lPrinter)
        Call ClosePrinter(lPrinter)
        Err.Raise vbPE_CantStartPage, "RAWPrinter", "Can't start page."
        Exit Sub

    End If
   
    lBytesWritten = 0
    lBytesSent = 0
    bJobStarted = True

End Sub

Public Sub KillDoc()

    Dim b As Long
   
    If bJobStarted Then
        b = SetJob(lPrinter, lJob, 0, ByVal 0&, JOB_CONTROL_CANCEL)
        Call EndPagePrinter(lPrinter)
        Call EndDocPrinter(lPrinter)
        Call ClosePrinter(lPrinter)
        bJobStarted = False
    Else
        b = 0

    End If
   
    If b <= 0 Then
        Err.Raise vbPE_KillDocFailed, "RAWPrinter", "Could not cancle the print job."

    End If

End Sub

Public Sub EndDoc()

    If Not bJobStarted Then
        Exit Sub

    End If
   
    If EndPagePrinter(lPrinter) <= 0 Then
        Call EndDocPrinter(lPrinter)
        Call ClosePrinter(lPrinter)
        bJobStarted = False
        Err.Raise vbPE_CantEndPage, "RAWPrinter", "Can't end page."
        Exit Sub

    End If
   
    If EndDocPrinter(lPrinter) <= 0 Then
        Call ClosePrinter(lPrinter)
        bJobStarted = False
        Err.Raise vbPE_CantEndDoc, "RAWPrinter", "Can't end print job."
        Exit Sub

    End If
   
    Call ClosePrinter(lPrinter)
   
    bJobStarted = False
   
    If lBytesWritten <> lBytesSent Then
        Err.Raise vbPE_UnSentBytes, "RAWPrinter", "Some data was not sent to the printer."

    End If

End Sub

Public Property Let DeviceName(Name As String)

    If bJobStarted Then
        Err.Raise vbPE_CantChangeDevice, "RAWPrinter", "Can't change device while printing."
    Else
        sDeviceName = Name

    End If

End Property

Public Property Get DeviceName() As String
    DeviceName = sDeviceName

End Property

'
' Bug... this doesn't work
'
Public Property Let DocumentName(DocName As String)

    Dim di As JOB_INFO_1
   
    If bJobStarted Then
        di.pDocument = DocName & vbNullChar

        If SetJob(lPrinter, lJob, 1, di, 0&) <= 0 Then
            Err.Raise vbPE_CantChangeName, "RAWPrinter", "Failed to change document name."
            Exit Property

        End If

    End If
   
    sDocName = DocName

End Property

Public Property Get DocumentName() As String
    DocumentName = sDocName

End Property

Public Sub PrintText(txt As String)

    Dim I As Long
   
    If Not bJobStarted Then
        NewDoc

    End If
   
    lBytesSent = lBytesSent + Len(txt)
   
    If WritePrinter(lPrinter, ByVal txt, Len(txt), I) = 0 Then
        Call EndPagePrinter(lPrinter)
        Call EndDocPrinter(lPrinter)
        Call ClosePrinter(lPrinter)
        bJobStarted = False
        Err.Raise vbPE_FailedWrite, "RAWPrinter", "Failed write to printer."
        Exit Sub

    End If
   
    lBytesWritten = lBytesWritten + I

End Sub

Public Sub PrintFile(fname As String)

    Dim fh     As Long

    Dim buffer As String

    Dim fl     As Long

    Dim R      As Long

    Dim I      As Long

    Dim bs     As Long
   
    If Not bJobStarted Then
        NewDoc

    End If
   
    fh = FreeFile(0)
    bs = 8192
    buffer = String(bs, 0)
   
    Open fname For Binary Access Read As fh
    fl = LOF(fh)
    R = 0
   
    If fl = 0 Then
        Close fh
        Exit Sub

    End If
   
    Do

        If fl - R < bs Then
            bs = fl - R
            buffer = String(bs, 0)

        End If
   
        Get fh, , buffer
        MsgBox buffer
        lBytesSent = lBytesSent + bs
        R = R + bs
      
        If WritePrinter(lPrinter, ByVal buffer, bs, I) = 0 Then
            Call EndPagePrinter(lPrinter)
            Call EndDocPrinter(lPrinter)
            Call ClosePrinter(lPrinter)
            bJobStarted = False

            On Error GoTo 0

            Err.Raise vbPE_FailedWrite, "RAWPrinter", "Failed write to printer."
            Exit Sub

        End If
      
        lBytesWritten = lBytesWritten + I
    Loop While R <> fl

    Close fh
    Exit Sub
   
PrintFileError:

    On Error Resume Next
   
    Call EndPagePrinter(lPrinter)
    Call EndDocPrinter(lPrinter)
    Call ClosePrinter(lPrinter)
    bJobStarted = False
    Close fh
   
    On Error GoTo 0

    Err.Raise vbPE_ReadFileError, "RAWPrinter", "Could not read from file."

End Sub

Private Sub Class_Initialize()
    sDocName = "Visual Basic - RAWPrinter"
    sDeviceName = Printer.DeviceName
    bJobStarted = False

End Sub

Private Sub Class_Terminate()

    If bJobStarted Then
        EndDoc

    End If

End Sub

Public Property Get hPrinter() As Long
    hPrinter = lPrinter

End Property

Public Property Get hJob() As Long
    hJob = lJob

End Property

Public Property Get Priority() As Long

    Dim di As String ' stores JOB_INFO_1

    Dim I  As Long
   
    Call GetJob(lPrinter, lJob, 1, ByVal di, 0, I)
    di = String(I, 0)
    Call CopyMemory(I, ByVal (Mid$(di, 33, 4)), 4)

    Priority = I

End Property

' Bug: Doesn't work?
Public Property Let Priority(ByVal I As Long)

    Dim di As JOB_INFO_1
   
    'JobId, pPrinterName, pMachineName, pDrivername,
    'Size, Submitted, and Time are ignored
    If I < MIN_PRIORITY Then
        I = DEF_PRIORITY
    ElseIf I > MAX_PRIORITY Then
        I = MAX_PRIORITY

    End If
   
    di.Priority = I
    di.Position = JOB_POSITION_UNSPECIFIED
    di.pUserName = vbNullString
    di.pDocument = vbNullString
    di.pDatatype = vbNullString
    di.pStatus = vbNullString
    di.Status = 0
    di.TotalPages = 0
    di.PagesPrinted = 0
   
    Call SetJob(lPrinter, lJob, 1, di, 0)

End Property

