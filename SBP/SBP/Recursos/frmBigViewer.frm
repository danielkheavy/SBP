VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form genver 
   Caption         =   "Viewer"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   12300
   Icon            =   "frmBigViewer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   12300
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdTextOK 
      Caption         =   "Finalizar..."
      Height          =   450
      Left            =   330
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4350
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Puerto Directos"
      Height          =   2175
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   11895
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   105
         TabIndex        =   5
         Top             =   285
         Width           =   11505
      End
   End
   Begin MSComDlg.CommonDialog cmdialog1 
      Left            =   1440
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click para cancelar Impresion..."
      Height          =   1095
      Left            =   960
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.PictureBox TextPict 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2115
      Left            =   0
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   0
      Top             =   0
      Width           =   2595
   End
   Begin VB.VScrollBar VBar 
      Height          =   2175
      Left            =   2640
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HBar 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2160
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InBufferSize    =   512
      OutBufferSize   =   64
   End
   Begin VB.Label VIENEDE 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Menu gy788 
      Caption         =   "&Archivo"
      Begin VB.Menu gu78h 
         Caption         =   "&1.Guardar Como.."
      End
   End
   Begin VB.Menu djim892 
      Caption         =   "&Imprime"
      Begin VB.Menu mnuArchivoArray 
         Caption         =   "Novisible"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu dkyuti92 
         Caption         =   "&1.Impresion directa por LPT,COM,X"
         Shortcut        =   {F5}
      End
      Begin VB.Menu cx89lo1 
         Caption         =   "&3.Impresion Usando Cola Impresion"
      End
   End
   Begin VB.Menu tamo9912 
      Caption         =   "&Tamaño"
   End
   Begin VB.Menu k8923 
      Caption         =   "&Correo"
   End
   Begin VB.Menu ldfo2323 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "genver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Large Textfile Display Form
'
'By Timothy R. Rude, timrude@hotmail.com

'Allows display of text files up to 32,768 lines
'and up to 32,768 characters per line

'Only VB-native controls are used, no OCX needed!

Public file As String   'Filename gets passed in by calling routine
'i.e.:  frmBigViewer.File = "filename"
'       frmBigViewer.Show vbModal

Private Declare Function TextOut _
                Lib "gdi32" _
                Alias "TextOutA" (ByVal hDC As Long, _
                                  ByVal X As Long, _
                                  ByVal Y As Long, _
                                  ByVal lpString As String, _
                                  ByVal ncount As Long) As Long

Private TextLine()    As String    'Array holds lines of text

Private MaxLineLength As Long   'Longest line in file

Private ScreenLines   As Long     'Number of lines displayable

Private ScreenWidth   As Long     'Number of characters displayable

Private CharWidth     As Single     'Width of single fixed-space character

Private CharHeight    As Single    'Height of single fixed-space character

Private Declare Function enumports _
                Lib "winspool.drv" _
                Alias "EnumPortsA" (ByVal pName As String, _
                                    ByVal Level As Long, _
                                    ByVal lpbPorts As Long, _
                                    ByVal cbBuf As Long, _
                                    pcbNeeded As Long, _
                                    pcReturned As Long) As Long

Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

Private Declare Sub CopyMem _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (pTo As Any, _
                                       uFrom As Any, _
                                       ByVal lSize As Long)

Private Declare Function HeapAlloc _
                Lib "kernel32" (ByVal hHeap As Long, _
                                ByVal dwFlags As Long, _
                                ByVal dwBytes As Long) As Long

Private Declare Function GetProcessHeap Lib "kernel32" () As Long

Private Declare Function HeapFree _
                Lib "kernel32" (ByVal hHeap As Long, _
                                ByVal dwFlags As Long, _
                                lpMem As Any) As Long

Private Type PORT_INFO_2

    pPortName As String
    pMonitorName As String
    pDescription As String
    fPortType As Long
    Reserved As Long

End Type

Private Type API_PORT_INFO_2

    pPortName As Long
    pMonitorName As Long
    pDescription As Long
    fPortType As Long
    Reserved As Long

End Type

Dim Ports(0 To 100) As PORT_INFO_2

Public Function CutString(strName As String) As String

    'Finds a null then trims the string
    Dim X As Integer

    X = InStr(strName, vbNullChar)

    If X > 0 Then CutString = Left(strName, X - 1) Else CutString = strName

End Function

Public Function LPSTRtoSTRING(ByVal lngPointer As Long) As String

    Dim lngLength As Long

    'number of characters
    lngLength = lstrlenW(lngPointer) * 2
    'Initialize the string
    LPSTRtoSTRING = String(lngLength, 0)
    'Copy the string
    CopyMem ByVal StrPtr(LPSTRtoSTRING), ByVal lngPointer, lngLength
    'Convert to Unicode
    LPSTRtoSTRING = CutString(StrConv(LPSTRtoSTRING, vbUnicode))

End Function

'You can specify a server name (example //WIN2KWKSTN) to get the ports of that machine
Public Function getports(ServerName As String) As Long

    Dim ret                   As Long

    Dim PortsStruct(0 To 100) As API_PORT_INFO_2

    Dim pcbNeeded             As Long

    Dim pcReturned            As Long

    Dim tmpbuffer             As Long

    Dim I                     As Integer

    'determine amount of bytes needed
    ret = enumports(ServerName, 2, tmpbuffer, 0, pcbNeeded, pcReturned)
    'use api to allocate the buffer
    tmpbuffer = HeapAlloc(GetProcessHeap(), 0, pcbNeeded)
    ret = enumports(ServerName, 2, tmpbuffer, pcbNeeded, pcbNeeded, pcReturned)

    If ret Then
        'convert string pointer value to vb-readable value
        CopyMem PortsStruct(0), ByVal tmpbuffer, pcbNeeded

        For I = 0 To pcReturned - 1
            Ports(I).pDescription = LPSTRtoSTRING(PortsStruct(I).pDescription)
            Ports(I).pPortName = LPSTRtoSTRING(PortsStruct(I).pPortName)
            Ports(I).pMonitorName = LPSTRtoSTRING(PortsStruct(I).pMonitorName)
            Ports(I).fPortType = PortsStruct(I).fPortType
        Next

    End If

    getports = pcReturned

    If tmpbuffer Then HeapFree GetProcessHeap(), 0, tmpbuffer

End Function

Private Sub Command1_Click()
    Command1.Visible = False
    opcion3 = 1

End Sub

Private Sub cx89lo1_Click()

    Dim found As Integer

    Dim sFile As String

    Dim oldprinter

    Dim oldprinter1

    oldprinter = Printer.DeviceName

    If Command1.Visible = True Then Exit Sub
    CommonDialog1.CancelError = True

    On Error Resume Next

    CommonDialog1.ShowPrinter

    If Err.Number = 32755 Then
        Exit Sub

    End If

    On Error GoTo 0

    If CommonDialog1.Orientation = cdlLandscape Then
        Printer.Orientation = cdlLandscape

    End If

    opcion3 = 0

    Command1.Visible = True
    oldprinter1 = Printer.DeviceName
    selecciona_impresoras (Trim(oldprinter1))

    sFile = globaldir & "\temporal\" & gusuario & ".txt"
    found = imprime_archivoj(sFile, 0, "" & tipoletra)
    Command1.Visible = False
    opcion3 = 0
    selecciona_impresoras (Trim(oldprinter))

End Sub

Sub imprime_en_cola()

    'Dim oldprinter
    '         oldprinter = Printer.DeviceName
    '         selecciona_impresoras (Trim(xbuf1))
    '         found = Imprime_archivojj(xbuf0, 0, "" & mytable11.Fields("tamanorden"), "" & mytable11.Fields("nombrefont"), "" & mytable11.Fields("BOLD"), "" & mytable11.Fields("letrainterna"))
    '         selecciona_impresoras (Trim(oldprinter))
    '
End Sub

Private Sub dkyuti92_Click()

    Dim found   As Integer

    Dim xpuerto As String

    Dim buf     As String

    xpuerto = "LPT1"
    'xpuerto = busca_usuario(gusuario)
    'If Len(xpuerto) = 0 Then
    '   MsgBox "Puerto no Definido ", 48, "Aviso"
    '   Exit Sub
    'End If
    buf = "Impresion directa: Debe de encontrarse definido en Personal el puerto por defecto de impresion " + Chr$(10) + Chr$(13)
    buf = buf & " El puerto definido es " + xpuerto + Chr$(10) + Chr$(13)
    buf = buf & "Desea imprimir "

    If MsgBox(buf, 1, "Aviso") <> 1 Then Exit Sub
    found = star_sp342(xpuerto, 0)
    found = corte_papel(xpuerto, 0)

End Sub

Private Sub dstr33_Click()

End Sub

Sub carga_puertos()

    Dim NumPorts As Long

    Dim I        As Integer

    Dim Item     As ListItem

    NumPorts = getports("")

    For I = 0 To NumPorts - 1

        With Ports(I)
            List1.AddItem .pPortName

            'List2.AddItem .fPortType
            'List3.AddItem .pDescription
            'List4.AddItem .pMonitorName
            'List5.AddItem .Reserved
        End With

    Next

End Sub

Private Sub Form_Activate()

    Dim I As Integer

    'File = globaldir & "\temporal\" & gusuario & ".txt"
    'Filename = globaldir & "\temporal\" & gusuario & ".txt"
    'apertura_automatico
    If Val(tipoletra) <= 6 Then
        tipoletra = "9"

    End If

    For I = 1 To mnuArchivoArray.count - 1
        Unload mnuArchivoArray(I)
    Next

    For I = 1 To 9

        If IsLptPortAvailable(I) Then
            agregar_menu "LPT" & I

        End If

    Next

    For I = 1 To 9

        If IsComPortAvailable(I) Then
            agregar_menu "COM" & I

        End If

    Next

End Sub

Function busca_usuario(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from vendedor where codigo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_usuario = "" & mytablex.Fields("puerto")

    End If

    mytablex.Close

End Function

Private Sub Form_Load()

    Dim FileHandle As Integer   'VB file handle number

    Dim Contents   As String      'Contents of entire text file

    Dim LineNumber As Long      'Line counter

    Dim LineLength As Long      'Length of current line
    
    TextPict.ScaleMode = vbPixels   'TextOut API requires Pixels
    
    'Using Terminal font - fixed spaced font required
    'Calculate height and width of single character
    CharWidth = TextPict.TextWidth("X")
    CharHeight = TextPict.TextHeight("X")
        
    Me.Caption = "Mirando: " & file
    
    'Read entire file into 'Contents' variable
    On Error GoTo View_Error

    Screen.MousePointer = vbHourglass
    FileHandle = FreeFile
    Open file For Binary As #FileHandle
        
    'Don't use this method to read the file - way too slow!
    'Contents = Input(LOF(FileHandle), #FileHandle) & vbCrLf
        
    'The following method is considerably faster!
    Contents = Space$(LOF(FileHandle) + 2)
    Get #FileHandle, , Contents
        
    'Added one more CR/LF at end so last line is completely shown in viewer
    Close #FileHandle

    On Error GoTo 0
    
    'Split individual lines into array
    TextLine() = Split(Contents, vbcrlf)

ViewOK:
    Contents = ""   'Release this memory - no longer needed
    'Determine length in characters of longest line
    MaxLineLength = 0

    For LineNumber = 0 To UBound(TextLine)
        LineLength = Len(TextLine(LineNumber))

        If LineLength > MaxLineLength Then MaxLineLength = LineLength
    Next LineNumber

    'Trigger a Resize so scrollbars are initialized
    Call Form_Resize
    'Turn on the OK button
    cmdTextOK.Visible = True
    cmdTextOK.Default = True
    cmdTextOK.Cancel = True
    Me.refresh
    Screen.MousePointer = vbNormal
    Exit Sub
    
View_Error:
    'Report the error and continue gracefully
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error: " & Err.Number
    Close #FileHandle
    ReDim TextLine(0)
    TextLine(0) = "===ERROR LECTURA FILE==="

    On Error GoTo 0

    Resume ViewOK

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    'Form.KeyPreview is True so we can respond to keypresses
    
    'Hotkeys:
    
    ' Up, Down, Left, Right - scroll one line/column
    ' Home, End - scroll to beginning or end of line
    ' PageUp, PageDown - scroll up or down one screenful
    
    ' Ctrl-Home, Ctrl-End - very beginning/end of document
    ' Ctrl-PageUp, Ctrl-PageDown - scroll to first/last line of document
    '                              leave horizontal position unchanged

    Select Case KeyCode

        Case vbKeyDown

            'Down one line
            If VBar.Value < VBar.max Then VBar.Value = VBar.Value + 1

        Case vbKeyUp

            'Up one line
            If VBar.Value > VBar.Min Then VBar.Value = VBar.Value - 1

        Case vbKeyPageDown

            'Down one screenful/bottom line
            If Shift = vbCtrlMask Then
                VBar.Value = VBar.max
            Else

                If VBar.Value <= VBar.max - VBar.LargeChange Then
                    VBar.Value = VBar.Value + VBar.LargeChange
                Else
                    VBar.Value = VBar.max

                End If

            End If

        Case vbKeyPageUp

            'Up one screenful/top line
            If Shift = vbCtrlMask Then
                VBar.Value = VBar.Min
            Else

                If VBar.Value - VBar.LargeChange >= VBar.Min Then
                    VBar.Value = VBar.Value - VBar.LargeChange
                Else
                    VBar.Value = VBar.Min

                End If

            End If

        Case vbKeyHome
            'Beginning of line/document
            HBar.Value = HBar.Min

            If Shift = vbCtrlMask Then VBar.Value = VBar.Min

        Case vbKeyEnd
            'End of line/document
            HBar.Value = HBar.max

            If Shift = vbCtrlMask Then VBar.Value = VBar.max

        Case vbKeyRight

            'Right one character
            If HBar.Value < HBar.max Then HBar.Value = HBar.Value + 1

        Case vbKeyLeft

            'Left one character
            If HBar.Value > HBar.Min Then HBar.Value = HBar.Value - 1

    End Select

End Sub

Private Sub Form_Resize()

    On Error GoTo cmd89067_err

    Dim UsableWidth  As Single   'Available width for text display area

    Dim UsableHeight As Single  'Available height for text display area

    If WindowState = vbMinimized Then Exit Sub

    'Don't let form go below a minimum set size
    '(this could be handled better with subclassing but this works ok)
    If Me.Width < 3000 Then Me.Width = 3000
    If Me.Height < 3000 Then Me.Height = 3000
    
    'Size and position OK button
    cmdTextOK.Move 100, Me.ScaleHeight - (cmdTextOK.Height + 100), Me.ScaleWidth - 200
    
    'Calculate remaining display area available for text
    UsableWidth = ScaleWidth - VBar.Width
    UsableHeight = ScaleHeight - HBar.Height - (cmdTextOK.Height + 200)
    
    'Size and position picturebox and scrollbars
    TextPict.Move 0, 0, UsableWidth, UsableHeight
    HBar.Move 0, UsableHeight, UsableWidth
    VBar.Move UsableWidth, 0, VBar.Width, UsableHeight
    
    'Set scroll bar properties
    ScreenWidth = TextPict.ScaleWidth / CharWidth

    With HBar
        .Min = 1
        .LargeChange = ScreenWidth - 1
        .max = MaxLineLength - HBar.LargeChange + 1

        If .max < .Min Then .max = .Min
        .SmallChange = 1

    End With

    ScreenLines = (TextPict.ScaleHeight / CharHeight)

    With VBar
        .Min = 0
        .LargeChange = ScreenLines - 1
        .max = UBound(TextLine) - ScreenLines + 1

        If .max < .Min Then .max = .Min
        .SmallChange = 1

    End With
    
    UpdateView  'Display a chunk of text
    Exit Sub
cmd89067_err:
    'MsgBox "Falta memoria ", 48, "Aviso"
    ejecutawor = 1
    'genver.Hide
    'Unload genver
    Exit Sub

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReDim TxtLine(0)    'redundant, but what the heck

End Sub

Private Sub cmdTextOK_Click()
    genver.Hide
    Unload genver

End Sub

Private Sub gu78h_Click()

    Dim snewdbase As String

    Dim antdir    As String

    Dim drive     As String

    snewdbase = sSelectDbase(CmDialog1, "NEW")

    If snewdbase = "" Then Exit Sub
    If Not bFileExists(snewdbase) Then
        copiar_archivo snewdbase, FileName
    Else
        MsgBox "YA EXISTE NOMBRE ARCHIVO", 24, "AVISO"

    End If

    Exit Sub

End Sub

Function sSelectDbase(cmdialog As CommonDialog, sMode As String) As String

    Dim dB As Database

    On Error Resume Next

    sMode = UCase$(sMode)
    cmdialog.DefaultExt = "txt"
    cmdialog.FileName = ""
    cmdialog.CancelError = True
    cmdialog.Filter = "Texto (*.txt)|*.txt|All files (*.*)|*.*|"
    cmdialog.Flags = &H4& Or &H1000& 'remove readonly checkbox

    Select Case sMode

        Case "NEW"
            cmdialog.DialogTitle = "GUARDAR EN ..."
            cmdialog.Action = 2

        Case "OPEN"
            cmdialog.DialogTitle = "ABRIR "
            cmdialog.Action = 1

    End Select

    If Err <> 32755 Then    'i.e not cancel
        sSelectDbase = cmdialog.FileName
    Else
        sSelectDbase = ""

    End If
    
End Function

Private Sub HBar_Change()
    UpdateView

End Sub

Private Sub HBar_Scroll()
    UpdateView

End Sub

Sub envio_correos(perfil As String)

    Dim txtserver     As String

    Dim txtusername   As String

    Dim txtpassword   As String

    Dim txtport       As String

    Dim txtto         As String

    Dim chkssl        As String

    Dim txtfromname   As String

    Dim txtfromemail  As String

    Dim txtattach     As String

    Dim txtsubject    As String

    Dim txtmsg        As String

    Dim retval        As String

    Dim txthtml       As String

    Dim txtselecciona As String

    'Dim txtselecciona As String
    Dim mytablex      As New ADODB.Recordset

    Dim buf           As String

    On Error GoTo cmd0905677_err

    '23/06/2017 kenyo CORRECCION envio de correo automatico
    mytablex.Open "select * from correos where cosms='11'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        MsgBox "Correo No Configurado", vbCritical, "Message"

    End If

    '23/06/2017 kenyo CORRECCION envio de correo automatico

    If mytablex.RecordCount > 0 Then
        txtserver = Trim("" & mytablex.Fields("txtserver"))
        txtusername = Trim("" & mytablex.Fields("txtusername"))
        txtpassword = Trim("" & mytablex.Fields("txtpassword"))
        txtfromname = Trim("" & mytablex.Fields("txtfromname"))
        txtfromemail = Trim("" & mytablex.Fields("txtfromemail"))
        txtport = Trim("" & mytablex.Fields("txtport"))
        txtselecciona = Trim("" & mytablex.Fields("txtselecciona"))
        'txtto = Trim("" & mytablex.Fields("txtto"))
        chkssl = Trim("" & mytablex.Fields("chkssl"))
        'txtfromname = Trim("" & nombre) 'Trim("" & mytablex.Fields("txtfromname"))
        txtto = Trim("" & mytablex.Fields("txtfromemail"))
        txtattach = FileName 'Trim("" & mytablex.Fields("txtattach"))
        txtsubject = Trim("" & mytablex.Fields("txtsubject"))
        txtmsg = Trim("" & mytablex.Fields("txtmsg"))
        txtmsg = txtmsg & Chr$(10) & Chr$(13) & ""
        txtmsg = txtmsg & Format(Now, "dd/mm/yyyy") + " " + Format(Now, "hh:mm:ss")

        If Len(Trim("" & mytablex.Fields("txtfromemail"))) > 0 Then
            txtto = Trim("" & mytablex.Fields("txtfromemail"))
            retval = SendMail(Trim$(txtto), Trim$(txtsubject), Trim$(txtfromname) & "<" & Trim$(txtfromemail) & ">", Trim$(txtmsg), Trim$(txtserver), CInt(Trim$(txtport)), Trim$(txtusername), Trim$(txtpassword), Trim$(txtattach), True, txtselecciona, txthtml)
   
        End If

        MsgBox "Correo Enviado ", 48, "Aviso"

    End If

    mytablex.Close

    Exit Sub
cmd0905677_err:
    MsgBox "No se Pudo enviar Correo... " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub k8923_Click()
    envio_correos "11"

    'tmail.adjunto.Caption = file
    'tmail.Show 1
End Sub

Private Sub ldfo2323_Click()
    cmdTextOK_Click

End Sub

Private Sub tamo9912_Click()

    Dim buf As String

    tipoletra = InputBox("Ingrese Letra", buf, tipoletra)

    If Val(tipoletra) = 0 Then
        tipoletra = "8"

    End If

End Sub

Private Sub VBar_Change()
    
    UpdateView

End Sub

Private Sub VBar_Scroll()
    
    UpdateView

End Sub

Private Sub UpdateView()

    Dim counter As Long

    Dim txt     As String
    
    'Output window of text into picturebox
    'This routine gets called anytime the user scrolls the text window
    
    'NOTE: This routine uses the TextOut API call to send the text
    'rather than the 'Print' command simply because some lower
    'ASCII characters are not 'Print'able but will be displayed
    'using TextOut. If this does not matter, the commented line
    'below could be used instead of the two statements following it.
    
    On Error Resume Next

    TextPict.Cls

    For counter = 0 To (ScreenLines - 1)
        txt = Mid$(TextLine(VBar.Value + counter), HBar.Value, ScreenWidth + 1)
        'TextPict.Print Txt
        TextOut TextPict.hDC, TextPict.CurrentX, TextPict.CurrentY, Mid$(TextLine(VBar.Value + counter), HBar.Value), Len(txt)
        TextPict.Print
    Next counter

    TextPict.refresh
    
End Sub

Sub agregar_menu(buf As String)
    Agregarm buf, mnuArchivoArray

End Sub

Sub Agregarm(TextoDeMenu As String, QueMenu As Object)

    Dim indice As Integer

    indice = QueMenu.count
    Load QueMenu(indice)
    QueMenu(indice).Caption = TextoDeMenu
    QueMenu(indice).Visible = True

End Sub

Function IsComPortAvailable(ByVal portNum As Integer) As Boolean

    Dim fnum As Integer

    On Error Resume Next

    fnum = FreeFile
    Open "COM" & CStr(portNum) For Binary Shared As #fnum

    If Err = 0 Then
        Close #fnum
        IsComPortAvailable = True

    End If

End Function

' Check whether a given LPT parallel port is available
Function IsLptPortAvailable(ByVal portNum As Integer) As Boolean

    Dim fnum As Integer

    On Error Resume Next

    fnum = FreeFile
    Open "LPT" & CStr(portNum) For Binary Shared As #fnum

    If Err = 0 Then
        Close #fnum
        IsLptPortAvailable = True

    End If

End Function

Sub mnuarchivoarray_click(Index As Integer)

    Dim found   As Integer

    Dim xpuerto As String

    Dim buf     As String

    buf = mnuArchivoArray(Index).Caption

    xpuerto = buf
    'xpuerto = busca_usuario(gusuario)
    'If Len(xpuerto) = 0 Then
    '   MsgBox "Puerto no Definido ", 48, "Aviso"
    '   Exit Sub
    'End If
    buf = "Impresion directa: Debe de encontrarse definido en Personal el puerto por defecto de impresion " + Chr$(10) + Chr$(13)
    buf = buf & " El puerto definido es " + xpuerto + Chr$(10) + Chr$(13)
    buf = buf & "Desea imprimir "

    If MsgBox(buf, 1, "Aviso") <> 1 Then Exit Sub
    found = star_sp342(xpuerto, 1)
    found = corte_papel(xpuerto, 1)

End Sub
