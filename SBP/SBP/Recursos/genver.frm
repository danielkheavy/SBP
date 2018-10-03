VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form genverx 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver Archivos"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   13650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   13650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Click para cancelar Impresion..."
      Height          =   1095
      Left            =   1320
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox txtfile 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   6240
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   1
      Top             =   2400
      Width           =   7335
   End
   Begin VB.VScrollBar vsbFile 
      Height          =   5175
      LargeChange     =   45
      Left            =   7320
      Max             =   100
      Min             =   1
      TabIndex        =   0
      Top             =   0
      Value           =   1
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu irm12 
      Caption         =   "&Imprime"
   End
   Begin VB.Menu jum343 
      Caption         =   "&Letra"
   End
   Begin VB.Menu dflo23 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "genverx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fileBuff()  As Variant       'array of buffers to store file
Dim numBuffTotal%, numBuffNow%
Dim numlinestotal%, saveLineNum%
Dim buffLines%, buffBytes%, extrabytes%

Private Sub Command1_Click()
Command1.Visible = False
opcion3 = 1
End Sub

Private Sub dflo23_Click()
If Command1.Visible = True Then
   opcion3 = 1
   Command1.Visible = False
   Exit Sub
End If
cerrar_archivo
genver.Hide
Unload genver
End Sub


Private Sub Form_Activate()


'Filename = "\rp_orion\001d\01\temporal\JOHNNY.txt"
Filename = globaldir & "\temporal\" & gusuario & ".txt"
apertura_automatico
End Sub

Private Sub Form_Load()
tipoletra = "9"
End Sub

Sub Form_Resize()
    Dim i%, numvisible%
    txtfile.Top = 0
    txtfile.Left = 0
    txtfile.Height = ScaleHeight
    txtfile.Width = ScaleWidth - vsbFile.Width
    
    vsbFile.Top = 0
    vsbFile.Left = txtfile.Width
    vsbFile.Height = txtfile.Height - vsbFile.Width  'for txtFile's hsb
    'numvisible = GetVisibleLines(txtfile)
    numvisible = 45
    vsbFile.LargeChange = numvisible - 1
    If vsbFile.LargeChange >= numlinestotal Then   'if all lines visible
        vsbFile.max = 1
    Else
        vsbFile.max = numlinestotal - vsbFile.LargeChange
    End If
End Sub

Sub apertura_automatico()
    Dim nextline$, ndx&, linenum%, buff$, Msg$, numErr%, avelength%
    If Dir(Filename) = "" Then Exit Sub
    If Filename = "" Then Exit Sub  'Open-Dialog Canceled
    avelength = AVELINELENGTH
    Screen.MousePointer = HOURGLASS

startNewFile1:

'extraBytes is number of bytes which will be added to each buffer
'from the next buffer in line, to be displayed in the textbox, below
'the "last" line of current buffer.
    extrabytes = MAXLINESVISIBLE * (avelength + 2)
    buffBytes = 30000 - extrabytes
    buffLines = buffBytes \ (avelength + 2)
    'reset before possible re-Open
    Close #1
    numBuffTotal = 0
    numlinestotal = 0
    Erase fileBuff
    Open Filename For Input As #1
    Do Until EOF(1)
        buff = Space$(buffBytes)
        linenum = 0
        ndx = 1
        On Error GoTo errorRead1
        Do Until linenum = buffLines Or EOF(1)
            linenum = linenum + 1
            Line Input #1, nextline
            nextline = nextline & Chr(13) & Chr(10)
            Mid$(buff, ndx, Len(nextline)) = nextline
            ndx = ndx + Len(nextline)
        Loop
        On Error GoTo 0
        numlinestotal = (numlinestotal + linenum)
        If linenum > 0 Then        'at least one line
            numBuffTotal = numBuffTotal + 1       'starts at one
            ReDim Preserve fileBuff(numBuffTotal)
            fileBuff(numBuffTotal - 1) = RTrim$(buff)
            buff = ""
        End If
    Loop
    Screen.MousePointer = DEFAULT
    numBuffNow = 1
    If vsbFile.LargeChange >= numlinestotal Then  'all lines visible
        vsbFile.max = 1           'disable the vert scroll bar
    Else
        vsbFile.max = numlinestotal - vsbFile.LargeChange
    End If

    If numBuffNow = numBuffTotal Then  ' if only one buffer
        txtfile.Text = fileBuff(numBuffNow - 1)
    Else
        txtfile.Text = fileBuff(numBuffNow - 1) & Left$(fileBuff(numBuffNow), extrabytes)
    End If
    Caption = Filename
    vsbFile.Value = 1
    'vsbFile.SetFocus
    saveLineNum = 1  'start at first line
    Exit Sub
errorRead1:
    numErr = numErr + 1
    Beep
    If Err = 5 And numErr <= 5 Then    'could not fit into file buffer
        avelength = 1.25 * avelength    'so try less lines per buffer
        Resume startNewFile1
    End If
    Msg = "ERROR During File Read !" & Chr(13) & Chr(10)
    Msg = Msg & "Attempts to adjust average line length failed" & Chr(13) & Chr(10)
    Msg = Msg & "HUGE line length? (Try adjusting Const AVELINELENGTH)"
    MsgBox Msg, 16
    End

End Sub



Private Sub Form_Unload(Cancel As Integer)
cerrar_archivo
End Sub

Private Sub jum343_Click()
Dim buf As String
tipoletra = InputBox("Ingrese Letra", buf, tipoletra)
If Val(tipoletra) = 0 Then
   tipoletra = "8"
End If
End Sub

Private Sub txtfile_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
    vsbFile.SetFocus

End Sub

Private Sub vsbFile_Change()
vsbFile_Scroll
End Sub

Private Sub vsbFile_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 37 Or KeyCode = 39 Then  'disable left, right arrows
        KeyCode = 0
    End If

End Sub
Sub vsbFile_Scroll()
    Dim numtoscroll&, l&, numbuffcorrect%

    numtoscroll = vsbFile.Value - saveLineNum   'started at 1
    saveLineNum = vsbFile.Value

    numbuffcorrect = (vsbFile.Value - 1) \ buffLines + 1

    If numBuffNow <> numbuffcorrect Then
        numBuffNow = numbuffcorrect
        If numBuffNow = numBuffTotal Then  ' if no more buffers
            txtfile.Text = fileBuff(numBuffNow - 1)
        Else
          txtfile.Text = fileBuff(numBuffNow - 1) & Left$(fileBuff(numBuffNow), extrabytes)
        End If
        numtoscroll = vsbFile.Value - ((numBuffNow - 1) * buffLines) - 1
    End If
    l = SendMessageByNum(txtfile.hWnd, EM_LINESCROLL, 0, numtoscroll)
End Sub




Private Sub irm12_Click()
Dim found As Integer
Dim sfile As String
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
sfile = globaldir & "\temporal\" & gusuario & ".txt"
found = imprime_archivoj(sfile, 0)
Command1.Visible = False
opcion3 = 0
dflo23_Click
End Sub
