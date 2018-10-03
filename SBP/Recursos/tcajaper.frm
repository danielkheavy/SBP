VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tcajaper 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Tabla de Aperturas de Caja"
   ClientHeight    =   9825
   ClientLeft      =   165
   ClientTop       =   15
   ClientWidth     =   14730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   14730
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   8775
      Left            =   30
      TabIndex        =   11
      Top             =   30
      Visible         =   0   'False
      Width           =   12495
      Begin VB.TextBox fechai 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   17
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox caja 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   14
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   10800
         Picture         =   "tcajaper.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Imprimir todo"
         Top             =   2040
         Width           =   1470
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   10800
         Picture         =   "tcajaper.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1080
         Width           =   1470
      End
      Begin VB.Label Label9 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HORA COMPUTADORA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   240
         TabIndex        =   29
         Top             =   5400
         Width           =   2415
      End
      Begin VB.Label bhora 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1935
         Left            =   2640
         TabIndex        =   28
         Top             =   5400
         Width           =   8775
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DD/MM/YYYY"
         Height          =   735
         Left            =   4560
         TabIndex        =   27
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label bfecha 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1935
         Left            =   2640
         TabIndex        =   26
         Top             =   3480
         Width           =   8775
      End
      Begin VB.Label Label6 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FECHA COMPUTADORA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   240
         TabIndex        =   25
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label mcaja 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   615
         Index           =   5
         Left            =   8400
         TabIndex        =   24
         Top             =   960
         Width           =   975
      End
      Begin VB.Label mcaja 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   615
         Index           =   4
         Left            =   7440
         TabIndex        =   23
         Top             =   960
         Width           =   975
      End
      Begin VB.Label mcaja 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   615
         Index           =   3
         Left            =   6480
         TabIndex        =   22
         Top             =   960
         Width           =   975
      End
      Begin VB.Label mcaja 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   615
         Index           =   2
         Left            =   5520
         TabIndex        =   21
         Top             =   960
         Width           =   975
      End
      Begin VB.Label mcaja 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   615
         Index           =   1
         Left            =   4560
         TabIndex        =   20
         Top             =   960
         Width           =   975
      End
      Begin VB.Label mcaja 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   615
         Index           =   0
         Left            =   3600
         TabIndex        =   19
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio Venta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   18
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Caja"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   2415
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   12435
      TabIndex        =   2
      Top             =   0
      Width           =   12495
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1200
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tcajaper.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Grabar registro"
         Top             =   0
         Width           =   1575
      End
      Begin VB.TextBox buffer 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         MaxLength       =   2
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   480
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7200
         Style           =   2  'Dropdown List
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   120
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000D&
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10560
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2760
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tcajaper.frx":23A6
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Borrar registro"
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   5760
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tcajaper.frx":35B8
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir"
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4320
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tcajaper.frx":47CA
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Imprimir"
         Top             =   -120
         Width           =   1455
      End
      Begin VB.CommandButton cmdAddEntry 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   0
         Picture         =   "tcajaper.frx":59DC
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label xusuario 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Height          =   195
         Left            =   3720
         TabIndex        =   16
         Top             =   120
         Width           =   45
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Consulta"
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   12495
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   7335
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   12938
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   35
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "Caja"
            Caption         =   "Caja"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "fechai"
            Caption         =   "Fechai"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   2910.047
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3344.882
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cierre Caja"
      Height          =   555
      Left            =   12585
      TabIndex        =   30
      Top             =   180
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Menu ajdu1 
      Caption         =   "&Add"
   End
   Begin VB.Menu f8443 
      Caption         =   "&Modifica"
   End
   Begin VB.Menu bo712 
      Caption         =   "&Borrar"
   End
   Begin VB.Menu fjh433 
      Caption         =   "&Zoom"
   End
   Begin VB.Menu djuer1 
      Caption         =   "&Reporte"
      Begin VB.Menu dk9893 
         Caption         =   "&0.GENERAL"
      End
      Begin VB.Menu mnuArchivoArray 
         Caption         =   "Novisible"
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tcajaper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim txempreca As New ADODB.Recordset

Private Sub ajdu1_Click()

    Dim mytablex As New ADODB.Recordset

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    inicializa
    fechai.Enabled = True
    Frame2.Visible = True
    Frame2.Caption = "Nuevo"
    cmdGuardar.Enabled = True
    habilita 1
    caja = buffer
    'caja.Enabled = True
    'turno.Enabled = True
    'mytablex.Open "select * from apertura where caja='" & Trim(caja) & "'", cn, adOpenStatic, adLockOptimistic
    'If mytablex.RecordCount > 0 Then
    'fechai = "" & mytablex.Fields("fechai")
    'fechaf = "" & mytablex.Fields("fechaf")
    'Else
    'End If
    'If IsDate(fechai) Then
    'fechai.Enabled = False
    'End If
    'mytablex.Close
    fechai.SetFocus

End Sub

Private Sub bo712_Click()

    Dim buf As String

    On Error GoTo cmd656_err

    buf = txempreca.Fields("caja")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    If MsgBox("Desea Borra " + txempreca.Fields("caja"), 1, "Aviso") <> 1 Then
        Exit Sub

    End If

    txempreca.Delete
    Command1_Click
    Exit Sub
cmd656_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub buffer_DblClick()
    Command1_Click

End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    Command1_Click

End Sub

Private Sub cmdAddEntry_Click()
    ajdu1_Click

End Sub

Private Sub cmdCerrar_Click()
    dlo132_Click

End Sub

Private Sub cmdDelete_Click()
    bo712_Click

End Sub

Private Sub cmdExit_Click()
    dlo132_Click

End Sub

Private Sub cmdHelp_Click()

End Sub

Private Sub cmdGuardar_Click()

    Dim found As Integer

    found = grabar()

End Sub

Private Sub cmdPrint_Click()

    'djuer1_Click
End Sub

Private Sub cmdSave_Click()
    f8443_Click

End Sub

Private Sub dk9893_Click()

    If Frame2.Visible = True Then Exit Sub
    reporgen.NAMETABLA = "apertura"
    reporgen.Show 1

End Sub

Sub prueba_reporte()

    'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\cajaesproducto.rpt", "")
End Sub

Private Sub CAJA_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        dlo132_Click
        Exit Sub

    End If

    If Len(caja) = 0 Then Exit Sub

    'turno.SetFocus
End Sub

Private Sub Command1_Click()
    Frame1.Visible = True
    Frame1.Enabled = True
    'buffer = ""
    opcion1 = "1"
    ejecuta 1

End Sub

Sub ejecuta(sw As Integer)

    Dim cad As String

    Dim buf As String

    If opcion1 = "1" Then  'bodega
        If Len(buffer) = 0 Then
            cad = "SELECT * from apertura  "

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT *  from apertura   where caja='" & buffer & "'"

        End If

        If txempreca.State = 1 Then txempreca.Close
        txempreca.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbGrid1.DataSource = txempreca

        'dbGrid1.columns(0).Width = 4000
        'dbGrid1.columns(1).Width = 2000
        If txempreca.RecordCount > 0 Then
            dbGrid1.SetFocus

        End If

    End If

End Sub

Private Sub Command2_Click()

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        'buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then

        'caja = dbGrid1.Columns(1)
        'Frame1.Visible = False
        'Frame1.Enabled = False
        'caja.SetFocus
        'caja_KeyPress 13
    End If

End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)

    Dim buf  As String

    Dim buf2 As String

    If KeyAscii <> 13 And KeyAscii <> 27 Then
        If KeyAscii = 8 Then
            If Len(buffer) > 0 Then
                buf = Mid$(buffer, 1, Len(buffer) - 1)
                buffer = buf
                KeyAscii = 0
            Else
                KeyAscii = 0
                Exit Sub

            End If

        End If

        buf = Chr(KeyAscii)

        If Chr(KeyAscii) = "*" Then
            buf = ""
            buffer = buf

        End If

        If KeyAscii <> 13 Then
            buffer = buffer + buf

        End If

        buf = buffer
        ejecuta 0
         
    End If

End Sub

Private Sub dlo132_Click()

    If Frame2.Visible = True Then
        habilita 0
        Frame2.Visible = False
        dbGrid1.Enabled = True
        Exit Sub

    End If

    tcajaper.Hide
    Unload tcajaper

End Sub

Private Sub f8443_Click()

    Dim buf As String

    On Error GoTo cmd456_err

    buf = txempreca.Fields("caja")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    inicializa
    Frame2.Visible = True
    Frame2.Caption = "Modifica"
    cmdGuardar.Enabled = True
    pone_registro
    habilita 1
    caja.Enabled = False
    'turno.Enabled = False
    fechai.Enabled = True
    'If IsDate(fechai) Then
    '   fechai.Enabled = False
    '   Exit Sub
    'End If
    fechai.SetFocus
    Exit Sub
cmd456_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fechai_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If Len(fechai) = 0 Then
        fechai = Format(Now, "dd/mm/yyyy")

        'fechaf = Format(Now, "dd/mm/yyyy")
    End If

    If Not IsDate(fechai) Then
        fechai.SetFocus
        Exit Sub

    End If

    found = busca_turno()
    'fechaf.SetFocus

End Sub

Private Sub fjh433_Click()

    Dim buf As String

    On Error GoTo cmd556_err

    buf = txempreca.Fields("caja")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    inicializa
    Frame2.Visible = True
    Frame2.Caption = "Zoom"
    cmdGuardar.Enabled = False
    pone_registro
    habilita 1
    caja.Enabled = False
    'turno.Enabled = False
    fechai.SetFocus
    Exit Sub
cmd556_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Form_Activate()
    'agregar_menus
    Frame2.Top = 10: Frame2.Left = 10
    Command1_Click

End Sub

Private Sub Form_Load()
    Combo1.Clear
    Combo1.AddItem "caja"
    Combo1.ListIndex = 0
    xusuario = gusuario
    'carga_caja
    'carga_turno
    bfecha = Format(Now, "dd/mm/yyyy")
    bhora = Format(Now, "hh:mm:ss")

End Sub

Sub inicializa()
    caja = ""
    'cajero = "" & xusuario
    'turno = ""
    fechai = ""
    'fechaf = ""
    'fechai = Format(Now, "dd/mm/yyyy")
    'fechaf = Format(Now, "dd/mm/yyyy")

End Sub

Sub pone_registro()
    'cajero = Trim("" & txempreca.Fields("cajero"))
    caja = Trim("" & txempreca.Fields("caja"))
    'turno = Trim("" & txempreca.Fields("turno"))
    fechai = Trim("" & txempreca.Fields("fechai"))

    'fechaf = Trim("" & txempreca.Fields("fechaf"))
End Sub

Sub grabando()
    txempreca.Fields("caja") = Trim(caja)
    'txempreca.Fields("turno") = Trim(turno)
    'txempreca.Fields("cajero") = Trim(cajero)
    txempreca.Fields("fechai") = Format(fechai, "dd/mm/yyyy")
    txempreca.Fields("fechaf") = Format(fechai, "dd/mm/yyyy")

End Sub

Private Sub grba1_Click()

End Sub

Function grabar()

    Dim found  As Integer

    Dim rbusca As New ADODB.Recordset

    found = valida()

    If found = 0 Then
        MsgBox "Campos invalidos", 48, "Aviso"
        Exit Function

    End If

    If Frame2.Caption = "Nuevo" Then
        If Len(caja) = 0 Then
            caja.SetFocus
            Exit Function

        End If

        rbusca.Open "select caja from apertura where  caja='" & caja & "'", cn, adOpenStatic, adLockOptimistic

        If rbusca.RecordCount > 0 Then
            rbusca.Close
            MsgBox "Ya existe caja  ", 48, "Aviso"
            Exit Function

        End If

        txempreca.AddNew
        'txempreca.Fields("caja") = caja
        grabando
        txempreca.Update
        'grabar en parametro
        'grabar_parameca
        dlo132_Click
        Exit Function

    End If

    If Frame2.Caption = "Modifica" Then
        txempreca.Fields("caja") = caja
        grabando
        txempreca.Update
        dlo132_Click
        Exit Function

    End If

End Function

Function valida()

    Dim found As Integer

    If Len(caja) = 0 Then
        caja.SetFocus
        Exit Function

    End If

    'If Len(turno) = 0 Then
    '   turno.SetFocus
    '   Exit Function
    'End If
    If Len(fechai) = 0 Then
        fechai.SetFocus
        Exit Function

    End If

    'If Len(fechaf) = 0 Then
    '   fechai.SetFocus
    '   Exit Function
    'End If

    If Not IsDate(fechai) Then
        fechai.SetFocus
        Exit Function

    End If

    'found = busca_turno()

    'If Not IsDate(fechaf) Then
    '   fechai.SetFocus
    '   Exit Function
    'End If
    'If Len(descripcio) = 0 Then
    '   descripcio.SetFocus
    '   Exit Function
    'End If
    valida = 1

End Function

Sub habilita(sw As Integer)

    If sw = 0 Then

        ajdu1.Enabled = True
        f8443.Enabled = True
        bo712.Enabled = True
        fjh433.Enabled = True
        djuer1.Enabled = True
        djuer1.Enabled = True
        Picture1.Enabled = True
        dbGrid1.Enabled = True
            
    End If

    If sw = 1 Then

        ajdu1.Enabled = False
        f8443.Enabled = False
        bo712.Enabled = False
        fjh433.Enabled = False
        djuer1.Enabled = False
        djuer1.Enabled = False
        Picture1.Enabled = False
        dbGrid1.Enabled = False
        dbGrid1.Enabled = False
           
    End If
      
End Sub

Sub agregar_menus()

    Dim I As Integer

    For I = 1 To mnuArchivoArray.count - 1
        Unload mnuArchivoArray(I)
    Next
     
    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from archivo where menu='caja' and   estado='S'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    Do

        If mytablex.EOF Then Exit Do
        Agregarm "" & mytablex.Fields("descripcio"), mnuArchivoArray
        mytablex.MoveNext
    Loop
    mytablex.Close

End Sub

Sub Agregarm(TextoDeMenu As String, QueMenu As Object)

    Dim indice As Integer

    'MsgBox QueMenu.count
    indice = QueMenu.count

    Load QueMenu(indice)

    QueMenu(indice).Caption = TextoDeMenu
    QueMenu(indice).Visible = True

End Sub

Private Sub Label7_Click()

    If Len(caja) = 0 Then
        caja.SetFocus
        Exit Sub

    End If

    'If Len(turno) = 0 Then
    '   turno.SetFocus
    '   Exit Sub
    'End If
    fechai = Format(Now, "dd/mm/yyyy")
    fechai_KeyPress 13

End Sub

Private Sub Label8_Click()

    Dim jcaja   As String

    'Dim jturno As String
    'Dim jcajero As String
    Dim jfechai As String

    'Dim jfechaf As String
    On Error GoTo cmd90012_err

    jcaja = "" & txempreca.Fields("caja")
    'jturno = "" & txempreca.Fields("turno")
    'jcajero = "" & txempreca.Fields("cajero")
    jfechai = "" & txempreca.Fields("fechai")
    'jfechaf = "" & txempreca.Fields("fechaf")
    
    opcion1 = "5"
    opcion2 = "1"
    opcion3 = "2"
    
    tcuadrc1.fechai.Enabled = True
    tcuadrc1.fechaf.Enabled = True
    usuariopos = gusuario
    tcuadrc1.tipoexterno.Visible = True
    tcuadrc1.numcuadre.Visible = False
    'tcuadrc1.flagdiario = "1"
    'tcuadrc1.cajero = jcajero
    tcuadrc1.caja = jcaja
    'tcuadrc1.turno = jturno
    
    tcuadrc1.fechai = Format(jfechai, "dd/mm/yyyy")
    'tcuadrc1.fechaf = Format(jfechaf, "dd/mm/yyyy")
    tcuadrc1.horai = "01"
    tcuadrc1.horaf = "24"
    tcuadrc1.Caption = "COPIA CIERRE DEL DIA"
    'tcuadrc1.pantalla = "PANTALLA"
    tcuadrc1.Show 1
    Command1_Click
    
    Exit Sub
cmd90012_err:
    MsgBox "No existen datos ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub mcaja_Click(Index As Integer)

    If caja.Enabled = True Then
        caja = mcaja(Index)
        busca_fechaapertura

    End If

End Sub

Sub mnuarchivoarray_click(Index As Integer)

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    buf = mnuArchivoArray(Index).Caption
    mytablex.Open "select * from archivo where menu='caja' and descripcio='" & Trim(buf) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close

    End If

    'busca el reporte
    buf = mytablex.Fields("archivo")
    mytablex.Close
    'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\" & buf, "")

End Sub

Private Sub mturno_Click(Index As Integer)

    'If turno.Enabled = True Then
    '   turno = mturno(Index)
    '   busca_turno
    'End If
End Sub

Function busca_turno()

    'Dim mytablex As New ADODB.Recordset
    'mytablex.Open "select * from turno where turno='" & turno & "'", cn, adOpenStatic, adLockOptimistic
    'If mytablex.RecordCount > 0 Then
    '   If "" & mytablex.Fields("flag") = "2" Then
    '     fechaf = Format(CVDate(fechai) + 1, "dd/mm/yyyy")
    '   End If
    '   If "" & mytablex.Fields("flag") = "1" Then
    '     fechaf = Format(CVDate(fechai), "dd/mm/yyyy")
    '   End If
    '   If "" & mytablex.Fields("flag") = "0" Then
    '     fechaf = Format(CVDate(fechai), "dd/mm/yyyy")
    '   End If
    '   busca_turno = 1
    'End If
    'mytablex.Close
End Function

Sub carga_caja()

    Dim mytablex As New ADODB.Recordset

    Dim I        As Integer

    I = 0

    For I = 0 To 5
        mcaja(I) = ""
    Next I

    I = 0
    mytablex.Open "select * from parameca ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        If "" & mytablex.Fields("terminal") = "C" Then
            If "" & mytablex.Fields("caja") <> "00" Then
                mcaja(I).Caption = "" & mytablex.Fields("caja")
                I = I + 1

            End If

            If I > 5 Then Exit Do

        End If

        mytablex.MoveNext
    Loop
    mytablex.Close

End Sub

Sub carga_turno()
    'Dim mytablex As New ADODB.Recordset
    'Dim i As Integer
    'i = 0
    'For i = 0 To 3
    '    mturno(i) = ""
    'Next i
    'i = 0
    'mytablex.Open "select * from turno ", cn, adOpenStatic, adLockOptimistic
    'Do
    'If mytablex.EOF Then Exit Do
    '      mturno(i).Caption = "" & mytablex.Fields("turno")
    '      i = i + 1
    '   If i > 3 Then Exit Do
    'mytablex.MoveNext
    'Loop
    'mytablex.Close

End Sub

Sub busca_fechaapertura()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from apertura where caja='" & caja & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        fechai = Format("" & mytablex.Fields("fecha"), "dd/mm/yyyy")

    End If

    mytablex.Close

End Sub

