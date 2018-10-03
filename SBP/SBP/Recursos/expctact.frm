VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form expctact 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimiento de Cuenta Corriente"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   14760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   14760
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Mensajes"
      ForeColor       =   &H00FFFFFF&
      Height          =   3360
      Left            =   3030
      TabIndex        =   39
      Top             =   2940
      Visible         =   0   'False
      Width           =   9015
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "ESPERE UN MOMENTO..PROCESANDO..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   480
         TabIndex        =   40
         Top             =   720
         Width           =   7935
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "expctact.frx":0000
      Height          =   6735
      Left            =   120
      OleObjectBlob   =   "expctact.frx":0014
      TabIndex        =   28
      Top             =   1320
      Width           =   14535
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   14160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   1275
      Left            =   120
      ScaleHeight     =   1215
      ScaleWidth      =   14700
      TabIndex        =   0
      Top             =   0
      Width           =   14760
      Begin VB.TextBox nombre 
         Height          =   375
         Left            =   2520
         MaxLength       =   11
         TabIndex        =   41
         Text            =   "%"
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox local1 
         Height          =   375
         Left            =   8040
         MaxLength       =   3
         TabIndex        =   26
         Text            =   "%"
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox codigo 
         Height          =   375
         Left            =   2520
         MaxLength       =   13
         TabIndex        =   24
         Text            =   "%"
         Top             =   120
         Width           =   1815
      End
      Begin VB.ComboBox ordenado 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10920
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   12960
         MaskColor       =   &H00FFFFFF&
         Picture         =   "expctact.frx":20E7
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.ComboBox tipo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10920
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox fechai 
         Height          =   375
         Left            =   5280
         MaxLength       =   10
         TabIndex        =   12
         Top             =   120
         Width           =   1815
      End
      Begin VB.TextBox fechaf 
         Height          =   375
         Left            =   5280
         MaxLength       =   10
         TabIndex        =   11
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox cajero 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8040
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox caja 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8040
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   120
         Width           =   1815
      End
      Begin VB.ComboBox turno 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10920
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   120
         Width           =   1815
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
         Height          =   615
         Left            =   720
         MaskColor       =   &H00E0E0E0&
         Picture         =   "expctact.frx":2895
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salir"
         Top             =   0
         Width           =   735
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
         Height          =   615
         Left            =   0
         MaskColor       =   &H00E0E0E0&
         Picture         =   "expctact.frx":3AA7
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprimir"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         TabIndex        =   42
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7080
         TabIndex        =   27
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         TabIndex        =   25
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ordenado Por"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9840
         TabIndex        =   19
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoDocto"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9840
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4320
         TabIndex        =   14
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label12 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4320
         TabIndex        =   13
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label11 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Turno"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9840
         TabIndex        =   10
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Caja"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7080
         TabIndex        =   9
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label9 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cajero"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7080
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Label Label21 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "US$"
      Height          =   375
      Left            =   4680
      TabIndex        =   38
      Top             =   8520
      Width           =   975
   End
   Begin VB.Label Label20 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S/."
      Height          =   375
      Left            =   4680
      TabIndex        =   37
      Top             =   8160
      Width           =   975
   End
   Begin VB.Label Label19 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cargo"
      Height          =   375
      Left            =   5640
      TabIndex        =   36
      Top             =   8520
      Width           =   975
   End
   Begin VB.Label Label18 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo"
      Height          =   375
      Left            =   10560
      TabIndex        =   35
      Top             =   8520
      Width           =   975
   End
   Begin VB.Label cargod 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6600
      TabIndex        =   34
      Top             =   8520
      Width           =   1335
   End
   Begin VB.Label Label16 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Abono"
      Height          =   375
      Left            =   7920
      TabIndex        =   33
      Top             =   8520
      Width           =   975
   End
   Begin VB.Label abonod 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   8880
      TabIndex        =   32
      Top             =   8520
      Width           =   1695
   End
   Begin VB.Label saldod 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11520
      TabIndex        =   31
      Top             =   8520
      Width           =   1455
   End
   Begin VB.Label saldos 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11520
      TabIndex        =   30
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Label abonos 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   8880
      TabIndex        =   29
      Top             =   8160
      Width           =   1695
   End
   Begin VB.Label Label15 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Abono"
      Height          =   375
      Left            =   7920
      TabIndex        =   23
      Top             =   8160
      Width           =   975
   End
   Begin VB.Label cargos 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6600
      TabIndex        =   22
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo"
      Height          =   375
      Left            =   10560
      TabIndex        =   21
      Top             =   8160
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cargo"
      Height          =   375
      Left            =   5640
      TabIndex        =   20
      Top             =   8160
      Width           =   975
   End
   Begin VB.Label afecta 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   13920
      TabIndex        =   4
      Top             =   7440
      Width           =   735
   End
   Begin VB.Label acu 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   13680
      TabIndex        =   3
      Top             =   7440
      Width           =   255
   End
   Begin VB.Menu dki222 
      Caption         =   "&Ejecuta"
   End
   Begin VB.Menu dki232312 
      Caption         =   "&Imprime"
   End
   Begin VB.Menu lfo3434 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "expctact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancelar_Click()

End Sub

Private Sub cmdDelete_Click()
    dbo912_Click

End Sub

Private Sub cmdGrabar_Click()

End Sub

Private Sub cmdSort_Click()

End Sub

Private Sub cmdExit_Click()
    lfo3434_Click

End Sub

Private Sub cmdPrint_Click()

    Dim found As Integer

    If Frame1.Visible = True Then Exit Sub

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    borrar_archivo FileName
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_documento1
    cuerpo_programa_documento1
    '------------------------------------
    Close #1
    cerrar_archivo
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1

End Sub

Private Sub Command1_Click()

    If Frame1.Visible = True Then Exit Sub
    Frame1.Visible = True
    xborrar
    sql_recibos
    Frame1.Visible = False

End Sub

Private Sub dbo912_Click()

End Sub

Private Sub dki9923_Click()

End Sub

Private Sub dnu823_Click()

End Sub

Private Sub dki222_Click()

    If Frame1.Visible = True Then Exit Sub
    Command1_Click

End Sub

Private Sub dki232312_Click()

    If Frame1.Visible = True Then Exit Sub
    cmdPrint_Click

End Sub

Private Sub Form_Activate()
    fechai = Format(Now, "dd/mm/yyyy") '"01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    fechaf = Format(Now, "dd/mm/yyyy")
    carga_inicial

End Sub

Sub carga_inicial()

    Dim mytablex As Table

    cajero.Clear
    cajero.AddItem "%"
    Set mytablex = mydbxglo.OpenTable("vendedor")
    Do

        If mytablex.EOF Then Exit Do
        cajero.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    mytablex.Close
    cajero.ListIndex = 0

    caja.Clear
    caja.AddItem "%"
    Set mytablex = mydbxglo.OpenTable("parameca")
    Do

        If mytablex.EOF Then Exit Do
        caja.AddItem "" & mytablex.Fields("caja") & "|" & mytablex.Fields("DESCRIPCIO")
        mytablex.MoveNext
    Loop
    mytablex.Close
    caja.ListIndex = 0

    turno.Clear
    turno.AddItem "%"
    Set mytablex = mydbxglo.OpenTable("turno")
    Do

        If mytablex.EOF Then Exit Do
        turno.AddItem "" & mytablex.Fields("turno") & "|" & mytablex.Fields("DESCRIPCIO")
        mytablex.MoveNext
    Loop
    mytablex.Close
    turno.ListIndex = 0

    tipo.Clear
    tipo.AddItem "%"
    Set mytablex = mydbxglo.OpenTable("tipo")
    Do

        If mytablex.EOF Then Exit Do
        tipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("DESCRIPCIO")
        mytablex.MoveNext
    Loop
    mytablex.Close
    tipo.ListIndex = 0

End Sub

Private Sub Form_Load()
    ordenado.Clear
    ordenado.AddItem "fecha"
    ordenado.AddItem "tipo"
    ordenado.AddItem "val(numero)"
    ordenado.AddItem "Codigo"
    ordenado.AddItem "Usuario"
    ordenado.AddItem "caja"
    ordenado.AddItem "turno"
    ordenado.AddItem "fpago"
    ordenado.AddItem "orden"
    ordenado.AddItem "observa"
    ordenado.AddItem "descripcio"
    ordenado.AddItem "nombre"
    ordenado.ListIndex = 0

End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

    If Frame1.Visible = True Then
        Frame1.Visible = False

    End If

End Sub

Private Sub lfo3434_Click()

    If Frame1.Visible = True Then
        Frame1.Visible = False
        Exit Sub

    End If

    expctact.Hide
    Unload expctact

End Sub

Sub sql_recibos()

    On Error GoTo cmd37_err

    Dim vr

    Dim found    As Integer

    Dim buf      As String

    Dim mytabley As Table

    Dim mytablex As Snapshot

    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub
    Set mytabley = mydbxglo.OpenTable("_b" + gusuario)
    buf = "select * from recibo where "
    buf = buf & "  fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"

    If local1 <> "%" Then
        buf = buf & " and local='" & local1 & "'"

    End If

    If cajero <> "%" Then
        buf = buf & " and usuario='" & extra_loquesea(cajero) & "'"

    End If

    If tipo <> "%" Then
        buf = buf & " and tipo='" & extra_loquesea(tipo) & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja='" & extra_loquesea(caja) & "'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno='" & extra_loquesea(turno) & "'"

    End If

    If codigo <> "%" Then
        buf = buf & " and codigo like '" & codigo & "'"

    End If

    If nombre <> "%" Then
        buf = buf & " and nombre like '" & nombre & "'"

    End If

    buf = buf & " and tipoclie='C' "
    buf = buf & " and estado='2' "
    buf = buf & " order by " & ordenado & ",str(numero)"
    Set mytablex = mydbxglo.CreateSnapshot(buf)
    Do

        If mytablex.EOF Then Exit Do
        vr = DoEvents()

        If Frame1.Visible = False Then
            Exit Do

        End If

        mytabley.AddNew
        mytabley.Fields("local") = "" & mytablex.Fields("local")
        mytabley.Fields("observa") = "" & mytablex.Fields("observa")
        mytabley.Fields("tipo") = "" & mytablex.Fields("tipo")
        mytabley.Fields("tipoclie") = "" & mytablex.Fields("tipoclie")
        mytabley.Fields("nombret") = busca_tipo("" & mytablex.Fields("tipo"))
        mytabley.Fields("serie") = "" & mytablex.Fields("serie")
        mytabley.Fields("numero") = "" & mytablex.Fields("numero")
        mytabley.Fields("codigo") = "" & mytablex.Fields("codigo")
        mytabley.Fields("nombre") = "" & mytablex.Fields("nombre")
        mytabley.Fields("fecha") = "" & mytablex.Fields("fecha")
        mytabley.Fields("moneda") = "" & mytablex.Fields("moneda")
        mytabley.Fields("usuario") = "" & mytablex.Fields("usuario")
        mytabley.Fields("caja") = "" & mytablex.Fields("caja")
        mytabley.Fields("turno") = "" & mytablex.Fields("turno")

        If "" & mytablex.Fields("acu") = "W" Then 'ingreso
            mytabley.Fields("acu") = "A"
            mytabley.Fields("abono") = Val("" & mytablex.Fields("total"))

        End If

        If "" & mytablex.Fields("acu") = "V" Then 'ingreso
            mytabley.Fields("acu") = "C"
            mytabley.Fields("cargo") = Val("" & mytablex.Fields("total"))

        End If

        mytabley.Update
sigamos:
        mytablex.MoveNext
    Loop
    carga_cuentac mytabley
    mytablex.Close
    mytabley.Close
    'xborrar

    Data2.Connect = "foxpro 2.5;"
    Data2.DatabaseName = globaldir
    Data2.RecordSource = "select * from _b" & gusuario
    Data2.refresh
               
    sumar_recibos
    'DBGrid2.SetFocus
    'MsgBox "xx"
               
    Frame1.Visible = False
    Exit Sub
cmd37_err:
    Frame1.Visible = False
    MsgBox "Error en select " & error$, 48, "Aviso"
    Exit Sub

End Sub

Sub carga_cuentac(mytabley As Table)

    Dim vr

    Dim found    As Integer

    Dim buf      As String

    Dim mytablex As Snapshot

    buf = "select * from cuentac where "
    buf = buf & "  fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"

    If local1 <> "%" Then
        buf = buf & " and local='" & local1 & "'"

    End If

    If cajero <> "%" Then
        buf = buf & " and usuario='" & extra_loquesea(cajero) & "'"

    End If

    If tipo <> "%" Then
        buf = buf & " and tipo='" & extra_loquesea(tipo) & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja='" & extra_loquesea(caja) & "'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno='" & extra_loquesea(turno) & "'"

    End If

    If codigo <> "%" Then
        buf = buf & " and codigo like '" & codigo & "'"

    End If

    If nombre <> "%" Then
        buf = buf & " and nombre like '" & nombre & "'"

    End If

    buf = buf & " order by " & ordenado & ",str(numero)"
    Set mytablex = mydbxglo.CreateSnapshot(buf)
    Do

        If mytablex.EOF Then Exit Do
        vr = DoEvents()

        If Frame1.Visible = False Then
            Exit Do

        End If

        mytabley.AddNew
        mytabley.Fields("local") = "" & mytablex.Fields("local")
        mytabley.Fields("observa") = "CREDITO"
        mytabley.Fields("tipo") = "" & mytablex.Fields("tipo")
        mytabley.Fields("tipoclie") = "" & mytablex.Fields("tipoclie")
        mytabley.Fields("nombret") = busca_tipo("" & mytablex.Fields("tipo"))
        mytabley.Fields("serie") = "" & mytablex.Fields("serie")
        mytabley.Fields("numero") = "" & mytablex.Fields("numero")
        mytabley.Fields("codigo") = "" & mytablex.Fields("codigo")
        mytabley.Fields("nombre") = "" & mytablex.Fields("nombre")
        mytabley.Fields("fecha") = "" & mytablex.Fields("fecha")
        mytabley.Fields("moneda") = "" & mytablex.Fields("moneda")
        mytabley.Fields("usuario") = "" & mytablex.Fields("usuario")
        mytabley.Fields("caja") = "" & mytablex.Fields("caja")
        mytabley.Fields("turno") = "" & mytablex.Fields("turno")
        mytabley.Fields("acu") = "C"
        mytabley.Fields("cargo") = Val("" & mytablex.Fields("total"))
        mytabley.Update
        mytablex.MoveNext
    Loop
    mytablex.Close

End Sub

Sub xborrar()

    On Error GoTo cmd112_err

    Data2.Database.Execute "DELETE FROM _b" & gusuario
    Exit Sub
cmd112_err:
    Exit Sub

End Sub

Sub sumar_recibos()

    Dim xcargos As Double

    Dim xabonos As Double

    Dim xcargod As Double

    Dim xabonod As Double

    xcargos = 0
    xabonos = 0
    xcargod = 0
    xabonod = 0

    Data2.refresh
    Do

        If Data2.Recordset.EOF Then Exit Do
        If "" & Data2.Recordset.Fields("moneda") = "S" Then
            xcargos = xcargos + Val("" & Data2.Recordset.Fields("cargo"))
            xabonos = xabonos + Val("" & Data2.Recordset.Fields("abono"))

        End If

        If "" & Data2.Recordset.Fields("moneda") = "D" Then
            xcargod = xcargod + Val("" & Data2.Recordset.Fields("cargo"))
            xabonod = xabonod + Val("" & Data2.Recordset.Fields("abono"))

        End If

        Data2.Recordset.MoveNext
    Loop
    cargos = Format(xcargos, "0.00")
    abonos = Format(xabonos, "0.00")
    saldos = Format(xcargos - xabonos, "0.00")
    cargod = Format(xcargod, "0.00")
    abonod = Format(xabonod, "0.00")
    saldod = Format(xcargod - xabonod, "0.00")

End Sub

Function busca_tipo(buf As String)

    Dim mytablex As Table

    Set mytablex = mydbxglo.OpenTable("tipo")
    mytablex.Index = "tipo"
    mytablex.Seek "=", buf

    If Not mytablex.NoMatch Then
        busca_tipo = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close

End Function

Sub cabecera_documento1()

    Dim buf   As String

    Dim I     As Integer

    Dim found As Integer

    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = "Movimiento de Caja  "
    found = formateaa(buf, 90, 2, 0)
    
    found = formateaa("Lo", 3, 0, 0)
    found = formateaa("Tp", 3, 0, 0)
    found = formateaa("Srie", 5, 0, 0)
    found = formateaa("Numero", 12, 0, 0)
    found = formateaa("Fecha", 11, 0, 0)
    found = formateaa("Nombre", 31, 0, 0)
    
    found = formateaa("Cargo ", 11, 0, 1)
    found = formateaa("Abono ", 11, 0, 1)
    found = formateaa("Saldo", 11, 2, 1)
        
    buf = String(150, "-")
    found = formateaa(buf, 150, 2, 0)

End Sub

Sub cuerpo_programa_documento1()

    Dim buf    As String

    Dim found  As Integer

    Dim xcargo As Double

    Dim xabono As Double

    On Error GoTo cmd78812_err

    xcargo = 0
    xabono = 0
    Data2.refresh
    Do

        If Data2.Recordset.EOF Then Exit Do
        buf = "" & Data2.Recordset.Fields("LOCAL")
        found = formateaa(buf, 2, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & Data2.Recordset.Fields("tipo")
        found = formateaa(buf, 2, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & Data2.Recordset.Fields("serie")
        found = formateaa(buf, 4, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & Data2.Recordset.Fields("numero")
        found = formateaa(buf, 11, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & Data2.Recordset.Fields("fecha")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & Data2.Recordset.Fields("nombre")
        found = formateaa(buf, 30, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = Format(Val("" & Data2.Recordset.Fields("cargo")), "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(Val("" & Data2.Recordset.Fields("abono")), "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 2, 0)
      
        nlineas
      
        xcargo = xcargo + Val("" & Data2.Recordset.Fields("cargo"))
        xabono = xabono + Val("" & Data2.Recordset.Fields("abono"))
      
        Data2.Recordset.MoveNext
    Loop

    found = formateaa("", 65, 0, 0)
    buf = Format(xcargo, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(xabono, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(xcargo - xabono, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)
      
    Exit Sub
cmd78812_err:
    Exit Sub

End Sub

Sub nlineas()
    contlin = contlin + 1

    If contlin > 45 Then
        cabecera_documento1

    End If

End Sub

Function escajachica(buf As String)

    Dim mytablex As Table

    Set mytablex = mydbxglo.OpenTable("tipo")
    mytablex.Index = "tipo"
    mytablex.Seek "=", buf

    If Not mytablex.NoMatch Then
        If "" & mytablex.Fields("cajachica") = "C" Then
            escajachica = 1

        End If

    End If

    mytablex.Close

End Function

Function busca_observa(mytabley As Table) As String

    Dim mytablex As Table

    Set mytablex = mydbxglo.OpenTable("recibo")
    mytablex.Index = "recibo"
    mytablex.Seek "=", "" & mytabley.Fields("local"), "" & mytabley.Fields("tipo"), "" & mytabley.Fields("serie"), "" & mytabley.Fields("numero")

    If Not mytablex.NoMatch Then
        busca_observa = "" & mytablex.Fields("observa")

    End If

    mytablex.Close

End Function
