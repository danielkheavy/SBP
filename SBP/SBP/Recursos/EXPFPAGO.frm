VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form expfpago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientode Formas de Pago"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   14760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   14760
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      Caption         =   "Consulta"
      Height          =   5055
      Left            =   4440
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   7335
      Begin VB.TextBox local1 
         Height          =   495
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   13
         Text            =   "%"
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox nombre 
         Height          =   495
         Left            =   1920
         MaxLength       =   11
         TabIndex        =   11
         Text            =   "%"
         Top             =   1920
         Width           =   3135
      End
      Begin VB.TextBox codigo 
         Height          =   495
         Left            =   1920
         MaxLength       =   13
         TabIndex        =   9
         Text            =   "%"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CommandButton cmdCancelar 
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
         Left            =   5760
         MaskColor       =   &H00FFFFFF&
         Picture         =   "EXPFPAGO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton cmdGrabar 
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
         Left            =   5760
         MaskColor       =   &H00FFFFFF&
         Picture         =   "EXPFPAGO.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   1695
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "EXPFPAGO.frx":0F5C
      Height          =   5175
      Left            =   120
      OleObjectBlob   =   "EXPFPAGO.frx":0F70
      TabIndex        =   3
      Top             =   1560
      Width           =   14535
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00E0E0E0&
      Height          =   1275
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   14700
      TabIndex        =   0
      Top             =   0
      Width           =   14760
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
         Left            =   10200
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox fpago 
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
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   30
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
         Left            =   13320
         MaskColor       =   &H00FFFFFF&
         Picture         =   "EXPFPAGO.frx":4267
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   0
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
         Left            =   10200
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox fechai 
         Height          =   495
         Left            =   4440
         MaxLength       =   10
         TabIndex        =   24
         Top             =   120
         Width           =   1815
      End
      Begin VB.TextBox fechaf 
         Height          =   495
         Left            =   4440
         MaxLength       =   10
         TabIndex        =   23
         Top             =   600
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
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   19
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
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   18
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
         Left            =   10200
         Style           =   2  'Dropdown List
         TabIndex        =   17
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
         Picture         =   "EXPFPAGO.frx":4A15
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
         Picture         =   "EXPFPAGO.frx":5C27
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprimir"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ordenado Por"
         Height          =   375
         Left            =   9120
         TabIndex        =   33
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FormaPago"
         Height          =   375
         Left            =   6240
         TabIndex        =   31
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoDocto"
         Height          =   375
         Left            =   9120
         TabIndex        =   28
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         Height          =   495
         Left            =   3120
         TabIndex        =   26
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         Height          =   495
         Left            =   3120
         TabIndex        =   25
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Turno"
         Height          =   375
         Left            =   9120
         TabIndex        =   22
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Caja"
         Height          =   375
         Left            =   6240
         TabIndex        =   21
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cajero"
         Height          =   375
         Left            =   6240
         TabIndex        =   20
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Otros"
      Height          =   375
      Left            =   2520
      TabIndex        =   41
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label qotros 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3480
      TabIndex        =   40
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Credito"
      Height          =   375
      Left            =   2520
      TabIndex        =   39
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label qcredito 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3480
      TabIndex        =   38
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label qdolares 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1080
      TabIndex        =   37
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dolares"
      Height          =   375
      Left            =   120
      TabIndex        =   36
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label qcontado 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1080
      TabIndex        =   35
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Efectivo"
      Height          =   375
      Left            =   120
      TabIndex        =   34
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label dolares 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   10080
      TabIndex        =   16
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Label soles 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11760
      TabIndex        =   15
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label afecta 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   12480
      TabIndex        =   8
      Top             =   7440
      Width           =   735
   End
   Begin VB.Label acu 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   12120
      TabIndex        =   7
      Top             =   7560
      Width           =   255
   End
   Begin VB.Menu dki9923 
      Caption         =   "&Consulta"
   End
   Begin VB.Menu lfo3434 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "expfpago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancelar_Click()
    lfo3434_Click

End Sub

Private Sub cmdDelete_Click()
    dbo912_Click

End Sub

Private Sub cmdGrabar_Click()
    sql_recibos
    lfo3434_Click

End Sub

Private Sub cmdSort_Click()

End Sub

Private Sub Command1_Click()
    sql_recibos

End Sub

Private Sub dbo912_Click()

End Sub

Private Sub dki9923_Click()
    Frame2.Visible = True
    fechai.SetFocus

End Sub

Private Sub dnu823_Click()

End Sub

Private Sub Form_Activate()
    fechai = Format(Now, "dd/mm/yyyy") '"01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    fechaf = Format(Now, "dd/mm/yyyy")
    carga_inicial
    sql_recibos

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

    fpago.Clear
    fpago.AddItem "%"
    Set mytablex = mydbxglo.OpenTable("fpago")
    Do

        If mytablex.EOF Then Exit Do
        fpago.AddItem "" & mytablex.Fields("fpago") & "|" & mytablex.Fields("DESCRIPCIO")
        mytablex.MoveNext
    Loop
    mytablex.Close
    fpago.ListIndex = 0

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

Private Sub lfo3434_Click()

    If Frame2.Visible = True Then
        Frame2.Visible = False
        Exit Sub

    End If

    expfpago.Hide
    Unload expfpago

End Sub

Sub sql_recibos()

    On Error GoTo cmd37_err

    Dim buf As String

    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub
    buf = "select * from fpagov where "
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

    If fpago <> "%" Then
        buf = buf & " and fpago='" & extra_loquesea(fpago) & "'"

    End If

    'extra_loquesea(bodega)
    'buf = buf & " and acu='" & acu & "'"
    'buf = buf & " and afecta='" & afecta & "'"
    If codigo <> "%" Then
        buf = buf & " and codigo like '" & codigo & "'"

    End If

    If nombre <> "%" Then
        buf = buf & " and nombre like '" & nombre & "'"

    End If

    buf = buf & " order by " & ordenado
    'MsgBox buf
    Data2.Connect = "foxpro 2.5;"
    Data2.DatabaseName = globaldir
    Data2.RecordSource = buf
    Data2.refresh
    sumar_recibos
    DBGrid2.SetFocus
    Exit Sub
cmd37_err:
    MsgBox "Error en select " & error$, 48, "Aviso"
    Exit Sub

End Sub

Sub sumar_recibos()

    Dim xsoles   As Double

    Dim xdolares As Double

    Dim contado  As Double

    Dim credito  As Double

    Dim vdolares As Double

    Dim otros    As Double

    contado = 0
    credito = 0
    otros = 0
    vdolares = 0

    xsoles = 0
    xdolares = 0
    soles = "0.00"
    dolares = "0.00"
    qcontado = ""
    qdolares = ""
    qcredito = ""
    qotros = ""

    Do

        If Data2.Recordset.EOF Then Exit Do
        If "" & Data2.Recordset.Fields("estado") = "2" Then

            If "" & Data2.Recordset.Fields("acu") = "A" Or "" & Data2.Recordset.Fields("acu") = "B" Or "" & Data2.Recordset.Fields("acu") = "C" Or "" & Data2.Recordset.Fields("acu") = "D" Or "" & Data2.Recordset.Fields("acu") = "G" Or "" & Data2.Recordset.Fields("acu") = "I" Or "" & Data2.Recordset.Fields("acu") = "W" Then
                If "" & Data2.Recordset.Fields("acufp") = "A" Then
                    contado = contado + Val("" & Data2.Recordset.Fields("total"))
                    GoTo popo12

                End If

                If "" & Data2.Recordset.Fields("acufp") = "B" Then
                    vdolares = vdolares + Val("" & Data2.Recordset.Fields("total"))
                    GoTo popo12

                End If

                If "" & Data2.Recordset.Fields("acufp") = "C" Then
                    credito = credito + Val("" & Data2.Recordset.Fields("total"))
                    GoTo popo12

                End If

                otros = otros + Val("" & Data2.Recordset.Fields("total"))
   
popo12:

                If "" & Data2.Recordset.Fields("moneda") = "S" Then
                    xsoles = xsoles + Val("" & Data2.Recordset.Fields("total"))

                End If

                If "" & Data2.Recordset.Fields("moneda") = "D" Then
                    xdolares = xdolares + Val("" & Data2.Recordset.Fields("total"))

                End If

            End If

            If "" & Data2.Recordset.Fields("acu") = "J" Or "" & Data2.Recordset.Fields("acu") = "K" Or "" & Data2.Recordset.Fields("acu") = "L" Or "" & Data2.Recordset.Fields("acu") = "M" Or "" & Data2.Recordset.Fields("acu") = "P" Or "" & Data2.Recordset.Fields("acu") = "V" Then
                If "" & Data2.Recordset.Fields("acufp") = "A" Then
                    contado = contado - Val("" & Data2.Recordset.Fields("total"))
                    GoTo popo13

                End If

                If "" & Data2.Recordset.Fields("acufp") = "B" Then
                    vdolares = vdolares - Val("" & Data2.Recordset.Fields("total"))
                    GoTo popo13

                End If

                If "" & Data2.Recordset.Fields("acufp") = "C" Then
                    credito = credito - Val("" & Data2.Recordset.Fields("total"))
                    GoTo popo13

                End If

                otros = otros - Val("" & Data2.Recordset.Fields("total"))
popo13:

                If "" & Data2.Recordset.Fields("moneda") = "S" Then
                    xsoles = xsoles - Val("" & Data2.Recordset.Fields("total"))

                End If

                If "" & Data2.Recordset.Fields("moneda") = "D" Then
                    xdolares = xdolares - Val("" & Data2.Recordset.Fields("total"))

                End If

            End If

        End If

        Data2.Recordset.MoveNext
    Loop
    qcontado = Format(contado, "0.00")
    qdolares = Format(vdolares, "0.00")
    qcredito = Format(credito, "0.00")
    qotros = Format(otros, "0.00")

    soles = Format(xsoles, "0.00")
    dolares = Format(xdolares, "0.00")

    soles = Format(xsoles, "0.00")
    dolares = Format(xdolares, "0.00")

End Sub
