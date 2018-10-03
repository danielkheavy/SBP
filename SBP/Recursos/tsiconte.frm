VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tsiconte 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Regularizaciones"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   14280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   14280
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Clave de actualizacion"
      Height          =   2415
      Left            =   2880
      TabIndex        =   12
      Top             =   2400
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton Command2 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   5040
         TabIndex        =   16
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   5040
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox clave 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "PROCESANDO...ESPERE...!!!!!!!!!!!!!!!!!!"
         Height          =   855
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clave"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00808080&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   14220
      TabIndex        =   0
      Top             =   0
      Width           =   14280
      Begin VB.ComboBox local1 
         Height          =   315
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Consul&Tar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9600
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tsiconte.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.ComboBox bodega 
         Height          =   315
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   2175
      End
      Begin VB.ComboBox vendedor 
         Height          =   315
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   120
         Width           =   2175
      End
      Begin VB.TextBox fechaf 
         Height          =   375
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   4
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox fechai 
         Height          =   375
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   2
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "local"
         Height          =   375
         Left            =   5880
         TabIndex        =   9
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Almacen"
         Height          =   375
         Left            =   2880
         TabIndex        =   6
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Usuario"
         Height          =   375
         Left            =   2880
         TabIndex        =   5
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Final"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin VB.Label fechaiw 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid dbGrid1 
      Height          =   7575
      Left            =   0
      TabIndex        =   18
      Top             =   960
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   13361
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   22
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "Numero"
         Caption         =   "Numero"
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
         DataField       =   "Fecha"
         Caption         =   "Fecha"
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
      BeginProperty Column02 
         DataField       =   "Local"
         Caption         =   "Local"
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
      BeginProperty Column03 
         DataField       =   "Bodega"
         Caption         =   "Almacen"
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
      BeginProperty Column04 
         DataField       =   "Vendedor"
         Caption         =   "Responsable"
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
      BeginProperty Column05 
         DataField       =   "Observa"
         Caption         =   "Observa"
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
      BeginProperty Column06 
         DataField       =   "Estado"
         Caption         =   "E"
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
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2610.142
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   4440.189
         EndProperty
         BeginProperty Column06 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "IMPORTANTE:LAS ACTUALIZACIONES EN EL STOCK ES CON LA FECHA DEL DOCUMENTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   8520
      Width           =   10455
   End
   Begin VB.Label yausado 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   8520
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Menu dkuwew 
      Caption         =   "&Add"
   End
   Begin VB.Menu mid8s 
      Caption         =   "&Modifica"
   End
   Begin VB.Menu dk8823 
      Caption         =   "&Imprime"
      Begin VB.Menu dk223 
         Caption         =   "&1.Normal"
      End
      Begin VB.Menu xclowew 
         Caption         =   "&2.Excell"
      End
   End
   Begin VB.Menu Kver612 
      Caption         =   "&Zoom"
   End
   Begin VB.Menu dkl8923 
      Caption         =   "Actua&Lizar"
   End
   Begin VB.Menu dlo2323 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tsiconte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbtconteo As New ADODB.Recordset

Private Sub clave_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If Len(clave) = 0 Then
        clave.SetFocus
        Exit Sub

    End If

    found = busca_clave()

    If found = 0 Then
        MsgBox "NO existe clave", 48, "Aviso"
        clave = ""
        clave.SetFocus
        Exit Sub

    End If

    grabar_conteo
    dlo2323_Click

End Sub

Private Sub cmdSave_Click()

End Sub

Private Sub cmdAddEntry_Click()

End Sub

Private Sub cmdExit_Click()

End Sub

Private Sub Command1_Click()
    dlo2323_Click

End Sub

Private Sub Command2_Click()
    Label6.Visible = True
    clave_KeyPress 13
    Label6.Visible = False

End Sub

Private Sub Command5_Click()
    sql_cabeza

End Sub

Private Sub dbgrid1_BeforeColEdit(ByVal ColIndex As Integer, _
                                  ByVal KeyAscii As Integer, _
                                  Cancel As Integer)

    If ColIndex <> 5 Then
        Cancel = True
        Exit Sub

    End If

End Sub

Private Sub dk223_Click()

    Dim sdx As String

    On Error GoTo cmd8_err

    sdx = "" & dbtconteo.Fields("numero")
    impresion1
    Exit Sub
cmd8_err:
    MsgBox "Elegir un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub dkl8923_Click()

    On Error GoTo cmd5612_err

    'grabar_conteoT
    'Exit Sub
    If "" & dbtconteo.Fields("estado") = "1" Then
        MsgBox "Documento ya actualizado ", 48, "Aviso"
        Exit Sub

    End If

    Frame1.Visible = True
    clave = ""
    clave.SetFocus
    'flag_clave1 = 0
    'tconcla.X = "C"
    'tconcla.Show 1
    'If flag_clave1 <> 1 Then  'si es descongela
    '   Exit Sub
    'End If
    Exit Sub
cmd5612_err:
    MsgBox "Seleccione un dato", 48, "Aviso"
    Exit Sub

End Sub

Private Sub dkuwew_Click()

    If local1 = "%" Then
        MsgBox "Seleccione un local ", 48, "Aviso"
        Exit Sub

    End If

    tconteoo.local1 = extra_loquesea(local1)
    tconteoo.modelo = "ADICIONA"
    tconteoo.Show 1

End Sub

Private Sub dlo2323_Click()

    If Frame1.Visible = True Then
        Frame1.Visible = False
        Exit Sub

    End If

    tsiconte.Hide
    Unload tsiconte

End Sub

Private Sub Form_Activate()

    'local1 = glocal
    If yausado = "" Then
        cargas_iniciales
        yausado = "1"

    End If

    sql_cabeza

End Sub

Private Sub Form_Load()
    fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    fechaf = Format(Now, "dd/mm/yyyy")

End Sub

Sub cargas_iniciales()

    Dim mytablex As New ADODB.Recordset

    local1.Clear
    local1.AddItem "%"
    mytablex.Open "select * from tlocal", cn, adOpenDynamic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        local1.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    local1.ListIndex = 0

    If local1.ListCount = 2 Then
        local1.ListIndex = 1

    End If

    mytablex.Close

    bodega.Clear
    bodega.AddItem "%"
    mytablex.Open "select * from bodega", cn, adOpenDynamic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        bodega.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    bodega.ListIndex = 0
    mytablex.Close
    vendedor.Clear
    vendedor.AddItem "%"
    mytablex.Open "select * from vendedor", cn, adOpenDynamic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        vendedor.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    vendedor.ListIndex = 0
    mytablex.Close

End Sub

Sub sql_cabeza()

    On Error GoTo cmd37_err

    Dim buf As String

    If Len(fechai) <> 10 Then Exit Sub
    If Len(fechaf) <> 10 Then Exit Sub
    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub
    'MsgBox cgusuario
    buf = "select * from cconteof where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    'buf = buf & " and tipoclie='" & tipoclie & "'"
    If vendedor <> "%" Then
        buf = buf & " and vendedor like '" & extra_loquesea(vendedor) & "'"

    End If

    If local1 <> "%" Then
        buf = buf & " and local like '" & extra_loquesea(local1) & "'"

    End If

    If bodega <> "%" Then
        buf = buf & " and bodega like '" & extra_loquesea(bodega) & "'"

    End If

    buf = buf & " order by fecha,numero"

    'MsgBox buf
    If dbtconteo.State = 1 Then dbtconteo.Close
    dbtconteo.Open buf, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = dbtconteo
    ir_ultimo

    If dbtconteo.EOF = True And dbtconteo.BOF = True Then
        Exit Sub

    End If

    dbGrid1.SetFocus
    Exit Sub
cmd37_err:
    MsgBox "Error en select " & error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub Kver612_Click()

    On Error GoTo cmd123_err

    tconteoo.Numero = "" & dbtconteo.Fields("numero")
    'tconteoo.observa = "" & Data1.Recordset.Fields("observa")
    tconteoo.vendedor.AddItem "" & dbtconteo.Fields("vendedor") & "|" & busca_xvendedor("" & dbtconteo.Fields("vendedor"))
    tconteoo.vendedor.ListIndex = 0
    tconteoo.local1 = "" & dbtconteo.Fields("local")
    'tconteoo.local1.ListIndex = 0
    tconteoo.bodega.AddItem "" & dbtconteo.Fields("bodega") & "|" & busca_xbodega("" & dbtconteo.Fields("bodega"))
    tconteoo.bodega.ListIndex = 0
    tconteoo.fecha = "" & dbtconteo.Fields("fecha")
    tconteoo.modelo = "SOLO VER"
    tconteoo.Show 1
    Exit Sub
cmd123_err:
    MsgBox "Seleccione un Dato", 48, "Aviso"
    Exit Sub

End Sub

Private Sub mid8s_Click()

    On Error GoTo cmd1_err

    If "" & dbtconteo.Fields("estado") = "1" Then
        If MsgBox("Ya fue actuaalizado,Desea Continuar", 1, "Modifica") <> 1 Then Exit Sub
        Exit Sub

    End If

    tconteoo.Numero = "" & dbtconteo.Fields("numero")
    'tconteoo.observa = "" & Data1.Recordset.Fields("observa")
    tconteoo.vendedor.AddItem "" & dbtconteo.Fields("vendedor") & "|" & busca_xvendedor("" & dbtconteo.Fields("vendedor"))
    tconteoo.vendedor.ListIndex = 0
    tconteoo.local1 = "" & dbtconteo.Fields("local")
    'tconteoo.local1.ListIndex = 0
    tconteoo.bodega.AddItem "" & dbtconteo.Fields("bodega") & "|" & busca_xbodega("" & dbtconteo.Fields("bodega"))
    tconteoo.bodega.ListIndex = 0
    tconteoo.fecha = "" & dbtconteo.Fields("fecha")
    tconteoo.modelo = "MODIFICA"
    tconteoo.Show 1
    Exit Sub
cmd1_err:
    MsgBox "Seleccione un Dato", 48, "Aviso"
    Exit Sub

End Sub

Private Sub impresion1()

    Dim found As Integer

    Dim buf   As String

    If MsgBox("Desea Imprimir", 1, "Aviso") <> 1 Then Exit Sub
    contpag = 0
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    'found = ir_primero1()
    
    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_documento
    cuerpo_programa_documento
    '------------------------------------
    Close #1
    cerrar_archivo
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1

End Sub

Sub cabecera_documento()

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
    buf = "Reporte de Conteos Fisicos  "
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Numero  :", 10, 0, 0)
    found = formateaa("" & dbtconteo.Fields("numero"), 10, 0, 0)
    found = formateaa("", 1, 2, 0)
    
    found = formateaa("Fecha  :", 10, 0, 0)
    found = formateaa("" & dbtconteo.Fields("fecha"), 10, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Local  :", 10, 0, 0)
    found = formateaa("" & dbtconteo.Fields("local"), 10, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Bodega :", 10, 0, 0)
    found = formateaa("" & dbtconteo.Fields("bodega"), 10, 0, 0)
    found = formateaa("", 1, 2, 0)
        
    buf = String(150, "-")
    found = formateaa(buf, 150, 2, 0)
    '------aqui van los registros----------------------
        
    found = formateaa("Producto", 10, 0, 0)
    found = formateaa("Descripcio", 40, 0, 0)
    found = formateaa("Stock ", 10, 0, 1)
    found = formateaa("Conteo ", 10, 0, 1)
    found = formateaa("Costo ", 10, 0, 1)
    found = formateaa("CantSobra ", 10, 0, 1)
    found = formateaa("ValoSobra ", 10, 0, 1)
    found = formateaa("CantFalta ", 10, 0, 1)
    found = formateaa("ValoFalta ", 10, 2, 1)
    '--------------------------------------------------
    buf = String(150, "-")
    found = formateaa(buf, 150, 2, 0)

End Sub

Sub cuerpo_programa_documento()

    Dim buf   As String

    Dim found As Integer

    Dim sdx   As Double

    Dim sdx1  As Double

    Dim costo As Double

    Dim sobrante, faltante As Double

    Dim saldoant As Double

    Dim saldoini As Double

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd788_err

    sdx = 0
    sdx1 = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from conteofi where numero='" & dbtconteo("numero") & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    Do

        'MsgBox "" & mytablex.Fields("producto")
        If mytablex.EOF Then Exit Do
        '-----------------------------------------
        buf = "" & mytablex.Fields("producto")
        found = formateaa(buf, 9, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = Mid$("" & mytablex.Fields("descripcio"), 1, 34) & " " & Mid$("" & mytablex.Fields("unidad"), 1, 6) & "x" & Mid$("" & mytablex.Fields("factor"), 1, 4)
        found = formateaa(buf, 39, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("cantidad")
        found = formateaa(buf, 9, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("saldoant")
        found = formateaa(buf, 9, 0, 1)
        found = formateaa("", 1, 0, 0)
        costo = busca_producto("" & mytablex.Fields("producto"))
        buf = Format(costo, "0.00")
        found = formateaa(buf, 9, 0, 1)
        found = formateaa("", 1, 0, 0)
        saldoini = "" & mytablex.Fields("cantidad")
        saldoant = "" & mytablex.Fields("saldoant")
        sobrante = 0
        faltante = 0

        If saldoini = saldoant Then  'igual

        End If

        If saldoini < saldoant Then  'sobrante
            sobrante = Abs(saldoini - saldoant)

        End If

        If saldoini > saldoant Then  'faltante
            faltante = Abs(saldoini - saldoant)

        End If

        buf = "" & sobrante
        found = formateaa(buf, 9, 0, 1)
        found = formateaa("", 1, 0, 0)
        suma1 = suma1 + sobrante

        sdx = costo * sobrante
        buf = "" & sdx
        found = formateaa(buf, 9, 0, 1)
        found = formateaa("", 1, 0, 0)
        suma2 = suma2 + sdx

        buf = "" & faltante
        found = formateaa(buf, 9, 0, 1)
        found = formateaa("", 1, 0, 0)
        suma3 = suma3 + faltante

        sdx = costo * faltante
        buf = "" & sdx
        found = formateaa(buf, 9, 0, 1)
        found = formateaa("", 1, 2, 0)
        suma4 = suma4 + sdx
        nlineas
        mytablex.MoveNext
    Loop

    found = formateaa("", 80, 0, 0)
    buf = suma1
    found = formateaa(buf, 9, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = suma2
    found = formateaa(buf, 9, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = suma3
    found = formateaa(buf, 9, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = suma4
    found = formateaa(buf, 9, 0, 1)
    found = formateaa("", 1, 2, 0)

    mytablex.Close

    Exit Sub
cmd788_err:
    Exit Sub

End Sub

Sub nlineas()
    contlin = contlin + 1

    If contlin > 45 Then
        cabecera_documento

    End If

End Sub

Function busca_producto(buf As String) As Double

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from producto where producto='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_producto = Val(Format(Val("" & mytablex.Fields("costou")), "0.00"))

    End If

    mytablex.Close

End Function

Function busca_xvendedor(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from vendedor where codigo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_xvendedor = "" & mytablex.Fields("nombre")

    End If

    mytablex.Close

End Function

Function busca_xbodega(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from bodega where codigo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_xbodega = "" & mytablex.Fields("nombre")

    End If

    mytablex.Close

End Function

Sub ir_ultimo()

    On Error GoTo cmd6711_err

    'Data2.Recordset.MoveLast
    Exit Sub
cmd6711_err:
    Exit Sub

End Sub

Function grabar_conteo()

    Dim mytablex As Table

    Dim found    As Integer

    Dim saldoa   As Double

    Dim vr

    On Error GoTo cmd3243_err

    If MsgBox("Desea actualizar ", 1, "Aviso") <> 1 Then Exit Function
    If "" & dbtconteo.Fields("estado") = "1" Then
        MsgBox "Documento ya actualizado ", 48, "Aviso"
        Exit Function

    End If

    '--------------- SE ANULO
    'Set mytablex = mydbxglo.OpenTable("conteofi")
    'mytablex.Index = "conteofi"
    'mytablex.Seek "=", "" & dbtconteo.fields("numero")
    'If Not mytablex.NoMatch Then
    'Do
    'If mytablex.EOF Then Exit Do
    'If "" & mytablex.Fields("numero") = "" & dbtconteo.fields("numero") Then
    '   saldoa = recalculo_saldos1(mytablex)
    '   found = grabarx(mytablex, saldoa)
    '   Else
    '   GoTo xx
    'End If
    'mytablex.MoveNext
    'Loop
    'endif
    'mytablex.Close
    'Data2.Recordset.Edit
    dbtconteo.Fields("estado") = "1"
    dbtconteo.Update
    MsgBox "Presione enter para continuar..", 48, "Aviso"
    Exit Function
cmd3243_err:
    MsgBox "Seleccione un documento " + error$, 48, "Aviso"
    Exit Function

End Function

Function busca_clave()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from vendedor where clave='" & clave & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_clave = 1

    End If

    mytablex.Close

End Function

Sub conteo_excell()

    Dim mytablex As New ADODB.Recordset

    Dim v, h As Integer

    Dim found       As Integer

    Dim I           As Integer

    Dim sdx         As Double
 
    Dim vprecios(7) As String

    Dim Heading(8)  As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    On Error GoTo cmd5612_err

    'Data1.Refresh
   
    mytablex.Open "select * from conteofi where numero='" & dbtconteo.Fields("numero") & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If
   
    Heading(1) = "Producto"
    Heading(2) = "Descripcio"
    Heading(3) = "Saldoa"
    Heading(4) = "Conteo"
   
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_Excel(4, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
    v = 5
    h = 1
    sdx = 0

    Do

        If mytablex.EOF Then Exit Do
     
        sdx = sdx + Val("" & mytablex.Fields("saldoant"))
        objExcel.ActiveSheet.Cells(v, h) = "" & mytablex.Fields("producto")
        objExcel.ActiveSheet.Cells(v, h + 1) = "" & mytablex.Fields("descripcio")
        objExcel.ActiveSheet.Cells(v, h + 2) = "" & mytablex.Fields("cantidad")
        objExcel.ActiveSheet.Cells(v, h + 3) = "" & mytablex.Fields("saldoant")
        v = v + 1
        mytablex.MoveNext
    Loop
 
    objExcel.ActiveSheet.Cells(v, h) = ""
    objExcel.ActiveSheet.Cells(v, h + 1) = ""
    objExcel.ActiveSheet.Cells(v, h + 2) = ""
    objExcel.ActiveSheet.Cells(v, h + 3) = "" & sdx
    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    mytablex.Close
    Exit Sub
cmd5612_err:
    MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub xclowew_Click()

    Dim sdx As String

    On Error GoTo cmd81_err

    sdx = "" & dbtconteo.Fields("numero")
    conteo_excell
    Exit Sub
cmd81_err:
    MsgBox "Elegir un dato ", 48, "Aviso"
    Exit Sub

End Sub
