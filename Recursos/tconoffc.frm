VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tconoffc 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Toma Inventario Pdt- No en Linea"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   14550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   14550
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Height          =   3975
      Left            =   2280
      TabIndex        =   26
      Top             =   2280
      Visible         =   0   'False
      Width           =   9015
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Clave de actualizacion"
      Height          =   2415
      Left            =   4200
      TabIndex        =   11
      Top             =   2520
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton Command2 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   5040
         TabIndex        =   15
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   5040
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox clave 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "PROCESANDO...ESPERE...!!!!!!!!!!!!!!!!!!"
         Height          =   855
         Left            =   120
         TabIndex        =   16
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
         TabIndex        =   12
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   14490
      TabIndex        =   0
      Top             =   0
      Width           =   14550
      Begin VB.ComboBox ordenado 
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
         Left            =   6720
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   480
         Width           =   2175
      End
      Begin VB.ComboBox local1 
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
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   120
         Width           =   2175
      End
      Begin VB.ComboBox conteo 
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
         Left            =   6720
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   120
         Width           =   735
      End
      Begin VB.ComboBox ubicacion 
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
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   840
         Visible         =   0   'False
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
         Picture         =   "tconoffc.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.ComboBox bodega 
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
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   2175
      End
      Begin VB.ComboBox vendedor 
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
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   840
         Width           =   1815
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
      Begin VB.Label Periodo 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   12120
         TabIndex        =   27
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ordenado"
         Height          =   375
         Left            =   5880
         TabIndex        =   25
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local1"
         Height          =   375
         Left            =   2880
         TabIndex        =   23
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Conteo"
         Height          =   375
         Left            =   5880
         TabIndex        =   21
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ubicacion"
         Height          =   375
         Left            =   2880
         TabIndex        =   19
         Top             =   840
         Visible         =   0   'False
         Width           =   855
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
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   975
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
      Height          =   7095
      Left            =   0
      TabIndex        =   17
      Top             =   1320
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   12515
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
      ColumnCount     =   9
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
         DataField       =   "ubicacion"
         Caption         =   "Ubicacion"
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
         DataField       =   "Conteo"
         Caption         =   "Conteo"
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
      BeginProperty Column07 
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
      BeginProperty Column08 
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
            ColumnWidth     =   629.858
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   824.882
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   675.213
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2789.858
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   4440.189
         EndProperty
         BeginProperty Column08 
         EndProperty
      EndProperty
   End
   Begin VB.Label yausado 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   10
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
   Begin VB.Menu Borrss 
      Caption         =   "&Borra"
   End
   Begin VB.Menu dk8823 
      Caption         =   "&Imprime"
      Begin VB.Menu dk223 
         Caption         =   "&1.Normal"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu xclowew 
         Caption         =   "&2.Excell"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu to885 
         Caption         =   "&3.Reporte Consolidado de Productos"
      End
      Begin VB.Menu fk9944 
         Caption         =   "&4.Generador"
      End
   End
   Begin VB.Menu Kver612 
      Caption         =   "&Zoom"
   End
   Begin VB.Menu dkl8923 
      Caption         =   "Actua&Lizar"
      Visible         =   0   'False
   End
   Begin VB.Menu dlo2323 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tconoffc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbtconteo As New ADODB.Recordset

Private Sub bodega_Click()

    If bodega <> "%" Then
        sql_cabeza

    End If

End Sub

Private Sub Borrss_Click()

    Dim buf As String

    On Error GoTo cmd4590_err

    buf = "" & dbtconteo.Fields("numero")

    If MsgBox("Desea Borrar " + buf, 1, "Aviso") <> 1 Then Exit Sub
    cn.Execute ("delete from pdtde where numero='" & buf & "'")
    cn.Execute ("delete from pdtca where numero='" & buf & "'")
    MsgBox "Archivo Borrado ", 48, "Aviso"
    sql_cabeza
    Exit Sub
cmd4590_err:
    MsgBox "Ejegir un Dato ", 48, "Aviso"
    Exit Sub

End Sub

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

Private Sub Command3_Click()

    If Command3.Visible = True Then
        Command3.Visible = False

    End If

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

Sub generar_saldoinicial()

    Dim sdx      As Double

    Dim mytabley As New ADODB.Recordset

    Dim vr

    Dim mytablex As New ADODB.Recordset

    cn.Execute ("delete from pdtdetmp where local='" & extra_loquesea(local1) & "'")
    dbtconteo.MoveFirst
    Command3.Visible = True
    sdx = 0
    Do

        If dbtconteo.EOF Then Exit Do
        Command3.Caption = "" & dbtconteo("numero")
        vr = DoEvents()

        If Command3.Visible = False Then
            MsgBox "Proceso Interrumpido ", 48, "Aviso"
            Exit Sub

        End If

        mytablex.Open "select * from pdtde where  numero='" & dbtconteo.Fields("numero") & "'", cn, adOpenDynamic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            Do

                If mytablex.EOF Or mytablex.BOF Then Exit Do
                mytabley.Open "select * from pdtdetmp where producto='" & mytablex.Fields("producto") & "' and local='" & extra_loquesea(local1) & "'", cn, adOpenDynamic, adLockOptimistic

                If mytabley.RecordCount = 0 Then
                    sdx = sdx + 1
                    mytabley.AddNew
                    mytabley.Fields("producto") = "" & mytablex.Fields("producto")
                    mytabley.Fields("descripcio") = "" & mytablex.Fields("descripcio")
                    mytabley.Fields("cantidad") = Val("" & mytabley.Fields("cantidad")) + Val("" & mytablex.Fields("saldoant"))
                    mytabley.Fields("local") = extra_loquesea(local1)
                    mytabley.Fields("saldo") = busca_saldoxx(extra_loquesea(local1), "" & mytablex.Fields("producto"), "" & mytablex.Fields("bodega"))
                    mytabley.Update
                Else
                    sdx = sdx + 1
                    mytabley.Fields("producto") = "" & mytablex.Fields("producto")
                    mytabley.Fields("cantidad") = Val("" & mytabley.Fields("cantidad")) + Val("" & mytablex.Fields("saldoant"))
                    mytabley.Update

                End If

                mytabley.Close
                Set mytabley = Nothing
                mytablex.MoveNext
            Loop

        End If

        mytablex.Close
        Set mytablex = Nothing
        dbtconteo.MoveNext
    Loop

End Sub

Private Sub dkuwew_Click()

    If local1 = "%" Then
        MsgBox "Seleccione un local ", 48, "Aviso"
        Exit Sub

    End If

    If bodega = "%" Then
        MsgBox "Seleccione una Bodega ", 48, "Aviso"
        Exit Sub

    End If

    tconoff.periodo = periodo
    tconoff.local1 = "" & extra_loquesea(local1) & ""
    tconoff.bodega = "" & extra_loquesea(bodega) & ""

    '' 03/07/2018 Conteo Fisico Sistema
    tconoff.ubicacion = "" & extra_loquesea(ubicacion) & ""
    '' 03/07/2018 Conteo Fisico Sistema

    tconoff.modelo = "ADICIONA"
    tconoff.Show 1

End Sub

Private Sub dlo2323_Click()

    If Frame1.Visible = True Then
        Frame1.Visible = False
        Exit Sub

    End If

    tconoffc.Hide
    Unload tconoffc

End Sub

Private Sub fk9944_Click()
    reporgen.NAMETABLA = "pdtde"
    reporgen.Show 1

End Sub

Private Sub Form_Activate()

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

    conteo.Clear
    conteo.AddItem "%"
    conteo.AddItem "1"
    conteo.AddItem "2"
    conteo.AddItem "3"
    conteo.ListIndex = 0

    ordenado.Clear
    ordenado.AddItem "%"
    ordenado.AddItem "Numero"
    ordenado.AddItem "Ubicacion"
    ordenado.AddItem "Local"
    ordenado.AddItem "Bodega"
    ordenado.AddItem "Conteo"
    ordenado.AddItem "Vendedor"
    ordenado.ListIndex = 0

    local1.Clear
    local1.AddItem "%"
    mytablex.Open "select * from tlocal", cn, adOpenDynamic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        local1.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    mytablex.Close
    local1.ListIndex = 0

    If local1.ListCount = 2 Then
        local1.ListIndex = 1

    End If

    ubicacion.Clear
    ubicacion.AddItem "%"
    ubicacion.ListIndex = 0

    bodega.Clear
    bodega.AddItem "%"
    mytablex.Open "select * from bodega", cn, adOpenDynamic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        bodega.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    bodega.ListIndex = 0

    '' 03/07/2018 Conteo Fisico Sistema
    If bodega.ListCount = 2 Then
        bodega.ListIndex = 1

    End If

    '' 03/07/2018 Conteo Fisico Sistema

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

    carga_ubicacion ""

End Sub

Sub sql_cabeza()

    On Error GoTo cmd37_err

    Dim buf As String

    If Len(fechai) <> 10 Then Exit Sub
    If Len(fechaf) <> 10 Then Exit Sub
    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub
    'MsgBox cgusuario
    buf = "select * from pdtca where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If vendedor <> "%" Then
        buf = buf & " and vendedor like '" & extra_loquesea(vendedor) & "'"

    End If

    If local1 <> "%" Then
        buf = buf & " and local like '" & extra_loquesea(local1) & "'"

    End If

    If ubicacion <> "%" Then
        buf = buf & " and ubicacion like '" & extra_loquesea(ubicacion) & "'"

    End If

    If conteo <> "%" Then
        buf = buf & " and conteo like '" & extra_loquesea(conteo) & "'"

    End If

    If bodega <> "%" Then
        buf = buf & " and bodega like '" & extra_loquesea(bodega) & "'"

    End If

    buf = buf & "and periodo='" & periodo & "' "

    If ordenado <> "%" Then
        buf = buf & " order by " & ordenado

    End If

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

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Kver612_Click()

    On Error GoTo cmd123_err

    tconoff.Numero = "" & dbtconteo.Fields("numero")
    'tconoff.observa = "" & Data1.Recordset.Fields("observa")
    tconoff.vendedor.AddItem "" & dbtconteo.Fields("vendedor") & "|" & busca_xvendedor("" & dbtconteo.Fields("vendedor"))
    tconoff.vendedor.ListIndex = 0
    tconoff.local1 = "" & dbtconteo.Fields("local")
    'tconoff.local1.ListIndex = 0
    tconoff.bodega = "" & dbtconteo.Fields("bodega")
    tconoff.fecha = "" & dbtconteo.Fields("fecha")
    tconoff.periodo = periodo
    tconoff.ubicacion = "" & dbtconteo.Fields("ubicacion")
    tconoff.conteo1.AddItem "" & dbtconteo.Fields("conteo")
    tconoff.conteo1.ListIndex = 0

    tconoff.modelo = "SOLO VER"
    tconoff.Show 1
    Exit Sub
cmd123_err:
    MsgBox "Seleccione un Dato", 48, "Aviso"
    Exit Sub

End Sub

Private Sub local1_Click()
    sql_cabeza

End Sub

Private Sub mid8s_Click()

    On Error GoTo cmd1_err

    If "" & dbtconteo.Fields("estado") = "1" Then
        If MsgBox("Ya fue actuaalizado,Desea Continuar", 1, "Modifica") <> 1 Then Exit Sub
        Exit Sub

    End If

    tconoff.Numero = "" & dbtconteo.Fields("numero")
    'tconoff.observa = "" & Data1.Recordset.Fields("observa")
    tconoff.ubicacion = "" & dbtconteo.Fields("ubicacion")
    tconoff.conteo1.AddItem "" & dbtconteo.Fields("conteo")
    tconoff.conteo1.ListIndex = 0
    tconoff.periodo = periodo
    tconoff.vendedor.AddItem "" & dbtconteo.Fields("vendedor") & "|" & busca_xvendedor("" & dbtconteo.Fields("vendedor"))
    tconoff.vendedor.ListIndex = 0
    tconoff.local1 = "" & dbtconteo.Fields("local")
    'tconoff.local1.ListIndex = 0
    tconoff.bodega = "" & dbtconteo.Fields("bodega")
    'tconoff.bodega.ListIndex = 0
    tconoff.fecha = "" & dbtconteo.Fields("fecha")
    tconoff.modelo = "MODIFICA"
    tconoff.Show 1
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
    mytablex.Open "select * from pdtde where numero='" & dbtconteo("numero") & "'", cn, adOpenDynamic, adLockOptimistic

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

    mytablex.Open "select * from vendedor where codigo='" & clave & "'", cn, adOpenDynamic, adLockOptimistic

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

    Dim buf         As String

    Dim Tmp         As String

    Dim sw          As Integer
 
    Dim vprecios(7) As String

    Dim Heading(8)  As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    On Error GoTo cmd5612_err

    'Data1.Refresh
    sw = 0
    Tmp = ""
   
    buf = "select Ubicacion,Producto,Conteo,sum(saldoant) AS Nro from pdtde where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If vendedor <> "%" Then
        buf = buf & " and vendedor like '" & extra_loquesea(vendedor) & "'"

    End If

    If local1 <> "%" Then
        buf = buf & " and local like '" & extra_loquesea(local1) & "'"

    End If

    If ubicacion <> "%" Then
        buf = buf & " and ubicacion like '" & extra_loquesea(ubicacion) & "'"

    End If

    If conteo <> "%" Then
        buf = buf & " and conteo like '" & extra_loquesea(conteo) & "'"

    End If

    If bodega <> "%" Then
        buf = buf & " and bodega like '" & extra_loquesea(bodega) & "'"

    End If

    buf = buf & " group by ubicacion,producto,conteo"
    MsgBox buf
   
    mytablex.Open buf, cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    Heading(1) = "Ubicacion"
    Heading(2) = "Producto"
    Heading(3) = "Conteo"
    Heading(4) = "Cantidad"

    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_Excel(4, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    v = 5
    h = 1
    sdx = 0
    Do

        If mytablex.EOF Then Exit Do
        objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("Ubicacion")
        objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("Producto")
        objExcel.ActiveSheet.Cells(v, h + 2) = "'" & mytablex.Fields("conteo")
        objExcel.ActiveSheet.Cells(v, h + 3) = "'" & mytablex.Fields("Nro")
        v = v + 1
        mytablex.MoveNext
    Loop
 
    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    mytablex.Close
    Exit Sub
cmd5612_err:
    MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub to885_Click()

    On Error GoTo cmd7866_err

    Dim vr

    Dim mytabley As New ADODB.Recordset

    Dim mytablex As New ADODB.Recordset

    If local1 = "%" Then
        MsgBox "Seleccione un local ", 48, "Aviso"
        Exit Sub

    End If

    If bodega = "%" Then
        MsgBox "Seleccione una Bodega ", 48, "Aviso"
        Exit Sub

    End If

    sql_cabeza

    If MsgBox("Desea Generar Archivo Excell de los conteos ", 1, "Aviso") <> 1 Then Exit Sub
    cn.Execute ("delete from pdtdetmp ")
    dbtconteo.MoveFirst
    Command3.Visible = True
    Do

        If dbtconteo.EOF Or dbtconteo.BOF Then Exit Do
        Command3.Caption = "" & dbtconteo("numero")
        vr = DoEvents()

        If Command3.Visible = False Then
            MsgBox "Proceso Interrumpido ", 48, "Aviso"
            Exit Sub

        End If

        mytablex.Open "select * from pdtde where  periodo='" & periodo & "' and numero='" & dbtconteo.Fields("numero") & "'", cn, adOpenDynamic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            Do

                If mytablex.EOF Or mytablex.BOF Then Exit Do
                mytabley.Open "select * from pdtdetmp where producto='" & mytablex.Fields("producto") & "'", cn, adOpenDynamic, adLockOptimistic

                If mytabley.RecordCount = 0 Then
                    mytabley.AddNew
                    mytabley.Fields("producto") = "" & mytablex.Fields("producto")
                    'mytabley.Fields("serie") = "" & mytablex.Fields("serie")
                    'mytabley.Fields("marca") = busca_xc("" & mytablex.Fields("producto"))
                    'mytabley.Fields("color") = busca_xc1("" & mytablex.Fields("producto"))
                    mytabley.Fields("descripcio") = "" & mytablex.Fields("descripcio")
                    mytabley.Fields("cantidad") = Val("" & mytabley.Fields("cantidad")) + Val("" & mytablex.Fields("saldoant"))
                    'mytabley.Fields("local") = extra_loquesea(LOCAL1)
                    mytabley.Fields("saldo") = busca_saldoxx(extra_loquesea(local1), "" & mytablex.Fields("producto"), "" & mytablex.Fields("bodega"))
                    mytabley.Update
                Else
                    mytabley.Fields("producto") = "" & mytablex.Fields("producto")
                    'mytabley.Fields("serie") = "" & mytablex.Fields("serie")
                    mytabley.Fields("cantidad") = Val("" & mytabley.Fields("cantidad")) + Val("" & mytablex.Fields("saldoant"))
                    mytabley.Update

                End If

                mytabley.Close
                Set mytabley = Nothing
                mytablex.MoveNext
            Loop

        End If

        mytablex.Close
        Set mytablex = Nothing
        dbtconteo.MoveNext
    Loop
    Command3.Visible = False
    MsgBox "Presione enter para generar reporte en excell ", 48, "Aviso"
    imprime_excellc

    Exit Sub
cmd7866_err:
    Command3.Visible = False
    MsgBox "Aviso en todos los productos " + error$, 48, "Aviso"
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

Function busca_ubicacion(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from ubicacion where ubicacion='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_ubicacion = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close

End Function

Sub imprime_excellc()

    Dim mytablex As New ADODB.Recordset

    Dim v, h As Double

    Dim found       As Integer

    Dim I           As Integer

    Dim sdx         As Double

    Dim buf         As String

    Dim Tmp         As String

    Dim sw          As Integer

    Dim sdx1        As Double

    Dim sdx2        As Double
 
    Dim vprecios(7) As String

    Dim Heading(12) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    On Error GoTo cmd5612_err

    'Data1.Refresh
   
    'cn.Execute ("update pdtdetmp set saldo=select sum(saldo) from almacen where producto where local='" & extra_loquesea(local1) & "'")

    buf = "select * from pdtdetmp "
    mytablex.Open buf, cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If
     
    '' 03/07/2018 Conteo Fisico Sistema
    'Heading(1) = "Nro"
    'Heading(2) = "Producto"
    'Heading(3) = "Descripcio"
    'Heading(4) = "Conteo"
    'Heading(5) = "Stock"
    'Heading(6) = "Diferencia"
    Heading(1) = "FECHA"
    Heading(2) = "N°"
    Heading(3) = "PRODUCTO"
    Heading(4) = "DESCRIPCION"
    Heading(5) = "CONTEO FÍSICO"
    Heading(6) = "STOCK SISTEMA"
    Heading(7) = "Sobrante"
    Heading(8) = "Faltante"
     
    Heading(9) = "Costo Ult"
    Heading(10) = "Costo Prom"
     
    '' 03/07/2018 Conteo Fisico Sistema
     
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
  
    '---------------------------
     
    '' 03/07/2018 Conteo Fisico Sistema
    Call Formato_Excel(10, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    objExcel.ActiveSheet.Cells(1, 3) = "                                LISTA DE CONTEO FÍSICO/SISTEMA DE PRODUCTOS"
    objExcel.ActiveSheet.Cells(1, 3).Font.bold = True
    objExcel.ActiveSheet.Cells(1, 3).Font.Size = 14
    objExcel.ActiveSheet.Cells(1, 3).Font.color = RGB(0, 112, 184)
    '' 03/07/2018 Conteo Fisico Sistema
     
    With objExcel.ActiveSheet
    
        For I = 1 To 11 Step 1
            .Cells(3, I) = Heading(I)
        Next I
       
        .columns("A").ColumnWidth = 10
        .columns("B").ColumnWidth = 10
        .columns("C").ColumnWidth = 10
        .columns("d").ColumnWidth = 50
        .columns("e").ColumnWidth = 20
        .columns("f").ColumnWidth = 20
        .columns("g").ColumnWidth = 15
        .columns("h").ColumnWidth = 15
        .columns("i").ColumnWidth = 15
        .columns("j").ColumnWidth = 15
        
    End With

    '---------------------------
     
    v = 4
    h = 1
    sdx = 0
    sdx2 = 0
    Do

        If mytablex.EOF Then Exit Do
        sdx1 = 0
        sdx2 = sdx2 + 1
            
        objExcel.ActiveSheet.Cells(v, h) = "'" & periodo
        objExcel.ActiveSheet.Cells(v, h + 1) = "'" & sdx2
        objExcel.ActiveSheet.Cells(v, h + 2) = "'" & mytablex.Fields("producto")
        objExcel.ActiveSheet.Cells(v, h + 3) = "'" & mytablex.Fields("descripcio")
        objExcel.ActiveSheet.Cells(v, h + 4) = "" & mytablex.Fields("cantidad")
        objExcel.ActiveSheet.Cells(v, h + 5) = "" & mytablex.Fields("saldo")
            
        '' 03/07/2018 Conteo Fisico Sistema
        'sdx1 = Val("" & mytablex.Fields("saldo")) - Val("" & mytablex.Fields("cantidad"))
        'objExcel.ActiveSheet.Cells(v, h + 5) = sdx1
        sdx1 = Val("" & mytablex.Fields("cantidad")) - Val("" & mytablex.Fields("saldo"))
                
        If Val(sdx1) > 0 Then
            objExcel.ActiveSheet.Cells(v, h + 6) = sdx1
        ElseIf Val(sdx1) < 0 Then
            objExcel.ActiveSheet.Cells(v, h + 7) = sdx1 * -1

        End If
            
        objExcel.ActiveSheet.Cells(v, h + 8) = busca_CostosProducto("" & mytablex.Fields("producto"), 1)
        objExcel.ActiveSheet.Cells(v, h + 9) = busca_CostosProducto("" & mytablex.Fields("producto"), 2)
            
        '' 03/07/2018 Conteo Fisico Sistema
            
        'If mytabley.State = 1 Then mytabley.Close
        'mytabley.Open "select saldo from almacen where local='" & LOCAL1 & "' and producto='", cn, adOpenDynamic, adLockOptimistic
        'mytablex.Open buf, cn, adOpenDynamic, adLockOptimistic
            
        '' 03/07/2018 Conteo Fisico Sistema
        ' sdx = sdx + Val("" & mytablex.Fields("cantidad"))
        '' 03/07/2018 Conteo Fisico Sistema
            
        v = v + 1
        mytablex.MoveNext
    Loop
    mytablex.Close
     
    '' 03/07/2018 Conteo Fisico Sistema
    'objExcel.ActiveSheet.Cells(v, h + 1) = "" & sdx2
    'objExcel.ActiveSheet.Cells(v, h + 5) = "" & sdx
    '' 03/07/2018 Conteo Fisico Sistema
     
    MsgBox "Proceso finalizado ", 48, "Aviso"
    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    Exit Sub
cmd5612_err:
    MsgBox "Error en impresion excell " + error$, 48, "Aviso"
    Exit Sub

End Sub

'' 03/07/2018 Conteo Fisico Sistema
Function busca_CostosProducto(buf As String, tipo As Integer) As Double

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select costou,costop from producto where producto='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If tipo = 1 Then 'Costo ultimo
            busca_CostosProducto = Val(Format(Val("" & mytablex.Fields("costou")), "0.00000"))
        ElseIf tipo = 2 Then 'Costo promedio
            busca_CostosProducto = Val(Format(Val("" & mytablex.Fields("costop")), "0.00000"))

        End If
  
    End If

    mytablex.Close

End Function

'' 03/07/2018 Conteo Fisico Sistema

Function busca_saldoxx(buf As String, buf1 As String, buf2 As String) As Double

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Double

    sdx = 0
    buf = "select * from almacen where local='" & buf & "' and producto='" & buf1 & "' and bodega='" & buf2 & "'"
    mytablex.Open buf, cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        Do

            If mytablex.EOF Then Exit Do
            sdx = sdx + Val("" & mytablex.Fields("SALDO"))
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close
    busca_saldoxx = sdx

End Function

Function busca_xc(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from producto where producto='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_xc = "" & mytablex.Fields("MARCA")

    End If

    mytablex.Close

End Function

Function busca_xc1(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from producto where producto='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_xc1 = "" & mytablex.Fields("COLOR")

    End If

    mytablex.Close

End Function

Sub carga_ubicacion(buf As String)

    Dim mytablex As New ADODB.Recordset

    ubicacion.Clear
    ubicacion.AddItem "%"
    mytablex.Open "select * from ubicacion ", cn, adOpenDynamic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        ubicacion.AddItem "" & mytablex.Fields("ubicacion") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    mytablex.Close
    ubicacion.ListIndex = 0

End Sub

