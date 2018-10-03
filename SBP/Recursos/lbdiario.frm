VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form lbdiario 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Libro Diario"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   12930
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   12930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox origen 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   120
      Width           =   3015
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   11640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5880
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Click Cancela"
      Height          =   2295
      Left            =   3240
      TabIndex        =   18
      Top             =   2160
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.TextBox nrolineas 
      BackColor       =   &H00FFFFFF&
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
      Left            =   8400
      MaxLength       =   30
      TabIndex        =   14
      Text            =   "45"
      Top             =   6480
      Width           =   1575
   End
   Begin VB.TextBox titulo 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   13
      Text            =   "Libro Diario del Mes:"
      Top             =   6480
      Width           =   3855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&ExportaExcel"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Periodo"
      Height          =   735
      Left            =   9960
      TabIndex        =   6
      Top             =   7680
      Visible         =   0   'False
      Width           =   2895
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FFFF00&
         Caption         =   "Acumulado"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFF00&
         Caption         =   "Mensual"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Tipo"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   5295
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Resumen"
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Analitico"
         Height          =   375
         Left            =   600
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "&Ejecutar"
      Height          =   375
      Left            =   9360
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "lbdiario.frx":0000
      Height          =   4455
      Left            =   120
      OleObjectBlob   =   "lbdiario.frx":0014
      TabIndex        =   0
      Top             =   960
      Width           =   12735
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   11640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5520
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF00&
      Caption         =   "Totales Us$"
      Height          =   375
      Left            =   4320
      TabIndex        =   23
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label tdebed 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   5760
      TabIndex        =   22
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label thaberd 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7320
      TabIndex        =   21
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Periodo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   20
      Top             =   480
      Width           =   855
   End
   Begin VB.Label periodo 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   17
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lineas x Pagina"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   16
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Titulo reporte"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Label thaber 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7320
      TabIndex        =   11
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label tdebe 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF00&
      Caption         =   "Totales S/."
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Origen"
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Menu imro9343 
      Caption         =   "&Imprime"
   End
   Begin VB.Menu lomj454 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "lbdiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim buf As String
Dim buf2 As String
If Command4.Visible = True Then Exit Sub
buf2 = Trim(extra_loquesea("" & origen))
If Option1.Value = True Then
   buf = "select T as Or,Vou,Ncuenta as Descripcio,Debe,Haber,Cuenta,Moneda as M,Tc,Fecha,Glosa as Concepto,Rut as Codigo,Rs as Razon_Social,Doc,Numero,Fechad as Emision,mes from mdh_vou where mes='" & Trim(periodo) & "'"
   If origen <> "*" Then
      If Len(buf2) > 0 Then
         buf = buf & " and  t='" & buf2 & "'"
      End If
   End If
   buf = buf & " order by t,val(vou)"
End If
If Option3.Value = True Then
   buf = "select T as Or,Cuenta,Ncuenta as Descripcio,Moneda as M,sum(debe) as Debito ,sum(haber) as Credito from mdh_vou where mes='" & periodo & "'"
   If origen <> "*" Then
      If Len(buf2) > 0 Then
         buf = buf & " and t='" & buf2 & "'"
      End If
   End If
   buf = buf & " group by T,Cuenta,Ncuenta,Moneda"
End If

Data1.Connect = "foxpro 2.5;"
Data1.DatabaseName = globalcont
Data1.RecordSource = buf
Data1.Refresh
If Option3.Value = True Then
suma_grid 1
End If
If Option3.Value = True Then
DBGrid1.Columns(0).Width = 400
DBGrid1.Columns(1).Width = 1500
DBGrid1.Columns(2).Width = 5000
End If
If Option1.Value = True Then
DBGrid1.Columns(0).Width = 400
DBGrid1.Columns(1).Width = 900
DBGrid1.Columns(2).Width = 3000
DBGrid1.Columns(3).Width = 900
DBGrid1.Columns(4).Width = 900
DBGrid1.Columns(5).Width = 1000
DBGrid1.Columns(6).Width = 400
DBGrid1.Columns(7).Width = 700
DBGrid1.Columns(8).Width = 1400
DBGrid1.Columns(9).Width = 3000
DBGrid1.Columns(10).Width = 1500
DBGrid1.Columns(11).Width = 3000
DBGrid1.Columns(12).Width = 400
DBGrid1.Columns(13).Width = 1500
DBGrid1.Columns(14).Width = 1300
suma_grid 0
End If

End Sub

Sub suma_grid(sw As Integer)

Dim xtdebes As Double
Dim xtdebed As Double
Dim xthabers As Double
Dim xthaberd As Double

Dim xdebes As Double
Dim xdebed As Double
Dim xhabers As Double
Dim xhaberd As Double
Dim tc As Double

   xdebes = 0
   xhabers = 0
   xdebed = 0
   xhaberd = 0

Do
If Data1.Recordset.EOF Then Exit Do
If sw = 0 Then
tc = Val("" & Data1.Recordset.Fields("tc"))
If tc <= 0 Then
   tc = 1
End If

   xtdebes = 0
   xthabers = 0
   xtdebed = 0
   xthaberd = 0
If "" & Data1.Recordset.Fields("m") = "S" Then
   xtdebes = Val("" & Data1.Recordset.Fields("debe"))
   xthabers = Val("" & Data1.Recordset.Fields("haber"))
   xtdebed = Val("" & Data1.Recordset.Fields("debe")) / tc
   xthaberd = Val("" & Data1.Recordset.Fields("haber")) / tc
End If
If "" & Data1.Recordset.Fields("m") = "D" Then
   xtdebes = Val("" & Data1.Recordset.Fields("debe")) * tc
   xthabers = Val("" & Data1.Recordset.Fields("haber")) * tc
   xtdebed = Val("" & Data1.Recordset.Fields("debe"))
   xthaberd = Val("" & Data1.Recordset.Fields("haber"))
End If
xdebes = xdebes + xtdebes
xhabers = xhabers + xthabers
xdebed = xdebed + xtdebed
xhaberd = xhaberd + xthaberd
End If
If sw = 1 Then
tc = 1
If tc <= 0 Then
   tc = 1
End If

   xtdebes = 0
   xthabers = 0
   xtdebed = 0
   xthaberd = 0
If "" & Data1.Recordset.Fields("m") = "S" Then
   xtdebes = Val("" & Data1.Recordset.Fields("debito"))
   xthabers = Val("" & Data1.Recordset.Fields("credito"))
   xtdebed = Val("" & Data1.Recordset.Fields("debito")) / tc
   xthaberd = Val("" & Data1.Recordset.Fields("credito")) / tc
End If
If "" & Data1.Recordset.Fields("m") = "D" Then
   xtdebes = Val("" & Data1.Recordset.Fields("debito")) * tc
   xthabers = Val("" & Data1.Recordset.Fields("credito")) * tc
   xtdebed = Val("" & Data1.Recordset.Fields("debito"))
   xthaberd = Val("" & Data1.Recordset.Fields("credito"))
End If
xdebes = xdebes + xtdebes
xhabers = xhabers + xthabers
xdebed = xdebed + xtdebed
xhaberd = xhaberd + xthaberd
End If
Data1.Recordset.MoveNext
Loop
'soles
tdebe = Format(xdebes, "0.00")
thaber = Format(xhabers, "0.00")
'Dolares
tdebed = Format(xdebed, "0.00")
thaberd = Format(xhaberd, "0.00")
End Sub


Private Sub Command2_Click()
If Command4.Visible = True Then Exit Sub
End Sub

Private Sub Command3_Click()

End Sub
Sub imprime_detalle()
Dim found As Integer
Dim buf As String
contlin = 0
suma1 = 0
suma2 = 0
suma3 = 0
ssuma1 = 0
ssuma2 = 0
ssuma3 = 0
If Command4.Visible = True Then Exit Sub
    Filename = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    borra_nombre Filename
    Open Filename For Append As #1
    '------------------------------------
    cabecera_documento
    cuerpo_programa_documento
    '------------------------------------
    Close #1
    cerrar_archivo
    genver.Show 1
End Sub
Sub imprime_resumen()
Dim found As Integer
Dim buf As String
contlin = 0
suma1 = 0
suma2 = 0
suma3 = 0
ssuma1 = 0
ssuma2 = 0
ssuma3 = 0
If Command4.Visible = True Then Exit Sub
    Filename = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    borra_nombre Filename
    Open Filename For Append As #1
    '------------------------------------
    cabecera_resumen
    cuerpo_programa_resumen
    '------------------------------------
    Close #1
    cerrar_archivo
    genver.Show 1

End Sub

Private Sub Command4_Click()
If Command4.Visible = True Then
   Command4.Visible = False
End If

End Sub

Private Sub Form_Load()
Dim found As Integer
carga_como
found = busca_parame()
Command1_Click

End Sub
Sub carga_como()
Dim mydbx As Database
Dim mytablex As Table
origen.Clear
origen.AddItem "*"
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("origen")
Do
If mytablex.EOF Then Exit Do
origen.AddItem "" & mytablex.Fields("origen") & "|" & "" & mytablex.Fields("descripcio")
mytablex.MoveNext
Loop
mytablex.Close
mydbx.Close
origen.ListIndex = 0

End Sub

Private Sub Label2_Click()

End Sub

Private Sub imro9343_Click()
If Option1.Value = True Then
   imprime_detalle
End If
If Option3.Value = True Then
   imprime_resumen
End If

End Sub

Private Sub lomj454_Click()
If Command4.Visible = True Then Exit Sub
lbdiario.Hide
Unload lbdiario
End Sub
Sub cabecera_resumen()
Dim buf As String
Dim i As Integer
Dim found As Integer
    If contlin > 0 Then
       buf = Chr$(12)
       found = formateaa(buf, Len(buf), 0, 0)
    End If
    contpag = contpag + 1
    contlin = 0
    'cabecera_tipico "" & menuipos!nempresa, "", "" & "" & usuariopos
    buf = titulo & " " & periodo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Libro Diario del Mes de : ", 25, 2, 0)
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)
    'found = formateaa("Mes", 7, 0, 0)
    found = formateaa("Cuenta", 11, 0, 0)
    found = formateaa("Nombre", 31, 0, 0)
    found = formateaa("Debito ", 11, 0, 1)
    found = formateaa("Credito ", 11, 2, 1)
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)

End Sub
Sub cabecera_documento()
Dim buf As String
Dim i As Integer
Dim found As Integer
    If contlin > 0 Then
       buf = Chr$(12)
       found = formateaa(buf, Len(buf), 0, 0)
    End If
    contpag = contpag + 1
    contlin = 0
    'cabecera_tipico "" & menuipos!nempresa, "", "" & "" & usuariopos
    buf = titulo & " " & periodo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Libro Diario del Mes de : ", 25, 2, 0)
        
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)
    found = formateaa("Fecha", 11, 0, 0)
    found = formateaa("Cuenta", 11, 0, 0)
    found = formateaa("Nombre", 31, 0, 0)
    found = formateaa("Concepto", 31, 0, 0)
    found = formateaa("Debito ", 11, 0, 1)
    found = formateaa("Credito ", 11, 2, 1)
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)
    

End Sub
Sub cuerpo_programa_documento()
Dim vr As String
Dim tmp As String
Dim tmp1 As String
Dim sw As Integer
Dim buf As String
Dim found As Integer
Dim sdx As Double
sdx = 0
sw = 0
suma1 = 0
suma2 = 0
ssuma1 = 0
ssuma2 = 0
ir_inicio
Command4.Visible = True
Do
vr = DoEvents()
If Command4.Visible = False Then Exit Do
If Data1.Recordset.EOF Then Exit Do

tmp1 = "" & Data1.Recordset.Fields("OR") & "" & Data1.Recordset.Fields("vou")
If sw = 0 Then
   buf = "" & Data1.Recordset.Fields("OR")
   found = formateaa(buf, 6, 0, 0)
   found = formateaa("", 1, 0, 0)
   buf = "" & Data1.Recordset.Fields("vou")
   found = formateaa(buf, 4, 0, 0)
   found = formateaa("", 1, 2, 0)
   nlineas
   sw = 1
   tmp = "" & Data1.Recordset.Fields("OR") & "" & Data1.Recordset.Fields("vou")
End If
If tmp <> tmp1 Then
   found = formateaa("", 84, 0, 0)
   buf = Format(suma1, "0.00")
   found = formateaa(buf, 10, 0, 1)
   found = formateaa("", 1, 0, 0)
   buf = Format(suma2, "0.00")
   found = formateaa(buf, 10, 0, 1)
   found = formateaa("", 1, 2, 0)
   nlineas
   
   buf = "" & Data1.Recordset.Fields("OR")
   found = formateaa(buf, 6, 0, 0)
   found = formateaa("", 1, 0, 0)
   buf = "" & Data1.Recordset.Fields("vou")
   found = formateaa(buf, 4, 0, 0)
   found = formateaa("", 1, 2, 0)
   nlineas
   tmp = "" & Data1.Recordset.Fields("OR") & "" & Data1.Recordset.Fields("vou")
   suma1 = 0
   suma2 = 0
End If
   buf = "" & Data1.Recordset.Fields("fecha")
   found = formateaa(buf, 10, 0, 0)
   found = formateaa("", 1, 0, 0)
   buf = "" & Data1.Recordset.Fields("cuenta")
   found = formateaa(buf, 10, 0, 0)
   found = formateaa("", 1, 0, 0)
   buf = busca_cuenta("" & Data1.Recordset.Fields("cuenta"))
   found = formateaa(buf, 30, 0, 0)
   found = formateaa("", 1, 0, 0)
   buf = "" & Data1.Recordset.Fields("concepto")  'glosa
   found = formateaa(buf, 30, 0, 0)
   found = formateaa("", 1, 0, 0)
   buf = "" & Data1.Recordset.Fields("debe")
   buf = Format(Val(buf), "0.00")
   If Val(buf) = 0 Then
      buf = ""
   End If
   
   found = formateaa(buf, 10, 0, 1)
   found = formateaa("", 1, 0, 0)
   buf = "" & Data1.Recordset.Fields("haber")
   buf = Format(Val(buf), "0.00")
   If Val(buf) = 0 Then
      buf = ""
   End If
   found = formateaa(buf, 10, 0, 1)
   found = formateaa("", 1, 2, 0)
   nlineas
   suma1 = suma1 + Val("" & Data1.Recordset.Fields("debe"))
   ssuma1 = ssuma1 + Val("" & Data1.Recordset.Fields("debe"))
   suma2 = suma2 + Val("" & Data1.Recordset.Fields("haber"))
   ssuma2 = ssuma2 + Val("" & Data1.Recordset.Fields("haber"))
Data1.Recordset.MoveNext
Loop
   found = formateaa("", 84, 0, 0)
   buf = Format(suma1, "0.00")
   found = formateaa(buf, 10, 0, 1)
   found = formateaa("", 1, 0, 0)
   buf = Format(suma2, "0.00")
   found = formateaa(buf, 10, 0, 1)
   found = formateaa("", 1, 2, 0)
   nlineas
   found = formateaa("Total Comprobante:", 84, 0, 0)
   buf = Format(ssuma1, "0.00")
   found = formateaa(buf, 10, 0, 1)
   found = formateaa("", 1, 0, 0)
   buf = Format(ssuma2, "0.00")
   found = formateaa(buf, 10, 0, 1)
   found = formateaa("", 1, 2, 0)
   
   Command4.Visible = False

End Sub
Sub cuerpo_programa_resumen()
Dim vr As String
Dim tmp As String
Dim tmp1 As String
Dim sw As Integer
Dim buf As String
Dim found As Integer
Dim sdx As Double
sdx = 0
sw = 0
suma1 = 0
suma2 = 0
ssuma1 = 0
ssuma2 = 0
ir_inicio
Command4.Visible = True
Do
vr = DoEvents()
If Command4.Visible = False Then Exit Do
If Data1.Recordset.EOF Then Exit Do
tmp1 = "" & Data1.Recordset.Fields("or")
If sw = 0 Then
   buf = "" & Data1.Recordset.Fields("or")
   found = formateaa(buf, 6, 0, 0)
   found = formateaa("", 1, 0, 0)
   buf = busca_origen("" & Data1.Recordset.Fields("or"))
   found = formateaa(buf, 30, 0, 0)
   found = formateaa("", 1, 2, 0)
   nlineas
   sw = 1
   tmp = "" & Data1.Recordset.Fields("or")
End If
If tmp <> tmp1 Then
   found = formateaa("", 42, 0, 0)
   buf = Format(suma1, "0.00")
   found = formateaa(buf, 10, 0, 1)
   found = formateaa("", 1, 0, 0)
   buf = Format(suma2, "0.00")
   found = formateaa(buf, 10, 0, 1)
   found = formateaa("", 1, 2, 0)
   nlineas
   buf = "" & Data1.Recordset.Fields("or")
   found = formateaa(buf, 6, 0, 0)
   found = formateaa("", 1, 0, 0)
   buf = busca_origen("" & Data1.Recordset.Fields("or"))
   found = formateaa(buf, 30, 0, 0)
   found = formateaa("", 1, 2, 0)
   nlineas
   tmp = "" & Data1.Recordset.Fields("or")
   suma1 = 0
   suma2 = 0
End If
   buf = "" & Data1.Recordset.Fields("cuenta")
   found = formateaa(buf, 10, 0, 0)
   found = formateaa("", 1, 0, 0)
   buf = busca_cuenta("" & Data1.Recordset.Fields("cuenta"))
   found = formateaa(buf, 30, 0, 0)
   found = formateaa("", 1, 0, 0)
   buf = "" & Data1.Recordset.Fields("debito")
   buf = Format(Val(buf), "0.00")
   If Val(buf) = 0 Then
      buf = ""
   End If
   found = formateaa(buf, 10, 0, 1)
   found = formateaa("", 1, 0, 0)
   buf = "" & Data1.Recordset.Fields("credito")
   buf = Format(Val(buf), "0.00")
   If Val(buf) = 0 Then
      buf = ""
   End If
   found = formateaa(buf, 10, 0, 1)
   found = formateaa("", 1, 2, 0)
   nlineas
   suma1 = suma1 + Val("" & Data1.Recordset.Fields("debito"))
   ssuma1 = ssuma1 + Val("" & Data1.Recordset.Fields("debito"))
   suma2 = suma2 + Val("" & Data1.Recordset.Fields("credito"))
   ssuma2 = ssuma2 + Val("" & Data1.Recordset.Fields("credito"))
Data1.Recordset.MoveNext
Loop
   found = formateaa("", 42, 0, 0)
   buf = Format(suma1, "0.00")
   found = formateaa(buf, 10, 0, 1)
   found = formateaa("", 1, 0, 0)
   buf = Format(suma2, "0.00")
   found = formateaa(buf, 10, 0, 1)
   found = formateaa("", 1, 2, 0)
   nlineas
   found = formateaa("Total Comprobante:", 42, 0, 0)
   buf = Format(ssuma1, "0.00")
   found = formateaa(buf, 10, 0, 1)
   found = formateaa("", 1, 0, 0)
   buf = Format(ssuma2, "0.00")
   found = formateaa(buf, 10, 0, 1)
   found = formateaa("", 1, 2, 0)
   
   Command4.Visible = False


End Sub
Sub nlineas()
    contlin = contlin + 1
    If contlin > Val(nrolineas) Then
       cabecera_documento
    End If
End Sub
Function busca_parame()
Dim sdx As Double
Dim mydbx As Database
Dim mytablex As Table
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("parame")
mytablex.Index = "codigo"
mytablex.Seek "=", "01"
If Not mytablex.NoMatch Then
   periodo = "" & mytablex.Fields("mesconta") & "" & mytablex.Fields("anoconta")
   busca_parame = 1
End If
'------------------------------------- ------------
mytablex.Close
mydbx.Close

End Function
Function busca_cuenta(buf As String) As String
Dim mydbx As Database
Dim mytablex As Table
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("mdh_plan")
mytablex.Index = "mdh_plan"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   busca_cuenta = "" & mytablex.Fields("nombre")
End If
'------------------------------------- ------------
mytablex.Close
mydbx.Close

End Function
Function busca_origen(buf As String) As String
Dim mydbx As Database
Dim mytablex As Table
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("origen")
mytablex.Index = "origen"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   busca_origen = "" & mytablex.Fields("descripcio")
End If
'------------------------------------- ------------
mytablex.Close
mydbx.Close

End Function
Sub ir_inicio()
On Error GoTo cmd1_err
Data1.Recordset.MoveFirst
Exit Sub
cmd1_err:
Exit Sub
End Sub
Function extra_loquesea(buf As String) As String
Dim j
Dim buf1 As String
buf1 = ""
If InStr(buf, "|") > 0 Then
j = InStr(buf, "|")
   buf1 = Mid$(buf, 1, j - 1)
End If
extra_loquesea = buf1
   
   
End Function



Private Sub Option1_Click()
Command1_Click
End Sub

Private Sub Option3_Click()
Command1_Click
End Sub

Private Sub Option6_Click()

End Sub

