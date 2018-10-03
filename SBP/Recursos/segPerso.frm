VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form segPerso 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seguimiento Entrada Salida"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   14625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   14625
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox ordenado 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10320
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "segPerso.frx":0000
      Height          =   6975
      Left            =   120
      OleObjectBlob   =   "segPerso.frx":0014
      TabIndex        =   9
      Top             =   1560
      Width           =   14295
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   11040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12600
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
   Begin VB.ComboBox personal 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3840
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   600
      Width           =   5055
   End
   Begin VB.TextBox fechaf 
      Height          =   375
      Left            =   5160
      MaxLength       =   10
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox fechai 
      Height          =   375
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ComboBox dpto 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ordenado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Personal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   240
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Departamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
   Begin VB.Menu fdlo232 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "segPerso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
sql_ingresos
End Sub

Private Sub fdlo232_Click()
segPerso.Hide
Unload segPerso

End Sub

Private Sub Form_Activate()
carga_inicial
End Sub
Sub sql_ingresos()
On Error GoTo cmd37_err
Dim buf As String
If Not IsDate(fechai) Then Exit Sub
If Not IsDate(fechaf) Then Exit Sub
buf = "select * from sisper where "
buf = buf & "  fechag>=" & "DateValue('" & fechai & "'" & ")"
buf = buf & " and  fechag<=" & "DateValue('" & fechaf & "'" & ")"
If dpto <> "%" Then
   buf = buf & " and dpto='" & dpto & "'"
End If
If personal <> "%" Then
   buf = buf & " and codigo='" & extra_loquesea(personal) & "'"
End If
'If ordenado <> "%" Then
'buf = buf & " order by " & ordenado & " ,fecha"
'  Else
 buf = buf & " order by fechag"
'End If
'MsgBox buf
               Data2.Connect = "foxpro 2.5;"
               Data2.DatabaseName = globaldir
               Data2.RecordSource = buf
               Data2.Refresh
               dbGrid1.SetFocus
Exit Sub
cmd37_err:
MsgBox "Error en select " & error$, 48, "Aviso"
Exit Sub

End Sub
Sub carga_inicial()
personal.Clear
personal.AddItem "%"
Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable("vendedor")
Do
If mytablex.EOF Then Exit Do
personal.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
mytablex.MoveNext
Loop
mytablex.Close
personal.ListIndex = 0
End Sub

Private Sub Form_Load()
fechai = "01/" + Format(Month(Now), "00") & "/" & Format(Year(Now), "000")
fechaf = Format(Now, "dd/mm/yyyy")
dpto.Clear
dpto.AddItem "%"
dpto.ListIndex = 0

ordenado.Clear
ordenado.AddItem "%"
ordenado.AddItem "Codigo"
ordenado.AddItem "Fecha"
ordenado.ListIndex = 0
End Sub
