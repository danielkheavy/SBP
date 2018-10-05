VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form tcrystal 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboprinter 
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
      Left            =   1680
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   2880
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Ejecutar"
      Height          =   1455
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   4575
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   8520
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Impresora"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label xtitulo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   6495
   End
   Begin VB.Label condicion 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   8985
   End
   Begin VB.Label archivoreporte 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
   Begin VB.Menu flo44 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tcrystal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Integer
Dim bfound As Boolean
If cboprinter.Text <> "None" Then
                bfound = False
                For i = 0 To Printers.count - 1
                    If Printers(i).DeviceName = cboprinter.Text Then
                        Set Printer = Printers(i)
                        bfound = True
                        Exit For
                    End If
                Next
End If
If bfound = True Then
   ejecutar_reporte
End If
End Sub

Private Sub flo44_Click()
tcrystal.Hide
Unload tcrystal
End Sub

Sub ejecutar_reporte()
On Error GoTo cmhj78_err
Screen.MousePointer = 11
'CrystalReport1.Connect = "DSN=xx;"
'CrystalReport1.Connect = "DSN=xx;UID=sa;PWD=;DBQ=<CRWDC>Database=calipso"
'CrystalReport1.Connect = "DSN=192.168.1.3;UID=sa;DSQ=calipso"
'CrystalReport1.Connect = "Provider=SQLOLEDB.1;Password=;Persist Security Info=;User ID=sa;Initial Catalog=NOMBREDELABASE;Data Source=IPDELSERVER;pass=;Network Library=dbmssocn;"
'CrystalReport1.Connect = "Provider=SQLOLEDB;Server=" & menup.vservidor & ";Database=calipso;UID=sa;PWD="
'CrystalReport1.Connect = "&quot;DSN=xx;UID=sa;PWD=;&quot;"
'CrystalReport1.Connect = "datasource=calipso;localtion=calipso;uid=sa;pwd=;"
'Dll de conexion
'estos dos sirven...
'"DSN= tuDSN_odbc ;UID=SA;PWD=;DSQ=dbo;"
CrystalReport1.Connect = "DSN=xx;UID=SA;PWD=;DSQ=calipso;"
'CrystalReport1.Connect = "DSN= xx ;UID=SA;PWD=;DSQ=calipso;"
'CrystalReport1.Connect = "Provider=SQLOLEDB;Server=" & menup.vservidor & ";Database=calipso;UID=sa;PWD="
'CrystalReport1.Connect = "Provider=SQLOLEDB;Server=" & Trim(menup.vservidor) & ";Database=calipso;UID=sa;PWD="


'CrystalReport1.LogOnServer "p2ssql.dll", menup.vservidor, "calipso", "sa", ""

CrystalReport1.ReportFileName = archivoreporte
CrystalReport1.DiscardSavedData = True
CrystalReport1.ProgressDialog = True
'CrystalReport1.SQLQuery = condicion
'DiarioVentas.RecordSelectionFormula = "{Facturacion.fecha} In Date (" & Format$(Desde, "yyyy,mm,dd" & " To Date (" & Format$(Hasta, "yyyy,mm,dd" & ""
'crtControl.ReplaceSelectionFormula "{Productos.CodigoRamo}=" & xcodigo
'MsgBox condicion
CrystalReport1.SelectionFormula = Trim(condicion)
CrystalReport1.Formulas(0) = "txttitulo='" & "" & xtitulo & "'"
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Action = 1
Screen.MousePointer = 0
Exit Sub
cmhj78_err:
MsgBox "No se puede ejecutar informe " + error$, 48, "Aviso"
Screen.MousePointer = 0
Exit Sub
End Sub

Private Sub Form_Load()
Dim i As Integer
On Error GoTo cmd9012_err
    cboprinter.Clear
    cboprinter.AddItem "None"
    'cboprinter.AddItem "File"
    Printer.Orientation = 1
    For i = 0 To Printers.count - 1
        cboprinter.AddItem Printers(i).DeviceName
    Next
    cboprinter.ListIndex = 0
    Exit Sub
cmd9012_err:
    MsgBox "NO existe impresoras valida al menos uno ", 48, "Aviso"
    'tcrystal.Hide
    'Unload tcrystal
    Exit Sub

End Sub
