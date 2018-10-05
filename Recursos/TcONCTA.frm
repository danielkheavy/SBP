VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form tconctaX 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tabla de Cuentas Contables"
   ClientHeight    =   6675
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Consultas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   11415
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   5160
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000D&
         Caption         =   "&Ejecutar"
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
         Left            =   8160
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox buffer 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "TcONCTA.frx":0000
         Height          =   5295
         Left            =   120
         OleObjectBlob   =   "TcONCTA.frx":0014
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1080
         Width           =   11055
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFF00&
      Caption         =   "Tipo de Analisis"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2640
      TabIndex        =   27
      Top             =   2520
      Width           =   4455
      Begin VB.OptionButton Option14 
         BackColor       =   &H00FFFF00&
         Caption         =   "Solo Detalle"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton Option13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Cuenta de Banco"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton Option12 
         BackColor       =   &H00FFFF00&
         Caption         =   "Por Documento"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton Option11 
         BackColor       =   &H00FFFF00&
         Caption         =   "Sin Analisis"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Tipo de Cuenta"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   19
      Top             =   2400
      Width           =   2415
      Begin VB.OptionButton Option10 
         BackColor       =   &H00FFFF00&
         Caption         =   "Mayor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   1680
         Width           =   1815
      End
      Begin VB.OptionButton Option9 
         BackColor       =   &H00FFFF00&
         Caption         =   "Orden"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   1440
         Width           =   1815
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00FFFF00&
         Caption         =   "Funcion"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   1815
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FFFF00&
         Caption         =   "Naturaleza"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FFFF00&
         Caption         =   "Resultado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFF00&
         Caption         =   "Pasivo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFF00&
         Caption         =   "Activo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Nivel Cuenta"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Width           =   6975
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Registro"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFF00&
         Caption         =   "Sub-Cuenta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Balance"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox thaber 
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
      Left            =   2640
      MaxLength       =   6
      TabIndex        =   13
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox tdebe 
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
      Left            =   2640
      MaxLength       =   6
      TabIndex        =   11
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox descripcio 
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
      Left            =   1560
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1200
      Width           =   5535
   End
   Begin VB.TextBox codigo 
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
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
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
      Height          =   615
      Left            =   1440
      MaskColor       =   &H00E0E0E0&
      Picture         =   "TcONCTA.frx":09DF
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Grabar registro"
      Top             =   0
      Width           =   735
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
      Height          =   615
      Left            =   0
      Picture         =   "TcONCTA.frx":1BF1
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Nuevo registro"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdHelp 
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
      Left            =   3600
      MaskColor       =   &H00E0E0E0&
      Picture         =   "TcONCTA.frx":2E03
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Ayuda"
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
      Left            =   2160
      MaskColor       =   &H00E0E0E0&
      Picture         =   "TcONCTA.frx":4015
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir"
      Top             =   0
      Width           =   735
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
      Left            =   4320
      MaskColor       =   &H00E0E0E0&
      Picture         =   "TcONCTA.frx":5227
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdDelete 
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
      Picture         =   "TcONCTA.frx":6439
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Borrar registro"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "TcONCTA.frx":764B
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Consulta"
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF00&
      Caption         =   "10 Balance 10.1 Subc.  10.1.10 Registro"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   6360
      Width           =   4335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF00&
      Caption         =   "Notas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      Caption         =   "1.balance 2.Subcuenta 3.registro"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   6120
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cuenta Amarre al Haber"
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
      TabIndex        =   14
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cuenta amarre al  Debe"
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
      TabIndex        =   12
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripcion"
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
      TabIndex        =   10
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo"
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
      TabIndex        =   9
      Top             =   840
      Width           =   1335
   End
   Begin VB.Menu ajdu1 
      Caption         =   "&Add"
   End
   Begin VB.Menu bo712 
      Caption         =   "&Borrar"
   End
   Begin VB.Menu grba1 
      Caption         =   "&Grabar"
   End
   Begin VB.Menu djuer1 
      Caption         =   "&Reporte"
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tconctaX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub ajdu1_Click()
If Frame1.Visible = True Then Exit Sub
inicializa
codigo = ""
codigo.SetFocus

End Sub

Private Sub bo712_Click()
Dim found As Integer
If Frame1.Visible = True Then Exit Sub
found = busca_registro()
If found = 0 Then
   MsgBox "No existe registro", 48, "Aviso"
   Exit Sub
End If
found = borra_registro()
If found = 0 Then Exit Sub
MsgBox "Ok,Registro Borrado", 48, "Aviso"
codigo = ""
inicializa
codigo.SetFocus
End Sub

Private Sub buffer_DblClick()
Command1_Click
End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   Frame1.Visible = False
   codigo.SetFocus
   Exit Sub
End If
Command1_Click

End Sub


Private Sub cmdAddEntry_Click()
ajdu1_Click
End Sub

Private Sub cmdDelete_Click()
bo712_Click
End Sub

Private Sub cmdExit_Click()
dlo132_Click

End Sub

Private Sub cmdPrint_Click()
djuer1_Click
End Sub

Private Sub cmdSave_Click()
grba1_Click
End Sub

Private Sub cmdSort_Click()
Frame1.Visible = True
buffer = ""
buffer.SetFocus
Command1_Click
End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   dlo132_Click
   Exit Sub
End If
If Len(codigo) = 0 Then Exit Sub
found = busca_registro()
If found = 0 Then
   inicializa
End If
descripcio.SetFocus
End Sub

Private Sub codigo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
cmdSort_Click
End If
End Sub

Private Sub Command1_Click()
Dim buf As String
If Len(buffer) = 0 Then
buf = "select Cuenta,Nombre,Rnf as T,Bd as N,Cta as B,Tdebe as TraDebe,Thaber as TraHaber from mdh_plan "
Else
buf = "select Cuenta,Nombre,Rnf as T,Bd as N,Cta as B,Tdebe as TraDebe,Thaber as TraHaber from mdh_plan where " & Combo1 & " like '" & buffer & "%'"
End If


               Data1.Connect = "foxpro 2.5;"
               Data1.DatabaseName = globalcont
               Data1.RecordSource = buf
               Data1.Refresh
               If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then
                  Data1.Recordset.Close
                  buffer.SetFocus
                  Exit Sub
               End If
               dbGrid1.Columns(0).Width = 1500
               dbGrid1.Columns(1).Width = 5000
               dbGrid1.Columns(2).Width = 600
               dbGrid1.Columns(3).Width = 600
               dbGrid1.Columns(4).Width = 600
               dbGrid1.Columns(5).Width = 1000
               dbGrid1.Columns(6).Width = 1000
               dbGrid1.SetFocus

End Sub




Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
 
  buffer.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
   codigo = dbGrid1.Columns(0)
   Frame1.Visible = False
   codigo.SetFocus
   codigo_KeyPress 13
End If
End Sub


Private Sub descripcio_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub


End Sub



Private Sub descripcio_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   codigo.SetFocus
   Exit Sub
End If

End Sub

Private Sub djuer1_Click()
If Frame1.Visible = True Then Exit Sub
reporgen.NAMETABLA = "mdh_plan"
reporgen.Show 1

End Sub

Private Sub dlo132_Click()
If Frame1.Visible = True Then
   Frame1.Visible = False
   Exit Sub
End If
tconctaX.Hide
Unload tconctaX
End Sub



Private Sub Form_Load()
Combo1.Clear
Combo1.AddItem "cuenta"
Combo1.AddItem "nombre"
Combo1.ListIndex = 0
End Sub
Sub inicializa()
descripcio = ""
tdebe = ""
thaber = ""
Option1.Value = False
Option2.Value = False
Option3.Value = False

Option4.Value = False
Option5.Value = False
Option6.Value = False
Option7.Value = False
Option8.Value = False
Option9.Value = False
Option10.Value = False

Option11.Value = False
Option12.Value = False
Option13.Value = False
Option14.Value = False


End Sub
Function borra_registro()

Dim mytablex As Table

Set mytablex = mydbzglo.OpenTable("mdh_plan")
mytablex.Index = "mdh_plan"
mytablex.Seek "=", codigo
If Not mytablex.NoMatch Then
   If MsgBox("Desea Borra el registro", 1, "Aviso") = "1" Then
      mytablex.Delete
      borra_registro = 1
   End If
End If
'------------------------------------- ------------
mytablex.Close
 

End Function
Function busca_registro()

Dim mytablex As Table

Set mytablex = mydbzglo.OpenTable("mdh_plan")
mytablex.Index = "mdh_plan"
mytablex.Seek "=", codigo
If Not mytablex.NoMatch Then
   pone_registro mytablex
   busca_registro = 1
End If
'------------------------------------- ------------
mytablex.Close
 
End Function
Sub pone_registro(mytablex As Table)
tdebe = "" & mytablex.Fields("tdebe")
thaber = "" & mytablex.Fields("thaber")
codigo = "" & mytablex.Fields("cuenta")
descripcio = "" & mytablex.Fields("nombre")
Option1.Value = False
Option2.Value = False
Option3.Value = False

Option4.Value = False
Option5.Value = False
Option6.Value = False
Option7.Value = False
Option8.Value = False
Option9.Value = False
Option10.Value = False

Option11.Value = False
Option12.Value = False
Option13.Value = False
Option14.Value = False


If "" & mytablex.Fields("bd") = "1" Then
   Option1.Value = True
End If
If "" & mytablex.Fields("bd") = "2" Then
   Option2.Value = True
End If
If "" & mytablex.Fields("bd") = "3" Then
   Option3.Value = True
End If


If "" & mytablex.Fields("rnf") = "A" Then
   Option4.Value = True
End If
If "" & mytablex.Fields("rnf") = "P" Then
   Option5.Value = True
End If
If "" & mytablex.Fields("rnf") = "R" Then
   Option6.Value = True
End If
If "" & mytablex.Fields("rnf") = "N" Then
   Option7.Value = True
End If
If "" & mytablex.Fields("rnf") = "F" Then
   Option8.Value = True
End If
If "" & mytablex.Fields("rnf") = "O" Then
   Option9.Value = True
End If
If "" & mytablex.Fields("rnf") = "M" Then
   Option10.Value = True
End If

If Trim("" & mytablex.Fields("cta")) = "" Then  '
   Option11.Value = True
   
End If
If Trim("" & mytablex.Fields("cta")) = "S" Then  '
   Option12.Value = True
End If
If Trim("" & mytablex.Fields("cta")) = "B" Then  '
   Option13.Value = True
End If
If Trim("" & mytablex.Fields("cta")) = "X" Then  '
   Option14.Value = True
End If


End Sub
Sub grabando(mytablex As Table)
mytablex.Fields("cuenta") = codigo
mytablex.Fields("nombre") = descripcio
If Option1.Value = True Then
   mytablex.Fields("bd") = "1" 'BALANCE
End If
If Option2.Value = True Then
   mytablex.Fields("bd") = "2" 'SUB CUENTA
End If
If Option3.Value = True Then
   mytablex.Fields("bd") = "3" 'REGISTRO
End If

If Option4.Value = True Then
   mytablex.Fields("rnf") = "A" 'ACTIVO
End If
If Option5.Value = True Then
   mytablex.Fields("rnf") = "P" 'PASIVO
End If
If Option6.Value = True Then
   mytablex.Fields("rnf") = "R" 'RESULTADO
End If
If Option7.Value = True Then
   mytablex.Fields("rnf") = "N" 'NATURALEZA
End If
If Option8.Value = True Then
   mytablex.Fields("rnf") = "F" 'FUNCTION
End If
If Option9.Value = True Then
   mytablex.Fields("rnf") = "O" 'ORDEN
End If
If Option10.Value = True Then
   mytablex.Fields("rnf") = "M" 'MAYOR
End If



If Option11.Value = True Then
   mytablex.Fields("cta") = "" 'sin analisis
End If
If Option12.Value = True Then
   mytablex.Fields("cta") = "S" 'por documentos
End If
If Option13.Value = True Then
   mytablex.Fields("cta") = "B" 'cuenta de banco
End If
If Option14.Value = True Then
   mytablex.Fields("cta") = "X" 'solo detalle
End If

End Sub

Private Sub grba1_Click()
Dim found As Integer
If Frame1.Visible = True Then Exit Sub
found = grabar()
If found = 0 Then Exit Sub
codigo.SetFocus

End Sub

Private Sub Label1_Click()
cmdSort_Click
End Sub


Function grabar()
Dim found As Integer
Dim mytablex As Table

found = valida()
If found = 0 Then
   MsgBox "Campos invalidos", 48, "Aviso"
   Exit Function
End If

Set mytablex = mydbzglo.OpenTable("mdh_plan")
mytablex.Index = "mdh_plan"
mytablex.Seek "=", codigo
If mytablex.NoMatch Then
   mytablex.AddNew
   grabando mytablex
   mytablex.Update
   grabar = 1
End If
If Not mytablex.NoMatch Then
   If MsgBox("Desea Reescribir?", 1, "Aviso") = 1 Then
   mytablex.Edit
   grabando mytablex
   mytablex.Update
   grabar = 1
   End If
End If
'------------------------------------- ------------
mytablex.Close
 
End Function

Function valida()
If Len(codigo) = 0 Then
   codigo.SetFocus
   Exit Function
End If
If Len(descripcio) = 0 Then
   descripcio.SetFocus
   Exit Function
End If
valida = 1
End Function

