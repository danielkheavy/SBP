VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tactfijo 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Control de Activos Fijos"
   ClientHeight    =   10650
   ClientLeft      =   165
   ClientTop       =   -45
   ClientWidth     =   18705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10650
   ScaleWidth      =   18705
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Frame2"
      Height          =   6615
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   12495
      Begin VB.TextBox etiquetas 
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
         Left            =   4560
         MaxLength       =   3
         TabIndex        =   34
         Top             =   5880
         Width           =   1335
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6960
         ScaleHeight     =   285
         ScaleWidth      =   1665
         TabIndex        =   33
         Top             =   4680
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox nsituacion 
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
         TabIndex        =   31
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   1560
         Picture         =   "tactfijo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   4680
         Width           =   1575
      End
      Begin VB.TextBox nconteo1 
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
         Left            =   2880
         MaxLength       =   6
         TabIndex        =   29
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox nconteo 
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
         TabIndex        =   27
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox nobserva 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   1560
         MaxLength       =   400
         TabIndex        =   25
         Top             =   2040
         Width           =   10575
      End
      Begin VB.TextBox nserie 
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
         TabIndex        =   23
         Top             =   1320
         Width           =   4215
      End
      Begin VB.TextBox ndescripcio 
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
         MaxLength       =   80
         TabIndex        =   21
         Top             =   600
         Width           =   10575
      End
      Begin VB.ComboBox bsituacion 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox nproducto 
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
         MaxLength       =   15
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Cerrar"
         Height          =   945
         Left            =   10680
         Picture         =   "tactfijo.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Imprimir todo"
         Top             =   4680
         Width           =   1470
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Imprime Codigo Barras"
         Height          =   975
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad Etiqueta"
         Height          =   375
         Left            =   3120
         TabIndex        =   35
         Top             =   5880
         Width           =   1455
      End
      Begin VB.Label nid 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3600
         TabIndex        =   32
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observaciones"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcio"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Producto"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Situacion"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFC0&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   14715
      TabIndex        =   2
      Top             =   0
      Width           =   14775
      Begin VB.ComboBox periodo 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   0
         Width           =   2535
      End
      Begin VB.ComboBox ubicacion 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   360
         Width           =   2055
      End
      Begin VB.ComboBox Seccion 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   0
         Width           =   2055
      End
      Begin VB.ComboBox division 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   720
         Width           =   2535
      End
      Begin VB.ComboBox Local1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   360
         Width           =   2535
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
         Left            =   7560
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Periodo"
         Height          =   375
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ubicacion"
         Height          =   375
         Left            =   3480
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Seccion"
         Height          =   375
         Left            =   3480
         TabIndex        =   10
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Division"
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
         Height          =   375
         Left            =   0
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consulta"
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   14775
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   7695
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   13573
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   17
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "Id"
            Caption         =   "Id"
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
         BeginProperty Column01 
            DataField       =   "Local"
            Caption         =   "Local"
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
            DataField       =   "Division"
            Caption         =   "Division"
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
         BeginProperty Column03 
            DataField       =   "Seccion"
            Caption         =   "Seccion"
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
            DataField       =   "Ubicacion"
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
            DataField       =   "Producto"
            Caption         =   "Producto"
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
            DataField       =   "Descripcio"
            Caption         =   "Descripcio"
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
            DataField       =   "Serie"
            Caption         =   "Serie"
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
            DataField       =   "Situacion"
            Caption         =   "Situacion"
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
         BeginProperty Column09 
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
         BeginProperty Column10 
            DataField       =   "Reconteo"
            Caption         =   "Reconteo"
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
         BeginProperty Column11 
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
         BeginProperty Column12 
            DataField       =   "Observa1"
            Caption         =   "Observa1"
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
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   555.024
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   3465.071
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   2534.74
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   5070.047
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   5385.26
            EndProperty
         EndProperty
      End
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
Attribute VB_Name = "tactfijo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim txactivo As New ADODB.Recordset
Private Sub ajdu1_Click()
If Frame2.Visible = True Then
   dbGrid1.SetFocus
   Exit Sub
End If
If periodo = "" Then
   MsgBox "No existe periodo ", 48, "Aviso"
   Exit Sub
End If
If Local1 = "" Then
    MsgBox "No existe local ", 48, "Aviso"
   Exit Sub
End If
If division = "" Then
    MsgBox "No existe division ", 48, "Aviso"
   Exit Sub
End If
If Seccion = "" Then
    MsgBox "No existe seccion ", 48, "Aviso"
   Exit Sub
End If
If ubicacion = "" Then
    MsgBox "No existe ubicacion ", 48, "Aviso"
   Exit Sub
End If

inicializa
Frame2.Visible = True
Frame2.Caption = "Nuevo"
Command2.Enabled = True
habilita 1
nid = ""
nproducto.SetFocus
End Sub

Private Sub bo712_Click()
Dim buf As String
On Error GoTo cmd656_err
buf = "" & txactivo.Fields("id")
If Frame2.Visible = True Then
   dbGrid1.SetFocus
   Exit Sub
End If
If MsgBox("Desea Borra " + "" & txactivo.Fields("situacion"), 1, "Aviso") <> 1 Then
   Exit Sub
End If
txactivo.Delete
Command1_Click



Exit Sub
cmd656_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub

End Sub



Private Sub cmdAddEntry_Click()
End Sub

Private Sub bsituacion_Click()
If bsituacion <> "" Then
   nsituacion = extra_loquesea1("" & bsituacion)
End If
End Sub

Private Sub cmdCerrar_Click()
dlo132_Click
End Sub

Private Sub cmdDelete_Click()
End Sub

Private Sub cmdExit_Click()

End Sub

Private Sub cmdHelp_Click()

End Sub

Private Sub cmdPrint_Click()
End Sub

Private Sub cmdSave_Click()
End Sub


Private Sub cmdGuardar_Click()
If Val(etiquetas) <= 0 Then
   MsgBox "Ingrese numero ", 48, "Aviso"
End If
ImprimirBarcodes
End Sub

Private Sub Command2_Click()
Dim found As Integer
found = grabar()
End Sub

Private Sub DBGrid1_DblClick()
Dim buf As String
On Error GoTo cmd90222_err
buf = "" & txactivo.Fields("producto")
tproduct.codigo = "" & txactivo.Fields("producto")
tproduct.ordename = "VER"
tproduct.Show 1
txactivo.Find "producto='" & buf & "'"
Exit Sub
cmd90222_err:
MsgBox "Aviso, Seleccione un dato ", 48, "Aviso"
Exit Sub

End Sub

Private Sub dk9893_Click()
If Frame2.Visible = True Then Exit Sub
reporgen.NAMETABLA = "activofijo"
reporgen.Show 1

End Sub
Sub prueba_reporte()
'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\situacionesproducto.rpt", "")
End Sub

Private Sub nconteo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
nconteo1.SetFocus

End Sub

Private Sub nconteo1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
nobserva.SetFocus

End Sub

Private Sub ndescripcio_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
nsituacion.SetFocus

End Sub

Private Sub nproducto_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
found = busca_producto()
If found = 0 Then
   MsgBox "No existe producto ", 48, "Aviso"
   nproducto.SetFocus
   Exit Sub
End If
ndescripcio.SetFocus
End Sub

Private Sub nproducto_KeyUp(KeyCode As Integer, Shift As Integer)
If Frame2.Caption = "Nuevo" Then
If KeyCode = &H76 Then  'f7
   xprodet.Show 1
End If
End If

End Sub

Private Sub nserie_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
nconteo.SetFocus

End Sub

Private Sub nsituacion_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
nserie.SetFocus

End Sub



Private Sub Command1_Click()
Frame1.Visible = True
Frame1.Enabled = True
opcion1 = "1"
ejecuta 1

End Sub
Sub ejecuta(sw As Integer)
Dim cad As String
Dim buffer As String
If opcion1 = "1" Then  'bodega
   cad = "SELECT * from activofijo    "
   cad = cad & " where local='" & Trim(extra_loquesea1(Local1)) & "'"
   cad = cad & " and periodo='" & Trim(extra_loquesea1(periodo)) & "'"
   cad = cad & " and seccion='" & Trim(extra_loquesea1(Seccion)) & "'"
   cad = cad & " and ubicacion='" & Trim(extra_loquesea1(ubicacion)) & "'"
   cad = cad & " and division='" & Trim(extra_loquesea1(division)) & "'"
   If txactivo.State = 1 Then txactivo.Close
   txactivo.Open cad, cn, adOpenStatic, adLockOptimistic
   Set dbGrid1.DataSource = txactivo
   'dbGrid1.columns(0).Width = 4000
   'dbGrid1.columns(1).Width = 2000
   If txactivo.RecordCount > 0 Then
     'dbGrid1.SetFocus
  End If
End If
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   'situacion = dbGrid1.Columns(1)
   'Frame1.Visible = False
   'Frame1.Enabled = False
   'situacion.SetFocus
   'situacion_KeyPress 13
End If
End Sub

Private Sub dlo132_Click()
If Frame2.Visible = True Then
   habilita 0
   Frame2.Visible = False
   dbGrid1.Enabled = True
   Exit Sub
End If
tactfijo.Hide
Unload tactfijo
End Sub


Private Sub f8443_Click()
Dim buf As String
On Error GoTo cmd456_err
buf = "" & txactivo.Fields("id")
If Frame2.Visible = True Then
   dbGrid1.SetFocus
   Exit Sub
End If

inicializa
Frame2.Visible = True
Frame2.Caption = "Modifica"
Command2.Enabled = True
pone_registro
habilita 1
'MsgBox "abc"
nproducto.SetFocus
Exit Sub
cmd456_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub

End Sub

Private Sub fjh433_Click()
Dim buf As String
On Error GoTo cmd556_err
buf = "" & txactivo.Fields("situacion")
If Frame2.Visible = True Then
   dbGrid1.SetFocus
   Exit Sub
End If

inicializa
Frame2.Visible = True
Frame2.Caption = "Zoom"
Command2.Enabled = False
pone_registro
habilita 1
nproducto.SetFocus
Exit Sub
cmd556_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub

End Sub

Private Sub Form_Activate()
'agregar_menus
Command1_Click
End Sub
Private Sub ImprimirBarcodes()
Dim mytablex As New ADODB.Recordset
Dim i As Integer
Dim Xpos As Double
Dim Ypos As Double
Dim HSpc As Double
Dim VSpc As Double
Dim altohoja As Double
Dim anchohoja As Double
On Error GoTo cmd999_err
mytablex.Open "select * from ETIQUETA", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
Xpos = mytablex!posx
Ypos = mytablex!posy
HSpc = mytablex!espacioh
VSpc = mytablex!espaciov
altohoja = mytablex!altohoja
anchohoja = mytablex!anchohoja
End If
mytablex.Close

Picture2.ScaleMode = vbCentimeters
Printer.Copies = 1
Printer.ScaleMode = vbCentimeters
Printer.CurrentX = Xpos
Printer.CurrentY = Ypos


Picture2.Cls
Picture2.ScaleMode = 3
Picture2.Height = Picture2.Height * (2.4 * 40 / Picture2.ScaleHeight)
Picture2.FontSize = 8
Call DrawBarcodeac("" & nproducto, Picture2, "" & ndescripcio & " - " & "" & nserie)

For i = 1 To Val(etiquetas)
If Printer.CurrentX <= anchohoja Then
Printer.PaintPicture Picture2, Printer.CurrentX, Printer.CurrentY
Printer.CurrentX = Printer.CurrentX + HSpc
Else

    Printer.CurrentX = Xpos
    Printer.CurrentY = Printer.CurrentY + VSpc
    Printer.PaintPicture Picture2, Printer.CurrentX, Printer.CurrentY
    Printer.CurrentX = Printer.CurrentX + HSpc


    If Printer.CurrentY >= altohoja And Printer.CurrentX >= anchohoja Then
    Printer.NewPage
    Printer.CurrentX = Xpos
    Printer.CurrentY = Ypos
    End If
End If
If Printer.CurrentY >= altohoja And Printer.CurrentX >= anchohoja Then
Printer.NewPage
Printer.CurrentX = Xpos
Printer.CurrentY = Ypos
End If
Next i
Printer.EndDoc
Exit Sub
cmd999_err:
MsgBox "Aviso en imprimir " + error$, 48, "Aviso"
Exit Sub
End Sub
Public Sub DrawBarcodeac(ByVal bc_string As String, OBJ As Control, strDescripcion As String)
    
    Dim Xpos!, Y1!, Y2!, dw%, th!, tw, new_string$
    Dim n As Integer
    Dim c As Integer
    Dim bc_pattern$
    Dim i As Integer
    'define barcode patterns
    Dim bc(90) As String
    bc(1) = "1 1221"            'pre-amble
    bc(2) = "1 1221"            'post-amble
    bc(48) = "11 221"           'digits
    bc(49) = "21 112"
    bc(50) = "12 112"
    bc(51) = "22 111"
    bc(52) = "11 212"
    bc(53) = "21 211"
    bc(54) = "12 211"
    bc(55) = "11 122"
    bc(56) = "21 121"
    bc(57) = "12 121"
                                'capital letters
    bc(65) = "211 12"           'A
    bc(66) = "121 12"           'B
    bc(67) = "221 11"           'C
    bc(68) = "112 12"           'D
    bc(69) = "212 11"           'E
    bc(70) = "122 11"           'F
    bc(71) = "111 22"           'G
    bc(72) = "211 21"           'H
    bc(73) = "121 21"           'I
    bc(74) = "112 21"           'J
    bc(75) = "2111 2"           'K
    bc(76) = "1211 2"           'L
    bc(77) = "2211 1"           'M
    bc(78) = "1121 2"           'N
    bc(79) = "2121 1"           'O
    bc(80) = "1221 1"           'P
    bc(81) = "1112 2"           'Q
    bc(82) = "2112 1"           'R
    bc(83) = "1212 1"           'S
    bc(84) = "1122 1"           'T
    bc(85) = "2 1112"           'U
    bc(86) = "1 2112"           'V
    bc(87) = "2 2111"           'W
    bc(88) = "1 1212"           'X
    bc(89) = "2 1211"           'Y
    bc(90) = "1 2211"           'Z
                                'Misc
    bc(32) = "1 2121"           'space
    bc(35) = ""                 '# cannot do!
    bc(36) = "1 1 1 11"         '$
    bc(37) = "11 1 1 1"         '%
    bc(43) = "1 11 1 1"         '+
    bc(45) = "1 1122"           '-
    bc(47) = "1 1 11 1"         '/
    bc(46) = "2 1121"           '.
    bc(64) = ""                 '@ cannot do!
    bc(65) = "1 1221"           '*
    
    
    
    bc_string = UCase(bc_string)
    
    strDescripcion = Left(strDescripcion, 25)
    
    'dimensions
    OBJ.ScaleMode = 3                               'pixels
    OBJ.Cls
    OBJ.Picture = Nothing
    dw = CInt(OBJ.ScaleHeight / 40)                 'space between bars
    If dw < 1 Then dw = 1
    'Debug.Print dw
    th = OBJ.TextHeight(bc_string & " " & strDescripcion)                  'text height
    tw = OBJ.textWidth(bc_string & " " & strDescripcion)                   'text width
    new_string = Chr$(1) & bc_string & Chr$(2)      'add pre-amble, post-amble
    
    Y1 = OBJ.ScaleTop
    Y2 = OBJ.ScaleTop + OBJ.ScaleHeight - 1.5 * th
    OBJ.Width = 1.1 * Len(new_string) * (15 * dw) * OBJ.Width / OBJ.ScaleWidth
    
    
    'draw each character in barcode string
    Xpos = OBJ.ScaleLeft
    For n = 1 To Len(new_string)
        c = Asc(Mid$(new_string, n, 1))
        If c > 90 Then c = 0
        bc_pattern$ = bc(c)
        
        'draw each bar
        For i = 1 To Len(bc_pattern$)
            Select Case Mid$(bc_pattern$, i, 1)
                Case " "
                    'space
                    OBJ.Line (Xpos, Y1)-(Xpos + 1 * dw, Y2), &HFFFFFF, BF
                    Xpos = Xpos + dw
                    
                Case "1"
                    'space
                    OBJ.Line (Xpos, Y1)-(Xpos + 1 * dw, Y2), &HFFFFFF, BF
                    Xpos = Xpos + dw
                    'line
                    OBJ.Line (Xpos, Y1)-(Xpos + 1 * dw, Y2), &H0&, BF
                    Xpos = Xpos + dw
                
                Case "2"
                    'space
                    OBJ.Line (Xpos, Y1)-(Xpos + 1 * dw, Y2), &HFFFFFF, BF
                    Xpos = Xpos + dw
                    'wide line
                    OBJ.Line (Xpos, Y1)-(Xpos + 2 * dw, Y2), &H0&, BF
                    Xpos = Xpos + 2 * dw
            End Select
        Next
    Next
    
    '1 more space
    OBJ.Line (Xpos, Y1)-(Xpos + 1 * dw, Y2), &HFFFFFF, BF
    Xpos = Xpos + dw
    OBJ.FontBold = False
    OBJ.Font.Size = 8
  
  
    'final size and text
    OBJ.Width = (Xpos + dw) * OBJ.Width / OBJ.ScaleWidth
    OBJ.CurrentX = (OBJ.ScaleWidth - tw) / 2
    OBJ.CurrentY = Y2 + 0.25 * th
    OBJ.Print bc_string & " " & strDescripcion
    
    
    'copy to clipboard
    OBJ.Picture = OBJ.Image
    Clipboard.Clear
    Clipboard.SetData OBJ.Image, 2



End Sub


Private Sub Form_Load()
carga_inicio

End Sub
Sub inicializa()
nproducto = ""
ndescripcio = ""
nsituacion = ""
nserie = ""
nconteo = ""
nconteo1 = ""
nobserva = ""
bsituacion.ListIndex = 0
End Sub
Sub pone_registro()
nsituacion = Trim("" & txactivo.Fields("situacion"))
ndescripcio = Trim("" & txactivo.Fields("descripcio"))
nproducto = Trim("" & txactivo.Fields("producto"))
nserie = Trim("" & txactivo.Fields("serie"))
nconteo = Trim("" & txactivo.Fields("conteo"))
nconteo1 = Trim("" & txactivo.Fields("conteo1"))
nobserva = Trim("" & txactivo.Fields("observa"))
End Sub
Sub grabando()
txactivo.Fields("periodo") = Trim(extra_loquesea1(periodo))
txactivo.Fields("local") = Trim(extra_loquesea1(Local1))
txactivo.Fields("division") = Trim(extra_loquesea1(division))
txactivo.Fields("seccion") = Trim(extra_loquesea1(Seccion))
txactivo.Fields("ubicacion") = Trim(extra_loquesea1(ubicacion))
txactivo.Fields("producto") = Trim(nproducto)
txactivo.Fields("descripcio") = Trim(ndescripcio)
txactivo.Fields("serie") = Trim(nserie)
txactivo.Fields("situacion") = Trim(nsituacion)
txactivo.Fields("conteo") = Val(nconteo)
txactivo.Fields("conteo1") = Val(nconteo1)
txactivo.Fields("observa") = Trim(nobserva)
End Sub

Private Sub grba1_Click()

End Sub

Function grabar()
Dim found As Integer
Dim rbuscap As New ADODB.Recordset
found = valida()
If found = 0 Then
   MsgBox "Campos invalidos", 48, "Aviso"
   Exit Function
End If
If Frame2.Caption = "Nuevo" Then
   rbuscap.Open "select * from activofijo where periodo='" & extra_loquesea1(periodo) & "' and producto='" & Trim(nproducto) & "'", cn, adOpenStatic, adLockOptimistic
   If rbuscap.RecordCount > 0 Then
      rbuscap.Close
      MsgBox "Ya existe Producto en el Periodo ", 48, "Aviso"
      Exit Function
   End If
   txactivo.AddNew
   grabando
   txactivo.Update
   dlo132_Click
   Exit Function
End If
If Frame2.Caption = "Modifica" Then
   grabando
   txactivo.Update
   dlo132_Click
   Exit Function
End If

End Function

Function valida()
If Len(nproducto) = 0 Then
   nproducto.SetFocus
   Exit Function
End If
If Len(ndescripcio) = 0 Then
   ndescripcio.SetFocus
   Exit Function
End If
If Len(nsituacion) = 0 Then
   nsituacion.SetFocus
   Exit Function
End If



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
Dim i As Integer
For i = 1 To mnuArchivoArray.count - 1
    Unload mnuArchivoArray(i)
Next
     
Dim mytablex As New ADODB.Recordset
   mytablex.Open "select * from archivo where menu='situacion' and   estado='S'", cn, adOpenStatic, adLockOptimistic
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
Sub mnuarchivoarray_click(Index As Integer)
Dim mytablex As New ADODB.Recordset
Dim buf As String
buf = mnuArchivoArray(Index).Caption
   mytablex.Open "select * from archivo where menu='situacion' and descripcio='" & Trim(buf) & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount = 0 Then
      mytablex.Close
   End If
   'busca el reporte
   buf = mytablex.Fields("archivo")
   mytablex.Close
   'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\" & buf, "")

End Sub
Sub carga_inicio()
Dim mytablex As New ADODB.Recordset
periodo.Clear
periodo.AddItem ""
mytablex.Open "select * from periodo ", cn, adOpenStatic, adLockOptimistic
Do
If mytablex.EOF Then Exit Do
periodo.AddItem Trim("" & mytablex.Fields("descripcio")) & "|" & Trim("" & mytablex.Fields("periodo"))
mytablex.MoveNext
Loop
mytablex.Close
periodo.ListIndex = 0


Local1.Clear
Local1.AddItem ""
mytablex.Open "select * from tlocal ", cn, adOpenStatic, adLockOptimistic
Do
If mytablex.EOF Then Exit Do
Local1.AddItem Trim("" & mytablex.Fields("nombre")) & "|" & Trim("" & mytablex.Fields("codigo"))
mytablex.MoveNext
Loop
mytablex.Close
Local1.ListIndex = 0

division.Clear
division.AddItem ""
mytablex.Open "select * from division ", cn, adOpenStatic, adLockOptimistic
Do
If mytablex.EOF Then Exit Do
division.AddItem Trim("" & mytablex.Fields("descripcio")) & "|" & Trim("" & mytablex.Fields("division"))
mytablex.MoveNext
Loop
mytablex.Close
division.ListIndex = 0


Seccion.Clear
Seccion.AddItem ""
mytablex.Open "select * from seccion ", cn, adOpenStatic, adLockOptimistic
Do
If mytablex.EOF Then Exit Do
Seccion.AddItem Trim("" & mytablex.Fields("descripcio")) & "|" & Trim("" & mytablex.Fields("seccion"))
mytablex.MoveNext
Loop
mytablex.Close
Seccion.ListIndex = 0

ubicacion.Clear
ubicacion.AddItem ""
mytablex.Open "select * from ubicacion ", cn, adOpenStatic, adLockOptimistic
Do
If mytablex.EOF Then Exit Do
ubicacion.AddItem Trim("" & mytablex.Fields("descripcio")) & "|" & Trim("" & mytablex.Fields("ubicacion"))
mytablex.MoveNext
Loop
mytablex.Close
ubicacion.ListIndex = 0

bsituacion.Clear
bsituacion.AddItem ""
mytablex.Open "select * from situacion ", cn, adOpenStatic, adLockOptimistic
Do
If mytablex.EOF Then Exit Do
bsituacion.AddItem Trim("" & mytablex.Fields("descripcio")) & "|" & Trim("" & mytablex.Fields("situacion"))
mytablex.MoveNext
Loop
mytablex.Close
bsituacion.ListIndex = 0
End Sub
Function busca_producto()
Dim mytablex As New ADODB.Recordset
mytablex.Open "select * from producto where producto='" & nproducto & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   ndescripcio = "" & mytablex.Fields("descripcio")
   'nsituacion = "" & mytablex.Fields("situacion")
   busca_producto = 1
End If
mytablex.Close
End Function

