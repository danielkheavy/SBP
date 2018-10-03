VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tcentrop 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Centro de Produccion"
   ClientHeight    =   9435
   ClientLeft      =   165
   ClientTop       =   -45
   ClientWidth     =   13290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   13290
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Consulta"
      Height          =   8625
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   12450
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Filtrar"
         Height          =   495
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   240
         Width           =   2175
      End
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
         Left            =   195
         Style           =   2  'Dropdown List
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   315
         Width           =   2295
      End
      Begin VB.TextBox Text1 
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
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin MSDataGridLib.DataGrid dbgrid13 
         Height          =   7410
         Left            =   120
         TabIndex        =   32
         Top             =   840
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   13070
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   29
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   6615
      Left            =   15
      TabIndex        =   11
      Top             =   735
      Visible         =   0   'False
      Width           =   12495
      Begin VB.TextBox subtablaproo 
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
         MaxLength       =   6
         TabIndex        =   26
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox volumendia 
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
         MaxLength       =   10
         TabIndex        =   24
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox tarifadia 
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
         MaxLength       =   10
         TabIndex        =   22
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox moneda 
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
         MaxLength       =   1
         TabIndex        =   20
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox cantidad 
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
         MaxLength       =   10
         TabIndex        =   18
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox tablaproo 
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
         MaxLength       =   6
         TabIndex        =   16
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox id 
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
         MaxLength       =   6
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   8400
         Picture         =   "tcentrop.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Imprimir todo"
         Top             =   3120
         Width           =   1470
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   8400
         Picture         =   "tcentrop.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2160
         Width           =   1470
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Subgrupo"
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
         TabIndex        =   27
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Volumen/Dia"
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
         TabIndex        =   25
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tarifa/Dia"
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
         TabIndex        =   23
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Moneda"
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
         TabIndex        =   21
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad"
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
         TabIndex        =   19
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Grupo"
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
         TabIndex        =   17
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Id"
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
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
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
         Height          =   615
         Left            =   720
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tcentrop.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Grabar registro"
         Top             =   0
         Width           =   735
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
         Height          =   375
         Left            =   7560
         MaxLength       =   10
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   120
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
         Left            =   5280
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
         Height          =   615
         Left            =   1440
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tcentrop.frx":23A6
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Borrar registro"
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
         Left            =   2760
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tcentrop.frx":35B8
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Left            =   2160
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tcentrop.frx":47CA
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Imprimir"
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
         Picture         =   "tcentrop.frx":59DC
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Consulta"
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   12495
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   7455
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   13150
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "Id"
            Caption         =   "Id"
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
            DataField       =   "Centroproduccion"
            Caption         =   "CentroProduccion"
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
            DataField       =   "Tablapro"
            Caption         =   "Grupo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13322
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Subtablapro"
            Caption         =   "Subgrupo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13322
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Cantidad"
            Caption         =   "Cantidad"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13322
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "moneda"
            Caption         =   "Moneda"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13322
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
         EndProperty
      End
      Begin VB.Label ncentroproduccion 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1440
         TabIndex        =   34
         Top             =   240
         Width           =   6255
      End
      Begin VB.Label centroproduccion 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1335
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
Attribute VB_Name = "tcentrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim txcentrox As New ADODB.Recordset

Private Sub ajdu1_Click()

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    inicializa
    Frame2.Visible = True
    Frame2.Caption = "Nuevo"
    cmdGuardar.Enabled = True
    habilita 1
    id.Enabled = True
    id = ""
    tablaproo.SetFocus

End Sub

Private Sub bo712_Click()

    Dim buf As String

    On Error GoTo cmd656_err

    buf = "" & txcentrox.Fields("centroproduccion")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    If MsgBox("Desea Borra " + txcentrox.Fields("centroproduccion"), 1, "Aviso") <> 1 Then
        Exit Sub

    End If

    cn.Execute ("delete from grupocentroproduccion where centroproduccion='" & txcentrox.Fields("centroproduccion") & "'")
    txcentrox.Delete
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

Private Sub dbgrid13_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim found As Integer

    If KeyCode = 27 Then
        Text1.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = "1" Then
            tablaproo = Trim("" & dbgrid13.columns(1))
            Frame3.Visible = False
            tablaproo.SetFocus

        End If

        If opcion1 = "2" Then
            subtablaproo = Trim("" & dbgrid13.columns(1))
            Frame3.Visible = False
            subtablaproo.SetFocus

        End If

    End If

End Sub

Private Sub dk9893_Click()

    If Frame2.Visible = True Then Exit Sub
    reporgen.NAMETABLA = "centroproduccion"
    reporgen.Show 1

End Sub

Sub prueba_reporte()

    'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\centroproduccionesproducto.rpt", "")
End Sub

Private Sub Command1_Click()
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    opcion1 = "1"
    ejecuta 1

End Sub

Sub ejecuta(sw As Integer)

    Dim cad As String

    If opcion1 = "1" Then  'bodega
        If Len(buffer) = 0 Then
            cad = "SELECT * from GRUPOcentroproduccion   where centroproduccion='" & centroproduccion & "'  order by tablapro,subtablapro"

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT *  from grupocentroproduccion   where centroproduccion='" & centroproduccion & "' and  " & Combo1 & " like '" & buffer & "%' order by tablapro,subtablapro"

        End If

        If txcentrox.State = 1 Then txcentrox.Close
        txcentrox.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbGrid1.DataSource = txcentrox

        'dbGrid1.columns(0).Width = 4000
        'dbGrid1.columns(1).Width = 2000
        If txcentrox.RecordCount > 0 Then
            dbGrid1.SetFocus

        End If

    End If

End Sub

Private Sub Command2_Click()
    filtro

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then

        'centroproduccion = dbGrid1.Columns(1)
        'Frame1.Visible = False
        'Frame1.Enabled = False
        'centroproduccion.SetFocus
        'centroproduccion_KeyPress 13
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

    If Frame3.Visible = True Then
        Frame3.Visible = False
        Exit Sub

    End If

    If Frame2.Visible = True Then
        habilita 0
        Frame2.Visible = False
        dbGrid1.Enabled = True
        Exit Sub

    End If

    tcentrop.Hide
    Unload tcentrop

End Sub

Private Sub f8443_Click()

    Dim buf As String

    On Error GoTo cmd456_err

    buf = "" & txcentrox.Fields("id")

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
    id.Enabled = False
    tablaproo.SetFocus
    Exit Sub
cmd456_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fjh433_Click()

    Dim buf As String

    On Error GoTo cmd556_err

    buf = "" & txcentrox.Fields("id")

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
    id.Enabled = False
    tablaproo.SetFocus
    Exit Sub
cmd556_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Form_Activate()
    'agregar_menus
    Command1_Click

End Sub

Private Sub Form_Load()
    Combo1.Clear
    'Combo1.AddItem "Descripcio"
    Combo1.AddItem "centroproduccion"
    Combo1.ListIndex = 0

End Sub

Sub inicializa()
    id = ""
    cantidad = ""
    moneda = "S"
    tarifadia = ""
    volumendia = ""
    'descripcio = ""
    tablaproo = ""
    subtablaproo = ""

End Sub

Sub pone_registro()
    id = Trim("" & txcentrox.Fields("id"))
    'descripcio = Trim("" & txcentrox.Fields("descripcio"))
    tablaproo = Trim("" & txcentrox.Fields("tablapro"))
    subtablaproo = Trim("" & txcentrox.Fields("subtablapro"))
    moneda = Trim("" & txcentrox.Fields("moneda"))
    volumendia = Trim("" & txcentrox.Fields("volumendia"))
    cantidad = Trim("" & txcentrox.Fields("cantidad"))
    tarifadia = Trim("" & txcentrox.Fields("tarifadia"))

End Sub

Sub grabando()
    txcentrox.Fields("centroproduccion") = Trim(centroproduccion)
    'txcentrox.Fields("descripcio") = Trim(descripcio)
    txcentrox.Fields("tablapro") = Trim(tablaproo)
    txcentrox.Fields("subtablapro") = Trim(subtablaproo)
    txcentrox.Fields("moneda") = Trim(moneda)
    txcentrox.Fields("volumendia") = Val(volumendia)
    txcentrox.Fields("cantidad") = Val(cantidad)
    txcentrox.Fields("tarifadia") = Val(tarifadia)

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
        'If Len(centroproduccion) = 0 Then
        '   centroproduccion.SetFocus
        '   Exit Function
        'End If
        'rbusca.Open "select centroproduccion from centroproduccion where centroproduccion='" & centroproduccion & "'", cn, adOpenStatic, adLockOptimistic
        'If rbusca.RecordCount > 0 Then
        '   rbusca.Close
        '   MsgBox "Ya existe centroproduccion ", 48, "Aviso"
        '   Exit Function
        'End If
        txcentrox.AddNew
        'txcentrox.Fields("centroproduccion") = centroproduccion
        grabando
        txcentrox.Update
        dlo132_Click
        Exit Function

    End If

    If Frame2.Caption = "Modifica" Then
        'txcentrox.Fields("centroproduccion") = centroproduccion
        grabando
        txcentrox.Update
        dlo132_Click
        Exit Function

    End If

End Function

Function valida()
    'If Len(centroproduccion) = 0 Then
    '   centroproduccion.SetFocus
    '   Exit Function
    'End If

    'If activo <> "S" And activo <> "N" Then
    '   activo.SetFocus
    '   Exit Function
    'End If
    'If Len(descripcio) = 0 Then
    '   descripcio.SetFocus
    '   Exit Function
    'End If
    If Len(Trim(tablaproo)) = 0 Then
        tablaproo.SetFocus
        Exit Function

    End If

    If Len(Trim(subtablaproo)) = 0 Then
        subtablaproo.SetFocus
        Exit Function

    End If

    If moneda <> "S" And moneda <> "D" Then
        moneda.SetFocus
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

    Dim I As Integer

    For I = 1 To mnuArchivoArray.count - 1
        Unload mnuArchivoArray(I)
    Next
     
    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from archivo where menu='centroproduccion' and   estado='S'", cn, adOpenStatic, adLockOptimistic

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

    Dim buf      As String

    buf = mnuArchivoArray(Index).Caption
    mytablex.Open "select * from archivo where menu='centroproduccion' and descripcio='" & Trim(buf) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close

    End If

    'busca el reporte
    buf = mytablex.Fields("archivo")
    mytablex.Close
    'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\" & buf, "")

End Sub

Private Sub subtablaproo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        If Len(Trim(tablaproo)) = 0 Then
            MsgBox "Debe Existir Grupo", 48, "Aviso"
            Exit Sub

        End If

        consulta_subgrupo

    End If

End Sub

Private Sub tablaproo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_grupo

    End If

End Sub

Sub consulta_grupo()
    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.AddItem "tablapro"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "1"
    Text1.SetFocus
    Command2_Click

End Sub

Sub consulta_subgrupo()
    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.AddItem "subtablapro"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "2"
    Text1.SetFocus
    Command2_Click

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    Command2_Click

End Sub

Sub filtro()

    Dim mytablex As New ADODB.Recordset

    Dim cad      As String

    If opcion1 = "1" Then  'producto
        If Len(Text1) = 0 Then
            cad = "select Descripcio,Tablapro from Tablapro "

        End If

        If Len(Text1) > 0 Then
            cad = "select Descripcio,Tablapro from tablapro where " & Combo1 & " like '%" & buffer & "%'"

        End If

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbgrid13.DataSource = mytablex
        dbgrid13.columns(0).Width = 5000
        dbgrid13.columns(1).Width = 2000
        'dbgrid13.columns(2).Width = 1000
        'dbgrid13.columns(3).Width = 1000
           
    End If

    If opcion1 = "2" Then  'producto
        If Len(Text1) = 0 Then
            cad = "select Descripcio,subTablapro,Tablapro from subTablapro where tablapro='" & tablaproo & "'"

        End If

        If Len(Text1) > 0 Then
            cad = "select Descripcio,SubTablapro,Tablapro from subtablapro where tablapro='" & tablaproo & "' and " & Combo1 & " like '%" & buffer & "%'"

        End If

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbgrid13.DataSource = mytablex
        dbgrid13.columns(0).Width = 5000
        dbgrid13.columns(1).Width = 2000
        'dbgrid13.columns(2).Width = 1000
        'dbgrid13.columns(3).Width = 1000
           
    End If
   
    If mytablex.RecordCount > 0 Then
        dbgrid13.SetFocus

    End If

    Exit Sub

End Sub

