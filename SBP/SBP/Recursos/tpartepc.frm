VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tpartepc 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parte Produccion"
   ClientHeight    =   10650
   ClientLeft      =   90
   ClientTop       =   -120
   ClientWidth     =   19020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10650
   ScaleWidth      =   19020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consulta"
      Height          =   9975
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   15015
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Filtrar"
         Height          =   495
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   25
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
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
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin MSDataGridLib.DataGrid dbgrid13 
         Height          =   8895
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   15690
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
      BackColor       =   &H00FFFFC0&
      Caption         =   "Frame2"
      Height          =   8775
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   15015
      Begin VB.TextBox ordentrabajo 
         Enabled         =   0   'False
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
         MaxLength       =   11
         TabIndex        =   13
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox tipomov 
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
         TabIndex        =   29
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox observa 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   2280
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   2400
         Width           =   5895
      End
      Begin VB.TextBox numero 
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
         TabIndex        =   16
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox fecha 
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
         TabIndex        =   15
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   12480
         Picture         =   "tpartepc.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Imprimir todo"
         Top             =   1320
         Width           =   1470
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   12480
         Picture         =   "tpartepc.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1470
      End
      Begin VB.TextBox bodega 
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
         TabIndex        =   11
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "(S)Entrada"
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
         Left            =   2880
         TabIndex        =   36
         Top             =   2040
         Width           =   3735
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nro Orden Trabajo"
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
         TabIndex        =   31
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoMovimiento"
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
         TabIndex        =   30
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Almacen"
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
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero"
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
         TabIndex        =   20
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha"
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
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observa"
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
         TabIndex        =   18
         Top             =   2400
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13320
      TabIndex        =   9
      Top             =   1320
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFC0&
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   18960
      TabIndex        =   2
      Top             =   0
      Width           =   19020
      Begin VB.ComboBox ordenado 
         Height          =   315
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox fechaf 
         Height          =   375
         Left            =   9120
         MaxLength       =   10
         TabIndex        =   35
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox fechai 
         Height          =   375
         Left            =   9120
         MaxLength       =   10
         TabIndex        =   33
         Top             =   0
         Width           =   1455
      End
      Begin VB.ComboBox almacen 
         Height          =   315
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   0
         Width           =   2415
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
         Left            =   720
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tpartepc.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Grabar registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000D&
         Caption         =   "&Filtrar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11520
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
         Picture         =   "tpartepc.frx":23A6
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
         Picture         =   "tpartepc.frx":35B8
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
         Picture         =   "tpartepc.frx":47CA
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
         Picture         =   "tpartepc.frx":59DC
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label vieneorden 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   10680
         TabIndex        =   39
         Top             =   360
         Width           =   105
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ordenado"
         Height          =   375
         Left            =   4080
         TabIndex        =   38
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaFinal"
         Height          =   375
         Left            =   7920
         TabIndex        =   34
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaInicio"
         Height          =   375
         Left            =   7920
         TabIndex        =   32
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Almacen"
         Height          =   375
         Left            =   4080
         TabIndex        =   28
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consulta"
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   13215
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   7935
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   13996
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
            DataField       =   "fecha"
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
            DataField       =   "Tipomov"
            Caption         =   "TipoMov"
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
            DataField       =   "ordentrabajo"
            Caption         =   "OrdenTrabajo"
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
            DataField       =   "Items"
            Caption         =   "Items"
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
            DataField       =   "Bodega"
            Caption         =   "Almacen"
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
         BeginProperty Column06 
            DataField       =   "Observa"
            Caption         =   "Observa"
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
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1635.024
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   7049.764
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
Attribute VB_Name = "tpartepc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim txformxupc As New ADODB.Recordset

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
    Numero.Enabled = False
    Numero = ""
    fecha.SetFocus

End Sub

Private Sub bo712_Click()

    Dim buf As String

    On Error GoTo cmd656_err

    buf = "" & txformxupc.Fields("numero")

    If borrar_notaingreso() = 1 Then
        MsgBox "Existen productos en Detalle Nota de Ingreso ", 48, "Aviso"
        Exit Sub

    End If

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    'If MsgBox("Desea Borra " + "" & txformxupc.Fields("numero"), 1, "Aviso") <> 1 Then
    '   Exit Sub
    'End If
    'borrar_notaingreso
    txformxupc.Delete
    Command1_Click

    Exit Sub
cmd656_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub bodega_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_bodega

    End If

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

Private Sub Command2_Click()

    On Error GoTo cmfd5611_err

    tpartepd.fecha = "" & txformxupc.Fields("fecha")
    tpartepd.idx = "" & txformxupc.Fields("numero")
    tpartepd.bodega = "" & txformxupc.Fields("bodega")
    tpartepd.ordentrabajo = "" & txformxupc.Fields("ordentrabajo")
    tpartepd.tipomov = "" & txformxupc.Fields("tipomov")
    tpartepd.Show 1
    sumar
    Command1_Click
    Exit Sub
cmfd5611_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()
    filtro

End Sub

Private Sub Command5_Click()

End Sub

Private Sub dbgrid13_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim found As Integer

    If KeyCode = 27 Then
        Text1.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = "1" Then
            bodega = Trim("" & dbgrid13.columns(1))
            Frame3.Visible = False
            bodega.SetFocus

        End If

        If opcion1 = "2" Then
            ordentrabajo = Trim("" & dbgrid13.columns(0))
            Frame3.Visible = False
            ordentrabajo.SetFocus

        End If

    End If

End Sub

Private Sub dk9893_Click()

    If Frame2.Visible = True Then Exit Sub
    reporgen.NAMETABLA = "parteproduccionc"
    reporgen.Show 1

End Sub

Sub prueba_reporte()

    'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\formulacionesproducto.rpt", "")
End Sub

Private Sub Command1_Click()
    'If almacen = "%" Then
    '   MsgBox "Seleccione Tipo Formula ", 48, "Aviso"
    '   Exit Sub
    'Exit Sub
    Frame1.Visible = True
    Frame1.Enabled = True

    'buffer = ""
    opcion1 = "1"
    ejecuta 1

End Sub

Sub ejecuta(sw As Integer)

    Dim cad As String

    If Not IsDate(fechai) Then
        Exit Sub

    End If

    If Not IsDate(fechaf) Then
        Exit Sub

    End If

    cad = "SELECT * from parteproduccionc where ordentrabajo=" & Val(vieneorden)
    cad = cad & " and  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    cad = cad & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
   
    If almacen <> "%" Then
        cad = cad & " and bodega='" & extra_loquesea1(almacen) & "' "

    End If

    If ordenado <> "%" Then
        cad = cad & " order by " & ordenado

    End If

    If txformxupc.State = 1 Then txformxupc.Close
    txformxupc.Open cad, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = txformxupc
    dbGrid1.SetFocus
 
End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then

        'buffer.SetFocus
        'Exit Sub
    End If

    If KeyCode = 13 Then

        'formulacion = dbGrid1.Columns(1)
        'Frame1.Visible = False
        'Frame1.Enabled = False
        'formulacion.SetFocus
        'formulacion_KeyPress 13
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

    tpartepc.Hide
    Unload tpartepc

End Sub

Private Sub f8443_Click()

    Dim buf As String

    On Error GoTo cmd456_err

    buf = "" & txformxupc.Fields("numero")

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
    'ordentrabajo.Enabled = False
    Numero.Enabled = False
    fecha.SetFocus
    Exit Sub
cmd456_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fjh433_Click()

    Dim buf As String

    On Error GoTo cmd556_err

    buf = "" & txformxupc.Fields("numero")

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
    Numero.Enabled = False
    'MsgBox "ABC"
    fecha.SetFocus
    Exit Sub
cmd556_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Form_Activate()
    Command1_Click

End Sub

Private Sub Form_Load()

    fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    fechaf = Format(Now, "dd/mm/yyyy")
    ordenado.Clear
    ordenado.AddItem "%"
    ordenado.AddItem "Numero"
    ordenado.AddItem "Fecha"
    ordenado.AddItem "ordentrabajo"
    ordenado.ListIndex = 0
    carga_almacen

End Sub

Sub inicializa()
    fecha = Format(Now, "dd/mm/yyyy")
    bodega = ""
    tipomov = "S"
    ordentrabajo = "" & vieneorden
    observa = ""

End Sub

Sub pone_registro()
    bodega = Trim("" & txformxupc.Fields("bodega"))
    tipomov = Trim("" & txformxupc.Fields("tipomov"))
    ordentrabajo = Trim("" & txformxupc.Fields("ordentrabajo"))
    fecha = Trim("" & txformxupc.Fields("fecha"))
    observa = Trim("" & txformxupc.Fields("observa"))
    Numero = Trim("" & txformxupc.Fields("numero"))

End Sub

Sub grabando()
    txformxupc.Fields("fecha") = Trim(fecha)
    txformxupc.Fields("bodega") = Trim(bodega)
    txformxupc.Fields("tipomov") = Trim(tipomov)
    txformxupc.Fields("ordentrabajo") = Val(ordentrabajo)
    txformxupc.Fields("observa") = Trim(observa)

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
        txformxupc.AddNew
        grabando
        txformxupc.Update
        dlo132_Click
        Exit Function

    End If

    If Frame2.Caption = "Modifica" Then
        grabando
        txformxupc.Update
        cn.Execute ("update parteproducciond set ordentrabajo=" & Val("" & txformxupc.Fields("ORDENTRABAJO")) & ",bodega='" & Trim("" & txformxupc.Fields("bodega")) & "',tipomov='" & txformxupc.Fields("tipomov") & "' where numero=" & txformxupc.Fields("numero"))
        dlo132_Click
        Exit Function

    End If

End Function

Function valida()

    If Not IsDate(fecha) Then
        fecha.SetFocus
        Exit Function

    End If

    If Len(Trim(bodega)) = 0 Then
        bodega.SetFocus
        Exit Function

    End If

    If Len(Trim(ordentrabajo)) = 0 Then
        'ordentrabajo.SetFocus
        Exit Function

    End If

    If Len(Trim(tipomov)) = 0 Then
        tipomov.SetFocus
        Exit Function

    End If

    If tipomov <> "S" Then
        tipomov.SetFocus
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

    mytablex.Open "select * from archivo where menu='formulacion' and   estado='S'", cn, adOpenStatic, adLockOptimistic

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
    mytablex.Open "select * from archivo where menu='formulacion' and descripcio='" & Trim(buf) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close

    End If

    'busca el reporte
    buf = mytablex.Fields("archivo")
    mytablex.Close
    'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\" & buf, "")

End Sub

Sub consulta_bodega()
    Combo2.Clear
    Combo2.AddItem "Nombre"
    Combo2.AddItem "Codigo"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "1"
    Text1.SetFocus
    Command4_Click

End Sub

Sub consulta_ordentrabajo()
    Combo2.Clear
    Combo2.AddItem "Ordentrabajo"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "2"
    Text1.SetFocus
    Command4_Click

End Sub

Private Sub ordentrabajo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_ordentrabajo

    End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    Command4_Click

End Sub

Sub filtro()

    Dim mytablex As New ADODB.Recordset

    Dim cad      As String

    If opcion1 = "1" Then  'producto
        If Len(Text1) = 0 Then
            cad = "select Nombre,Codigo from bodega "

        End If

        If Len(Text1) > 0 Then
            cad = "select Nombre,Codigo from bodega where " & Combo2 & " like '" & Text1 & "%'"

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
            cad = "select Ordentrabajo,Fecha,FechaI,fechae,estado,Aprobado from ordentrabajoc where aprobado='S' "

        End If

        If Len(Text1) > 0 Then
            cad = "select Ordentrabajo,Fecha,FechaI,fechae,estado,Aprobado from ordentrabajoc  where aprobado='S' and " & Combo2 & "=" & Text1 & ""

        End If

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbgrid13.DataSource = mytablex
        'dbgrid13.columns(0).Width = 5000
        'dbgrid13.columns(1).Width = 2000
        'dbgrid13.columns(2).Width = 1000
        'dbgrid13.columns(3).Width = 1000
           
    End If
   
    If mytablex.RecordCount > 0 Then
        dbgrid13.SetFocus

    End If

    Exit Sub

End Sub

Sub carga_almacen()

    Dim mytablex As New ADODB.Recordset

    almacen.Clear
    almacen.AddItem "%"
    mytablex.Open "select * from bodega", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        almacen.AddItem Trim("" & mytablex.Fields("nombre")) & "|" & Trim("" & mytablex.Fields("codigo"))
        mytablex.MoveNext
    Loop
    mytablex.Close
    almacen.ListIndex = 0

End Sub

Sub sumar()

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Double

    Dim sdx1     As Double

    Dim sdx2     As Double

    On Error GoTo cmd86712_err

    sdx = 0
    sdx1 = 0
    sdx2 = 0
    mytablex.Open "select * from parteproducciond where numero=" & txformxupc("numero"), cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        sdx = sdx + 1
        sdx1 = sdx1 + Val("" & mytablex.Fields("cantidad"))
        sdx2 = sdx2 + Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("precio"))
        mytablex.MoveNext
    Loop
    txformxupc.Fields("items") = sdx
    txformxupc.Fields("total") = sdx2
    txformxupc.Update
    Exit Sub
cmd86712_err:
    Exit Sub

End Sub

Function borrar_notaingreso()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from parteproducciond where numero=" & txformxupc("numero"), cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        borrar_notaingreso = 1
        mytablex.Close
        Exit Function

    End If

    mytablex.Close
    Exit Function

End Function
