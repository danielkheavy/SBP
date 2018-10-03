VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tasiento 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asientos Contables"
   ClientHeight    =   10230
   ClientLeft      =   150
   ClientTop       =   0
   ClientWidth     =   14970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10230
   ScaleWidth      =   14970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consulta"
      Height          =   10215
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   14535
      Begin VB.TextBox cadena 
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
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command3 
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
         Left            =   5520
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid dbGrid3 
         Height          =   9255
         Left            =   240
         TabIndex        =   29
         Top             =   840
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   16325
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   25
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
            Weight          =   700
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
               LCID            =   3082
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
               LCID            =   3082
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
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Frame2"
      Height          =   10215
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   14655
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   10800
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   1320
         Width           =   3615
      End
      Begin VB.CommandButton image8 
         BackColor       =   &H00FFFF80&
         Caption         =   "Refresca"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   9240
         Picture         =   "tasiento.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   9120
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ir Detalle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   12600
         Picture         =   "tasiento.frx":175E
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox libro 
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
         TabIndex        =   44
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox tipo 
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
         TabIndex        =   42
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox cod_asien 
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
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Ir Cabecera"
         Enabled         =   0   'False
         Height          =   945
         Left            =   12360
         Picture         =   "tasiento.frx":1A68
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Imprimir todo"
         Top             =   9120
         Width           =   1470
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00FFFF80&
         Caption         =   "&GuardarEnElDiario"
         Enabled         =   0   'False
         Height          =   975
         Left            =   10800
         Picture         =   "tasiento.frx":2332
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   9120
         Width           =   1470
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
         TabIndex        =   11
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox referencia 
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
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1320
         Width           =   6375
      End
      Begin VB.TextBox fuente 
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
         TabIndex        =   9
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox comproba 
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
         MaxLength       =   20
         TabIndex        =   8
         Top             =   2040
         Width           =   1935
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "tasiento.frx":2BFC
         Height          =   3975
         Left            =   120
         OleObjectBlob   =   "tasiento.frx":2C10
         TabIndex        =   38
         Top             =   3960
         Width           =   14415
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cabecera"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   14415
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Detalle"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   54
         Top             =   3600
         Width           =   14415
      End
      Begin VB.Label nruc 
         BackColor       =   &H00FFFF80&
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
         Left            =   1320
         TabIndex        =   52
         Top             =   8280
         Width           =   4335
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ruc"
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
         TabIndex        =   51
         Top             =   8280
         Width           =   1215
      End
      Begin VB.Line Line1 
         X1              =   10440
         X2              =   14280
         Y1              =   8520
         Y2              =   8520
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ccostos"
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
         TabIndex        =   50
         Top             =   7920
         Width           =   1215
      End
      Begin VB.Label nccosto 
         BackColor       =   &H00FFFF80&
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
         Left            =   1320
         TabIndex        =   49
         Top             =   7920
         Width           =   4335
      End
      Begin VB.Label librocorre 
         BackColor       =   &H00FFFF80&
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
         Left            =   2280
         TabIndex        =   48
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Correlativo"
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
         TabIndex        =   47
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label nlibro 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   4560
         TabIndex        =   46
         Top             =   2400
         Width           =   4095
      End
      Begin VB.Image Image3 
         BorderStyle     =   1  'Fixed Single
         Height          =   360
         Left            =   1680
         Picture         =   "tasiento.frx":3E2B
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   660
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Libro"
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
         TabIndex        =   45
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   360
         Left            =   1680
         Picture         =   "tasiento.frx":42CF
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   660
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   360
         Left            =   1680
         Picture         =   "tasiento.frx":4773
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   660
      End
      Begin VB.Label ntipo 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   4560
         TabIndex        =   43
         Top             =   1680
         Width           =   4095
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
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
         TabIndex        =   41
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Carga Grupo Cuentas"
         Enabled         =   0   'False
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
         Left            =   10800
         TabIndex        =   40
         Top             =   960
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F1.Consulta Cuentas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   8640
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label nfuente 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   4560
         TabIndex        =   24
         Top             =   3120
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Asiento Nro"
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
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha  DD/MM/AAAA"
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
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Referencia"
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
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fuente"
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
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label debe 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10440
         TabIndex        =   19
         Top             =   8040
         Width           =   1815
      End
      Begin VB.Label haber 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12240
         TabIndex        =   18
         Top             =   8040
         Width           =   2055
      End
      Begin VB.Label diferencia 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   10440
         TabIndex        =   17
         Top             =   8640
         Width           =   3855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Diferencia"
         Height          =   375
         Left            =   9360
         TabIndex        =   16
         Top             =   8640
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nro.Documento"
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
         TabIndex        =   15
         Top             =   2040
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFC0&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   14475
      TabIndex        =   0
      Top             =   0
      Width           =   14535
      Begin VB.TextBox btipo 
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
         Left            =   8760
         MaxLength       =   6
         TabIndex        =   59
         Text            =   "%"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox blibro 
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
         Left            =   8760
         MaxLength       =   6
         TabIndex        =   57
         Text            =   "%"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox bcuenta 
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
         Left            =   8760
         MaxLength       =   6
         TabIndex        =   34
         Text            =   "%"
         Top             =   0
         Width           =   1935
      End
      Begin VB.TextBox fechaf 
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
         Left            =   5280
         MaxLength       =   10
         TabIndex        =   32
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox Fechai 
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
         Left            =   5280
         MaxLength       =   10
         TabIndex        =   30
         Top             =   360
         Width           =   1935
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
         Picture         =   "tasiento.frx":4C17
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   5280
         MaxLength       =   11
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   1935
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
         Left            =   12480
         TabIndex        =   4
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
         Picture         =   "tasiento.frx":5E29
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Left            =   2160
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tasiento.frx":703B
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salir"
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
         Picture         =   "tasiento.frx":824D
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
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
         Left            =   7200
         TabIndex        =   60
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Libro Contable"
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
         Left            =   7200
         TabIndex        =   58
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nro.Documento"
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
         Left            =   3720
         TabIndex        =   37
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuenta Nro"
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
         Left            =   7200
         TabIndex        =   35
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Final"
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
         Left            =   3720
         TabIndex        =   33
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
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
         Left            =   3720
         TabIndex        =   31
         Top             =   360
         Width           =   1575
      End
   End
   Begin MSDataGridLib.DataGrid dbGrid1 
      Height          =   8895
      Left            =   0
      TabIndex        =   36
      Top             =   1200
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   15690
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12648447
      HeadLines       =   2
      RowHeight       =   16
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
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "cod_asien"
         Caption         =   "NroAsiento"
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
         DataField       =   "Fecha_asi"
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
         DataField       =   "Cod_libro"
         Caption         =   "Cod_libro"
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
         DataField       =   "Tipo"
         Caption         =   "Tipo"
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
      BeginProperty Column04 
         DataField       =   "Cuenta"
         Caption         =   "Cuenta"
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
      BeginProperty Column05 
         DataField       =   "Descripcio"
         Caption         =   "Descripcio"
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
      BeginProperty Column06 
         DataField       =   "Debito"
         Caption         =   "Debito"
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
      BeginProperty Column07 
         DataField       =   "Credito"
         Caption         =   "Credito"
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
      BeginProperty Column08 
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   824.882
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1649.764
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   5160.189
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1425.26
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   945.071
         EndProperty
      EndProperty
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
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu k8844 
         Caption         =   "&0.AsientoContable"
      End
      Begin VB.Menu dk89442 
         Caption         =   "&1.Libro Diario"
      End
      Begin VB.Menu j783 
         Caption         =   "&2.General"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu tr666 
         Caption         =   "&3.Reported"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tasiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim txasiento As New ADODB.Recordset
Private Sub ajdu1_Click()
Dim found As Integer
If Frame3.Visible = True Then Exit Sub
If Frame2.Visible = True Then Exit Sub
   
found = sql_ingreso()
If found = 0 Then
   MsgBox "Tasiento " + error$, 48, "Aviso"
   Exit Sub
End If

DBGrid2.Enabled = False
inicializa
Frame2.Visible = True
Frame2.Caption = "Nuevo"
DBGrid2.Enabled = False
cmdCerrar.Enabled = False
cmdGuardar.Enabled = False
'combo_dpreasiento
Combo1.Enabled = True
habilita 1
cod_asien.Enabled = False
cod_asien = ""
'cod_asien.SetFocus
Label16.Visible = True
ir_final
habilita_cabeza 0
fecha.SetFocus
End Sub

Private Sub bo712_Click()
Dim buf As String
On Error GoTo cmd656_err
If Frame3.Visible = True Then Exit Sub
If Frame2.Visible = True Then Exit Sub
buf = "" & txasiento.Fields("cod_asien")
If Frame2.Visible = True Then
   DBGrid1.SetFocus
   Exit Sub
End If
If MsgBox("Desea Borra " + "" & txasiento.Fields("cod_asien"), 1, "Aviso") <> 1 Then
   Exit Sub
End If
cn.Execute ("delete from asientos where cod_asien=" & txasiento.Fields("cod_asien"))
Command1_Click
Exit Sub
cmd656_err:
MsgBox "Seleccione un dato " + error$, 48, "Aviso"
Exit Sub

End Sub


Private Sub buffer_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
Command1_Click
End Sub

Private Sub cadena_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 13 Then
   ejecuta2 1
   Exit Sub
End If
If KeyAscii = 27 Then
   If opcion1 = "1" Then
      Frame3.Visible = False
      Frame3.Enabled = False
      fuente.SetFocus
      Exit Sub
   End If
   If opcion1 = "5" Then
      Frame3.Visible = False
      Frame3.Enabled = False
      tipo.SetFocus
      Exit Sub
   End If
   If opcion1 = "500" Then
      Frame3.Visible = False
      Frame3.Enabled = False
      libro.SetFocus
      Exit Sub
   End If
   
   If opcion1 = "2" Then
      Frame3.Visible = False
      Frame3.Enabled = False
      DBGrid2.SetFocus
      Exit Sub
   End If
   If opcion1 = "3" Then
      Frame3.Visible = False
      Frame3.Enabled = False
      DBGrid2.SetFocus
      Exit Sub
   End If
   If opcion1 = "300" Then
      Frame3.Visible = False
      Frame3.Enabled = False
      DBGrid2.SetFocus
      Exit Sub
   End If

End If
End Sub

Private Sub cmdAddEntry_Click()
ajdu1_Click
End Sub

Private Sub cmdCerrar_Click()
habilita_cabeza 0
'Combo1.Enabled = True
DBGrid2.Enabled = False
cmdCerrar.Enabled = False
cmdGuardar.Enabled = False
fecha.SetFocus

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
'k8844_Click
End Sub

Private Sub cmdSave_Click()
f8443_Click
End Sub


Private Sub cod_asien_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   dlo132_Click
   Exit Sub
End If
If Len(cod_asien) = 0 Then Exit Sub
fecha.SetFocus

End Sub



Private Sub Combo1_Click()
If Trim(Combo1) = "" Or Trim(Combo1) = "%" Then
   Exit Sub
End If
If MsgBox("Desea Cargar Cuenta Predefinida " & Combo1, 1, "Aviso") <> 1 Then Exit Sub
carga_dpreasiento

End Sub

Private Sub Command1_Click()
'Frame1.Visible = True
'Frame1.Enabled = True
'buffer = ""
opcion1 = "1"
ejecuta 1

End Sub
Sub ejecuta(sw As Integer)
Dim cad As String
If Not IsDate(fechai) Then
   fechai.SetFocus
   Exit Sub
End If
If Not IsDate(fechaf) Then
   fechaf.SetFocus
   Exit Sub
End If

If Len(buffer) > 0 Then
   If Not IsNumeric(buffer) Then
      MsgBox "Referencia Numerico ", 48, "Aviso"
      buffer = ""
      buffer.Enabled = True
      buffer.SetFocus
      Exit Sub
   End If
End If

cad = "SELECT *  from asientos   where  "
cad = cad & " fecha_asi>='" & Format(fechai, "YYYYMMDD") & "'"
cad = cad & " and fecha_asi<='" & Format(fechaf, "YYYYMMDD") & "' "
If Val(buffer) > 0 Then
   cad = cad & " and cod_asien=" & buffer
End If
If bcuenta <> "%" Then
   cad = cad & " and cuenta like '" & bcuenta & "'"
End If
If btipo <> "%" Then
   cad = cad & " and tipo like '" & btipo & "'"
End If
If blibro <> "%" Then
   cad = cad & " and cod_libro like '" & blibro & "'"
End If

cad = cad & " order by cod_asien ,id,fecha_asi,comproba"

   
   If txasiento.State = 1 Then txasiento.Close
   txasiento.Open cad, cn, adOpenStatic, adLockOptimistic
   Set DBGrid1.DataSource = txasiento
   'dbGrid1.columns(0).Width = 2000
   'dbGrid1.columns(1).Width = 2000
   'dbGrid1.columns(2).Width = 4000
   If txasiento.RecordCount > 0 Then
     DBGrid1.SetFocus
  End If
'End If
End Sub

Sub ejecuta1(sw As Integer)
Dim buf As String
Dim mytablex As New ADODB.Recordset
If opcion1 = "1" Then
   If Len(cadena) = 0 Then
   buf = "select Descripcio,tipoclie from tipoclie "
   Else
   buf = "select Descripcio,tipoclie from tipoclie where " & Combo2 & " like '" & cadena & "%'"
   End If
End If
If opcion1 = "2" Then
   If Len(cadena) = 0 Then
   buf = "select Nombre,Codigo from Vendedor "
   Else
   buf = "select Nombre,Codigo from Vendedor where " & Combo2 & " like '" & cadena & "%'"
   End If
End If



   If mytablex.State = 1 Then mytablex.Close
   mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
   Set dbgrid3.DataSource = mytablex
   dbgrid3.columns(0).Width = 4000
   dbgrid3.columns(1).Width = 2000
   If mytablex.RecordCount = 0 Then
      mytablex.Close
      cadena.SetFocus
      Exit Sub
   End If
   dbgrid3.SetFocus
End Sub
Sub ejecuta2(sw As Integer)
Dim buf As String
Dim mytablex As New ADODB.Recordset
If opcion1 = "5" Then
If Len(cadena) = 0 Then
buf = "select Descripcio,Docta from Docta "
Else
buf = "select Descripcio,Docta from Docta where " & Combo2 & " like '" & cadena & "%'"
End If
End If
If opcion1 = "500" Then
If Len(cadena) = 0 Then
buf = "select Descripcio,LibroAuxiliar from LibroAuxiliar "
Else
buf = "select Descripcio,LibroAuxiliar from LibroAuxiliar where " & Combo2 & " like '" & cadena & "%'"
End If
End If
If opcion1 = "1" Then
If Len(cadena) = 0 Then
buf = "select Descripcio,Fuente from fuente "
Else
buf = "select Descripcio,Fuente from fuente where " & Combo2 & " like '" & cadena & "%'"
End If
End If
If opcion1 = "2" Then
If Len(cadena) = 0 Then
buf = "select Descripcio,Codcta,Nivel_cta,Flag_ruc from cuentas ORDER by codcta"
Else
buf = "select Descripcio,Codcta,Nivel_cta,Flag_ruc from cuentas where " & Combo2 & " like '" & cadena & "%' order by codcta "
End If
End If
If opcion1 = "3" Then
If Len(cadena) = 0 Then
buf = "select Nombre,Codigo from clientes "
Else
buf = "select Nombre,Codigo from Clientes where " & Combo2 & " like '" & cadena & "%'"
End If
End If
If opcion1 = "300" Then
If Len(cadena) = 0 Then
buf = "select Descripcio,CCosto from ccosto "
Else
buf = "select Descripcio,Ccosto from ccosto where " & Combo2 & " like '" & cadena & "%'"
End If
End If

'MsgBox buf


   If mytablex.State = 1 Then mytablex.Close
   mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
   Set dbgrid3.DataSource = mytablex
   dbgrid3.columns(0).Width = 4000
   dbgrid3.columns(1).Width = 2000
   If mytablex.RecordCount = 0 Then
      mytablex.Close
      cadena.SetFocus
      Exit Sub
   End If
   dbgrid3.SetFocus
End Sub

Private Sub Command2_Click()
Dim found As Integer
found = validar()
If found = 0 Then Exit Sub
habilita_cabeza 1
'Combo1.Enabled = False
DBGrid2.Enabled = True
cmdCerrar.Enabled = True
cmdGuardar.Enabled = True
ir_final
End Sub

Private Sub Command3_Click()
ejecuta2 1
End Sub

Private Sub comproba_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
libro.SetFocus
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   buffer.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
End If
End Sub



Private Sub DBGrid2_AfterColUpdate(ByVal ColIndex As Integer)
Select Case ColIndex
       Case 4
            ir_final
       Case 5
            ir_final
End Select
End Sub

Private Sub DBGrid2_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)

 Select Case ColIndex
Case 0
     'MsgBox "" & DBGrid2.Columns(0)
     If Len("" & DBGrid2.columns(0)) > 0 Then  'si ya existe no cambiar
        Cancel = True
        Exit Sub
     End If
     
Case 1
     If Len("" & DBGrid2.columns(0)) = 0 Then  'si ya existe no cambiar
        Cancel = True
        Exit Sub
     End If
     
     
Case 2
     If Len("" & DBGrid2.columns(0)) = 0 Then  'si ya existe no cambiar
        Cancel = True
        Exit Sub
     End If
Case 3
     If Len("" & DBGrid2.columns(0)) = 0 Then  'si ya existe no cambiar
        Cancel = True
        Exit Sub
     End If
     
Case 4
     If Len("" & DBGrid2.columns(0)) = 0 Then  'si ya existe no cambiar
        Cancel = True
        Exit Sub
     End If
     
     
Case 5
     If Len("" & DBGrid2.columns(0)) = 0 Then  'si ya existe no cambiar
        Cancel = True
        Exit Sub
     End If
     
     
     
     

End Select
End Sub

Private Sub DBGrid2_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim found As Integer
Select Case ColIndex

Case 4
      If Len(DBGrid2.columns(0)) = 0 Then
        Cancel = True
        Exit Sub
      End If
     
     If Not IsNumeric("" & DBGrid2.columns("debito")) Then  'si ya existe no cambiar
        Cancel = True
        Exit Sub
     End If
     
Case 5
      If Len(DBGrid2.columns(0)) = 0 Then
        Cancel = True
        Exit Sub
      End If
     
     If Not IsNumeric("" & DBGrid2.columns("credito")) Then  'si ya existe no cambiar
        Cancel = True
        Exit Sub
     End If
     
Case 2
      If Len(DBGrid2.columns(0)) = 0 Then
        Cancel = True
        Exit Sub
      End If
      If Len(DBGrid2.columns("ccosto")) > 0 Then
         found = busca_ccosto("" & DBGrid2.columns("ccosto"))
         If found = 0 Then
            MsgBox "Debe Existir CCosto ", 48, "Aviso"
            Cancel = True
            Exit Sub
         End If
      End If

     
Case 3
      If Len(DBGrid2.columns(0)) = 0 Then
        Cancel = True
        Exit Sub
      End If
      If Len(DBGrid2.columns("nro_ruc")) > 0 Then
         found = busca_tercero("" & DBGrid2.columns("Nro_ruc"))
         If found = 0 Then
            MsgBox "Debe Existir Ruc ", 48, "Aviso"
            Cancel = True
            Exit Sub
         End If
      End If
Case 0
      If Len(DBGrid2.columns(0)) = 0 Then
        Cancel = True
        Exit Sub
      End If
      found = busca_cuenta("" & DBGrid2.columns(0))
      If found = 0 Then
         MsgBox "Cuenta No existe ", 48, "Aviso"
         Cancel = True
         Exit Sub
      End If
      
End Select
End Sub

Private Sub dbgrid2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim found As Integer
found = busca_ccosto("" & DBGrid2.columns("ccosto"))
found = busca_tercero("" & DBGrid2.columns("Nro_ruc"))
Exit Sub
If KeyCode = 13 Then
Select Case DBGrid2.Col
       Case 4, 5
       If Val(DBGrid2.columns(4)) > 0 Or Val(DBGrid2.columns(5)) > 0 Then
       DBGrid2.Col = 0
       DBGrid2.Row = DBGrid2.VisibleRows - 1
       DBGrid2.SetFocus
       End If
End Select
End If
End Sub

Private Sub DBGrid2_KeyUp(KeyCode As Integer, Shift As Integer)
Dim found As Integer
On Error GoTo cmd45_err
'----F1
found = busca_ccosto("" & DBGrid2.columns("ccosto"))
found = busca_tercero("" & DBGrid2.columns("Nro_ruc"))

If KeyCode = &H70 Then  'f1
   If DBGrid2.Col = 0 Then
      Combo2.Clear
      Combo2.AddItem "Descripcio"
      Combo2.AddItem "codcta"
      Combo2.ListIndex = 1
      opcion1 = "2"
      Frame3.Visible = True
      Frame3.Enabled = True
      cadena = ""
      cadena.SetFocus
      Command3_Click
   End If
   If DBGrid2.Col = 3 Then
      Combo2.Clear
      Combo2.AddItem "Nombre"
      Combo2.AddItem "Codigo"
      Combo2.ListIndex = 0
      opcion1 = "3"
      Frame3.Visible = True
      Frame3.Enabled = True
      cadena = ""
      cadena.SetFocus
      Command3_Click
   End If
   If DBGrid2.Col = 2 Then  'CCOSTO
      Combo2.Clear
      Combo2.AddItem "Descripcio"
      Combo2.AddItem "ccosto"
      Combo2.ListIndex = 0
      opcion1 = "300"
      Frame3.Visible = True
      Frame3.Enabled = True
      cadena = ""
      cadena.SetFocus
      Command3_Click
   End If
   Exit Sub
End If

'--- FIN F1

If KeyCode = &H2E Then  'borrar linea
If DBGrid2.Row = -1 Then
        MsgBox "No hay ningún registro para eliminar", vbInformation
        Exit Sub
End If
If MsgBox("Se va a eliminar el registro : está seguro ", _
   vbExclamation + vbYesNo, "Eliminar") = vbYes Then
   Data1.Recordset.Delete
   ir_final
   Exit Sub
End If
End If
Exit Sub
cmd45_err:
MsgBox "Aviso en keyup " + error$, 48, "Aviso"
Exit Sub

End Sub

Private Sub DBGrid3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   cadena.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
  
   If opcion1 = "1" Then
      fuente = Trim(dbgrid3.columns(1))
      nfuente = Trim(dbgrid3.columns(0))
      Frame3.Visible = False
      Frame3.Enabled = False
      fuente.SetFocus
   End If
   If opcion1 = "5" Then
      tipo = Trim(dbgrid3.columns(1))
      ntipo = Trim(dbgrid3.columns(0))
      Frame3.Visible = False
      Frame3.Enabled = False
      tipo.SetFocus
   End If
   If opcion1 = "500" Then
      libro = Trim(dbgrid3.columns(1))
      nlibro = Trim(dbgrid3.columns(0))
      Frame3.Visible = False
      Frame3.Enabled = False
      libro.SetFocus
   End If
   
   If opcion1 = "2" Then
      'If Trim("" & dbGrid3.columns("Nivel_cta")) <> "D" Then
      '   Exit Sub
      'End If
      DBGrid2.columns(0) = Trim(dbgrid3.columns(1))
      DBGrid2.columns(6) = Trim(dbgrid3.columns(0))
      'ncuenta = Trim(dbGrid3.columns(0))
      Frame3.Visible = False
      Frame3.Enabled = False
      DBGrid2.SetFocus
   End If
   If opcion1 = "3" Then
      DBGrid2.columns(3) = Trim(dbgrid3.columns(1))
      nruc = Trim(dbgrid3.columns(0))
      'DBGrid2.columns(6) = Trim(dbGrid3.columns(0))
      Frame3.Visible = False
      Frame3.Enabled = False
      DBGrid2.SetFocus
   End If
   If opcion1 = "300" Then 'ccosto
      DBGrid2.columns(2) = Trim(dbgrid3.columns(1))
      nccosto = Trim(dbgrid3.columns(0))
      Frame3.Visible = False
      Frame3.Enabled = False
      DBGrid2.SetFocus
   End If


End If
End Sub

Private Sub dbgrid3_KeyPress(KeyAscii As Integer)
Dim buf As String
Dim buf2 As String
If KeyAscii <> 13 And KeyAscii <> 27 Then
  
         If KeyAscii = 8 Then
            If Len(cadena) > 0 Then
               buf = Mid$(cadena, 1, Len(cadena) - 1)
               cadena = buf
               KeyAscii = 0
               Else
               KeyAscii = 0
               Exit Sub
            End If
         End If
         buf = Chr(KeyAscii)
         If Chr(KeyAscii) = "*" Then
            buf = ""
            cadena = buf
         End If
         If KeyAscii <> 13 Then
            cadena = cadena + buf
         End If
         buf = cadena
         ejecuta2 0
         
End If
End Sub

Private Sub dk89442_Click()
 Dim found As Integer
 Dim i As Integer
 Dim v As Long
 Dim R As Long
 Dim ih As Integer
 Dim H As Integer
 Dim cad As String
 Dim tmp1 As String
 
 Dim sdx As Double
 Dim xdebito As Double
 Dim xcredito As Double
 Dim mytablex As New ADODB.Recordset
 Dim sw As Integer
 Dim Tmp As String
    Dim Heading(8) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion
    On Error GoTo cmd445612_err
If MsgBox("Desea Exportar excel", 1, "Aviso") <> 1 Then Exit Sub
    'Heading(1) = "Cod.Asiento"
    'Heading(2) = "Cuenta"
    'Heading(3) = "Descripcio"
    'Heading(4) = "Debitos"
    'Heading(5) = "Creditos"
    'Heading(6) = "Motivo"
    
    If txasiento.RecordCount = 0 Then Exit Sub
    txasiento.MoveFirst
    
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    
    '------------------------------------------------
    With objExcel.ActiveSheet
        '.Range(.Cells(1, 1), .Cells(10, 3)).Borders.LineStyle = xlContinuous
        .Range(.Cells(4, 1), .Cells(4, 20)).Borders.LineStyle = xlContinuous
       
        .columns("A").ColumnWidth = 15
        .columns("B").ColumnWidth = 15
        .columns("C").ColumnWidth = 30
        .columns("D").ColumnWidth = 15
        .columns("E").ColumnWidth = 15
        .columns("f").ColumnWidth = 15
    
End With
    'cabecera
    mytablex.Open "select * from empresa where codigo='01'", cn, adOpenStatic, adLockOptimistic
    If mytablex.RecordCount > 0 Then
    objExcel.ActiveSheet.Cells(2, 1) = "'" & mytablex.Fields("nombre")
    End If
    mytablex.Close
    objExcel.ActiveSheet.Cells(2, 5) = "'" & Format(Now, "dd/mm/yyyy")
    objExcel.ActiveSheet.Cells(3, 2) = "'Reporte de Libro Diario:" + "Desde:" + fechai + " Hasta:" + fechaf + " en Soles(S/.)"
    
    
   
    
    '------------------------------------------------
    
    
    objExcel.ActiveSheet.Cells(4, 1) = "'Cod.Asiento"
    objExcel.ActiveSheet.Cells(4, 2) = "'Cod.Cta"
    objExcel.ActiveSheet.Cells(4, 3) = "'NombreCta"
    objExcel.ActiveSheet.Cells(4, 4) = "'Debito"
    objExcel.ActiveSheet.Cells(4, 5) = "'Credito"
    objExcel.ActiveSheet.Cells(4, 6) = "'Referencia"
    
    
    '------------------------------------------------
v = 5
H = 1
    xdebito = 0
    xcredito = 0
    sw = 0
    
    Do
         If txasiento.EOF Then Exit Do
          tmp1 = "" & txasiento.Fields("cod_asien") & "" & txasiento.Fields("fecha_asi") & "" & txasiento.Fields("comproba")
         If sw = 0 Then
            sw = 1
            Tmp = "" & txasiento.Fields("cod_asien") & "" & txasiento.Fields("fecha_asi") & "" & txasiento.Fields("comproba")
         
         objExcel.ActiveSheet.Cells(v, 1) = "Fecha:" & txasiento.Fields("fecha_asi") & " Comproba:" & txasiento.Fields("comproba") & " Tipo:" & txasiento.Fields("tipo") & " Referencia:" & txasiento.Fields("referencia")
         v = v + 1
                
         End If
         If Tmp <> tmp1 Then
            objExcel.ActiveSheet.Cells(v, 4) = xdebito
            objExcel.ActiveSheet.Cells(v, 5) = xcredito
            xdebito = 0
            xcredito = 0
            Tmp = "" & txasiento.Fields("cod_asien") & "" & txasiento.Fields("fecha_asi") & "" & txasiento.Fields("comproba")
            objExcel.ActiveSheet.Cells(v, 1) = "Fecha:" & txasiento.Fields("fecha_asi") & " Comproba:" & txasiento.Fields("comproba") & " Tipo:" & txasiento.Fields("tipo") & " Referencia:" & txasiento.Fields("referencia")
            v = v + 1
         End If
         
   objExcel.ActiveSheet.Cells(v, 1) = "'" & txasiento.Fields("cod_asien")
   objExcel.ActiveSheet.Cells(v, 2) = "'" & txasiento.Fields("cuenta")
   objExcel.ActiveSheet.Cells(v, 3) = "'" & txasiento.Fields("Descripcio")
   If "" & txasiento.Fields("tipo_cta") = "D" Then
      objExcel.ActiveSheet.Cells(v, 4) = "" & txasiento.Fields("cantidad")
      xdebito = xdebito + Val("" & txasiento.Fields("cantidad"))
   End If
   If "" & txasiento.Fields("tipo_cta") = "H" Then
      objExcel.ActiveSheet.Cells(v, 5) = "" & txasiento.Fields("cantidad")
      xcredito = xcredito + Val("" & txasiento.Fields("cantidad"))
   End If
   objExcel.ActiveSheet.Cells(v, 6) = "" & txasiento.Fields("motivo")
   
   v = v + 1
txasiento.MoveNext
Loop

objExcel.ActiveSheet.Cells(v, 4) = xdebito
objExcel.ActiveSheet.Cells(v, 5) = xcredito
Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
Exit Sub
cmd445612_err:
MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
Exit Sub

End Sub

Private Sub dlo132_Click()
If Frame3.Visible = True Then
   cadena_KeyPress 27
   Exit Sub  '999751845
End If

If Frame2.Visible = True Then
   habilita 0
   Frame2.Visible = False
   DBGrid1.Enabled = True
   Exit Sub
End If
tasiento.Hide
Unload tasiento
End Sub


Private Sub f8443_Click()
Dim buf As String
Dim found As Integer
On Error GoTo cmd456_err
If Frame3.Visible = True Then Exit Sub
If Frame2.Visible = True Then Exit Sub
buf = txasiento.Fields("cod_asien")
If Frame2.Visible = True Then
   DBGrid1.SetFocus
   Exit Sub
End If

found = sql_ingreso()
If found = 0 Then
   MsgBox "Tasiento " + error$, 48, "Aviso"
   Exit Sub
End If


inicializa
Frame2.Visible = True
Frame2.Caption = "Modifica"
DBGrid2.Enabled = False
DBGrid2.Enabled = False
cmdCerrar.Enabled = False
cmdGuardar.Enabled = False
'combo_dpreasiento
Combo1.Enabled = False
pone_registro
habilita 1
Label16.Visible = True
habilita_cabeza 0
cod_asien.Enabled = False
ir_final

fecha.SetFocus
Exit Sub
cmd456_err:
MsgBox "Seleccione un dato " + error$, 48, "Aviso"
Exit Sub

End Sub

Private Sub fecha_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
referencia.SetFocus

End Sub

Private Sub fjh433_Click()
Dim found As Integer
Dim buf As String
On Error GoTo cmd556_err
If Frame3.Visible = True Then Exit Sub
If Frame2.Visible = True Then Exit Sub

buf = txasiento.Fields("cod_asien")
If Frame2.Visible = True Then
   DBGrid1.SetFocus
   Exit Sub
End If
found = sql_ingreso()
If found = 0 Then
   MsgBox "Tasiento " + error$, 48, "Aviso"
   Exit Sub
End If


inicializa
Frame2.Visible = True
Frame2.Caption = "Zoom"
DBGrid2.Enabled = False
cmdCerrar.Enabled = False
cmdGuardar.Enabled = False
'combo_dpreasiento
Combo1.Enabled = False
pone_registro
habilita 1
Label16.Visible = False
cod_asien.Enabled = False
ir_final
habilita_cabeza 0

Command2.Enabled = False
DBGrid2.Enabled = False
fecha.SetFocus
Exit Sub
cmd556_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub

End Sub

Private Sub Form_Activate()
combo_dpreasiento
Command1_Click
End Sub
Function sql_ingreso()
On Error GoTo cmd7812_err
               Data1.Connect = "foxpro 2.5;"
               Data1.DatabaseName = globaldat
               Data1.RecordSource = "SELECT * FROM tasiento where len(cuenta)>0"
               Data1.refresh
               borrar_data1
               Data1.refresh
               DBGrid2.refresh
               sql_ingreso = 1
               Exit Function
cmd7812_err:
MsgBox "Aviso en sql_ingreso " + error$, 48, "Aviso"
Exit Function
End Function
Sub borrar_data1()
On Error GoTo cmd212_err
inicio1:
Data1.Recordset.MoveFirst
Data1.Recordset.Delete
GoTo inicio1
Exit Sub
cmd212_err:
Exit Sub

End Sub
Sub inicializa()
'descripcio = ""
fecha = Format(Now, "dd/mm/yyyy")
referencia = ""
comproba = ""
fuente = ""
tipo = ""
ntipo = ""
libro = ""
librocorre = ""
debe = ""
haber = ""
diferencia = ""
nccosto = ""
nruc = ""
ntipo = ""
nlibro = ""
nfuente = ""
End Sub
Sub pone_registro()
Dim found As Integer
cod_asien = Trim("" & txasiento.Fields("cod_asien"))
fecha = Trim("" & txasiento.Fields("fecha_asi"))
referencia = Trim("" & txasiento.Fields("referencia"))
comproba = Trim("" & txasiento.Fields("comproba"))
fuente = Trim("" & txasiento.Fields("fuente"))
tipo = Trim("" & txasiento.Fields("tipo"))
libro = Trim("" & txasiento.Fields("cod_libro"))
librocorre = Trim("" & txasiento.Fields("corre_libro"))
found = busca_tipo()
found = busca_libro("" & libro)
found = busca_tipo()
carga_documento cod_asien
End Sub
Sub grabando()
txasiento.Fields("cod_asien") = Trim(cod_asien)
txasiento.Fields("fecha_asi") = Format(fecha, "dd/mm/yyyy")
txasiento.Fields("referencia") = Trim(referencia)
txasiento.Fields("comproba") = Trim(comproba)
txasiento.Fields("fuente") = Trim(fuente)
txasiento.Fields("tipo") = Trim(tipo)
txasiento.Fields("cod_libro") = Trim(libro)

End Sub

Private Sub grba1_Click()

End Sub

Function grabar()
Dim found As Integer
Dim sdx As Double
Dim mytablex As New ADODB.Recordset
Dim rbusca As New ADODB.Recordset
ir_final
If Val(debe) <= 0 Then
   MsgBox "No existen datos del debe ", 48, "Aviso"
   DBGrid2.SetFocus
   Exit Function
End If
If Val(haber) <= 0 Then
   MsgBox "No existen datos del haber ", 48, "Aviso"
   DBGrid2.SetFocus
   Exit Function
End If

If Val(diferencia) > 0 Then
   MsgBox "Cantidades no cuadran ", 48, "Aviso"
   DBGrid2.SetFocus
   Exit Function
End If

'-----------VALIDAMOS SI ES OBLIGATORIO INGRESAR RUC
   Data1.refresh
   Do
   If Data1.Recordset.EOF Then Exit Do
      If obliga_ruc() = 1 Then
         MsgBox "Es necesario el Numero de Ruc ", 48, "Aviso"
         Exit Do
      End If
      If obliga_ccosto() = 1 Then
         MsgBox "Es necesario el Ccosto ", 48, "Aviso"
         Exit Do
      End If
   Data1.Recordset.MoveNext
   Loop


'found = valida()
'If found = 0 Then
'   MsgBox "Campos invalidos", 48, "Aviso"
'   Exit Function
'End If
If Frame2.Caption = "Nuevo" Then
   'If Len(cod_asien) = 0 Then
   '   cod_asien.SetFocus
   '   Exit Function
   'End If
   mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount = 0 Then
      mytablex.Close
      Exit Function
   End If
   sdx = Val("" & mytablex.Fields("asientos")) + 1
amiga:
   cod_asien = "" & sdx
   rbusca.Open "select cod_asien from asientos where cod_asien=" & cod_asien & "", cn, adOpenStatic, adLockOptimistic
   If rbusca.RecordCount > 0 Then
      rbusca.Close
      sdx = sdx + 1
      GoTo amiga
   End If
   rbusca.Close
   mytablex.Fields("asientos") = "" & cod_asien
   mytablex.Update
   mytablex.Close
   
   Data1.refresh
   Do
   If Data1.Recordset.EOF Then Exit Do
   If Len("" & Data1.Recordset.Fields("cuenta")) > 0 Then
      txasiento.AddNew
      txasiento.Fields("cod_asien") = cod_asien
      grabando
      txasiento.Fields("fecha_asi") = Format(fecha, "dd/mm/yyyy")
      txasiento.Fields("nro_seq") = ""
      txasiento.Fields("cuenta") = "" & Data1.Recordset.Fields("cuenta")
      txasiento.Fields("ccosto") = "" & Data1.Recordset.Fields("ccosto")
      txasiento.Fields("debito") = Val("" & Data1.Recordset.Fields("debito"))
      txasiento.Fields("credito") = Val("" & Data1.Recordset.Fields("credito"))
      'txasiento.Fields("tipo_cta") = "" & Data1.Recordset.Fields("tipo_cta")
      txasiento.Fields("cod_libro") = "" & libro
      txasiento.Fields("corre_libro") = Val("" & librocorre)
      txasiento.Fields("motivo") = "" & Data1.Recordset.Fields("motivo")
      txasiento.Fields("referencia") = "" & referencia
      txasiento.Fields("comproba") = "" & comproba
      txasiento.Fields("fuente") = "" & fuente
      txasiento.Fields("nro_ruc") = "" & Data1.Recordset.Fields("nro_ruc")
      txasiento.Fields("vrbase") = 0
      txasiento.Fields("cod_cdec") = 0
      txasiento.Fields("descripcio") = "" & Data1.Recordset.Fields("descripcio")
      'txasiento.Fields("nombre") = "" & Data1.Recordset.Fields("nombre")
      txasiento.Update
   End If
   Data1.Recordset.MoveNext
   Loop
   dlo132_Click
   Exit Function
End If
If Frame2.Caption = "Modifica" Then
   ejecuta 1
   cn.Execute ("delete from asientos where cod_asien='" & cod_asien & "'")
   Data1.refresh
   Do
   If Data1.Recordset.EOF Then Exit Do
   If Len("" & Data1.Recordset.Fields("debito")) Or Val("" & Data1.Recordset.Fields("credito")) > 0 Then
      txasiento.AddNew
      txasiento.Fields("cod_asien") = cod_asien
      grabando
      txasiento.Fields("fecha_asi") = Format(fecha, "dd/mm/yyyy")
      txasiento.Fields("nro_seq") = ""
      txasiento.Fields("cuenta") = "" & Data1.Recordset.Fields("cuenta")
      txasiento.Fields("debito") = Val("" & Data1.Recordset.Fields("debito"))
      txasiento.Fields("credito") = Val("" & Data1.Recordset.Fields("credito"))
      'txasiento.Fields("tipo_cta") = "" & Data1.Recordset.Fields("tipo_cta")
      txasiento.Fields("cod_libro") = "" & libro
      txasiento.Fields("corre_libro") = Val("" & librocorre)
      txasiento.Fields("motivo") = "" & Data1.Recordset.Fields("motivo")
      txasiento.Fields("referencia") = "" & referencia
      txasiento.Fields("comproba") = "" & comproba
      txasiento.Fields("fuente") = "" & fuente
      txasiento.Fields("nro_ruc") = "" & Data1.Recordset.Fields("nro_ruc")
      txasiento.Fields("vrbase") = 0
      txasiento.Fields("cod_cdec") = 0
      txasiento.Fields("descripcio") = "" & Data1.Recordset.Fields("descripcio")
      txasiento.Fields("cuenta") = "" & Data1.Recordset.Fields("cuenta")
      txasiento.Update
   End If
   Data1.Recordset.MoveNext
   Loop
   ejecuta 1
   dlo132_Click
   Exit Function
End If

End Function

Function valida()
Dim found As Integer
'If Len(cod_asien) = 0 Then
'   cod_asien.SetFocus
'   Exit Function
'End If


If Len(fecha) <> 10 Then
   fecha.SetFocus
   Exit Function
End If
If Not IsDate(fecha) Then
   fecha.SetFocus
   Exit Function
End If
'found = busca_periodo()
'If found = 0 Then
'   MsgBox "Periodo Contable no Existe o no es el Mes", 48, "Aviso"
'   Exit Function
'End If
If Len(referencia) = 0 Then
   referencia.SetFocus
   Exit Function
End If
If Len(tipo) > 0 Then
   ntipo = ""
   found = busca_tipo()
   If found = 0 Then
      tipo.SetFocus
      Exit Function
   End If
   If Len(comproba) = 0 Then
      comproba.SetFocus
      Exit Function
   End If
End If

If Len(fuente) = 0 Then
   fuente.SetFocus
   Exit Function
End If
found = busca_fuente()
If found = 0 Then
   fuente.SetFocus
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
            'dbGrid1.Enabled = True

            
End If
If sw = 1 Then

            ajdu1.Enabled = False
            f8443.Enabled = False
            bo712.Enabled = False
            fjh433.Enabled = False
            djuer1.Enabled = False
            djuer1.Enabled = False
            Picture1.Enabled = False
            'dbGrid1.Enabled = False
'dbGrid1.Enabled = False

            
End If

      
End Sub

Private Sub Frame1_DragDrop(Source As control, x As Single, Y As Single)

End Sub

Private Sub Form_Load()
fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
fechaf = "30/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
End Sub

Private Sub fuente_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
found = busca_fuente()
If found = 0 Then
   MsgBox "No existe Fuente ", 48, "Aviso"
   fuente.SetFocus
   Exit Sub
End If
If Command2.Enabled = False Then Exit Sub
Command2_Click
End Sub
Sub ir_final()
Dim sdx As Double
Dim sdx1 As Double
Dim sdx2 As Double
On Error GoTo cmd12_err
sdx = 0
sdx1 = 0
diferencia = ""
debe = ""
haber = ""
Data1.refresh
Do
If Data1.Recordset.EOF Then Exit Do
sdx = sdx + Val("" & Data1.Recordset.Fields("debito"))
sdx1 = sdx1 + Val("" & Data1.Recordset.Fields("credito"))
Data1.Recordset.MoveNext
Loop
debe = Format(sdx, "0.00")
haber = Format(sdx1, "0.00")
sdx2 = sdx1 - sdx
diferencia = Format(Abs(sdx2), "0.00")
'Data1.Recordset.MoveLast
DBGrid2.Col = 0
DBGrid2.Row = DBGrid2.VisibleRows - 1
If DBGrid2.Enabled = True Then
   DBGrid2.SetFocus
End If
Exit Sub
cmd12_err:
MsgBox "Aviso en ir_final " + error$, 48, "Aviso"
If DBGrid2.Enabled = True Then
   DBGrid2.SetFocus
End If
Exit Sub
End Sub

Private Sub fuente_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   Combo2.Clear
   Combo2.AddItem "Descripcio"
   Combo2.AddItem "Fuente"
   Combo2.ListIndex = 0
   opcion1 = "1"
   Frame3.Visible = True
   Frame3.Enabled = True
   cadena = ""
   cadena.SetFocus
   Command3_Click
   Exit Sub
End If

End Sub

Private Sub Image1_Click()
   Combo2.Clear
   Combo2.AddItem "Descripcio"
   Combo2.AddItem "docta"
   Combo2.ListIndex = 0
   opcion1 = "5"
   Frame3.Visible = True
   Frame3.Enabled = True
   cadena = ""
   cadena.SetFocus
   Command3_Click
   Exit Sub

End Sub

Private Sub image2_Click()
   Combo2.Clear
   Combo2.AddItem "Descripcio"
   Combo2.AddItem "Fuente"
   Combo2.ListIndex = 0
   opcion1 = "1"
   Frame3.Visible = True
   Frame3.Enabled = True
   cadena = ""
   cadena.SetFocus
   Command3_Click
   Exit Sub

End Sub

Private Sub image3_Click()
   Combo2.Clear
   Combo2.AddItem "Descripcio"
   Combo2.AddItem "LibroAuxliar"
   Combo2.ListIndex = 0
   opcion1 = "500"
   Frame3.Visible = True
   Frame3.Enabled = True
   cadena = ""
   cadena.SetFocus
   Command3_Click
   Exit Sub

End Sub

Private Sub Image8_Click()
ir_final
End Sub

Private Sub j783_Click()
Dim sdx As String
On Error GoTo cmd8_err
sdx = "" & txasiento.Fields("cod_asien")
impresion1
Exit Sub
cmd8_err:
MsgBox "Elegir un dato ", 48, "Aviso"
Exit Sub

End Sub
Private Sub impresion1()
Dim found As Integer
Dim buf As String
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
Dim buf As String
Dim i As Integer
Dim found As Integer
    If contlin > 0 Then
       buf = Chr$(12)
       found = formateaa(buf, Len(buf), 0, 0)
    End If
    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = "Reporte de Asiento Contable en Soles(S/.)  "
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Numero      :", 13, 0, 0)
    found = formateaa("" & txasiento.Fields("cod_asien"), 10, 2, 0)
        
    found = formateaa("Fecha       :", 13, 0, 0)
    found = formateaa("" & txasiento.Fields("fecha_asi"), 10, 2, 0)
    
    found = formateaa("Referencia  :", 13, 0, 0)
    found = formateaa("" & txasiento.Fields("Referencia"), 10, 2, 0)
    
    found = formateaa("Comprobante :", 13, 0, 0)
    found = formateaa("" & txasiento.Fields("comproba"), 10, 2, 0)
    
    found = formateaa("Fuente      :", 13, 0, 0)
    found = formateaa("" & txasiento.Fields("fuente"), 10, 2, 0)
    
        
    buf = String(150, "-")
    found = formateaa(buf, 150, 2, 0)
    '------aqui van los registros----------------------
    
        
    found = formateaa("Cuenta", 15, 0, 0)
    found = formateaa("Descripcio", 41, 0, 0)
    found = formateaa("Debito ", 11, 0, 0)
    found = formateaa("Credito ", 11, 0, 0)
    found = formateaa("Comentario(Glosa) ", 30, 2, 0)
    '--------------------------------------------------
    buf = String(150, "-")
    found = formateaa(buf, 150, 2, 0)
    

End Sub
Sub cuerpo_programa_documento()
Dim xdebito As Double
Dim xcredito As Double
Dim buf As String
Dim found As Integer
Dim mytablex As New ADODB.Recordset
On Error GoTo cmd788_err
suma1 = 0
suma2 = 0
xdebito = 0
xcredito = 0
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "select * from asientos where cod_asien='" & txasiento("cod_asien") & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount = 0 Then
   mytablex.Close
   Exit Sub
End If


Do
'MsgBox "" & mytablex.Fields("producto")
If mytablex.EOF Then Exit Do
'-----------------------------------------
buf = "" & mytablex.Fields("cuenta")
found = formateaa(buf, 14, 0, 0)
found = formateaa("", 1, 0, 0)
buf = "" & mytablex.Fields("descripcio")
found = formateaa(buf, 40, 0, 0)
found = formateaa("", 1, 0, 0)

xdebito = 0
xcredito = 0
If "" & mytablex.Fields("tipo_cta") = "D" Then
   xdebito = Val("" & mytablex.Fields("cantidad"))
End If
If "" & mytablex.Fields("tipo_cta") = "H" Then
   xcredito = Val("" & mytablex.Fields("cantidad"))
End If

buf = Format(xdebito, "0.00")
If Val(buf) = 0 Then
   buf = ""
End If
found = formateaa(buf, 10, 0, 1)
found = formateaa("", 1, 0, 0)

buf = Format(xcredito, "0.00")
If Val(buf) = 0 Then
   buf = ""
End If

found = formateaa(buf, 10, 0, 1)
found = formateaa("", 1, 0, 0)

buf = "" & mytablex.Fields("motivo")
found = formateaa(buf, 30, 0, 0)
found = formateaa("", 1, 2, 0)

suma1 = suma1 + xdebito
suma2 = suma2 + xcredito
nlineas
mytablex.MoveNext
Loop
mytablex.Close
found = formateaa("Total ", 56, 0, 1)
buf = Format(suma1, "0.00")
If Val(buf) = 0 Then
   buf = ""
End If

found = formateaa(buf, 10, 0, 1)
found = formateaa("", 1, 0, 0)

buf = Format(suma2, "0.00")
If Val(buf) = 0 Then
   buf = ""
End If
found = formateaa(buf, 10, 0, 1)
found = formateaa("", 1, 2, 0)
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

Private Sub k8844_Click()
 Dim found As Integer
 Dim i As Integer
 Dim v As Long
 Dim R As Long
 Dim ih As Integer
 Dim H As Integer
 Dim cad As String
 Dim Tmp As String
 Dim sw As Integer
 Dim sdx As Double
 Dim xdebito As Double
 Dim xcredito As Double
 Dim buf As String
 Dim mytablex As New ADODB.Recordset
    Dim Heading(8) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion
    On Error GoTo cmd45612_err
    'If Not IsNumeric(buffer) Or Val(buffer) < 0 Then
    '   MsgBox "Poner un Numero de referencia ", 48, "Aviso"
    '   Exit Sub
    'End If
    buf = Trim("" & txasiento.Fields("cod_asien"))
    
If MsgBox("Desea Exportar excel", 1, "Aviso") <> 1 Then Exit Sub
    Heading(1) = "Codigo de Cta."
    Heading(2) = "Descripcion de la Cuenta "
    Heading(3) = "Clasificacion Cta"
    Heading(4) = "Cuenta"
    Heading(5) = dicruc
    'Command1_Click
    If txasiento.RecordCount = 0 Then Exit Sub
    'txasiento.MoveFirst
    
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    
    '------------------------------------------------
    With objExcel.ActiveSheet
        '.Range(.Cells(1, 1), .Cells(10, 3)).Borders.LineStyle = xlContinuous
        .Range(.Cells(9, 1), .Cells(9, 20)).Borders.LineStyle = xlContinuous
       
        .columns("A").ColumnWidth = 15
        .columns("B").ColumnWidth = 30
        .columns("C").ColumnWidth = 15
        .columns("D").ColumnWidth = 15
        .columns("E").ColumnWidth = 25
   
    
End With
    'cabecera
    mytablex.Open "select * from empresa where codigo='01'", cn, adOpenStatic, adLockOptimistic
    If mytablex.RecordCount > 0 Then
    objExcel.ActiveSheet.Cells(2, 1) = "'" & mytablex.Fields("nombre")
    End If
    mytablex.Close
    objExcel.ActiveSheet.Cells(2, 5) = "'" & Format(Now, "dd/mm/yyyy")
    objExcel.ActiveSheet.Cells(3, 2) = "'Reporte de Asiento en Soles(S/.)"
    
    
   
    
    '------------------------------------------------
    objExcel.ActiveSheet.Cells(4, 1) = "'Numero"
    objExcel.ActiveSheet.Cells(5, 1) = "'Fecha"
    objExcel.ActiveSheet.Cells(6, 1) = "'Referencia"
    objExcel.ActiveSheet.Cells(7, 1) = "'Comprobante"
    objExcel.ActiveSheet.Cells(8, 1) = "'Fuente"
    
    objExcel.ActiveSheet.Cells(4, 2) = "'" & txasiento.Fields("cod_asien")
    objExcel.ActiveSheet.Cells(5, 2) = "'" & txasiento.Fields("fecha_asi")
    objExcel.ActiveSheet.Cells(6, 2) = "'" & txasiento.Fields("referencia")
    objExcel.ActiveSheet.Cells(7, 2) = "'" & txasiento.Fields("comproba")
    objExcel.ActiveSheet.Cells(7, 3) = "'" & txasiento.Fields("tipo")
    objExcel.ActiveSheet.Cells(8, 2) = "'" & txasiento.Fields("Fuente")
    
    
    objExcel.ActiveSheet.Cells(9, 1) = "'Cuenta"
    objExcel.ActiveSheet.Cells(9, 2) = "'NombreCta"
    objExcel.ActiveSheet.Cells(9, 3) = "'Debito"
    objExcel.ActiveSheet.Cells(9, 4) = "'Credito"
    objExcel.ActiveSheet.Cells(9, 5) = "'Motivo"
    objExcel.ActiveSheet.Cells(9, 6) = "'Comproba"
    objExcel.ActiveSheet.Cells(9, 7) = "'Ruc"
    
    
    '------------------------------------------------
v = 10
H = 1
    xdebito = 0
    xcredito = 0
    
    Do
         If txasiento.EOF Then Exit Do
         If Trim("" & txasiento.Fields("cod_asien")) <> buf Then Exit Do
   objExcel.ActiveSheet.Cells(v, 1) = "'" & txasiento.Fields("cuenta")
   objExcel.ActiveSheet.Cells(v, 2) = "'" & txasiento.Fields("Descripcio")
   If "" & txasiento.Fields("tipo_cta") = "D" Then
      objExcel.ActiveSheet.Cells(v, 3) = "" & txasiento.Fields("cantidad")
      xdebito = xdebito + Val("" & txasiento.Fields("cantidad"))
   End If
   If "" & txasiento.Fields("tipo_cta") = "H" Then
      objExcel.ActiveSheet.Cells(v, 4) = "" & txasiento.Fields("cantidad")
      xcredito = xcredito + Val("" & txasiento.Fields("cantidad"))
   End If
   objExcel.ActiveSheet.Cells(v, 5) = "" & txasiento.Fields("motivo")
   objExcel.ActiveSheet.Cells(v, 6) = "" & txasiento.Fields("comproba")
   objExcel.ActiveSheet.Cells(v, 7) = "" & txasiento.Fields("Nro_ruc")
   
   
   v = v + 1
txasiento.MoveNext
Loop

objExcel.ActiveSheet.Cells(v, 3) = xdebito
objExcel.ActiveSheet.Cells(v, 4) = xcredito
Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
Exit Sub
cmd45612_err:
MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
Exit Sub

End Sub



Private Sub Label18_Click()
   Combo2.Clear
   Combo2.AddItem "Descripcio"
   Combo2.AddItem "docta"
   Combo2.ListIndex = 0
   opcion1 = "5"
   Frame3.Visible = True
   Frame3.Enabled = True
   cadena = ""
   cadena.SetFocus
   Command3_Click
   Exit Sub

End Sub

Private Sub libro_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Len(libro) > 0 Then
   nlibro = ""
   found = busca_libro("" & libro)
   If found = 0 Then
      libro.SetFocus
      Exit Sub
   End If
End If
fuente.SetFocus
End Sub

Private Sub libro_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   Combo2.Clear
   Combo2.AddItem "Descripcio"
   Combo2.AddItem "LibroAuxliar"
   Combo2.ListIndex = 0
   opcion1 = "500"
   Frame3.Visible = True
   Frame3.Enabled = True
   cadena = ""
   cadena.SetFocus
   Command3_Click
   Exit Sub

End If

End Sub

Private Sub referencia_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
tipo.SetFocus
End Sub
Function busca_fuente()
Dim mytablex As New ADODB.Recordset
nfuente = ""
mytablex.Open "select * from fuente where fuente='" & fuente & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   busca_fuente = 1
   nfuente = "" & mytablex.Fields("descripcio")
End If
mytablex.Close
End Function
Function busca_cuenta(buf As String)
Dim mytablex As New ADODB.Recordset
mytablex.Open "select * from cuentas where codcta='" & buf & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
      busca_cuenta = 1
      DBGrid2.columns(6) = "" & mytablex.Fields("descripcio")
End If
mytablex.Close
End Function
Function busca_libro(buf As String)
Dim mytablex As New ADODB.Recordset
mytablex.Open "select * from libroauxiliar where libroauxiliar='" & buf & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   busca_libro = 1
   nlibro = "" & mytablex.Fields("descripcio")
End If
mytablex.Close
End Function
Function obliga_ruc()
Dim mytablex As New ADODB.Recordset
mytablex.Open "select * from cuentas where codcta='" & "" & Data1.Recordset.Fields("cuenta") & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   If "" & mytablex.Fields("flag_ruc") = "S" Then
      If Len(Trim("" & Data1.Recordset.Fields("ruc"))) = 0 Then
         obliga_ruc = 1
      End If
   End If
End If
mytablex.Close
End Function
Function obliga_ccosto()
Dim mytablex As New ADODB.Recordset
mytablex.Open "select * from ccosto where ccosto='" & "" & Data1.Recordset.Fields("ccosto") & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   If "" & mytablex.Fields("ccosto") = "S" Then
      If Len(Trim("" & Data1.Recordset.Fields("ccosto"))) = 0 Then
         obliga_ccosto = 1
      End If
   End If
End If
mytablex.Close
End Function

Function busca_tercero(buf As String)
Dim mytablex As New ADODB.Recordset
mytablex.Open "select * from clientes where codigo='" & buf & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   nruc = "" & mytablex.Fields("nombre")
   busca_tercero = 1
End If
mytablex.Close
End Function
Sub carga_documento(buf As String)
On Error GoTo cmd7878_err
Dim mytablex As New ADODB.Recordset
mytablex.Open "select * from asientos where cod_asien=" & buf & "", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount = 0 Then
   mytablex.Close
   Exit Sub
End If
Do
If mytablex.EOF Then Exit Do
Data1.Recordset.AddNew
Data1.Recordset.Fields("cuenta") = "" & mytablex.Fields("cuenta")
Data1.Recordset.Fields("motivo") = "" & mytablex.Fields("motivo")
Data1.Recordset.Fields("debito") = Val("" & mytablex.Fields("debito"))
Data1.Recordset.Fields("credito") = Val("" & mytablex.Fields("credito"))
Data1.Recordset.Fields("nro_ruc") = Trim("" & mytablex.Fields("nro_ruc"))
Data1.Recordset.Fields("descripcio") = Trim("" & mytablex.Fields("descripcio"))
Data1.Recordset.Fields("ccosto") = Trim("" & mytablex.Fields("ccosto"))



'Data1.Recordset.Fields("cod_asien") = Val("" & mytablex.Fields("cod_asien"))
'Data1.Recordset.Fields("fecha_asi") = "" & mytablex.Fields("fecha_asi")
'Data1.Recordset.Fields("nro_seq") = "" & mytablex.Fields("nro_seq")
'Data1.Recordset.Fields("cuenta") = "" & mytablex.Fields("cuenta")
'Data1.Recordset.Fields("debito") = Val("" & mytablex.Fields("debito"))
'Data1.Recordset.Fields("credito") = Val("" & mytablex.Fields("credito"))
'Data1.Recordset.Fields("ccosto") = "" & mytablex.Fields("ccosto")
'Data1.Recordset.Fields("tipo_cta") = "" & mytablex.Fields("tipo_cta")
'Data1.Recordset.Fields("cod_libro") = "" & mytablex.Fields("cod_libro")
'Data1.Recordset.Fields("motivo") = "" & mytablex.Fields("motivo")
'Data1.Recordset.Fields("referencia") = "" & mytablex.Fields("referencia")
'Data1.Recordset.Fields("comproba") = "" & mytablex.Fields("comproba")
'Data1.Recordset.Fields("fuente") = "" & mytablex.Fields("fuente")
'Data1.Recordset.Fields("nro_ruc") = "" & mytablex.Fields("nro_ruc")
'Data1.Recordset.Fields("vrbase") = Val("" & mytablex.Fields("vrbase"))
'Data1.Recordset.Fields("cod_cdec") = Val("" & mytablex.Fields("cod_cdec"))
'Data1.Recordset.Fields("descripcio") = "" & mytablex.Fields("descripcio")
'Data1.Recordset.Fields("Tipo") = "" & mytablex.Fields("tipo")

'Data1.Recordset.Fields("nombre") = "" & mytablex.Fields("nombre")
Data1.Recordset.Update
mytablex.MoveNext
Loop
mytablex.Close
Exit Sub
cmd7878_err:
MsgBox "Aviso en carga_documento " + error$, 48, "Aviso"
Exit Sub
End Sub

Private Sub tipo_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Len(tipo) > 0 Then
   ntipo = ""
   found = busca_tipo()
   If found = 0 Then
      tipo.SetFocus
      Exit Sub
   End If
End If
comproba.SetFocus

End Sub

Private Sub tipo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   Combo2.Clear
   Combo2.AddItem "Descripcio"
   Combo2.AddItem "Docta"
   Combo2.ListIndex = 0
   opcion1 = "5"
   Frame3.Visible = True
   Frame3.Enabled = True
   cadena = ""
   cadena.SetFocus
   Command3_Click
   Exit Sub
End If

End Sub

Private Sub tr666_Click()
If Frame3.Visible = True Then Exit Sub
If Frame2.Visible = True Then Exit Sub
reporgen.NAMETABLA = "asientos"
reporgen.Show 1

End Sub
Function busca_tipo()
Dim mytablex As New ADODB.Recordset
mytablex.Open "select * from docta where docta='" & tipo & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   ntipo = "" & mytablex.Fields("descripcio")
   busca_tipo = 1
End If
mytablex.Close
End Function
Function busca_ccosto(buf As String)
Dim mytablex As New ADODB.Recordset
mytablex.Open "select * from ccosto where ccosto='" & buf & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   nccosto = "" & mytablex.Fields("descripcio")
   busca_ccosto = 1
End If
mytablex.Close
End Function

Function busca_periodo()
Dim mytablex As New ADODB.Recordset

mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   If IsDate("" & mytablex.Fields("periodocontable")) Then
      If Year(mytablex.Fields("periodocontable")) = Year(fecha) And Month(mytablex.Fields("periodocontable")) = Month(fecha) Then
         busca_periodo = 1
      End If
   End If
   
End If
mytablex.Close
End Function
Function validar()
Dim found As Integer
   found = busca_tipo()
   If found = 0 Then
      MsgBox "No existe Tipo ", 48, "Aviso"
      tipo.SetFocus
      Exit Function
   End If
   found = busca_libro("" & libro)
   If found = 0 Then
      MsgBox "No existe Libro ", 48, "Aviso"
      libro.SetFocus
      Exit Function
   End If
   found = busca_fuente()
   If found = 0 Then
   MsgBox "No existe Fuente ", 48, "Aviso"
   fuente.SetFocus
   Exit Function
   End If
   validar = 1


End Function
Sub habilita_cabeza(sw As Integer)
Dim xsw
If sw = 0 Then
   xsw = True
End If
If sw = 1 Then
   xsw = False
End If
cod_asien.Enabled = xsw
fecha.Enabled = xsw
referencia.Enabled = xsw
tipo.Enabled = xsw
comproba.Enabled = xsw
libro.Enabled = xsw
fuente.Enabled = xsw
image1.Enabled = xsw
image2.Enabled = xsw
image3.Enabled = xsw
Command2.Enabled = xsw

End Sub
Sub combo_dpreasiento()
Dim mytablex As New ADODB.Recordset
mytablex.Open "select * from preasiento", cn, adOpenDynamic, adLockOptimistic
Combo1.Clear
Combo1.AddItem "%"
Do
If mytablex.EOF Then Exit Sub
Combo1.AddItem "" & mytablex.Fields("preasiento") & "|" & mytablex.Fields("descripcio")
mytablex.MoveNext
Loop
mytablex.Close
Combo1.ListIndex = 0
End Sub
Sub carga_dpreasiento()
Dim mytablex As New ADODB.Recordset
Dim found As Integer
ir_final
If Val(debe) > 0 Or Val(haber) > 0 Then
   MsgBox "Existen Datos Ingresados ", 48, "Aviso"
   Exit Sub
End If
mytablex.Open "select * from dpreasiento where preasiento='" & extra_loquesea(Combo1) & "' ", cn, adOpenStatic, adLockOptimistic
Do
If mytablex.EOF Then Exit Do
If Len(Trim("" & mytablex.Fields("cuenta"))) > 0 Then

'cod_asien = Trim("" & txasiento.Fields("cod_asien"))
fecha = Trim("" & mytablex.Fields("fecha_asi"))
referencia = Trim("" & mytablex.Fields("referencia"))
comproba = Trim("" & mytablex.Fields("comproba"))
fuente = Trim("" & mytablex.Fields("fuente"))
tipo = Trim("" & mytablex.Fields("tipo"))
libro = Trim("" & mytablex.Fields("cod_libro"))
librocorre = Trim("" & mytablex.Fields("corre_libro"))
found = busca_tipo()
found = busca_libro("" & libro)
found = busca_tipo()


Data1.Recordset.AddNew
Data1.Recordset.Fields("cuenta") = "" & mytablex.Fields("cuenta")
Data1.Recordset.Fields("motivo") = "" & mytablex.Fields("motivo")
Data1.Recordset.Fields("debito") = Val("" & mytablex.Fields("debito"))
Data1.Recordset.Fields("credito") = Val("" & mytablex.Fields("credito"))
Data1.Recordset.Fields("nro_ruc") = Trim("" & mytablex.Fields("nro_ruc"))
Data1.Recordset.Fields("descripcio") = Trim("" & mytablex.Fields("descripcio"))
Data1.Recordset.Fields("ccosto") = Trim("" & mytablex.Fields("ccosto"))
Data1.Recordset.Update
End If
mytablex.MoveNext
Loop
mytablex.Close
ir_final

End Sub
