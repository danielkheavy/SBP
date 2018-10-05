VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form explorap 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Explorador de Documentos"
   ClientHeight    =   9750
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   15375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9750
   ScaleWidth      =   15375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrmConsultaSunat 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CONSULTA_SUNAT"
      Height          =   3615
      Left            =   9000
      TabIndex        =   148
      Top             =   3240
      Visible         =   0   'False
      Width           =   7575
      Begin VB.TextBox txtacu 
         Height          =   375
         Left            =   4920
         MaxLength       =   11
         TabIndex        =   170
         Text            =   "Acu"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtRuc 
         Height          =   375
         Left            =   4920
         MaxLength       =   11
         TabIndex        =   169
         Text            =   "Ruc"
         Top             =   2400
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   4680
         TabIndex        =   167
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox TxtTipo 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2410
         Style           =   2  'Dropdown List
         TabIndex        =   166
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtEstado 
         Height          =   375
         Left            =   7080
         MaxLength       =   11
         TabIndex        =   165
         Top             =   2880
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TxtLocal 
         Height          =   375
         Left            =   2400
         MaxLength       =   11
         TabIndex        =   162
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtEstadoSunat 
         Height          =   375
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   160
         Top             =   2880
         Width           =   1935
      End
      Begin VB.TextBox TxtNumero 
         Height          =   375
         Left            =   2400
         MaxLength       =   11
         TabIndex        =   158
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox TxtFecha 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2400
         MaxLength       =   13
         TabIndex        =   153
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CommandButton Command17 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5520
         MaskColor       =   &H00FFFFFF&
         Picture         =   "explorap.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   152
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.CommandButton Command16 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5520
         MaskColor       =   &H00FFFFFF&
         Picture         =   "explorap.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   151
         Top             =   1440
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.TextBox TxtSerie 
         Height          =   375
         Left            =   2400
         MaxLength       =   11
         TabIndex        =   150
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox TxtTotal 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2400
         MaxLength       =   11
         TabIndex        =   149
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label49 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   163
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label46 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ESTADO DOCUMENTO >"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   161
         Top             =   2880
         Width           =   2220
      End
      Begin VB.Label Label45 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Número"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   159
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label44 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   157
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label50 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   156
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label48 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   155
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label47 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   154
         Top             =   2520
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   6435
      Left            =   -3120
      TabIndex        =   74
      Top             =   9600
      Visible         =   0   'False
      Width           =   14910
      Begin VB.TextBox buffer 
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
         Left            =   7380
         MaxLength       =   10
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   990
         Width           =   135
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H8000000D&
         Caption         =   "&Aceptar"
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
         Left            =   7620
         TabIndex        =   145
         TabStop         =   0   'False
         Top             =   975
         Width           =   1575
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
         Left            =   6045
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   285
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   5295
         Left            =   195
         TabIndex        =   78
         Top             =   960
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   9340
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   22
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
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Presione ENTER para continuar!!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   146
         Top             =   630
         Width           =   3510
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "PedidosParaGenerar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   8055
      Left            =   4800
      TabIndex        =   140
      Top             =   9480
      Visible         =   0   'False
      Width           =   10980
      Begin VB.CommandButton Command14 
         Caption         =   "Close"
         Height          =   615
         Left            =   9480
         TabIndex        =   144
         Top             =   240
         Width           =   1335
      End
      Begin VB.ListBox List2 
         Height          =   6495
         Left            =   8040
         TabIndex        =   143
         Top             =   240
         Width           =   3135
      End
      Begin VB.CommandButton Command13 
         Caption         =   "GenerarOrden"
         Height          =   615
         Left            =   8040
         TabIndex        =   142
         Top             =   240
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid dbgrid12 
         Height          =   7215
         Left            =   120
         TabIndex        =   141
         Top             =   330
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   12726
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
   Begin VB.Frame Frame8 
      Caption         =   "Emisiones"
      Height          =   5475
      Left            =   7440
      TabIndex        =   135
      Top             =   9600
      Visible         =   0   'False
      Width           =   8760
      Begin VB.CommandButton Command12 
         Caption         =   "Cerrrar"
         Height          =   735
         Left            =   13320
         TabIndex        =   137
         Top             =   360
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid dbgrid5 
         Height          =   7650
         Left            =   210
         TabIndex        =   136
         Top             =   285
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   13494
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
   Begin VB.Frame Frame7 
      BackColor       =   &H00808080&
      Caption         =   "Actualizaciones de Fechas"
      Height          =   3735
      Left            =   9720
      TabIndex        =   115
      Top             =   9120
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox clavefecha 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   121
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FFFF80&
         Caption         =   "Procesar"
         Height          =   735
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   119
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00FFFF80&
         Caption         =   "Cerrrar"
         Height          =   735
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox hfechai 
         Height          =   495
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   117
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clave de Proceso"
         Height          =   495
         Left            =   120
         TabIndex        =   122
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         Caption         =   "dd/mm/yyyy"
         Height          =   495
         Left            =   3600
         TabIndex        =   120
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cambiar Fecha Apertura"
         Height          =   495
         Left            =   120
         TabIndex        =   116
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Forma Pago"
      Height          =   6135
      Left            =   13080
      TabIndex        =   111
      Top             =   5280
      Visible         =   0   'False
      Width           =   11055
      Begin MSDataGridLib.DataGrid dbgrid33 
         Height          =   5535
         Left            =   240
         TabIndex        =   113
         Top             =   360
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   9763
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
         ColumnCount     =   10
         BeginProperty Column00 
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
         BeginProperty Column01 
            DataField       =   "Tipo"
            Caption         =   "Tipo"
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
         BeginProperty Column02 
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
         BeginProperty Column03 
            DataField       =   "NUmero"
            Caption         =   "Numero"
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
            DataField       =   "Fpago"
            Caption         =   "Fpago"
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
         BeginProperty Column06 
            DataField       =   "Moneda"
            Caption         =   "M"
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
            DataField       =   "Total"
            Caption         =   "Total"
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
            DataField       =   "Recibe"
            Caption         =   "Entrega"
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
            DataField       =   "Saldos"
            Caption         =   "Saldos"
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
               ColumnWidth     =   494.929
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   390.047
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   494.929
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   555.024
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   299.906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column09 
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Cerrrar"
         Height          =   735
         Left            =   9480
         TabIndex        =   112
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Uso de la Tienda-Tickets"
      Height          =   9735
      Left            =   -2160
      TabIndex        =   79
      Top             =   9480
      Visible         =   0   'False
      Width           =   14895
      Begin VB.Label oturno 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   4440
         TabIndex        =   95
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label32 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Turno"
         Height          =   615
         Left            =   2760
         TabIndex        =   94
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label ocaja 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   4440
         TabIndex        =   93
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label31 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Caja"
         Height          =   615
         Left            =   2760
         TabIndex        =   92
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label ocajero 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   4440
         TabIndex        =   91
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label30 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cajero"
         Height          =   615
         Left            =   2760
         TabIndex        =   90
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label ofechaf 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   4440
         TabIndex        =   89
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label29 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Final"
         Height          =   615
         Left            =   2760
         TabIndex        =   88
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label ofechai 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   4440
         TabIndex        =   87
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label28 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         Height          =   615
         Left            =   2760
         TabIndex        =   86
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Close"
         Height          =   735
         Left            =   8040
         TabIndex        =   83
         Top             =   3960
         Width           =   2415
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Copia Cierre Dia"
         Height          =   735
         Left            =   5520
         TabIndex        =   82
         Top             =   3960
         Width           =   2415
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuadre Parcial"
         Height          =   735
         Left            =   3000
         TabIndex        =   81
         Top             =   3960
         Width           =   2415
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reinpresion"
         Height          =   735
         Left            =   480
         TabIndex        =   80
         Top             =   3960
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Sumar"
      Height          =   615
      Left            =   12720
      TabIndex        =   69
      Top             =   8760
      Width           =   2175
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00808080&
      Caption         =   "Generacion de Documentos Normal"
      Height          =   7095
      Left            =   0
      TabIndex        =   58
      Top             =   1320
      Visible         =   0   'False
      Width           =   8415
      Begin VB.TextBox tipod 
         Height          =   375
         Left            =   5160
         TabIndex        =   168
         Text            =   "tipod"
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Grabar"
         Height          =   615
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   360
         Width           =   1335
      End
      Begin VB.ListBox List1 
         Height          =   3960
         Left            =   240
         TabIndex        =   73
         Top             =   2040
         Width           =   5655
      End
      Begin VB.TextBox gacu 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         MaxLength       =   11
         TabIndex        =   66
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Salir"
         Height          =   615
         Left            =   6885
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox gnumero 
         Height          =   375
         Left            =   1560
         MaxLength       =   11
         TabIndex        =   63
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox gserie 
         Height          =   375
         Left            =   1560
         MaxLength       =   11
         TabIndex        =   61
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox gtipo 
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Flag"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   67
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   64
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   62
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoDocumento"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   60
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Clave de Acceso"
      Height          =   5055
      Left            =   240
      TabIndex        =   51
      Top             =   8160
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3480
         MaskColor       =   &H00FFFFFF&
         Picture         =   "explorap.frx":0F5C
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   4080
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3480
         MaskColor       =   &H00FFFFFF&
         Picture         =   "explorap.frx":170A
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   3240
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.TextBox clave 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   1200
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1800
         Width           =   3015
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ingrese su clave para realizar esta Accion"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   600
         TabIndex        =   55
         Top             =   735
         Width           =   4215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Consultas-Condiciones"
      Height          =   5175
      Left            =   960
      TabIndex        =   11
      Top             =   8400
      Visible         =   0   'False
      Width           =   7695
      Begin VB.ComboBox placa 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   133
         Top             =   3240
         Width           =   1935
      End
      Begin VB.CheckBox saldoini 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SaldoInicial"
         Height          =   375
         Left            =   240
         TabIndex        =   123
         Top             =   3960
         Width           =   2295
      End
      Begin VB.ComboBox servicio 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   2880
         Width           =   1935
      End
      Begin VB.ComboBox moneda 
         Height          =   360
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox nombre 
         Height          =   375
         Left            =   2400
         MaxLength       =   11
         TabIndex        =   21
         Text            =   "%"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox serie 
         Height          =   375
         Left            =   2400
         MaxLength       =   11
         TabIndex        =   20
         Text            =   "%"
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5520
         MaskColor       =   &H00FFFFFF&
         Picture         =   "explorap.frx":1EB8
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1440
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.CommandButton cmdGrabar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5520
         MaskColor       =   &H00FFFFFF&
         Picture         =   "explorap.frx":2666
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.ComboBox estado 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox codigo 
         Height          =   375
         Left            =   2400
         MaxLength       =   13
         TabIndex        =   12
         Text            =   "%"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label41 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Proceso"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   134
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Servicio"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   71
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Moneda"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estado"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00808080&
      Height          =   1275
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   15315
      TabIndex        =   5
      Top             =   0
      Width           =   15375
      Begin VB.CheckBox chkMostrarSoloAnulados 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Por Dar de Baja"
         Height          =   375
         Left            =   1920
         TabIndex        =   147
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox ve 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13080
         MaxLength       =   1
         TabIndex        =   138
         Top             =   480
         Width           =   375
      End
      Begin VB.ComboBox ordenado 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   11280
         Style           =   2  'Dropdown List
         TabIndex        =   109
         Top             =   840
         Width           =   3615
      End
      Begin VB.TextBox numero 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13080
         MaxLength       =   11
         TabIndex        =   96
         Text            =   "%"
         Top             =   120
         Width           =   855
      End
      Begin VB.ComboBox turno 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   11280
         Style           =   2  'Dropdown List
         TabIndex        =   84
         Top             =   120
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   11280
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "No mostrar Tipo 5"
         Height          =   255
         Left            =   0
         TabIndex        =   49
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox cajero 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9120
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   840
         Width           =   1695
      End
      Begin VB.ComboBox vendedor 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6600
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   840
         Width           =   1695
      End
      Begin VB.ComboBox bodegaf 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox tipo 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9120
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   480
         Width           =   1695
      End
      Begin VB.ComboBox caja 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9120
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   120
         Width           =   1695
      End
      Begin VB.ComboBox bodega 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Filtrar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   14040
         MaskColor       =   &H00FFFFFF&
         Picture         =   "explorap.frx":2E14
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox fechaf 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   30
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox fechai 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   28
         Top             =   120
         Width           =   1095
      End
      Begin VB.ComboBox llocal1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdSort 
         BackColor       =   &H00C0C0C0&
         Height          =   855
         Left            =   1800
         MaskColor       =   &H00E0E0E0&
         Picture         =   "explorap.frx":35C2
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Consulta"
         Top             =   0
         Width           =   615
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
         Height          =   855
         Left            =   1200
         MaskColor       =   &H00E0E0E0&
         Picture         =   "explorap.frx":47D4
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Borrar registro"
         Top             =   0
         Width           =   615
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
         Height          =   855
         Left            =   2400
         MaskColor       =   &H00E0E0E0&
         Picture         =   "explorap.frx":59E6
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salir"
         Top             =   0
         Width           =   615
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
         Height          =   855
         Left            =   600
         MaskColor       =   &H00E0E0E0&
         Picture         =   "explorap.frx":6BF8
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Imprimir"
         Top             =   0
         Width           =   615
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
         Height          =   855
         Left            =   0
         Picture         =   "explorap.frx":7E0A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label42 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "V"
         Height          =   375
         Left            =   12480
         TabIndex        =   139
         Top             =   480
         Width           =   615
      End
      Begin VB.Label sginicio 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   15000
         TabIndex        =   114
         Top             =   120
         Width           =   105
      End
      Begin VB.Label Label37 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Orden"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   10800
         TabIndex        =   110
         Top             =   840
         Width           =   495
      End
      Begin VB.Label importacion 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   840
         TabIndex        =   99
         Top             =   945
         Width           =   255
      End
      Begin VB.Label Label33 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   12480
         TabIndex        =   97
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Turnoxx 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Turno"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   10800
         TabIndex        =   85
         Top             =   120
         Width           =   495
      End
      Begin VB.Label tinterno 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   13560
         TabIndex        =   68
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label34 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   10800
         TabIndex        =   57
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label26 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cajero"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   8280
         TabIndex        =   46
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vendedor"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5640
         TabIndex        =   44
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label25 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AlmaFin"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3120
         TabIndex        =   43
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoDoc"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   8280
         TabIndex        =   38
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Caja"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   8280
         TabIndex        =   36
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AlmaIni"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3120
         TabIndex        =   34
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Final"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5640
         TabIndex        =   31
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5640
         TabIndex        =   29
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3120
         TabIndex        =   27
         Top             =   120
         Width           =   615
      End
   End
   Begin MSDataGridLib.DataGrid dbgrid2 
      Height          =   6300
      Left            =   0
      TabIndex        =   50
      Top             =   1320
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   11113
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   19
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
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   46
      BeginProperty Column00 
         DataField       =   "dflag"
         Caption         =   "X"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Yausado"
         Caption         =   "A"
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
         DataField       =   "Estado"
         Caption         =   "E"
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
      BeginProperty Column04 
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
      BeginProperty Column05 
         DataField       =   "Serie"
         Caption         =   "Serie"
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
      BeginProperty Column07 
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
      BeginProperty Column08 
         DataField       =   "Fechae"
         Caption         =   "FechaE"
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
      BeginProperty Column09 
         DataField       =   "Hora"
         Caption         =   "Hora"
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
      BeginProperty Column10 
         DataField       =   "Tipoclie"
         Caption         =   "T"
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
      BeginProperty Column11 
         DataField       =   "Codigo"
         Caption         =   "Codigo"
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
      BeginProperty Column12 
         DataField       =   "Nombre"
         Caption         =   "Nombre"
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
      BeginProperty Column13 
         DataField       =   "Moneda"
         Caption         =   "M"
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
      BeginProperty Column14 
         DataField       =   "Total"
         Caption         =   "Total"
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
      BeginProperty Column15 
         DataField       =   "Acuenta"
         Caption         =   "Acuenta"
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
      BeginProperty Column16 
         DataField       =   "Adetotal"
         Caption         =   "Saldo"
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
      BeginProperty Column17 
         DataField       =   "bodega"
         Caption         =   "BodI"
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
      BeginProperty Column18 
         DataField       =   "bodegaf"
         Caption         =   "Bodf"
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
      BeginProperty Column19 
         DataField       =   "Localf"
         Caption         =   "LocF"
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
      BeginProperty Column20 
         DataField       =   "Nro_items"
         Caption         =   "Items"
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
      BeginProperty Column21 
         DataField       =   "Usuario"
         Caption         =   "Cajero"
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
      BeginProperty Column22 
         DataField       =   "Placa"
         Caption         =   "Proceso"
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
      BeginProperty Column23 
         DataField       =   "Caja"
         Caption         =   "Caja"
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
      BeginProperty Column24 
         DataField       =   "Turno"
         Caption         =   "T"
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
      BeginProperty Column25 
         DataField       =   "vendedor"
         Caption         =   "Vendedor"
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
      BeginProperty Column26 
         DataField       =   "Observa"
         Caption         =   "Observa"
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
      BeginProperty Column27 
         DataField       =   "Acu"
         Caption         =   "Acu"
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
      BeginProperty Column28 
         DataField       =   "Servicio"
         Caption         =   "S"
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
      BeginProperty Column29 
         DataField       =   "Local1"
         Caption         =   "Local1"
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
      BeginProperty Column30 
         DataField       =   "tipo1"
         Caption         =   "Tipo1"
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
      BeginProperty Column31 
         DataField       =   "Serie1"
         Caption         =   "Serie1"
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
      BeginProperty Column32 
         DataField       =   "numero1"
         Caption         =   "Numero1"
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
      BeginProperty Column33 
         DataField       =   "retipo1"
         Caption         =   "retipo1"
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
      BeginProperty Column34 
         DataField       =   "renumero3"
         Caption         =   "renumero3"
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
      BeginProperty Column35 
         DataField       =   "renumero1"
         Caption         =   "renumero1"
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
      BeginProperty Column36 
         DataField       =   "renumero2"
         Caption         =   "renumero2"
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
      BeginProperty Column37 
         DataField       =   "Neto"
         Caption         =   "Neto"
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
      BeginProperty Column38 
         DataField       =   "Descuento"
         Caption         =   "Descuento"
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
      BeginProperty Column39 
         DataField       =   "Subtotal"
         Caption         =   "Subtotal"
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
      BeginProperty Column40 
         DataField       =   "Impuesto"
         Caption         =   "Impuesto"
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
      BeginProperty Column41 
         DataField       =   "Tipoimp"
         Caption         =   "Tipoimp"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column42 
         DataField       =   "Acu1"
         Caption         =   "Acu1"
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
      BeginProperty Column43 
         DataField       =   "Acu"
         Caption         =   "Acu"
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
      BeginProperty Column44 
         DataField       =   "Fechae"
         Caption         =   "FechaE"
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
      BeginProperty Column45 
         DataField       =   "Hora"
         Caption         =   "Hora"
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
         MarqueeStyle    =   2
         BeginProperty Column00 
            ColumnWidth     =   180.283
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   180.283
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   209.764
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   360
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   404.787
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   450.142
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   945.071
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   180.283
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   2009.764
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   225.071
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   404.787
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   390.047
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   450.142
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   390.047
         EndProperty
         BeginProperty Column24 
            ColumnWidth     =   195.024
         EndProperty
         BeginProperty Column25 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column26 
         EndProperty
         BeginProperty Column27 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column28 
            ColumnWidth     =   345.26
         EndProperty
         BeginProperty Column29 
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column30 
            ColumnWidth     =   569.764
         EndProperty
         BeginProperty Column31 
            ColumnWidth     =   615.118
         EndProperty
         BeginProperty Column32 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column33 
         EndProperty
         BeginProperty Column34 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column35 
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column36 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column37 
         EndProperty
         BeginProperty Column38 
         EndProperty
         BeginProperty Column39 
         EndProperty
         BeginProperty Column40 
         EndProperty
         BeginProperty Column41 
         EndProperty
         BeginProperty Column42 
         EndProperty
         BeginProperty Column43 
         EndProperty
         BeginProperty Column44 
         EndProperty
         BeginProperty Column45 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label51 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F2 - Ver Detalle Sunat"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   164
      Top             =   7630
      Width           =   2220
   End
   Begin VB.Label totald 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   13560
      TabIndex        =   132
      Top             =   8400
      Width           =   1695
   End
   Begin VB.Label impuestod 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   10920
      TabIndex        =   131
      Top             =   8400
      Width           =   1455
   End
   Begin VB.Label subtotald 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   9480
      TabIndex        =   130
      Top             =   8400
      Width           =   1455
   End
   Begin VB.Label netod 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6720
      TabIndex        =   129
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Label serviciod 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   8040
      TabIndex        =   128
      Top             =   8400
      Width           =   1455
   End
   Begin VB.Label percepciond 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   12360
      TabIndex        =   127
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Label percepcions 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   12360
      TabIndex        =   126
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Label servicios 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   8040
      TabIndex        =   125
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Label netos 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6720
      TabIndex        =   124
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Label Label36 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NoGravados"
      Height          =   375
      Left            =   1920
      TabIndex        =   108
      Top             =   8760
      Width           =   1095
   End
   Begin VB.Label Label35 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Anulados"
      Height          =   375
      Left            =   2280
      TabIndex        =   107
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ventas"
      Height          =   375
      Left            =   2280
      TabIndex        =   106
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Label lvvendido 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3360
      TabIndex        =   105
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Label lvanulado 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3360
      TabIndex        =   104
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Label lvnogravado 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3000
      TabIndex        =   103
      Top             =   8760
      Width           =   1215
   End
   Begin VB.Label lnogravado 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4200
      TabIndex        =   102
      Top             =   8760
      Width           =   1455
   End
   Begin VB.Label lanulado 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4560
      TabIndex        =   101
      Top             =   8400
      Width           =   1455
   End
   Begin VB.Label lvendido 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4560
      TabIndex        =   100
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Label flag_estado 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   960
      TabIndex        =   72
      Top             =   9720
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label zooma 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   13800
      TabIndex        =   48
      Top             =   8280
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label YacaRGA 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   13800
      TabIndex        =   41
      Top             =   8040
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label nbodega1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   600
      TabIndex        =   40
      Top             =   8400
      Width           =   1695
   End
   Begin VB.Label nbodega 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   600
      TabIndex        =   39
      Top             =   8040
      Width           =   1695
   End
   Begin VB.Label tipoclie 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   360
      TabIndex        =   25
      Top             =   6960
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "US$."
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6000
      TabIndex        =   17
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label subtotals 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   9480
      TabIndex        =   4
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Label impuestos 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   10920
      TabIndex        =   3
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Label acu 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   13320
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label totals 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   13560
      TabIndex        =   1
      Top             =   8040
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S/."
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   8040
      Width           =   735
   End
   Begin VB.Menu djku232 
      Caption         =   "&Nuevo"
   End
   Begin VB.Menu agt62323 
      Caption         =   "&Borrar"
   End
   Begin VB.Menu mio8923 
      Caption         =   "&Modifica"
   End
   Begin VB.Menu anulier 
      Caption         =   "&Anular"
   End
   Begin VB.Menu dkiw232 
      Caption         =   "&Imprimir"
      Begin VB.Menu dkifor 
         Caption         =   "&1.FormatoDefinido"
      End
      Begin VB.Menu dkiewre 
         Caption         =   "&2.Reporteador"
      End
      Begin VB.Menu dl89er 
         Caption         =   "&3.Excell Impresion Total"
      End
      Begin VB.Menu dki889343 
         Caption         =   "&4.Excell Impresion solo seleccionado"
      End
      Begin VB.Menu impso02 
         Caption         =   "&5.Excell Impresion solo Documentos"
      End
      Begin VB.Menu fk8844 
         Caption         =   "&6.Reinpresion de Tickets"
      End
   End
   Begin VB.Menu mit56232 
      Caption         =   "&Zoom"
   End
   Begin VB.Menu fdl89234 
      Caption         =   "&Validar"
   End
   Begin VB.Menu djbu232 
      Caption         =   "&Consulta"
   End
   Begin VB.Menu fk4844 
      Caption         =   "&Generar"
   End
   Begin VB.Menu Flo881 
      Caption         =   "&Fpago"
   End
   Begin VB.Menu dki844 
      Caption         =   "&OrdenTrabajo"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu dlo8923 
      Caption         =   "&Procesos"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu dj8844 
         Caption         =   "&1.CambiarFecha"
      End
   End
   Begin VB.Menu nmur41 
      Caption         =   "&Emisiones"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu GTTR 
      Caption         =   "&PROCC"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu ju7881 
      Caption         =   "&Enviar"
   End
   Begin VB.Menu DarBaja 
      Caption         =   "&DarBaja"
      Visible         =   0   'False
   End
   Begin VB.Menu DetalleSunat 
      Caption         =   "&DetalleSunat"
      Visible         =   0   'False
   End
   Begin VB.Menu ldo33 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "explorap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' integracion proyecto paty
Dim mytablex        As New ADODB.Recordset

Dim my_nombre_local As String

Dim mio

Dim my_esunat

' integracion proyecto paty

Dim rexplorap As New ADODB.Recordset

Private Sub agt62323_Click()

    Dim buf1 As String

    On Error GoTo cmd6_err

    'If Frame4.Visible = True Then Exit Sub
    If Frame8.Visible = True Then Exit Sub

    If Frame7.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If "" & rexplorap.Fields("estado") <> "0" Then
        MsgBox "Para Borrar el documento debe estar en estado=0", 48, "Aviso"
        Exit Sub

    End If

    If MsgBox("Desea Borrar Documento", 1, "Aviso") <> 1 Then Exit Sub
    'MsgBox cgusuario
    buf1 = " and acu='" & "" & rexplorap.Fields("acu") & "'"
    cn.Execute "DELETE FROM  " & dgusuariog & "   where  local='" & "" & rexplorap.Fields("local") & "' and tipo='" & "" & rexplorap.Fields("tipo") & "' and serie='" & "" & rexplorap.Fields("serie") & "' and  numero='" & "" & rexplorap.Fields("numero") & "'" & buf1
    cn.Execute "DELETE FROM  fpagov   where  local='" & "" & rexplorap.Fields("local") & "' and tipo='" & "" & rexplorap.Fields("tipo") & "' and serie='" & "" & rexplorap.Fields("serie") & "' and  numero='" & "" & rexplorap.Fields("numero") & "'" & buf1
    cn.Execute "DELETE FROM  " & cgusuario & "   where  local='" & "" & rexplorap.Fields("local") & "' and tipo='" & "" & rexplorap.Fields("tipo") & "' and serie='" & "" & rexplorap.Fields("serie") & "' and  numero='" & "" & rexplorap.Fields("numero") & "'" & buf1
    'cn.Execute "DELETE FROM  facturagasto   where   tipo='" & "" & rexplorap.Fields("tipo") & "' and serie='" & "" & rexplorap.Fields("serie") & "' and  numero='" & "" & rexplorap.Fields("numero") & "'" & buf1

    If ACU = "3" Then
        cn.Execute "DELETE FROM  serviciotecnico   where  local='" & "" & rexplorap.Fields("local") & "' and tipo='" & "" & rexplorap.Fields("tipo") & "' and serie='" & "" & rexplorap.Fields("serie") & "' and  numero='" & "" & rexplorap.Fields("numero") & "'"

    End If

    MsgBox "Ok,Documento Borrado", 24, "Aviso"
    sql_cabeza
    Exit Sub
cmd6_err:
    MsgBox "Aviso en Borrar " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub anulier_Click()

    Dim buf1 As String

    Dim buf  As String

    Dim Msg  As String

    On Error GoTo cmd8_err

    'If Frame4.Visible = True Then Exit Sub
    If Frame8.Visible = True Then Exit Sub

    If Frame7.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    Msg = "Ojo.. Esta opcion de anular permite poner el documento en modo de anulacion ,luego de realizacion" + Chr$(10) + Chr$(13)
    Msg = Msg + "No puede ya reversar.... " + Chr$(10) + Chr$(13)

    If MsgBox(Msg, 1, "Aviso") <> 1 Then Exit Sub
    If "" & rexplorap.Fields("estado") = "2" Then
        MsgBox "Para anular el documento debe estar en estado=0 or estado=1", 48, "Aviso"
        Exit Sub

    End If

    If MsgBox("Desea Anular Documento,Quedara inmodificable ", 1, "Aviso") <> 1 Then Exit Sub
    buf = "1"

    If "" & rexplorap.Fields("estado") = "1" Then
        buf = "0"

    End If

    'Data2.Recordset.Edit
    'Data2.Recordset.Fields("estado") = buf
    'Data2.Recordset.Update
    'MsgBox cgusuario
    buf1 = " and acu='" & "" & rexplorap.Fields("acu") & "'"

    If importacion = "IMPORTACION" Then
        cn.Execute ("update  gastofactura set estado='" & buf & "'  where  local='" & "" & rexplorap.Fields("local") & "' and tipo='" & "" & rexplorap.Fields("tipo") & "' and serie='" & "" & rexplorap.Fields("serie") & "' and  numero='" & "" & rexplorap.Fields("numero") & "'" & buf1)

    End If

    cn.Execute ("update  " & dgusuariog & " set estado='" & buf & "'  where  local='" & "" & rexplorap.Fields("local") & "' and tipo='" & "" & rexplorap.Fields("tipo") & "' and serie='" & "" & rexplorap.Fields("serie") & "' and  numero='" & "" & rexplorap.Fields("numero") & "'" & buf1)
    cn.Execute ("update  fpagov  set estado='" & buf & "'  where  local='" & "" & rexplorap.Fields("local") & "' and tipo='" & "" & rexplorap.Fields("tipo") & "' and serie='" & "" & rexplorap.Fields("serie") & "' and  numero='" & "" & rexplorap.Fields("numero") & "'" & buf1)
    cn.Execute ("update  " & cgusuario & " set estado='" & buf & "'  where  local='" & "" & rexplorap.Fields("local") & "' and tipo='" & "" & rexplorap.Fields("tipo") & "' and serie='" & "" & rexplorap.Fields("serie") & "' and  numero='" & "" & rexplorap.Fields("numero") & "'" & buf1)
    MsgBox "Ok,Documento Anulado", 24, "Aviso"
    sql_cabeza
    Exit Sub
cmd8_err:
    MsgBox "Aviso en anular documento " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub buffer_DblClick()
    Command1_Click

End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        ldo33_Click
        Exit Sub

    End If

    Command1_Click

End Sub

Private Sub clave_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(clave) = 0 Then
        clave.SetFocus

    End If

    Command4_Click

End Sub

Private Sub cmdAddEntry_Click()
    djku232_Click

End Sub

Private Sub cmdExit_Click()
    ldo33_Click

End Sub

Private Sub cmdGrabar_Click()
    sql_cabeza

End Sub

Private Sub cmdPrint_Click()
    dkifor_Click

End Sub

Private Sub cmdSort_Click()
    djbu232_Click

End Sub

Private Sub Codigo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_codigo

    End If

End Sub

Private Sub Command1_Click()
    ejecuta 1

End Sub

Sub ejecuta(sw As Integer)

    Dim buf       As String

    Dim rconsulta As New ADODB.Recordset

    Dim cad       As String

    If opcion1 = "1" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo from clientes "
        Else
            buf = "select Nombre,Codigo from clientes where " & Combo1 & " like '" & buffer & "%'"

        End If

    End If

    If opcion1 = "6100" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo from tlocal "
        Else
            buf = "select Nombre,Codigo from tlocal where " & Combo1 & " like '" & buffer & "%'"

        End If

    End If

    If opcion1 = "2" Then
        buf = "select Producto,Descripcio,Unidad as Und,Factor as Fac,Precio,Cantidad as Cant,Total,Local,Deslipo as Dscto from  " & dgusuariog & " where local='" & "" & rexplorap.Fields("local") & "' and serie='" & "" & rexplorap.Fields("serie") & "' and numero='" & "" & rexplorap.Fields("numero") & "'"

    End If

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open buf, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True And rconsulta.BOF = True Then
        rconsulta.Close
        buffer.SetFocus
        Exit Sub

    End If

    Set dbGrid1.DataSource = rconsulta

    If opcion1 = "1" Or opcion1 = "6100" Then
        dbGrid1.columns(0).Width = 4000
        dbGrid1.columns(1).Width = 2000

    End If

    If opcion1 = "2" Then
        dbGrid1.columns(0).Width = 1500
        dbGrid1.columns(1).Width = 5000
        dbGrid1.columns(2).Width = 900
        dbGrid1.columns(3).Width = 900
        dbGrid1.columns(4).Width = 900
        dbGrid1.columns(5).Width = 900
        dbGrid1.columns(6).Width = 1500
        dbGrid1.columns(7).Width = 900
        dbGrid1.columns(8).Width = 700

    End If
   
    If sw = 1 Then
        dbGrid1.SetFocus

    End If

End Sub

Private Sub Command10_Click()
    Frame7.Visible = False

End Sub

Private Sub Command11_Click()

    Dim vr

    If Not IsDate(hfechai) Then
        MsgBox "Fecha no valido ", 48, "Aviso"
        hfechai = ""
        hfechai.SetFocus
        Exit Sub

    End If

    If Len(Trim(clavefecha)) = 0 Then
        clavefecha.SetFocus
        Exit Sub

    End If

    If Trim(clavefecha) <> "CAMBIAR" Then
        MsgBox "Solo personal autorizado ", 48, "Aviso"
        clavefecha.SetFocus
        Exit Sub

    End If

    'rexplorap.Requery
    Do

        If rexplorap.EOF Then Exit Do
        vr = DoEvents()
        rexplorap.Fields("fecha") = Format(hfechai, "dd/mm/yyyy")
        rexplorap.Update
        cn.Execute ("update detalle set fecha='" & Format(hfechai, "dd/mm/yyyy") & "' where local='" & "" & rexplorap.Fields("local") & "' and tipo='" & "" & rexplorap.Fields("tipo") & "' and serie='" & "" & rexplorap.Fields("serie") & "' and numero='" & "" & rexplorap.Fields("numero") & "'")
        cn.Execute ("update fpagov set fecha='" & Format(hfechai, "dd/mm/yyyy") & "' where local='" & "" & rexplorap.Fields("local") & "' and tipo='" & "" & rexplorap.Fields("tipo") & "' and serie='" & "" & rexplorap.Fields("serie") & "' and numero='" & "" & rexplorap.Fields("numero") & "'")
        rexplorap.MoveNext
    Loop
    rexplorap.Requery
    MsgBox "Proceso Terminado ", 48, "Aviso"
    Frame7.Visible = False

End Sub

Private Sub Command12_Click()
    ldo33_Click

End Sub

Private Sub Command14_Click()
    Frame9.Visible = False
    sql_cabeza

End Sub

Private Sub Command15_Click()

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    If opcion1 = "1" Then
        codigo = dbGrid1.columns(1)
        Frame1.Visible = False
        Frame1.Enabled = False
        codigo.SetFocus

    End If

    If opcion1 = "6100" Then
        mytablex.Open "SELECT * FROM userlocal where codigo='" & gusuario & "' and local='" & Trim(dbGrid1.columns(1)) & "'", cn, adOpenKeyset, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.Close
            MsgBox "Usuario No autorizado,utilizar este local ", 48, "Aviso"
            Exit Sub

        End If

        mytablex.Close
   
        buf = Trim(dbGrid1.columns(1))
        Frame1.Visible = False
        Frame1.Enabled = False
        'buscar almacen que pertenece----
        'mytablex.Open "SELECT * FROM bodega where local='" & buf & "'", cn, adOpenKeyset, adLockOptimistic
        'If mytablex.RecordCount > 0 Then
        'tfactura.bodega = Trim("" & mytablex.Fields("codigo"))
        'End If
        'mytablex.Close
        menu_nuevo buf

        'codigo.SetFocus
    End If

End Sub

Private Sub Command16_Click()
    ' Testing Proyecto Facturacion Electronica 05/04/2018
    FrmConsultaSunat.Visible = False

    ' Testing Proyecto Facturacion Electronica 05/04/2018
End Sub

' Testing Proyecto Facturacion Electronica 05/04/2018
Private Sub Command17_Click()
    VerDetalleSunat

End Sub

' Testing Proyecto Facturacion Electronica 05/04/2018
 
Private Sub Command2_Click()
    ldo33_Click

End Sub

Private Sub Command3_Click()
    ldo33_Click

End Sub

Private Sub Command4_Click()

    On Error GoTo cmd7_err

    Dim found As Integer

    Dim buf   As String

    If Len(clave) = 0 Then
        clave.SetFocus
        Exit Sub

    End If

    found = valida_clave("" & clave)

    If found = 0 Then
        MsgBox "Clave no valida para realizar este proceso ", 48, "Aviso"
        clave = ""
        clave.SetFocus
        Exit Sub

    End If

    'MsgBox ""
    If Frame2.Caption = "DESMARCA" Then
        If MsgBox("Desea Desmarca el Documento", 1, "Aviso") <> 1 Then Exit Sub
        If Trim("" & rexplorap.Fields("acu")) = "A" Or Trim("" & rexplorap.Fields("acu")) = "B" Or Trim("" & rexplorap.Fields("acu")) = "C" Or Trim("" & rexplorap.Fields("acu")) = "D" Or Trim("" & rexplorap.Fields("acu")) = "G" Then  'ventas
            buf = "cuentacd"
            'MsgBox ""
            found = verificar_recibo(buf, Trim(rexplorap.Fields("local")), Trim(rexplorap.Fields("tipo")), Trim(rexplorap.Fields("serie")), Trim(rexplorap.Fields("numero")))

            If found = 1 Then
                MsgBox "Ya existe recibo ", 48, "Aviso"
                Exit Sub

            End If

            'MsgBox ""
        End If

        If Trim("" & rexplorap.Fields("acu")) = "J" Or Trim("" & rexplorap.Fields("acu")) = "K" Or Trim("" & rexplorap.Fields("acu")) = "L" Or Trim("" & rexplorap.Fields("acu")) = "M" Or Trim("" & rexplorap.Fields("acu")) = "P" Then  'ventas
            buf = "cuentaPd"
            'MsgBox ""
            found = verificar_recibo(buf, Trim(rexplorap.Fields("local")), Trim(rexplorap.Fields("tipo")), Trim(rexplorap.Fields("serie")), Trim(rexplorap.Fields("numero")))

            If found = 1 Then
                MsgBox "Ya existe recibo ", 48, "Aviso"
                Exit Sub

            End If

        End If

        'MsgBox ""
        desmarca_documento

    End If

    Frame2.Visible = False
    Exit Sub
cmd7_err:
    MsgBox "Seleccione un dato " + error$, 48, "Aviso"
    Frame2.Visible = False
    Exit Sub

End Sub

Private Sub Command5_Click()
    sql_cabeza

End Sub

Private Sub Command6_Click()
    Frame6.Visible = False

End Sub

Private Sub Command7_Click()
    Frame4.Visible = False

End Sub

Private Sub Command9_Click()

    Dim buf      As String

    Dim bufca    As String

    Dim I        As Integer

    Dim sw       As Integer

    Dim sdx      As Double

    Dim bufde    As String

    Dim mytablez As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim mytablex As New ADODB.Recordset

    If gtipo = "%" Then Exit Sub
    If Len(gserie) = 0 Then
        gserie = ""
        Exit Sub

    End If

    If Len(gnumero) = 0 Then
        gnumero = ""
        Exit Sub

    End If

    If Len(gacu) = 0 Then
        gnumero = ""
        Exit Sub

    End If

    ' Testing Proyecto Facturacion Electronica Nota de Credito 06/2018
    'Actualizado estado de Comprobante Electronico

    If gacu = "E" Or ACU = "F" Then '  SI SE GENERA NC O ND
        VerDetalleSunat
        FrmConsultaSunat.Visible = False

    End If

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
    Dim valor As Boolean

    If gacu = "E" Or gacu = "F" Then   ' NOTA DE CREDITO Ó DÉBITO
        Call Busca_comprobanteRelacionado_sunat(rexplorap.Fields("local"), rexplorap.Fields("SERIE"), rexplorap.Fields("numero"), rexplorap.Fields("tipo"), valor)

        If valor = False Then
            Exit Sub

        End If

    End If

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

    Select Case Trim(gacu)

        Case "1", "A", "B", "C", "D", "E", "G", "F" 'VENTAS
            bufca = "factura"
            bufde = "detalle"

        Case "H"  'COTIZACION
            bufca = "ccotizav"
            bufde = "dcotizav"

        Case "I"  'PEDIDO
            bufca = "cpedidov"
            bufde = "dpedidov"

        Case "J", "K", "L", "M", "P", "N", "O"  'COMPRAS
            bufca = "factura"
            bufde = "detalle"

        Case "Q"  'PEDIDO COMPRA
            bufca = "cpedidoc"
            bufde = "dpedidoc"

        Case "R"  'ORDEN COMPRA
            bufca = "cordenc"
            bufde = "dordenc"

        Case "S", "T"
            bufca = "factura"
            bufde = "detalle"

        Case "Z"
            bufca = "ctraslad"
            bufde = "dtraslad"

        Case Else
            Exit Sub

    End Select

    'cabecera
    sw = 0

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

    '01/08/2018 Testing Facturacion Electronica
    'rexplorap.MoveFirst
    '01/08/2018 Testing Facturacion Electronica
        
    'Do
    '  If rexplorap.EOF Then Exit Do
    ' If rexplorap.Fields("dflag") = "S" Then
    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

    If sw = 0 Then
ax:

        If mytablex.State = 1 Then mytablex.Close
        buf = "SELECT * FROM " & bufca & " where local='" & rexplorap.Fields("local") & "' and tipo='" & extra_loquesea(gtipo) & "' and serie='" & gserie & "' and numero='" & gnumero & "'"
        mytablex.Open buf, cn, adOpenKeyset, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            sdx = Val(gnumero) + 1
            gnumero = "" & sdx
            GoTo ax
            Exit Sub

        End If

        mytablez.Open "SELECT * FROM tipo where tipo='" & extra_loquesea(gtipo) & "'", cn, adOpenKeyset, adLockOptimistic

        If mytablez.RecordCount > 0 Then
            mytablez.Fields("numero") = gnumero
            mytablez.Update

        End If

        mytablez.Close
        mytablex.AddNew

        ' Testing Proyecto Facturacion Electronica Nota de Credito 06/2018
        For I = 0 To rexplorap.Fields.count - 1
            mytablex.Fields(I) = rexplorap.Fields(I)
        Next I

        'mytablex.Fields("tipoclie") = rexplorap.Fields("tipoclie")
        'mytablex.Fields("codigo") = rexplorap.Fields("codigo")
        'mytablex.Fields("partida") = rexplorap.Fields("partida")
        'mytablex.Fields("destino") = rexplorap.Fields("destino")
        'mytablex.Fields("moneda") = rexplorap.Fields("moneda")
        'mytablex.Fields("vendedor") = rexplorap.Fields("vendedor")
        'mytablex.Fields("transporte") = rexplorap.Fields("transporte")
        'mytablex.Fields("fpago") = rexplorap.Fields("fpago")
        'mytablex.Fields("paridad") = rexplorap.Fields("paridad")
        'mytablex.Fields("dias") = rexplorap.Fields("dias")
        'mytablex.Fields("bodega") = rexplorap.Fields("bodega")
        'mytablex.Fields("bodegaf") = rexplorap.Fields("bodegaf")
        'mytablex.Fields("observa") = rexplorap.Fields("observa")
        'mytablex.Fields("usuario") = rexplorap.Fields("usuario")
        'mytablex.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
        'mytablex.Fields("hora") = Format(Now, "hh:MM:SS")
        'mytablex.Fields("nombre") = rexplorap.Fields("nombre")
        'mytablex.Fields("total") = rexplorap.Fields("total")
        'mytablex.Fields("descuento") = rexplorap.Fields("descuento")
        'mytablex.Fields("neto") = rexplorap.Fields("neto")
        'mytablex.Fields("impuesto") = rexplorap.Fields("impuesto")
        'mytablex.Fields("subtotal") = rexplorap.Fields("subtotal")
        'mytablex.Fields("flage") = rexplorap.Fields("flage")
        'mytablex.Fields("local") = rexplorap.Fields("local")
        '
        'mytablex.Fields("c1") = rexplorap.Fields("c1")
        'mytablex.Fields("c2") = rexplorap.Fields("c2")
        'mytablex.Fields("c3") = rexplorap.Fields("c3")
        'mytablex.Fields("c4") = rexplorap.Fields("c4")
        'mytablex.Fields("c5") = rexplorap.Fields("c5")
        'mytablex.Fields("c6") = rexplorap.Fields("c6")
        'mytablex.Fields("c7") = rexplorap.Fields("c7")
        'mytablex.Fields("c8") = rexplorap.Fields("c8")
        'mytablex.Fields("c9") = rexplorap.Fields("c9")
        'mytablex.Fields("nro_items") = rexplorap.Fields("nro_items")
        'mytablex.Fields("yausado") = rexplorap.Fields("yausado")
        'mytablex.Fields("caja") = rexplorap.Fields("caja")
        'mytablex.Fields("turno") = rexplorap.Fields("turno")
        'mytablex.Fields("servicio") = rexplorap.Fields("servicio")

        '
        'mytablex.Fields("percepcion") = rexplorap.Fields("percepcion")
        ' Testing Proyecto Facturacion Electronica Nota de Credito 06/2018

        mytablex.Fields("fecha") = Format(Now, "dd/mm/yyyy")
        mytablex.Fields("fechae") = Format(Now, "dd/mm/yyyy")
        mytablex.Fields("fechasunat") = Format(Now, "dd/mm/yyyy")
        mytablex.Fields("tipo") = extra_loquesea(gtipo)
        mytablex.Fields("serie") = gserie
        mytablex.Fields("numero") = gnumero
        mytablex.Fields("acu") = gacu
        mytablex.Fields("estado") = "0"
        mytablex.Fields("dflag") = ""
        mytablex.Fields("tipo1") = "" & rexplorap.Fields("tipo")
        mytablex.Fields("serie1") = rexplorap.Fields("serie")
        mytablex.Fields("numero1") = rexplorap.Fields("numero")
        ' Testing Proyecto Facturacion Electronica Nota de Credito 06/2018
        'mytablex.Fields("total") = rexplorap.Fields("total") * -1
        'mytablex.Fields("descuento") = rexplorap.Fields("descuento") * -1
        'mytablex.Fields("neto") = rexplorap.Fields("neto") * -1
        'mytablex.Fields("impuesto") = rexplorap.Fields("impuesto") * -1
        'mytablex.Fields("subtotal") = rexplorap.Fields("subtotal") * -1
        'mytablex.Fields("percepcion") = rexplorap.Fields("percepcion") * -1

        mytablex.Fields("ESTADO_SUNAT") = ""
        ' Testing Proyecto Facturacion Electronica Nota de Credito 06/2018

        ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

        If "" & rexplorap.Fields("acu") = "Q" Then
            If gacu = "Z" Then
                mytablex.Fields("bodega") = "" & rexplorap.Fields("bodegaf")
                mytablex.Fields("bodegaf") = "" & rexplorap.Fields("bodega")
                mytablex.Fields("local") = "" & rexplorap.Fields("localf")
                mytablex.Fields("localf") = "" & rexplorap.Fields("local")

            End If

        End If

        '''21/08/2017 kenyo Guia de Salida con Factura
        If gacu = "T" Then
            mytablex.Fields("acu1") = "S"
        Else
            mytablex.Fields("acu1") = "" & rexplorap.Fields("acu")

        End If

        '''21/08/2017 kenyo Guia de Salida con Factura

        ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
        ' mytablex.Fields("cdr") = ""
        ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

        mytablex.Update
        mytablex.Close
        sw = 1

    End If

    'Genera solo documentos sin recetas
    '19/03/2018 Correcion de error al generar documento con recetas como guia de salida
    'mytabley.Open "SELECT * FROM " & dgusuariog & " where local='" & rexplorap.Fields("local") & "' and tipo='" & rexplorap.Fields("tipo") & "' and serie='" & rexplorap.Fields("serie") & "' and numero='" & rexplorap.Fields("numero") & "'", cn, adOpenKeyset, adLockOptimistic
    mytabley.Open "SELECT * FROM " & dgusuariog & " where  dua is null and local='" & rexplorap.Fields("local") & "' and tipo='" & rexplorap.Fields("tipo") & "' and serie='" & rexplorap.Fields("serie") & "' and numero='" & rexplorap.Fields("numero") & "'", cn, adOpenKeyset, adLockOptimistic
    '19/03/2018 Correcion de error al generar documento con recetas como guia de salida

    If mytabley.RecordCount > 0 Then
        mytablex.Open "SELECT * FROM " & bufde & " where  local='" & mytabley.Fields("local") & "' and tipo='" & extra_loquesea(gtipo) & "' and serie='" & gserie & "' and numero='" & gnumero & "'", cn, adOpenKeyset, adLockOptimistic
        Do

            If mytabley.EOF Then Exit Do
            mytablex.AddNew

            ' Testing Proyecto Facturacion Electronica Nota de Credito 06/2018
            For I = 0 To mytabley.Fields.count - 1
                mytablex.Fields(I) = mytabley.Fields(I)
            Next I

            '
            'mytablex.Fields("total") = mytabley.Fields("total")
            'mytablex.Fields("descuento") = mytabley.Fields("descuento")
            'mytablex.Fields("neto") = mytabley.Fields("neto")
            'mytablex.Fields("impuesto") = mytabley.Fields("impuesto")
            'mytablex.Fields("subtotal") = mytabley.Fields("subtotal")
            'mytablex.Fields("percepcion") = mytabley.Fields("percepcion")
            ' Testing Proyecto Facturacion Electronica Nota de Credito 06/2018

            If gacu = "E" Then
                mytablex.Fields("CANTIDAD") = mytabley.Fields("CANTIDAD") * -1
            Else
            
                mytablex.Fields("CANTIDAD") = mytabley.Fields("CANTIDAD")

            End If
            
            mytablex.Fields("tipo") = extra_loquesea(gtipo)
            mytablex.Fields("serie") = gserie
            mytablex.Fields("numero") = gnumero
            mytablex.Fields("acu") = gacu
            mytablex.Fields("estado") = "0"
            mytablex.Fields("fecha") = Format(Now, "dd/mm/yyyy")
            mytablex.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
            
            '07/08/2018 No descuenta stock en guia de remision
            If gacu = "T" Then
                If busca_DescuentaStock(extra_loquesea(gtipo)) = "N" Then
                    mytablex.Fields("L4") = "N"

                End If

            End If
            
            '07/08/2018 No descuenta stock en guia de remision

            ' Testing Proyecto Facturacion Electronica Nota de Credito 06/2018
            ''mytablex.Fields("CANTIDAD") = mytabley.Fields("CANTIDAD") * -1
            ''mytablex.Fields("total") = mytabley.Fields("total") * -1
            ''mytablex.Fields("descuento") = mytabley.Fields("descuento") * -1
            ''mytablex.Fields("neto") = mytabley.Fields("neto") * -1
            ''mytablex.Fields("impuesto") = mytabley.Fields("impuesto") * -1
            ''mytablex.Fields("subtotal") = mytabley.Fields("subtotal") * -1
            ''mytablex.Fields("percepcion") = mytabley.Fields("percepcion") * -1
            ' Testing Proyecto Facturacion Electronica Nota de Credito 06/2018

            'mytablex.Fields("acu1") = "" & rexplorap.Fields("acu")

            ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
            ' HORA ACTUAL PARA TODOS LOS DOCUMENTOS
            '25/06/2018 Testing Almacen General

            'mytablex.Fields("acu1") = ""
            '25/06/2018 Testing Almacen General
            'If gacu = "E" Or gacu = "F" Then
            '    mytablex.Fields("acu1") = ""
            'Else
            '    mytablex.Fields("acu1") = "" & rexplorap.Fields("acu")
            'End If

            'mytablex.Fields("hora") = Format(Now, "hh:MM:SS")
            ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

            If "" & rexplorap.Fields("acu") = "Q" Then
                If Trim(gacu) = "Z" Then
                    mytablex.Fields("bodega") = "" & mytabley.Fields("bodegaf")
                    mytablex.Fields("bodegaf") = "" & mytabley.Fields("bodega")
                    mytablex.Fields("local") = "" & mytabley.Fields("localf")
                    mytablex.Fields("localf") = "" & mytabley.Fields("local")

                End If

            End If

            mytablex.Update
            mytabley.Fields("acu1") = ""
            mytabley.MoveNext
        Loop

    End If

    mytabley.Close
    mytablex.Close
    rexplorap.Fields("yausado") = "1"
    rexplorap.Fields("dflag") = ""
    rexplorap.Fields("tipo1") = extra_loquesea(gtipo)
    rexplorap.Fields("serie1") = gserie
    rexplorap.Fields("numero1") = gnumero
    rexplorap.Fields("acu1") = ""
    rexplorap.Update

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

    MsgBox "Proceso Realizado", 48, "Aviso"
    
    '01/08/2018 Testing Facturacion Electronica
    'End If
    ' rexplorap.MoveNext
    'Loop
    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
    '01/08/2018 Testing Facturacion Electronica

    sql_cabeza
    Frame4.Visible = False
    Exit Sub

End Sub

Private Sub Command8_Click()
    SUMAR_CABEZA

End Sub

' Testing Proyecto Facturacion Electronica

Private Sub DarBaja_Click()

    Dim salida           As Boolean

    Dim my_ruc           As String

    Dim my_CDR           As String

    Dim file             As String

    Dim encontro         As Boolean

    Dim input_file       As String

    Dim my_estadosunat   As String

    Dim my_numerointerno As String

    Dim hastaCuanto      As Integer

    Dim nuevoDato        As String

    Dim myDato           As String
                
    my_local = "" & rexplorap.Fields("local")
    my_tipo = "" & rexplorap.Fields("tipo")
    my_serie = "" & rexplorap.Fields("serie")
    my_numero = "" & rexplorap.Fields("numero")
    my_estado = "" & rexplorap.Fields("estado")
    my_acu = "" & rexplorap.Fields("acu")
    my_caja = "" & rexplorap.Fields("caja")
    my_tipo1 = "" & rexplorap.Fields("tipo1")
 
    ' Testing Proyecto Facturacion Electronica Hash
    my_CDR = "" & rexplorap.Fields("cdr")
    ' Testing Proyecto Facturacion Electronica Hash

    ' Testing Proyecto Facturacion Electronica 12/04/2018
    If my_tipo = "3" Or my_tipo = "4" Or my_tipo = "5" Then  'En caso sea boleta manual, factura manual o ticket
        MsgBox "No permitido", vbCritical
        Exit Sub

    End If
 
    ' Testing Proyecto Facturacion Electronica Hash
    If my_CDR = "" Then
        Call Datos_Empresa(my_struc_datos_empresa(), my_local, salida, 0)
        my_ruc = my_struc_datos_empresa(0).codigo1
    
        hastaCuanto = 8 - Len(my_numero)
        myDato = my_numero
        Call E_llenar_zero(hastaCuanto, myDato, my_numero)
        my_numero = myDato
                
        If my_tipo = "1" Then
            file = my_ruc & "_03" & "_" & my_serie & "-" & my_numero & ".INPUT.TXT"
        ElseIf my_tipo = "2" Then
            file = my_ruc & "_01" & "_" & my_serie & "-" & my_numero & ".INPUT.TXT"

        End If
       
        Do While encontro = False
        
            salida = FileExists("D:\ce_Input\FIRMADO\R_" & Left(file, (Len(file) - 10)) & ".txt")

            If salida = True Then
                Call busca_respuesta_electronica(file, "D:\ce_Input\FIRMADO\R_")
                input_file = "D:\ce_Input\FIRMADO\R_" & Left(file, (Len(file) - 10)) & ".txt"
                encontro = True

            End If

            If salida = False Then
                MsgBox "Document No encontrado", vbCritical
                Exit Sub

            End If

        Loop
   
        my_numerointerno = "" & rexplorap.Fields("numero")
        Call read_save_electronico(input_file, my_local, my_serie, my_numerointerno, my_tipo, ACU, "")
   
    End If

    ' Testing Proyecto Facturacion Electronica Hash
 
    ' Testing Proyecto Facturacion Electronica 12/04/2018
        
    If MsgBox("¿SEGURO DAR DE BAJA?: Documento " & "" & my_serie & "-" & my_numero, 1, "Aviso") <> 1 Then Exit Sub
    If (my_tipo1 = "1" Or my_tipo1 = "2" Or my_tipo = "1" Or my_tipo = "2") Then
   
        my_estadosunat = obtiene_EstadoSunat(cgusuario, my_local, my_tipo, my_serie, my_numero)

        If my_estadosunat = "BAJA" Or my_estadosunat = "PENDIENTE_BAJA" Then
            MsgBox "Documento ya declarado como baja", vbCritical
            Exit Sub

        End If
   
        If my_estado = "1" Then
            Call Busca_comprobante_sunat(my_local, my_serie, my_numero, my_tipo, salida)

            If salida = True Then 'si existe factura y/o Boleta
                Call Datos_Empresa(my_struc_datos_empresa(), my_local, salida, 0)
                my_ruc = my_struc_datos_empresa(0).codigo1
            
                ' Testing Proyecto Facturacion Electronica
                
                hastaCuanto = 8 - Len(my_numero)
                myDato = my_numero
                Call E_llenar_zero(hastaCuanto, myDato, my_numero)
                my_numero = myDato

                If my_tipo = "2" Then ' 01 FACTURA 03 BOLETA
                    file = my_ruc & "_01" & "_" & my_serie & "-" & my_numero & ".INPUT.TXT"
                ElseIf my_tipo = "1" Then
                    file = my_ruc & "_03" & "_" & my_serie & "-" & my_numero & ".INPUT.TXT"

                End If
                
                ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
                If my_tipo1 = "1" Or my_tipo1 = "2" Then ' nc
                    If my_acu = "E" Then 'NOTAS DE CREDITO
                        file = my_ruc & "_07" & "_" & my_serie & "-" & my_numero & ".INPUT.TXT"
                    ElseIf my_acu = "F" Then 'NOTAS DE DEBITO
                        file = my_ruc & "_08" & "_" & my_serie & "-" & my_numero & ".INPUT.TXT"

                    End If

                End If

                ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
                
                input_file = "D:\ce_Input\PROCESADO\R_" & Left(file, (Len(file) - 10)) & ".txt"
                salida = FileExists(input_file)

                If salida = False Then
                    MsgBox "Documento debe ser declarado", vbCritical
                    Exit Sub
                Else
                    Call verifica_estado_electronico(input_file)

                    If salida = False Then 'Si documnento de respuesta tiene error
                        MsgBox "Documento declarado invalido", vbCritical
                        Exit Sub

                    End If

                End If

                ' Testing Proyecto Facturacion Electronica
                 
                my_numerointerno = "" & rexplorap.Fields("numero")
                Call estrae_baja(my_ruc, Mid(my_local, 1, 2), my_tipo, my_serie, my_numerointerno, rexplorap.Fields("fecha"), file, salida)

                If salida = True Then

                    'aqui seria la busqueda del file
                    Do While encontro = False
                        salida = FileExists("D:\ce_Input\FIRMADO\R_" & Left(file, (Len(file) - 10)) & ".txt")

                        If salida = True Then
                            Call busca_respuesta_electronica(file, "D:\ce_Input\FIRMADO\R_")
                            input_file = "D:\ce_Input\FIRMADO\R_" & Left(file, (Len(file) - 10)) & ".txt"
                            encontro = True

                        End If

                        salida = FileExists("D:\ce_Input\ERROR\R_" & Left(file, (Len(file) - 10)) & ".txt")

                        If salida = True Then
                            Call busca_respuesta_electronica(file, "D:\ce_Input\ERROR\R_")
                            input_file = "D:\ce_Input\ERROR\R_" & Left(file, (Len(file) - 10)) & ".txt"
                            encontro = True

                        End If

                    Loop
                   
                    my_numerointerno = "" & rexplorap.Fields("numero")
                   
                    Call Importacion_cotizacion("1", my_local, my_tipo, my_serie, my_numerointerno, my_acu)
                    Call Actualiza_Estado_Sunat("PENDIENTE_BAJA", my_local, my_tipo, my_serie, my_numerointerno, my_acu)
                    
                    MsgBox "Proceso Elaborado"
                    Call sql_cabeza

                End If

            Else
                Exit Sub
        
            End If

        Else
            MsgBox "Documento debe estar anulado", vbCritical

        End If

    Else
        MsgBox "Verificar Documento"

    End If

End Sub

' Testing Proyecto Facturacion Electronica

' Testing Proyecto Facturacion Electronica
Sub ExtraeDAtos()

End Sub

' Testing Proyecto Facturacion Electronica

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = "1" Then
            codigo = dbGrid1.columns(1)
            Frame1.Visible = False
            Frame1.Enabled = False
            codigo.SetFocus

        End If

        If opcion1 = "6100" Then
            mytablex.Open "SELECT * FROM userlocal where codigo='" & gusuario & "' and local='" & Trim(dbGrid1.columns(1)) & "'", cn, adOpenKeyset, adLockOptimistic

            If mytablex.RecordCount = 0 Then
                mytablex.Close
                MsgBox "Usuario No autorizado,utilizar este local ", 48, "Aviso"
                Exit Sub

            End If

            mytablex.Close
   
            buf = Trim(dbGrid1.columns(1))
            Frame1.Visible = False
            Frame1.Enabled = False
            'buscar almacen que pertenece----
            'mytablex.Open "SELECT * FROM bodega where local='" & buf & "'", cn, adOpenKeyset, adLockOptimistic
            'If mytablex.RecordCount > 0 Then
            'tfactura.bodega = Trim("" & mytablex.Fields("codigo"))
            'End If
            'mytablex.Close
            menu_nuevo buf

            'codigo.SetFocus
        End If

    End If

End Sub

Private Sub dbgrid2_BeforeColEdit(ByVal ColIndex As Integer, _
                                  ByVal KeyAscii As Integer, _
                                  Cancel As Integer)

    If ColIndex <> 13 Then
        Cancel = True
        Exit Sub

    End If

    Select Case ColIndex

        Case 13

            If Len("" & rexplorap.Fields("local")) = 0 Then
                Cancel = True
                Exit Sub

            End If

    End Select

End Sub

Private Sub dbgrid2_BeforeColUpdate(ByVal ColIndex As Integer, _
                                    OldValue As Variant, _
                                    Cancel As Integer)

    Dim found As Integer

    Select Case ColIndex

        Case 13

            If Len(rexplorap.Fields("local")) = 0 Then
                Cancel = True
                Exit Sub

            End If

            If Len("" & rexplorap.Fields(13)) = 0 Then
                rexplorap.Fields(13) = rexplorap.Fields("fecha")
                Exit Sub

            End If

            found = valida_fecha("" & rexplorap.Fields(13))

            If found = 0 Then
                Cancel = True
                Exit Sub

            End If

    End Select

End Sub

Private Sub DBGrid2_DblClick()

    On Error GoTo cmd8966_err

    If "" & rexplorap.Fields("dflag") = "S" Then
        rexplorap.Fields("DFLAG") = ""
        rexplorap.Update
        Exit Sub

    End If

    rexplorap.Fields("DFLAG") = "S"
    rexplorap.Update
    Exit Sub
cmd8966_err:
    MsgBox "Seleccione un Dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub DBGrid2_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'consultando
        consulta_detalle

    End If

    ' Testing Proyecto Facturacion Electronica 05/04/2018
    If KeyCode = 113 Then  'consultando
        VerDetalleSunat

    End If

    ' Testing Proyecto Facturacion Electronica 05/04/2018

End Sub

' Testing Proyecto Facturacion Electronica 05/04/2018

Sub CargaTipoDocumento()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM tipo where TIPODOC='A'  OR TIPODOC='B'  OR TIPODOC='C'  OR TIPODOC='D' OR TIPODOC='N'  OR TIPODOC='O'  OR TIPODOC='E'  OR TIPODOC='F'  ", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        TxtTipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    TxtTipo.ListIndex = 0
    mytablex.Close

End Sub

Function CargaDescripcionDocumento(ByRef tipo As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM tipo  where tipo='" & tipo & "'", cn, adOpenKeyset, adLockOptimistic
    CargaDescripcionDocumento = mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")
    mytablex.Close

End Function

Sub VerDetalleSunat()

    Dim salida           As Boolean

    Dim my_ruc           As String

    Dim file             As String

    Dim encontro         As Boolean

    Dim input_file       As String

    Dim hastaCuanto      As Integer

    Dim my_estadosunat   As String

    Dim my_numerointerno As String

    Dim my_tipointerno   As String
    
    Dim rpta             As String

    Dim buf              As String

    If rexplorap.RecordCount = 0 Then Exit Sub
    my_tipo = rexplorap.Fields("tipo")

    If my_tipo = "3" Or my_tipo = "4" Or my_tipo = "5" Then  'En caso sea boleta manual, factura manual o ticket
        MsgBox "No permitido", vbCritical
        Exit Sub

    End If

    If FrmConsultaSunat.Visible = False Then
        CargaTipoDocumento
        TxtLocal.Text = rexplorap.Fields("local")
        TxtTipo.Text = CargaDescripcionDocumento(rexplorap.Fields("tipo"))
        
        TxtSerie.Text = rexplorap.Fields("serie")
        TxtNumero.Text = rexplorap.Fields("numero")
        TxtLocal.Text = rexplorap.Fields("local")
     
        TxtTotal.Text = rexplorap.Fields("total")
        TxtFecha.Text = rexplorap.Fields("fecha")
        txtEstado.Text = rexplorap.Fields("estado")
        
        my_local = rexplorap.Fields("local")
        my_tipo = CargaDescripcionDocumento(rexplorap.Fields("tipo"))
        my_serie = rexplorap.Fields("serie")
        my_numero = rexplorap.Fields("numero")
        my_acu = rexplorap.Fields("acu")
        txtEstadoSunat = "" & obtiene_DatosComprobante(4, cgusuario, my_local, extra_loquesea(TxtTipo), TxtSerie.Text, TxtNumero.Text)

    End If
    
    If FrmConsultaSunat.Visible = True Then
        
        my_local = TxtLocal.Text
        my_tipo = extra_loquesea(TxtTipo)
        my_serie = TxtSerie.Text
        my_numero = TxtNumero.Text
        my_acu = txtacu.Text
    
        buf = verificar_existencia(cgusuario, my_local, my_tipo, my_serie, my_numero)

        If buf = "0" Then
            MsgBox ("Documento No Existe")
            TxtTotal.Text = ""
            TxtFecha.Text = ""
            txtEstado.Text = ""
            txtEstadoSunat = ""
            txtruc.Text = ""
            TxtNumero.SetFocus
            Exit Sub

        End If
        
        TxtFecha = obtiene_DatosComprobante(1, cgusuario, my_local, my_tipo, my_serie, my_numero)
        TxtTotal = obtiene_DatosComprobante(2, cgusuario, my_local, my_tipo, my_serie, my_numero)
        txtEstado = obtiene_DatosComprobante(3, cgusuario, my_local, my_tipo, my_serie, my_numero)
        txtacu = obtiene_DatosComprobante(5, cgusuario, my_local, my_tipo, my_serie, my_numero)
      
        txtEstadoSunat = "" & obtiene_DatosComprobante(4, cgusuario, my_local, my_tipo, my_serie, my_numero)
        
    End If
           
    ' Testing Proyecto Facturacion Electronica FE 21/05/2018
        
    'If txtEstadoSunat = "PENDIENTE" Or txtEstadoSunat = "ACEPTADO" Then
    Call Datos_Empresa(my_struc_datos_empresa(), my_local, True, 0)
    my_ruc = my_struc_datos_empresa(0).codigo1
               
    If my_acu = "D" Then
        my_tipointerno = "01"
    ElseIf my_acu = "E" Then
        my_tipointerno = "07"
    ElseIf my_acu = "F" Then
        my_tipointerno = "08"

    End If
                 
    my_numerointerno = my_numero
                     
    hastaCuanto = 8 - Len(my_numerointerno) '**en la tabla
    my_numerointerno = my_numerointerno
    Call E_llenar_zero(hastaCuanto, my_numerointerno, my_numerointerno)
         
    If txtEstadoSunat = "PENDIENTE" Then
        file = my_ruc & "_" & my_tipointerno & "_" & my_serie & "-" & my_numerointerno & ".INPUT.TXT"
    ElseIf txtEstadoSunat = "PENDIENTE_BAJA" Then
        file = my_ruc & "_RA_" & my_serie & "-" & my_numerointerno & ".INPUT.TXT"
    Else
        FrmConsultaSunat.Visible = True
        TxtNumero.SetFocus
        Exit Sub

    End If
        
    input_file = "D:\ce_Input\PROCESADO\R_" & Left(file, (Len(file) - 10)) & ".txt"
    salida = FileExists(input_file)
        
    If salida = False Then
        MsgBox "Documento NO ENCONTRADO", vbCritical
    Else
        Call verifica_estado_electronicoXDocumento(input_file)

    End If
        
    txtEstadoSunat = "" & obtiene_DatosComprobante(4, cgusuario, my_local, extra_loquesea(TxtTipo), TxtSerie.Text, TxtNumero.Text)
                 
    ' End If
        
    ' Testing Proyecto Facturacion Electronica FE 21/05/2018
        
    rpta = obtiene_EstadoSunat(cgusuario, my_local, my_tipo, my_serie, my_numero)
    FrmConsultaSunat.Visible = True
    TxtNumero.SetFocus
    
End Sub

' Testing Proyecto Facturacion Electronica 05/04/2018

' Testing Proyecto Facturacion Electronica 05/04/2018
Function obtiene_DatosComprobante(ByRef tipo As String, _
                                  cgusuario As String, _
                                  my_local As String, _
                                  my_tipo As String, _
                                  my_serie As String, _
                                  my_numero As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT fecha,total,estado,estado_sunat,acu FROM  " & cgusuario & "  where  local='" & my_local & "' and tipo='" & my_tipo & "' and serie='" & my_serie & "' and numero='" & my_numero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
          
        If tipo = 1 Then
            obtiene_DatosComprobante = mytablex.Fields("fecha")
        ElseIf tipo = 2 Then
            obtiene_DatosComprobante = mytablex.Fields("total")
        ElseIf tipo = 3 Then
            obtiene_DatosComprobante = mytablex.Fields("estado")
        ElseIf tipo = 4 Then
            obtiene_DatosComprobante = mytablex.Fields("estado_sunat")
        ElseIf tipo = 5 Then
            obtiene_DatosComprobante = mytablex.Fields("acu")

        End If

    Else
        obtiene_DatosComprobante = "NO EXISTE"

    End If

    mytablex.Close

End Function

' Testing Proyecto Facturacion Electronica 05/04/2018

' Testing Proyecto Facturacion Electronica 05/04/2018
Private Sub DetalleSunat_Click()
    VerDetalleSunat

End Sub

' Testing Proyecto Facturacion Electronica 05/04/2018

Private Sub dj8844_Click()

    Dim buf As String

    On Error GoTo cmd94512_err

    If Frame8.Visible = True Then Exit Sub

    If Frame7.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    hfechai = ""
    'hfechaf = ""
    buf = "" & rexplorap.Fields("tipo")
    Frame7.Visible = True
    hfechai.SetFocus
    Exit Sub
cmd94512_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub djbu232_Click()

    'If Frame4.Visible = True Then Exit Sub
    If Frame8.Visible = True Then Exit Sub

    If Frame7.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    Frame3.Visible = True

End Sub

Private Sub djku232_Click()

    If Frame8.Visible = True Then Exit Sub

    If Frame7.Visible = True Then Exit Sub

    If Frame5.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    consulta_local

End Sub

Sub menu_nuevo(buf As String)

    Dim found As Integer

    On Error GoTo cmd28_err

    'tfactura.importacion = importacion
    'If importacion = "GASTOS" Then
    'tfactura.Label51.Visible = True
    'tfactura.cgasto.Visible = True
    'Exit Sub
    'End If
    'If importacion = "IMPORTACION" Then
    'tfactura.Label56.Visible = True
    'tfactura.Label57.Visible = True
    'tfactura.agencia.Visible = True
    'tfactura.dua.Visible = True
   
    'borratempdao
    'cn.Execute ("select * into _i" & gusuario & " from tgastofactura")
    'cn.Execute ("delete from _i" & gusuario)
    'End If
    'borratempdao
    tfactura.local1 = buf

    If ACU = "Z" Then
        'tfactura.local1 = "01"
        tfactura.codigo = "01"
        tfactura.Label2.Caption = "Codigo"
        tfactura.Label14.Visible = True
        tfactura.Label38.Visible = True
        tfactura.localf.Visible = True
        tfactura.bodegaf.Visible = True
        tfactura.localf.Enabled = True
        tfactura.bodegaf.Enabled = True
        tfactura.Label14.Caption = "LocalDestino"
        tfactura.Label38.Caption = "AlmacenDestino"

    End If

    If ACU = "Q" Then
        tfactura.codigo = "01"
        tfactura.Label2.Caption = "Codigo"
        tfactura.Label38.Visible = True
        'tfactura.label13.Caption = "AlmacenDestino"
        'tfactura.Label38.Caption = "Alm:origen"
   
        tfactura.bodegaf.Visible = True
        tfactura.tipoclie = tipoclie
        tfactura.Label14.Visible = True
        tfactura.localf.Visible = True
        'tfactura.Label14.Caption = "Loc:Origen"
        tfactura.localf.Enabled = True
        tfactura.bodegaf.Enabled = True
 
        'tfactura.Label14.Caption = "LocalOrigen"
        ' tfactura.Label38.Caption = "AlmacenOrigen"
 
    End If

    If ACU = "V" Then
        found = copiar_temporal()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
            Exit Sub

        End If

        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Facturacion x Ventas"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        'tfactura.caja = "00"
        tfactura.bandera = "Nuevo"
        tfactura.ACU = "V"
        tfactura.tipoclie = tipoclie

        'tfactura.local1=local
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "H" Then
        found = copiar_temporal()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
            Exit Sub

        End If

        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Cotizaciones Ventas"
        cgusuario = "CCOTIZAV"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "dcotizav"
        tfactura.bandera = "Nuevo"
        tfactura.ACU = "H"
        tfactura.tipoclie = tipoclie
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "I" Then
        found = copiar_temporal()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
            Exit Sub

        End If

        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Cotizaciones Ventas"
        cgusuario = "Cpedidov"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "dpedidov"
        tfactura.bandera = "Nuevo"
        tfactura.ACU = "I"
        tfactura.tipoclie = tipoclie
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "3" Then
        found = copiar_temporal()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
            Exit Sub

        End If

        found = copiar_servicio()

        If found = 0 Then
            MsgBox "Error al copiar servicio", 24, "Aviso"
            Exit Sub

        End If

        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Servicio Tecnico"
        cgusuario = "Cservicio"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "dservicio"
        sgusuario = "_s" & gusuario  'servicio tecnico
        tfactura.bandera = "Nuevo"
        tfactura.ACU = "3"
        tfactura.tipoclie = tipoclie
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "T" Then
        found = copiar_temporal()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
            Exit Sub

        End If

        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Guia Salida"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.bandera = "Nuevo"
        tfactura.ACU = "T"
        tfactura.tipoclie = tipoclie
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "E" Then
        found = copiar_temporal()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
            Exit Sub

        End If

        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Nota Credito Ventas"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.bandera = "Nuevo"
        tfactura.ACU = "E"
        tfactura.tipoclie = tipoclie
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "F" Then
        found = copiar_temporal()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
            Exit Sub

        End If

        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Nota debito Ventas"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        tfactura.bandera = "Nuevo"
        dgusuariog = "DETALLE"
        tfactura.ACU = "F"
        tfactura.tipoclie = tipoclie
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "R" Then
        found = copiar_temporal()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   
            Exit Sub

        End If

        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Orden de Compra"
        cgusuario = "CORDENC"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DORDENC"
        tfactura.ACU = "R"
        tfactura.bandera = "Nuevo"
        tfactura.tipoclie = tipoclie
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "S" Then
        found = copiar_temporal()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   
            Exit Sub

        End If

        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Guia Remision Entrada"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.tipoclie = tipoclie
        tfactura.ACU = "S"
        tfactura.bandera = "Nuevo"
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "C" Then
        found = copiar_temporal()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   
            Exit Sub

        End If

        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Factura de Compra"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.ACU = "C"
        tfactura.bandera = "Nuevo"
        tfactura.tipoclie = tipoclie
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "N" Then
        found = copiar_temporal()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
            Exit Sub

        End If

        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Nota Credito Compra"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.ACU = "N"
        tfactura.bandera = "Nuevo"
        tfactura.tipoclie = tipoclie
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "O" Then
        found = copiar_temporal()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
            Exit Sub

        End If

        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Nota debito de Compra"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.ACU = "O"
        tfactura.bandera = "Nuevo"
        tfactura.tipoclie = tipoclie
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "Q" Then
        found = copiar_temporal()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
            Exit Sub

        End If

        tfactura.bodegaf.Visible = True
        tfactura.Label38.Visible = True
        'tfactura.label13.Caption = "AlmacenDestino"
        'tfactura.Label38.Caption = "DesdAlm."
        tfactura.bodegaf.Visible = True

        tfactura.localf.Enabled = True
        tfactura.bodegaf.Enabled = True

        'tfactura.Label14.Caption = "LocalOrigen"
        '   tfactura.Label38.Caption = "AlmacenOrigen"

        tfactura.Label2 = "Codigo"
        tfactura.Caption = "Pedido Almacen"
        cgusuario = "CREQUISA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DREQUISA"
        tfactura.ACU = "Q"
        tfactura.ttipo = "Q"
        tfactura.bandera = "Nuevo"
        tfactura.tipoclie = tipoclie
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "Z" Then
        found = copiar_temporal()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
            End
            Exit Sub

        End If

        'tfactura.Label2 = "Cod.Inicio"
        tfactura.Caption = "Traslado entre almacen de un mismo establecimiento"
        cgusuario = "CTRASLAD"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DTRASLAD"
        tfactura.ACU = "Z"
        tfactura.bandera = "Nuevo"
        tfactura.tipoclie = tipoclie
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    MsgBox "Presione tecla para continuar..", 48, "Aviso"
    sql_cabeza
    Exit Sub
cmd28_err:
    MsgBox "Nuevo:Seleccione un dato", 48, "Aviso"
    Exit Sub

End Sub

Private Sub dki844_Click()

    Dim sw As Integer

    On Error GoTo cmd90765_err

    If Frame8.Visible = True Then Exit Sub
    If Frame7.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub
    Frame9.Visible = True
    '-------------------------------------------------

    If "" & rexplorap.Fields("estado") <> "2" Then
        MsgBox "Estado debe estar en 2 para realizar esta operacion ", 48, "Aviso"
        Exit Sub

    End If

    sw = 0
    List2.Clear
    rexplorap.MoveFirst
    Do

        If rexplorap.EOF Then Exit Do
        If "" & rexplorap.Fields("dflag") = "S" Then
            List2.AddItem "" & rexplorap.Fields("tipo") & " " & rexplorap.Fields("serie") & " " & rexplorap.Fields("numero") & " " & rexplorap.Fields("nombre")
            sw = 1

        End If

        rexplorap.MoveNext
    Loop

    If sw = 0 Then
        MsgBox "No ha seleccionado ", 48, "Aviso"
        Exit Sub

    End If

    List2.ListIndex = 0
    '-----------------------------------------------
    proceso_seleccion
    Exit Sub
cmd90765_err:
    MsgBox "Seleccione un dato " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub proceso_seleccion()

    Dim sw       As Integer

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    sw = 0
    cn.Execute ("delete from tmpedido")
    rexplorap.MoveFirst
    Do

        If rexplorap.EOF Then Exit Do
        If rexplorap.Fields("dflag") = "S" Then
            '----------------------
            buf = "select * from dpedidov where local='" & "" & rexplorap.Fields("local") & "' and tipo='" & "" & rexplorap.Fields("tipo") & "' and serie='" & "" & rexplorap.Fields("serie") & "' and numero='" & "" & rexplorap.Fields("numero") & "'"
            mytablex.Open buf, cn, adOpenKeyset, adLockOptimistic
            Do

                If mytablex.EOF Then Exit Do
                mytabley.Open "select * from tmpedido where producto='" & "" & mytablex.Fields("producto") & "'", cn, adOpenKeyset, adLockOptimistic

                If mytabley.RecordCount = 0 Then
                    mytabley.AddNew
                    mytabley.Fields("producto") = Trim("" & mytablex.Fields("producto"))
                    mytabley.Fields("descripcio") = Trim("" & mytablex.Fields("descripcio"))
                    mytabley.Fields("cantidad") = Val("" & mytablex.Fields("cantidad"))
                    mytabley.Fields("unidad") = Trim("" & mytablex.Fields("unidad"))
                    mytabley.Fields("factor") = Val("" & mytablex.Fields("factor"))
                    mytabley.Update
                Else
                    mytabley.Fields("cantidad") = Val("" & mytabley.Fields("cantidad")) + Val("" & mytablex.Fields("cantidad"))
                    mytabley.Update

                End If

                mytabley.Close
                mytablex.MoveNext
            Loop
            mytablex.Close

            '----------------------
        End If

        rexplorap.MoveNext
    Loop
    mytablex.Open "select * from tmpedido", cn, adOpenKeyset, adLockOptimistic
    Set dbgrid12.DataSource = mytablex
    dbgrid12.refresh

End Sub

Private Sub dki889343_Click()

    If Frame8.Visible = True Then Exit Sub

    If Frame7.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If MsgBox("Desea Exportar Excell", 1, "Aviso") <> 1 Then Exit Sub
    conteo_excell_uno

End Sub

Private Sub dkiewre_Click()

    'If Frame4.Visible = True Then Exit Sub
    If Frame8.Visible = True Then Exit Sub

    If Frame7.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub

    reporgen.NAMETABLA = cgusuario
    reporgen.Show 1

End Sub

Private Sub dkifor_Click()

    'If Frame4.Visible = True Then Exit Sub
    If Frame8.Visible = True Then Exit Sub

    If Frame7.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    proceso_impresion1

End Sub

Private Sub dl89er_Click()

    If Frame8.Visible = True Then Exit Sub

    If Frame7.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If MsgBox("Desea Exportar Excell", 1, "Aviso") <> 1 Then Exit Sub
    menu_excell

End Sub

Private Sub fdl89234_Click()

    Dim buf As String

    On Error GoTo cmd45112_err

    If Frame8.Visible = True Then Exit Sub
    If Frame7.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub
    buf = "" & rexplorap.Fields("local")

    If Trim("" & rexplorap.Fields("estado")) <> "2" Then
        MsgBox "Para este fin el estado debe estar en 2", 48, "Aviso"
        Exit Sub

    End If

    'Select Case acu
    '       Case "Z", "S", "T"
    '       Case Else: Exit Sub
    'End Select
    'MsgBox rexplorap.Fields(1)
    If Trim("" & rexplorap.Fields("yausado")) = "0" Then
        If MsgBox("Estado Actual:Pendiente " & Chr$(10) & Chr$(13) & "Cambiar a Atendido", 1, "Aviso") = 1 Then
            rexplorap.Fields("yausado") = "1"
            Exit Sub

        End If

    End If

    If Trim("" & rexplorap.Fields("yausado")) = "1" Then
        If MsgBox("Estado Actual:Atendido " & Chr$(10) & Chr$(13) & "Cambiar a Pendiente", 1, "Aviso") = 1 Then
            rexplorap.Fields("yausado") = "0"

        End If

    End If

    MsgBox "Proceso Terminado ", 48, "Aviso"
    Exit Sub
cmd45112_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fk4844_Click()

    On Error GoTo cmd8912_err

    Dim sw       As Integer

    Dim mytablex As New ADODB.Recordset

    If Frame8.Visible = True Then Exit Sub
    If Frame7.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub

    ' Testing Proyecto Facturacion Electronica Nota de Credito 06/2018
    'If "" & rexplorap.Fields("estado") <> "2" Then
    '   MsgBox "Estado debe estar en 2 para realizar esta operacion ", 48, "Aviso"
    '   Exit Sub
    'End If
    ' Testing Proyecto Facturacion Electronica Nota de Credito 06/2018

    'If "" & rexplorap.Fields("yausado") = "1" Then
    '   MsgBox "Ya fue Utilizado ", 48, "Aviso"
    '   Exit Sub
    'End If
    sw = 0
    List1.Clear

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
    'rexplorap.MoveFirst
    'Do
    'If rexplorap.EOF Then Exit Do
    'If "" & rexplorap.Fields("dflag") = "S" Then
    '   List1.AddItem "" & rexplorap.Fields("tipo") & " " & rexplorap.Fields("serie") & " " & rexplorap.Fields("numero") & " " & rexplorap.Fields("nombre")
    '   sw = 1
    'End If
    'rexplorap.MoveNext
    'Loop

    List1.AddItem "" & rexplorap.Fields("tipo") & " " & rexplorap.Fields("serie") & " " & rexplorap.Fields("numero") & " " & rexplorap.Fields("nombre")
    sw = 1
    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

    If sw = 0 Then
        MsgBox "No ha seleccionado ", 48, "Aviso"
        Exit Sub

    End If

    'If Len("" & rexplorap.Fields("serie1")) > 0 Then
    '   MsgBox "Documento ya generado " & "" & rexplorap.Fields("tipo1") & " " & rexplorap.Fields("serie1") & " " & rexplorap.Fields("numero1"), 48, "Aviso"
    '   Exit Sub
    'End If
    List1.ListIndex = 0

    gacu = ""
    gserie = ""
    gnumero = ""
    gtipo.Clear
    gtipo.AddItem "%"
    mytablex.Open "SELECT * FROM tipo  ", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        If ACU = "T" Then  'Guia remision
            If "" & mytablex.Fields("tipodoc") = "1" Or "" & mytablex.Fields("tipodoc") = "A" Or "" & mytablex.Fields("tipodoc") = "B" Or "" & mytablex.Fields("tipodoc") = "C" Or "" & mytablex.Fields("tipodoc") = "D" Or "" & mytablex.Fields("tipodoc") = "G" Or "" & mytablex.Fields("tipodoc") = "E" Or "" & mytablex.Fields("tipodoc") = "F" Then
                gtipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")

            End If

        End If

        '''21/08/2017 kenyo Guia de Salida con Factura
        ' If ACU = "V" Then  'Cotizacion
        '    If "" & mytablex.Fields("tipodoc") = "E" Or "" & mytablex.Fields("tipodoc") = "F" Or "" & mytablex.Fields("tipodoc") = "B" Or "" & mytablex.Fields("tipodoc") = "T" Then
        '   gtipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")
        '    End If
        'End If

        If ACU = "V" Then  'Cotizacion
    
            If "" & mytablex.Fields("tipodoc") = "E" Or "" & mytablex.Fields("tipodoc") = "F" Then
                If Mid(rexplorap.Fields("SERIE"), 1, 1) = Mid(mytablex.Fields("SERIE"), 1, 1) Then
                    gtipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")

                End If

            End If
    
            If "" & mytablex.Fields("tipodoc") = "B" Or "" & mytablex.Fields("tipodoc") = "T" Then
                gtipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")

            End If
    
        End If

        '''21/08/2017 kenyo Guia de Salida con Factura
        ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

        ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
        If ACU = "E" Then  'Nota de Credito
            If "" & mytablex.Fields("tipodoc") = "T" Then
                gtipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")

            End If

        End If

        If ACU = "F" Then  'Nota de Débito
            If "" & mytablex.Fields("tipodoc") = "S" Then
                gtipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")

            End If

        End If

        ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

        If ACU = "H" Then  'Cotizacion
            If "" & mytablex.Fields("tipodoc") = "1" Or "" & mytablex.Fields("tipodoc") = "A" Or "" & mytablex.Fields("tipodoc") = "B" Or "" & mytablex.Fields("tipodoc") = "C" Or "" & mytablex.Fields("tipodoc") = "D" Or "" & mytablex.Fields("tipodoc") = "G" Or "" & mytablex.Fields("tipodoc") = "E" Or "" & mytablex.Fields("tipodoc") = "F" Or "" & mytablex.Fields("tipodoc") = "I" Or "" & mytablex.Fields("tipodoc") = "T" Then
                gtipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")

            End If

        End If

        If ACU = "I" Then  'pedido
            If "" & mytablex.Fields("tipodoc") = "T" Or "" & mytablex.Fields("tipodoc") = "1" Or "" & mytablex.Fields("tipodoc") = "A" Or "" & mytablex.Fields("tipodoc") = "B" Or "" & mytablex.Fields("tipodoc") = "C" Or "" & mytablex.Fields("tipodoc") = "D" Or "" & mytablex.Fields("tipodoc") = "G" Or "" & mytablex.Fields("tipodoc") = "E" Or "" & mytablex.Fields("tipodoc") = "F" Or "" & mytablex.Fields("tipodoc") = "T" Then
                gtipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")

            End If

        End If

        If ACU = "R" Then  'orden de compra
            If "" & mytablex.Fields("tipodoc") = "J" Or "" & mytablex.Fields("tipodoc") = "K" Or "" & mytablex.Fields("tipodoc") = "L" Or "" & mytablex.Fields("tipodoc") = "M" Or "" & mytablex.Fields("tipodoc") = "P" Or "" & mytablex.Fields("tipodoc") = "N" Or "" & mytablex.Fields("tipodoc") = "O" Or "" & mytablex.Fields("tipodoc") = "S" Then
                gtipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")

            End If

        End If

        If ACU = "S" Then  'guia de compra
            If "" & mytablex.Fields("tipodoc") = "J" Or "" & mytablex.Fields("tipodoc") = "K" Or "" & mytablex.Fields("tipodoc") = "L" Or "" & mytablex.Fields("tipodoc") = "M" Or "" & mytablex.Fields("tipodoc") = "P" Or "" & mytablex.Fields("tipodoc") = "N" Or "" & mytablex.Fields("tipodoc") = "O" Then
                gtipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")

            End If

        End If

        If ACU = "C" Then  'factura de compra
            If "" & mytablex.Fields("tipodoc") = "O" Or "" & mytablex.Fields("tipodoc") = "N" Then
                gtipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")

            End If

        End If

        If ACU = "Q" Then  'nota pedido almacen
            If "" & mytablex.Fields("tipodoc") = "S" Or "" & mytablex.Fields("tipodoc") = "T" Or "" & mytablex.Fields("tipodoc") = "Z" Or "" & mytablex.Fields("tipodoc") = "R" Then
                gtipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")

            End If

        End If

        'If acu = "Q" Then  'nota pedido almacen
        '   If "" & mytablex.Fields("tipodoc") = "Z" Then 'Or "" & mytablex.Fields("tipodoc") = "S" Then
        '      gtipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")
        '
        '   End If
        'End If
        mytablex.MoveNext
    Loop
    mytablex.Close
    gtipo.ListIndex = 0
    Frame4.Visible = True
    'Command6.Caption = "Selecciona"
    gtipo.Enabled = True
    gtipo.SetFocus
    Exit Sub
cmd8912_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fk8844_Click()

    If Frame8.Visible = True Then Exit Sub

    If Frame7.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    Frame5.Visible = True
    ofechai = fechai
    ofechaf = fechaf
    ocajero = extra_loquesea(cajero)
    ocaja = extra_loquesea(caja)
    oturno = turno

End Sub

Private Sub Flo881_Click()

    Dim buf As String

    On Error GoTo cmd7900_err

    Dim mytablex As New ADODB.Recordset

    If Frame8.Visible = True Then Exit Sub

    If Frame7.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
   
    buf = " select * from fpagov where local='" & rexplorap.Fields("local") & "'"
    buf = buf & " and tipo = '" & rexplorap.Fields("tipo") & "'"
    buf = buf & " and serie = '" & rexplorap.Fields("serie") & "'"
    buf = buf & " and numero = '" & rexplorap.Fields("numero") & "'"
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    Set dbgrid33.DataSource = mytablex
    dbgrid33.refresh
    Frame6.Visible = True
    Exit Sub
cmd7900_err:
    MsgBox "Aviso en consulta Forma Pagos " + error$, 48, "Aviso"
    Frame6.Visible = False
    Exit Sub

End Sub

Private Sub Form_Activate()
    Frame1.Top = 10: Frame1.Left = 10
    Frame9.Top = 10: Frame9.Left = 10

    ' Testing Proyecto Facturacion Electronica 05/04/2018
    Frame8.Top = 10: Frame8.Left = 10
    Frame7.Top = 10: Frame7.Left = 10
    Frame6.Top = 10: Frame6.Left = 10
    Frame5.Top = 10: Frame5.Left = 10
    Frame4.Top = 10: Frame4.Left = 10
    Frame3.Top = 10: Frame3.Left = 10
    Frame2.Top = 10: Frame2.Left = 10
    FrmConsultaSunat.Top = 10: FrmConsultaSunat.Left = 10

    ' Testing Proyecto Facturacion Electronica 05/04/2018

    If Trim(flag_estado) <> "S" Then
        carga_iniciales

    End If

    flag_estado = "S"
    Check1.Visible = False

    'MsgBox acu
    Select Case ACU

        Case "V"
            Check1.Visible = True

    End Select

    If zooma = "Zomm" Then
        Frame3.Visible = False
        zooma = ""
        Exit Sub

    End If

    zooma = ""
    'If YacaRGA = "" Then
    sql_cabeza
    'End If

End Sub

Sub color_cambio()

End Sub

Private Sub Form_Load()

    Dim mytablex As New ADODB.Recordset

    PLACA.Clear
    PLACA.AddItem "%"
    PLACA.AddItem "ENTREGADO"
    PLACA.AddItem "PENDIENTE"
    PLACA.ListIndex = 0

    ordenado.Clear
    ordenado.AddItem "%"
    ordenado.AddItem "Local,Tipo,Serie,Numero"
    ordenado.AddItem "Local,Tipo,Serie,str(Numero)"
    ordenado.AddItem "Local,Tipo,Serie,Numero,Codigo"
    ordenado.AddItem "Fecha,Local,Tipo,Serie,Numero"
    ordenado.AddItem "Fecha,Local,Tipo,Serie,str(Numero)"
    ordenado.ListIndex = 0

    servicio.Clear
    servicio.AddItem "%"
    'servicio.AddItem "Deliveri"
    'servicio.AddItem "Autoservicio"
    'servicio.AddItem "Comanda"
    'servicio.ListIndex = 0
    turno.Clear
    turno.AddItem "%"
    turno.AddItem "1"
    turno.AddItem "2"
    turno.AddItem "3"
    turno.AddItem "4"
    turno.ListIndex = 0

    mytablex.Open "SELECT * FROM servicio ", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        servicio.AddItem "" & mytablex.Fields("servicio") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    servicio.ListIndex = 0

    Combo2.Clear
    Combo2.AddItem "Pendiente"
    Combo2.AddItem "Atendido"
    Combo2.AddItem "%"
    Combo2.ListIndex = 2

    moneda.Clear
    moneda.AddItem "%"
    moneda.AddItem "S"
    moneda.AddItem "D"
    moneda.ListIndex = 0
    estado.Clear
    estado.AddItem "%"
    estado.AddItem "2"
    estado.AddItem "1"
    estado.AddItem "0"

    fechai = Format(Now, "dd/mm/yyyy")
    fechaf = Format(Now, "dd/mm/yyyy")
    estado.ListIndex = 0
    'cmdGrabar_Click

End Sub

Sub carga_iniciales()

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    cajero.Clear
    cajero.AddItem "%"
    vendedor.Clear
    vendedor.AddItem "%"
    caja.Clear
    caja.AddItem "%"
    tipo.Clear
    tipo.AddItem "%"
    bodega.Clear
    bodega.AddItem "%"
    llocal1.Clear
    llocal1.AddItem "%"
    bodegaf.Clear
    bodegaf.AddItem "%"

    mytablex.Open "SELECT * FROM vendedor order by codigo", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        vendedor.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        cajero.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    vendedor.ListIndex = 0
    cajero.ListIndex = 0
    mytablex.Close

    mytablex.Open "SELECT * FROM tlocal order by codigo ", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        llocal1.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    llocal1.ListIndex = 0

    If llocal1.ListCount = 2 Then
        llocal1.ListIndex = 1

    End If

    mytablex.Close
    'MsgBox acu

    mytablex.Open "SELECT * FROM tipo " & buf, cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        If "" & mytablex.Fields("grupo") = ACU Then
            tipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")
        End If '

        mytablex.MoveNext
    Loop
    tipo.ListIndex = 0
    mytablex.Close

    mytablex.Open "SELECT * FROM bodega ", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        bodega.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        bodegaf.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")

        mytablex.MoveNext
    Loop
    bodega.ListIndex = 0
    bodegaf.ListIndex = 0

    mytablex.Close

    If llocal1 <> "%" Then
        buf = " where local='" & extra_loquesea(llocal1) & "'"

    End If

    mytablex.Open "SELECT * FROM parameca order by caja ", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        If "" & mytablex.Fields("terminal") = "C" Then
            caja.AddItem "" & mytablex.Fields("caja") & "|" & mytablex.Fields("descripcio")

        End If

        mytablex.MoveNext
    Loop
    caja.ListIndex = 0
    mytablex.Close

End Sub

Private Sub gtipo_Click()

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Double

    If gtipo = "%" Then
        gacu = ""
        gserie = ""
        gnumero = ""
        gacu = ""
        Exit Sub

    End If

    mytablex.Open "SELECT * FROM tipo where tipo='" & extra_loquesea(gtipo) & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        gserie = ""
        gnumero = ""
        gacu = ""
        mytablex.Close
        'gserie.SetFocus
        Exit Sub

    End If

    gacu = "" & mytablex.Fields("tipodoc")

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
    'gserie = "" & mytablex.Fields("serie")
    If rexplorap.Fields("tipo") = "1" Then
        gserie = "" & mytablex.Fields("serie")
    ElseIf rexplorap.Fields("tipo") = "2" Then
        gserie = "" & mytablex.Fields("serie")
    Else
        gserie = "" & mytablex.Fields("serie")

    End If

    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

    sdx = Val("" & mytablex.Fields("numero")) + 1
    gnumero = "" & sdx
    mytablex.Close
    'gtipo.Enabled = False
    Exit Sub

End Sub

Private Sub GTTR_Click()
    sihue

End Sub

Private Sub impso02_Click()

    If Frame8.Visible = True Then Exit Sub

    If Frame7.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub
    menu_excell1

End Sub

Private Sub ju7881_Click()

    If ACU = "Q" Then
        tload.Label12.Caption = "REQUERIMIENTO"
        tload.fechai = fechai
        tload.Show 1

    End If
        
End Sub

Private Sub Label22_Click()

    If Not IsDate(ofechai) Then Exit Sub
    If Not IsDate(ofechai) Then Exit Sub
    If ocajero = "%" Then Exit Sub
    If ocaja = "%" Then Exit Sub
    If oturno = "%" Then Exit Sub
    proceso_impresion2

End Sub

Private Sub Label23_Click()

    If Not IsDate(ofechai) Then Exit Sub
    If Not IsDate(ofechai) Then Exit Sub
    If ocajero = "%" Then Exit Sub
    If ocaja = "%" Then Exit Sub
    If oturno = "%" Then Exit Sub

    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub
    If cajero = "%" Then Exit Sub
    If caja = "%" Then Exit Sub
    If turno = "%" Then Exit Sub

    opcion2 = "1"
    opcion1 = "1"
    opcion3 = "2"
    tcuadrc1.fechai.Enabled = True
    tcuadrc1.fechaf.Enabled = True

    usuariopos = gusuario
    tcuadrc1.flagdiario = "1"
    tcuadrc1.cajero = extra_loquesea(cajero)
    tcuadrc1.caja = extra_loquesea(caja)
    tcuadrc1.turno = turno
    tcuadrc1.fechai = Format(fechai, "dd/mm/yyyy")
    tcuadrc1.fechaf = Format(fechaf, "dd/mm/yyyy")
    tcuadrc1.horai = "01"
    tcuadrc1.horaf = "24"
    tcuadrc1.Caption = "CUADRE PARCIAL DEL DIA"
    tcuadrc1.Show 1

End Sub

Private Sub Label24_Click()

    If Not IsDate(ofechai) Then Exit Sub
    If Not IsDate(ofechai) Then Exit Sub
    If ocajero = "%" Then Exit Sub
    If ocaja = "%" Then Exit Sub
    If oturno = "%" Then Exit Sub

    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub
    If cajero = "%" Then Exit Sub
    If caja = "%" Then Exit Sub
    If turno = "%" Then Exit Sub

    opcion1 = "5"
    opcion2 = "1"
    opcion3 = "2"
    
    tcuadrc1.fechai.Enabled = True
    tcuadrc1.fechaf.Enabled = True
    usuariopos = gusuario
    tcuadrc1.tipoexterno.Visible = True
    tcuadrc1.numcuadre.Visible = True
    'tcuadrc1.flagdiario = "1"
    tcuadrc1.cajero = extra_loquesea(cajero)
    tcuadrc1.caja = extra_loquesea(caja)
    tcuadrc1.turno = turno
    
    tcuadrc1.fechai = Format(fechai, "dd/mm/yyyy")
    tcuadrc1.fechaf = Format(fechaf, "dd/mm/yyyy")
    tcuadrc1.horai = "01"
    tcuadrc1.horaf = "24"
    tcuadrc1.Caption = "COPIA CIERRE DEL DIA"
    tcuadrc1.pantalla = "PANTALLA"
    tcuadrc1.Show 1

End Sub

Private Sub Label27_Click()
    Frame5.Visible = False

End Sub

Private Sub ldo33_Click()

    If Frame8.Visible = True Then
        Frame8.Visible = False
        dbgrid2.SetFocus
        Exit Sub

    End If

    If Frame7.Visible = True Then
        Frame7.Visible = False
        dbgrid2.SetFocus
        Exit Sub

    End If

    If Frame6.Visible = True Then
        Frame6.Visible = False
        dbgrid2.SetFocus
        Exit Sub

    End If

    If Frame5.Visible = True Then
        Frame5.Visible = False
        dbgrid2.SetFocus
        Exit Sub

    End If

    If Frame3.Visible = True Then
        Frame3.Visible = False
        dbgrid2.SetFocus
        Exit Sub

    End If

    If Frame2.Visible = True Then
        Frame2.Visible = False
        dbgrid2.SetFocus
        Exit Sub

    End If

    If opcion1 = "1" Then
        If Frame1.Visible = True Then
            Frame1.Visible = False
            codigo.SetFocus
            Exit Sub

        End If

    End If

    If opcion1 = "6100" Then
        If Frame1.Visible = True Then
            Frame1.Visible = False
            Frame1.Enabled = False
            'codigo.SetFocus
            Exit Sub

        End If

    End If

    If opcion1 = "2" Then
        If Frame1.Visible = True Then
            Frame1.Visible = False
            dbgrid2.SetFocus
            Exit Sub

        End If

    End If

    explorap.Hide
    Unload explorap

End Sub

Sub sql_cabeza()

    Dim buf  As String

    Dim buf2 As String

    On Error GoTo cmd921_err

    'MsgBox caja
    'MsgBox fechai
    If Len(fechai) <> 10 Then Exit Sub
    If Len(fechaf) <> 10 Then Exit Sub
    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub
    'MsgBox cgusuario
    buf = "select * from " & cgusuario & " where "

    If ve = "V" Then
        buf = buf & "  fechae>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fechae<='" & Format(fechaf, "YYYYMMDD") & "' "
    Else
        buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & "  and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    End If

    buf2 = ""

    buf = buf & buf2

    If Trim(llocal1) <> "%" Then
        buf = buf & " and local like '" & extra_loquesea(llocal1) & "'"

    End If

    If Trim(tipo) <> "%" Then
        buf = buf & " and tipo like '" & extra_loquesea(tipo) & "'"

    End If

    If Trim(caja) <> "%" Then
        buf = buf & " and caja like '" & extra_loquesea(caja) & "'"

    End If

    If Trim(turno) <> "%" Then
        buf = buf & " and turno like '" & turno & "'"

    End If

    If serie <> "%" Then
        buf = buf & " and serie like '" & serie & "'"

    End If

    If Numero <> "%" Then
        buf = buf & " and numero like '" & Numero & "'"

    End If

    If codigo <> "%" Then
        buf = buf & " and codigo like '" & codigo & "'"

    End If

    If nombre <> "%" Then
        buf = buf & " and nombre like '" & nombre & "'"

    End If

    If moneda <> "%" Then
        buf = buf & " and moneda like '" & moneda & "'"

    End If

    If estado <> "%" Then
        buf = buf & " and estado like '" & estado & "'"

    End If

    If Trim(PLACA) <> "%" Then
        buf = buf & " and placa='" & Trim(PLACA) & "'"

    End If

    ' Testing Proyecto Facturacion Electronica 05/04/2018
    'my_estado = "%"
    If chkMostrarSoloAnulados.Value = 1 Then
        buf = buf & " and estado=1 and  estado_sunat='PENDIENTE'"

    End If

    ' Testing Proyecto Facturacion Electronica 05/04/2018

    If Trim(vendedor) <> "%" Then
        buf = buf & " and vendedor like '" & extra_loquesea(vendedor) & "'"

    End If

    If Trim(cajero) <> "%" Then
        buf = buf & " and usuario like '" & extra_loquesea(cajero) & "'"

    End If

    If Trim(bodega) <> "%" Then
        buf = buf & " and bodega like '" & extra_loquesea(bodega) & "'"

    End If

    If Trim(bodegaf) <> "%" Then
        buf = buf & " and bodegaf like '" & extra_loquesea(bodegaf) & "'"

    End If

    If Trim(servicio) <> "%" Then
        'If servicio = "Deliveri" Then
        buf = buf & " and  servicio='" & extra_loquesea(servicio) & "'"

        'End If
        'If servicio = "Comanda" Then
        '   buf = buf & " and  servicio='C' "
        'End If
        'If servicio = "Autoservicio" Then
        '   buf = buf & " and  servicio='A' "
        'End If
    End If

    If saldoini.Value = 1 Then
        buf = buf & " and nop='S' "

    End If

    If ACU <> "C" And ACU <> "V" Then
        buf = buf & " and acu='" & ACU & "'"

    End If

    If Combo2 <> "%" Then
        If Combo2 = "Atendido" Then
            buf = buf & " and  yausado='1'"

        End If

        If Combo2 = "Pendiente" Then
            buf = buf & " and  yausado='0'"

        End If

    End If

    If ACU = "V" Then

        '19/06/2017 kenyo NOTA DE CREDITO
        'buf = buf & " and (acu='A' OR acu='B' OR acu='C' OR acu='D' OR acu='G' )"
        buf = buf & " and (acu='A' OR acu='B' OR acu='C' OR acu='D' OR acu='G' or acu='E' )"

        '19/06/2017 kenyo NOTA DE CREDITO
   
        If Check1.Value = 1 Then
            buf = buf & " and tipo<>'5'"

        End If

    End If

    If ACU = "C" Then
        buf = buf & " and (acu='J' OR acu='K' OR acu='L' OR acu='M' OR acu='P' )"

        'If Check1.Value = 1 Then
        '   buf = buf & " and tipo<>'5'"
        'End If
    End If

    If ACU <> "Z" Then

        'buf = buf & " and importacio<>'S' "
    End If

    If ACU = "Q" Then

        'buf = buf & " and tipoclie='V'"
    End If

    If tinterno = "S" Then

        'buf = buf & " and tipoclie='V'"
        'Else
        'buf = buf & " and tipoclie<>'I'"
    End If

    'If importacion = "COMERCIAL" Then
    '   buf = buf & " and tipoimp='C'"
    'End If
    'If importacion = "GASTO" Then
    '   buf = buf & " and tipoimp='G'"
    'End If
    If importacion = "IMPORTACION" Then
        buf = buf & " and tipoimp='I'"

    End If

    If importacion = "GASTOS" Then
        buf = buf & " and tipoimp='G'"

    End If

    If importacion = "COMERCIAL" Then
        buf = buf & " and (tipoimp='C' or tipoimp is null) "

    End If

    If ordenado <> "%" Then
        buf = buf & "order by " & ordenado

    End If

    'MsgBox buf
    If rexplorap.State = 1 Then rexplorap.Close
    rexplorap.Open buf, cn, adOpenStatic, adLockOptimistic
    'MsgBox ""
    Set dbgrid2.DataSource = rexplorap
    dbgrid2.refresh
   
    'If rexplorap.EOF = True And rexplorap.BOF = True Then
    'rconsulta.Close
    'buffer.SetFocus
    'Exit Sub
    'End If
   
    If rexplorap.RecordCount > 0 Then

        '   dbgrid2.Col = 0
        '   dbgrid2.Row = dbgrid2.VisibleRows - 1
        '   dbgrid2.SetFocus
    End If

    'Data2.Connect = "foxpro 2.5;"
    'Data2.DatabaseName = globaldir
    'Data2.RecordSource = buf
    'Data2.Refresh
    'SUMAR_CABEZA rexplorap
    'ir_ultimo
               
    'MsgBox "xxx"
    'If rexplorap.RecordCount > 0 Then
    '   dbgrid2.Col = 0
    '   dbgrid2.Row = dbgrid2.VisibleRows - 1
    '   dbgrid2.SetFocus
    'End If
               
    Exit Sub
cmd921_err:
    MsgBox "aviso en sql_cabeza   " & error$, 48, "Aviso"
    Exit Sub

End Sub

Sub SUMAR_CABEZA()

    Dim sdx         As Double

    Dim sdx1        As Double

    Dim sdx2        As Double

    Dim sdx3        As Double

    Dim sdx4        As Double

    Dim sdx5        As Double

    Dim sdx6        As Double

    Dim ssdx1       As Double

    Dim ssdx2       As Double

    Dim ssdx3       As Double

    Dim ssdx4       As Double

    Dim ssdx5       As Double

    Dim ssdx6       As Double

    Dim xanulado    As Double

    Dim xvendido    As Double

    Dim xnogravado  As Double

    Dim xmanulado   As Double

    Dim xmvendido   As Double

    Dim xmnogravado As Double

    On Error GoTo cmd7812_err

    xanulado = 0
    xvendido = 0
    xnogravado = 0

    xmanulado = 0
    xmvendido = 0
    xmnogravado = 0

    sdx1 = 0
    sdx2 = 0
    sdx3 = 0
    sdx4 = 0
    sdx5 = 0
    sdx6 = 0

    ssdx1 = 0
    ssdx2 = 0
    ssdx3 = 0
    ssdx4 = 0
    ssdx5 = 0
    ssdx6 = 0

    netos = ""
    lanulado = ""
    lvendido = ""
    lnogravado = ""

    servicios = ""
    percepcions = ""
    netos = ""
    subtotals = ""
    impuestos = ""
    totals = ""

    serviciod = ""
    percepciond = ""
    netod = ""
    subtotald = ""
    impuestod = ""
    totald = ""

    If rexplorap.RecordCount = 0 Then
        Exit Sub

    End If

    rexplorap.MoveFirst
    Do

        If rexplorap.EOF Or rexplorap.BOF Then Exit Do
        If "" & rexplorap.Fields("estado") = "1" Then
            xanulado = xanulado + 1
            xmanulado = xmanulado + Val("" & rexplorap.Fields("total"))

        End If

        If "" & rexplorap.Fields("estado") = "0" Then
            xnogravado = xnogravado + 1
            xmnogravado = xmnogravado + Val("" & rexplorap.Fields("total"))

        End If

        If "" & rexplorap.Fields("estado") = "2" Then
            xvendido = xvendido + 1
            xmvendido = xmvendido + Val("" & rexplorap.Fields("total"))

            If Trim("" & rexplorap.Fields("moneda")) = "S" Then
                sdx1 = sdx1 + Val("" & rexplorap.Fields("neto"))
                sdx2 = sdx2 + Val("" & rexplorap.Fields("impuesto"))
                sdx3 = sdx3 + Val("" & rexplorap.Fields("subtotal"))
                sdx4 = sdx4 + Val("" & rexplorap.Fields("servicioco"))
                sdx5 = sdx5 + Val("" & rexplorap.Fields("total"))
                sdx6 = sdx6 + Val("" & rexplorap.Fields("percepcion"))

            End If

            If "" & rexplorap.Fields("moneda") = "D" Then
                ssdx1 = ssdx1 + Val("" & rexplorap.Fields("neto"))
                ssdx2 = ssdx2 + Val("" & rexplorap.Fields("impuesto"))
                ssdx3 = ssdx3 + Val("" & rexplorap.Fields("subtotal"))
                ssdx4 = ssdx4 + Val("" & rexplorap.Fields("servicioco"))
                ssdx5 = ssdx5 + Val("" & rexplorap.Fields("total"))
                ssdx6 = ssdx6 + Val("" & rexplorap.Fields("percepcion"))

            End If

        End If

        rexplorap.MoveNext
    Loop
    servicios = Format(sdx4, "0.00")
    percepcions = Format(sdx6, "0.00")
    netos = Format(sdx1, "0.00")
    subtotals = Format(sdx3, "0.00")
    impuestos = Format(sdx2, "0.00")
    totals = Format(sdx5, "0.00")

    serviciod = Format(ssdx4, "0.00")
    percepciond = Format(ssdx6, "0.00")
    netod = Format(ssdx1, "0.00")
    subtotald = Format(ssdx3, "0.00")
    impuestod = Format(ssdx2, "0.00")
    totald = Format(ssdx5, "0.00")

    lanulado = "" & xanulado
    lvendido = "" & xvendido
    lnogravado = "" & xnogravado

    lvanulado = "" & xmanulado
    lvvendido = "" & xmvendido
    lvnogravado = "" & xmnogravado

    Exit Sub
cmd7812_err:
    MsgBox "Error en Suma" & error$, 48, "Aviso"
    Exit Sub

End Sub

Sub ir_inicio()

End Sub

Sub consulta_codigo()
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Telefono"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "1"
    Command1_Click

End Sub

Sub consulta_local()
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    buffer = ""
    Frame1.Enabled = True
    buffer.SetFocus
    opcion1 = "6100"
    Command1_Click

End Sub

Sub consulta_detalle()
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Telefono"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "2"
    Command1_Click

End Sub

Sub ir_ultimo()

    On Error GoTo cmd123_err

    'Data2.Recordset.MoveLast
    Exit Sub
cmd123_err:
    Exit Sub

End Sub

Sub proceso_impresion1()

    Dim found    As Integer

    Dim archivot As String

    Dim ttipo    As String

    Dim tserie   As String

    Dim local1   As String

    Dim tnumero  As String

    On Error GoTo cmd6_err:

    local1 = "" & rexplorap.Fields("local")
    ttipo = "" & rexplorap.Fields("tipo")
    tserie = "" & rexplorap.Fields("serie")
    tnumero = "" & rexplorap.Fields("numero")
    cerrar_archivo
    factura_formato local1, "" & ttipo, "" & tserie, "" & tnumero, "", 0
    cerrar_archivo
    'MsgBox gusuario
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1
    Exit Sub
cmd6_err:
    MsgBox "Mensaje, Error al iniciar Impresion " & error$
    Exit Sub

End Sub

Sub proceso_impresion2()

    Dim found    As Integer

    Dim archivot As String

    Dim ttipo    As String

    Dim tserie   As String

    Dim local1   As String

    Dim tnumero  As String

    Dim FileName As String

    On Error GoTo cmd66_err:

    If rexplorap.RecordCount = 0 Then Exit Sub
    rexplorap.MoveFirst
   
    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    found = borra_nombre("" & FileName)
    gofpago = "fpagov"
    dgusuariog = "detalle"
    cgusuario = "factura"
       
    Do

        If rexplorap.EOF Then Exit Do
        local1 = "" & rexplorap.Fields("local")
        ttipo = "" & rexplorap.Fields("tipo")
        tserie = "" & rexplorap.Fields("serie")
        tnumero = "" & rexplorap.Fields("numero")
        cerrar_archivo
        factura_formato "" & local1, "" & ttipo, "" & tserie, "" & tnumero, "", 1
        rexplorap.MoveNext
    Loop
    cerrar_archivo
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1
    Exit Sub
cmd66_err:
    MsgBox "Mensaje, Error al iniciar Impresion " & error$
    Exit Sub

End Sub

Sub desmarca_documento()

    Dim buf1  As String

    Dim te    As String

    Dim ts    As String

    Dim found As Integer

    On Error GoTo cmd57_err

    found = valida_flag("" & rexplorap.Fields("acu"))

    If found = 0 Then

    End If

    If found = 1 Or found = 2 Then
        ' 05/06/207 kenyo cambio orden de compra inventario
        ' If Len(Trim("" & rexplorap.Fields("tipo1"))) = 0 And Len(Trim("" & rexplorap.Fields("serie1"))) = 0 And Len(Trim("" & rexplorap.Fields("numero1"))) = 0 Then
        descarga_saldo Trim("" & rexplorap.Fields("local")), Trim("" & rexplorap.Fields("tipo")), Trim("" & rexplorap.Fields("serie")), Trim("" & rexplorap.Fields("numero")), 1, "", "" & rexplorap.Fields("tipo1")

        ' End If
    End If

    If found = 3 Then  'si es traslado
        If Len(Trim("" & rexplorap.Fields("tipo1"))) = 0 And Len(Trim("" & rexplorap.Fields("serie1"))) = 0 And Len(Trim("" & rexplorap.Fields("numero1"))) = 0 Then
      
            descarga_saldo Trim("" & rexplorap.Fields("local")), "TE", Trim("" & rexplorap.Fields("serie")), Trim("" & rexplorap.Fields("numero")), 1, "1", "" & rexplorap.Fields("local")
            descarga_saldo Trim("" & rexplorap.Fields("local")), "TS", Trim("" & rexplorap.Fields("serie")), Trim("" & rexplorap.Fields("numero")), 1, "1", "" & rexplorap.Fields("local")
   
            borra_detalle Trim("" & rexplorap.Fields("local")), "TE", Trim("" & rexplorap.Fields("serie")), Trim("" & rexplorap.Fields("numero"))
            borra_detalle Trim("" & rexplorap.Fields("local")), "TS", Trim("" & rexplorap.Fields("serie")), Trim("" & rexplorap.Fields("numero"))
   
            ''''04/10/2017 kenyo Correcion duplicidad de traslados
            descarga_saldo Trim("" & rexplorap.Fields("localf")), "TE", Trim("" & rexplorap.Fields("serie")), Trim("" & rexplorap.Fields("numero")), 1, "1", "" & rexplorap.Fields("localf")
            descarga_saldo Trim("" & rexplorap.Fields("localf")), "TS", Trim("" & rexplorap.Fields("serie")), Trim("" & rexplorap.Fields("numero")), 1, "1", "" & rexplorap.Fields("localf")
   
            borra_detalle Trim("" & rexplorap.Fields("localf")), "TE", Trim("" & rexplorap.Fields("serie")), Trim("" & rexplorap.Fields("numero"))
            borra_detalle Trim("" & rexplorap.Fields("localf")), "TS", Trim("" & rexplorap.Fields("serie")), Trim("" & rexplorap.Fields("numero"))
   
            ''''04/10/2017 kenyo Correcion duplicidad de traslados
   
        End If

    End If

    'MsgBox ""

    'buf1 = " and acu='" & Trim("" & rexplorap.fields("acu")) & "'"
    buf1 = "update  " & dgusuariog & " set estado='0'  where local='" & Trim("" & rexplorap.Fields("local")) & "' and  tipo='" & Trim("" & rexplorap.Fields("tipo")) & "' and serie='" & Trim("" & rexplorap.Fields("serie")) & "' and  numero='" & Trim("" & rexplorap.Fields("numero")) & "'" & " and acu='" & Trim("" & rexplorap.Fields("acu")) & "'"
    cn.Execute (buf1)

    'borra importaciones
    If importacion = "IMPORTACION" Then
        'borratempdao
        'cn.Execute ("select * into _i" & gusuario & " from gastofactura where   tipo='" & Trim("" & rexplorap.Fields("tipo")) & "' and serie='" & Trim("" & rexplorap.Fields("serie")) & "' and  numero='" & Trim("" & rexplorap.Fields("numero")) & "'")
        cn.Execute ("update gastofactura set estado='0' where   tipo='" & Trim("" & rexplorap.Fields("tipo")) & "' and serie='" & Trim("" & rexplorap.Fields("serie")) & "' and  numero='" & Trim("" & rexplorap.Fields("numero")) & "'")

    End If

    'adicionamos la desmarcacion de las guias
    desmarca_yausado "" & rexplorap.Fields("LOCAL"), "" & rexplorap.Fields("tipo"), "" & rexplorap.Fields("SERIE"), "" & rexplorap.Fields("numero")
    'MsgBox ""
    cn.Execute ("update  fpagov  set estado='0'  where  local='" & Trim("" & rexplorap.Fields("local")) & "' and  tipo='" & Trim("" & rexplorap.Fields("tipo")) & "' and serie='" & Trim("" & rexplorap.Fields("serie")) & "' and  numero='" & Trim("" & rexplorap.Fields("numero")) & "'" & " and acu='" & Trim("" & rexplorap.Fields("acu")) & "'")
    cn.Execute ("update  recibo  set usado='N'  where  local='" & "" & rexplorap.Fields("local") & "' and tipo='" & "" & rexplorap.Fields("retipo1") & "' and numero='" & "" & rexplorap.Fields("renumero1") & "'")
    cn.Execute ("update  recibo  set usado='N'  where  local='" & "" & rexplorap.Fields("local") & "' and tipo='" & "" & rexplorap.Fields("retipo1") & "' and numero='" & "" & rexplorap.Fields("renumero2") & "'")
    cn.Execute ("update  recibo  set usado='N'  where  local='" & "" & rexplorap.Fields("local") & "' and tipo='" & "" & rexplorap.Fields("retipo1") & "' and numero='" & "" & rexplorap.Fields("renumero3") & "'")

    'MsgBox ""
    If ACU = "Z" Then
        cn.Execute ("DELETE FROM detallE where local='" & "" & rexplorap.Fields("local") & "' and tipo='" & te & "'and serie='" & "" & rexplorap.Fields("serie") & "'  and numero='" & "" & rexplorap.Fields("numero") & "TE" & "'")
        cn.Execute ("DELETE FROM detallE where local='" & "" & rexplorap.Fields("local") & "' and tipo='" & ts & "'and serie='" & "" & rexplorap.Fields("serie") & "'  and numero='" & "" & rexplorap.Fields("numero") & "TS" & "'")

    End If

    If "" & rexplorap.Fields("acu") = "I" Then
        graba_acumulado_clientes "" & rexplorap.Fields("codigo"), -1, Val("" & rexplorap.Fields("total"))

    End If
 
    If valida_flag("" & "" & rexplorap.Fields("acu")) = 1 Or valida_flag("" & "" & rexplorap.Fields("acu")) = 2 Then  'compras o ventas
        found = desgraba_cuentac()

    End If

    MsgBox "Desmarcacion SatisfactoriaK ", 48, "Aviso"

    sql_cabeza

    Exit Sub
cmd57_err:
    MsgBox "Aviso en desmarca documento " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub graba_acumulado_clientes(buf As String, signo As Double, sumador As Double)

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Double

    mytablex.Open "SELECT * FROM clientes where  codigo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        sdx = Val("" & mytablex.Fields("pedido")) + signo * sumador
        mytablex.Fields("pedido") = sdx
        mytablex.Update

    End If

    mytablex.Close

End Sub

Function valida_flag(buf As String)

    Select Case buf

        Case "Z"
            valida_flag = 3

        Case "T", "A", "B", "C", "D", "G", "E", "F"
            valida_flag = 1

        Case "S", "J", "K", "L", "M", "P", "N", "O"
            valida_flag = 2

    End Select

End Function

Function busca_tipo1(sw As Integer) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM tipo where  tipo='" & Trim("" & rexplorap.Fields("tipo")) & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        If sw = 0 Then
            busca_tipo1 = "" & mytablex.Fields("te")

        End If

        If sw = 1 Then
            busca_tipo1 = "" & mytablex.Fields("ts")

        End If

    End If

    mytablex.Close

End Function

Sub descarga_saldo(xlocal As String, _
                   xtipo As String, _
                   xserie As String, _
                   xnumero As String, _
                   sw As Integer, _
                   tipoarch As String, _
                   xtipo1 As String)

    Dim sdx       As Double

    Dim signo     As Double

    Dim sww       As Integer

    Dim mytablefa As New ADODB.Recordset

    Dim mytablex  As New ADODB.Recordset

    Dim mytabley  As New ADODB.Recordset

    Dim buf       As String

    Dim found     As Integer

    On Error GoTo cmd19_err

    sww = 0
    'AHORA HAY QUE VALIDAR QUE no existe ya cruzado el documento----
    mytablefa.Open "SELECT * FROM " & cgusuario & " where  local='" & xlocal & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablefa.RecordCount > 0 Then  'si existe
   
        '25/06/2018 Testing Almacen General
        'If Len(xtipo1) > 0 Then
        '   found = ve_descarga(xtipo1)
        '     If found = 1 Then
        '     sww = 1
        '     End If
        'End If
        If Len(xtipo1) > 0 Then
            found = ve_descarga(xtipo, xtipo1)

            If found = 1 Then
                sww = 1
            Else
                sww = 0

            End If

        End If

        '25/06/2018 Testing Almacen General

        ''' 29/11/2017 Correción  General del Stock
        If busca_tipoacu(xtipo1) = "T" Then  ' T es Guia de Salida
            sww = 0

        End If

        ''' 29/11/2017 Correción  General del Stock

    End If

    buf = dgusuariog

    If tipoarch = "1" Then
        buf = "detalle"

    End If

    mytablex.Open "SELECT * FROM " & buf & " where  local='" & xlocal & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then 'si existe
        Exit Sub

    End If

    'MsgBox ""
    'If permite_entrada_salida("" & mytablex.Fields("acu1")) = 1 Then 'si existe acu1 no descontar
    '   Exit Sub
    'End If
    Do

        If mytablex.EOF Then Exit Do
        '-------------------------------------------------
        signo = 1

        Select Case "" & mytablex.Fields("acu")

            Case "S", "J", "K", "L", "M", "P", "E"
                signo = 1

                '25/06/2018 Testing Almacen General
                'Case "T", "A", "B", "C", "D", "G", "N"
            Case "T", "A", "B", "C", "D", "G", "N", "F"
                '25/06/2018 Testing Almacen General
                signo = -1

        End Select

        'MsgBox signo
        If "" & mytablex.Fields("acu") = "J" Or "" & mytablex.Fields("acu") = "K" Or "" & mytablex.Fields("acu") = "L" Or "" & mytablex.Fields("acu") = "M" Then 'compras varia el precios y costo

            'graba_costos mytablex
        End If
      
        '-------------------------------------------------
        'busden:
        If sww = 0 Then
            If mytabley.State = 1 Then mytabley.Close
            mytabley.Open "select * from almacen where local='" & Trim("" & mytablex.Fields("local")) & "' and producto='" & Trim("" & mytablex.Fields("producto")) & "' and bodega='" & Trim("" & mytablex.Fields("bodega")) & "'", cn, adOpenDynamic, adLockOptimistic 'adOpenKeyset, adLockOptimistic

            'MsgBox mytabley.RecordCount
            If mytabley.RecordCount = 0 Then 'si existe
                'MsgBox ""
                mytabley.AddNew
                mytabley.Fields("local") = "" & mytablex.Fields("local")
                mytabley.Fields("producto") = "" & mytablex.Fields("producto")
                mytabley.Fields("bodega") = "" & mytablex.Fields("bodega")
     
                sdx = Val("" & mytabley.Fields("saldo")) + signo * Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
      
                'MsgBox sdx
                mytabley.Fields("saldo") = sdx
                mytabley.Update
            Else

                If sw = 0 Then
                    'mytabley.Edit
                    'MsgBox ""
                    sdx = Val("" & mytabley.Fields("saldo")) + signo * Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
                    'MsgBox sdx
                    mytabley.Fields("saldo") = sdx
                    decarga_saldo_talla mytabley, mytablex, signo
                    mytabley.Update

                End If

                If sw = 1 Then
                    'mytabley.Edit
         
                    sdx = Val("" & mytabley.Fields("saldo")) - signo * Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
                    mytabley.Fields("saldo") = sdx
                    decarga_saldo_talla mytabley, mytablex, signo
                    mytabley.Update

                End If

                '-------------------------------
            End If

        End If 'fin sw sw

        '-------------------------------------------------
        mytablex.MoveNext
    Loop
    Exit Sub
cmd19_err:
    MsgBox "Aviso en descarga saldo " + error$, 48, "Aviso"
    Exit Sub

End Sub

'' 29/11/2017 Correción  General del Stock
Function busca_tipoacu(ByRef valor As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "Select tipodoc from tipo where tipo='" & "" & valor & "'  ", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_tipoacu = "" & mytablex.Fields("tipodoc")

    End If

    mytablex.Close

End Function

'' 29/11/2017 Correción  General del Stock

Sub borra_detalle(xlocal As String, xtipo As String, xserie As String, xnumero As String)
    cn.Execute ("delete from detalle where local='" & xlocal & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'")

End Sub

Sub desmarca_yausado(buf0 As String, buf1 As String, buf2 As String, buf3 As String)

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    On Error GoTo cmd333_err

    buf = "update " & cgusuario & " set estado='0' where local='" & buf0 & "' and tipo='" & buf1 & "' and serie='" & buf2 & "' and numero='" & buf3 & "'"
    cn.Execute (buf)

    mytablex.Open "SELECT * FROM " & cgusuario & " where  local='" & buf0 & "' and tipo='" & buf1 & "' and serie='" & buf2 & "' and numero='" & buf3 & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        descarga_el_uso "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo1"), "" & mytablex.Fields("serie1"), "" & mytablex.Fields("numero1"), "0"
        descarga_el_uso "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo1"), "" & mytablex.Fields("serie2"), "" & mytablex.Fields("numero2"), "0"
        descarga_el_uso "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo1"), "" & mytablex.Fields("serie3"), "" & mytablex.Fields("numero3"), "0"
        descarga_el_uso "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo1"), "" & mytablex.Fields("serie4"), "" & mytablex.Fields("numero4"), "0"
        descarga_el_uso "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo1"), "" & mytablex.Fields("serie5"), "" & mytablex.Fields("numero5"), "0"
        descarga_el_uso "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo1"), "" & mytablex.Fields("serie6"), "" & mytablex.Fields("numero6"), "0"
        descarga_el_uso "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo1"), "" & mytablex.Fields("serie7"), "" & mytablex.Fields("numero7"), "0"

    End If

    '------------------------------------- ------------
    mytablex.Close
    Exit Sub
cmd333_err:
    MsgBox "Aviso en desmarca ya usado " + error$, 48, "Aviso"
    Exit Sub
 
End Sub

Sub descarga_el_uso(buf0 As String, _
                    buf1 As String, _
                    buf2 As String, _
                    buf3 As String, _
                    xsw As String)

    If Len(buf1) = 0 Then Exit Sub
    If Len(buf2) = 0 Then Exit Sub
    If Len(buf3) = 0 Then Exit Sub
    cn.Execute ("update " & cgusuario & " set yausado=" & xsw & " where local='" & buf0 & "' and tipo='" & buf1 & "' and serie='" & buf2 & "' and numero='" & buf3 & "'")

End Sub

Sub decarga_saldo_talla(mytablex As ADODB.Recordset, _
                        mytabley As ADODB.Recordset, _
                        signo As Double)

    Dim sdx As Double

    sdx = Val("" & mytablex.Fields("t1")) + signo * Val("" & mytabley.Fields("t1"))
    mytablex.Fields("t1") = sdx
    sdx = Val("" & mytablex.Fields("t2")) + signo * Val("" & mytabley.Fields("t2"))
    mytablex.Fields("t2") = sdx
    sdx = Val("" & mytablex.Fields("t3")) + signo * Val("" & mytabley.Fields("t3"))
    mytablex.Fields("t3") = sdx
    sdx = Val("" & mytablex.Fields("t4")) + signo * Val("" & mytabley.Fields("t4"))
    mytablex.Fields("t4") = sdx
    sdx = Val("" & mytablex.Fields("t5")) + signo * Val("" & mytabley.Fields("t5"))
    mytablex.Fields("t5") = sdx
    sdx = Val("" & mytablex.Fields("t6")) + signo * Val("" & mytabley.Fields("t6"))
    mytablex.Fields("t6") = sdx
    sdx = Val("" & mytablex.Fields("t7")) + signo * Val("" & mytabley.Fields("t7"))
    mytablex.Fields("t7") = sdx
    sdx = Val("" & mytablex.Fields("t8")) + signo * Val("" & mytabley.Fields("t8"))
    mytablex.Fields("t8") = sdx
    sdx = Val("" & mytablex.Fields("t9")) + signo * Val("" & mytabley.Fields("t9"))
    mytablex.Fields("t9") = sdx
    sdx = Val("" & mytablex.Fields("t10")) + signo * Val("" & mytabley.Fields("t10"))
    mytablex.Fields("t10") = sdx
    sdx = Val("" & mytablex.Fields("t11")) + signo * Val("" & mytabley.Fields("t11"))
    mytablex.Fields("t11") = sdx
    sdx = Val("" & mytablex.Fields("t12")) + signo * Val("" & mytabley.Fields("t12"))
    mytablex.Fields("t12") = sdx
    sdx = Val("" & mytablex.Fields("t13")) + signo * Val("" & mytabley.Fields("t13"))
    mytablex.Fields("t13") = sdx
    sdx = Val("" & mytablex.Fields("t14")) + signo * Val("" & mytabley.Fields("t14"))
    mytablex.Fields("t14") = sdx
    sdx = Val("" & mytablex.Fields("t15")) + signo * Val("" & mytabley.Fields("t15"))
    mytablex.Fields("t15") = sdx
    sdx = Val("" & mytablex.Fields("t16")) + signo * Val("" & mytabley.Fields("t16"))
    mytablex.Fields("t16") = sdx

End Sub

Private Sub mio8923_Click()

    Dim found    As Integer

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd27_err

    'If Frame4.Visible = True Then Exit Sub
    If Frame8.Visible = True Then Exit Sub

    If Frame7.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub

    If ACU = "I" Then 'pedidos
        buf = " select Local,Tipo,Serie,Numero,Fecha,Fechae,Hora,Codigo,Nombre,Moneda as M,Total,Estado as E,Tipo1,Serie1,Numero1,Usuario,Caja,Turno from factura where local='" & rexplorap.Fields("local") & "'"
        buf = buf & " and tipo1 = '" & rexplorap.Fields("tipo") & "'"
        buf = buf & " and serie1 = '" & rexplorap.Fields("serie") & "'"
        buf = buf & " and numero1 = '" & rexplorap.Fields("numero") & "'"
        mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            MsgBox "Pedido tiene otras Transacciones,Boleta,Guias ", 48, "Aviso"
            mytablex.Close
            Exit Sub

        End If

        mytablex.Close

    End If

    If importacion = "COMERCIAL" Then

        'If Trim("" & rexplorap.Fields("tipoimp")) <> "C" Then
        'MsgBox "Solo puede realizar por Comercial ", 48, "Aviso"
        'Exit Sub
        'End If
    End If

    ' 01/08/2018 Testing Facturacion Electronica
    If ACU = "V" Then
        If "" & rexplorap.Fields("tipo") <> "1" And "" & rexplorap.Fields("tipo") <> "2" And "" & rexplorap.Fields("tipo") <> "3" Then
            MsgBox "En este módulo no se modifica este documento", vbCritical
            Exit Sub

        End If

    End If

    ' 01/08/2018 Testing Facturacion Electronica
 
    If "" & rexplorap.Fields("estado") <> "0" Then
        'MsgBox "Estado debe estar =0", 48, "Aviso"
        xxdesmarca
        Exit Sub

    End If

    'If Trim("" & rexplorap.Fields(0)) = "A" Then
    '   MsgBox "Modo atendido,no se puede modificar ", 48, "Aviso"
    '   dbgrid2.SetFocus
    '   Exit Sub
    'End If

    found = copiar_temporal()

    If found = 0 Then
        MsgBox "Error al copiar archivo temporal", 24, "Aviso"
        Exit Sub

    End If

    borratempdao

    If ACU = "Z" Then
        tfactura.Label14.Visible = True
        tfactura.Label38.Visible = True
        tfactura.localf.Visible = True
        tfactura.bodegaf.Visible = True
        tfactura.Label2.Caption = "Cod.Int."
        tfactura.localf.Enabled = True
        tfactura.bodegaf.Enabled = True
        tfactura.Label14.Caption = "LocalDestino"
        tfactura.Label38.Caption = "AlmacenDestino"

    End If

    If ACU = "Q" Then
        tfactura.codigo = "01"
        tfactura.Label2.Caption = "Codigo"
        tfactura.Label38.Visible = True
        'tfactura.label13.Caption = "AlmacenDestino"
        'tfactura.Label38.Caption = "Alm:origen"
        tfactura.bodegaf.Visible = True
        tfactura.Label14.Visible = True
        tfactura.localf.Visible = True
        'tfactura.Label14.Caption = "Loc:Origen"
        tfactura.localf.Enabled = True
        tfactura.bodegaf.Enabled = True
   
        'tfactura.Label14.Caption = "LocalOrigen"
        'tfactura.Label38.Caption = "AlmacenOrigen"
   
    End If

    tfactura.zlocal = "" & rexplorap.Fields("local")
    tfactura.ztipo = "" & rexplorap.Fields("tipo")
    tfactura.zserie = "" & rexplorap.Fields("serie")
    tfactura.znumero = "" & rexplorap.Fields("numero")

    If ACU = "V" Then
        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Facturacion x Ventas"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.bandera = "Modifica"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.ACU = "V"
        tfactura.tipoclie = tipoclie
        'MsgBox "" & DBGrid1.Columns(2)
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.importacion = importacion
        tfactura.Show 1
        'sql_cabeza

    End If

    If ACU = "H" Then
        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Cotizacion x Ventas"
        cgusuario = "ccotizav"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "dcotizav"
        tfactura.bandera = "Modifica"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.ACU = "H"
        tfactura.tipoclie = tipoclie
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "I" Then
        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Pedidos x Ventas"
        cgusuario = "cpedidov"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "dpedidov"
        tfactura.bandera = "Modifica"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.ACU = "I"
        tfactura.tipoclie = tipoclie
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "3" Then
        found = copiar_servicio()

        If found = 0 Then
            MsgBox "Error al copiar servicio", 24, "Aviso"
            Exit Sub

        End If

        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Servicio Tecnico"
        cgusuario = "cservicio"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        sgusuario = "_s" & gusuario  'servicio tecnico
        dgusuariog = "dservicio"
        tfactura.bandera = "Modifica"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.ACU = "3"
        tfactura.tipoclie = tipoclie
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "T" Then
        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Guia Remision x Ventas"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.bandera = "Modifica"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.ACU = "T"
        tfactura.tipoclie = tipoclie
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "E" Then
        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Nota Credito x Ventas"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.bandera = "Modifica"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.ACU = "E"
        tfactura.tipoclie = tipoclie
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "R" Then
        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Orden Compra"
        cgusuario = "CORDENC"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DORDENC"
        tfactura.bandera = "Modifica"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.ACU = "R"
        tfactura.tipoclie = tipoclie
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "F" Then
        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Nota Debito x Ventas"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"

        tfactura.ACU = "F"
        tfactura.tipoclie = tipoclie
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.bandera = "Modifica"
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "S" Then
        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Guia Remision Entrada"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.tipoclie = tipoclie
        tfactura.ACU = "S"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.bandera = "Modifica"
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "C" Then
        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Factura de Compra"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.ACU = "C"
        tfactura.tipoclie = tipoclie
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.bandera = "Modifica"
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "N" Then
        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Nota Credito Compra"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.ACU = "N"
        tfactura.tipoclie = tipoclie
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.bandera = "Modifica"
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "O" Then
        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Nota debito de Compra"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.ACU = "O"
        tfactura.tipoclie = tipoclie
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.bandera = "Modifica"
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "Q" Then
        tfactura.Label2 = "Codigo"
        tfactura.Caption = "Requerimiento Almacen"
        cgusuario = "CREQUISA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DREQUISA"
        tfactura.ACU = "Q"
        tfactura.Label38.Visible = True
        tfactura.bodegaf.Visible = True
        'tfactura.label13.Caption = "AlmacenDestino"
        'tfactura.Label38.Caption = "DesdAlm."
        tfactura.localf.Enabled = True
        tfactura.bodegaf.Enabled = True

        'tfactura.Label14.Caption = "LocalOrigen"
        '   tfactura.Label38.Caption = "AlmacenOrigen"

        tfactura.tipoclie = tipoclie
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.bandera = "Modifica"
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "Z" Then
        tfactura.Caption = "Traslado entre almacen de un mismo establecimiento"
        cgusuario = "CTRASLAD"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DTRASLAD"
        tfactura.ACU = "Z"
        tfactura.tipoclie = tipoclie
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.bandera = "Modifica"
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    'sql_cabeza

    MsgBox "Presione tecla para continuar..", 48, "Aviso"
    sql_cabeza

    Exit Sub
cmd27_err:
    MsgBox "Seleccione un dato  ", 48, "Aviso"
    Exit Sub

End Sub

Sub pone_registro()

End Sub

Private Sub mit56232_Click()

    If Frame8.Visible = True Then Exit Sub

    If Frame7.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    zooma = "Zomm"
    visualizar_zoom
    MsgBox "Presione tecla para continuar..", 48, "Aviso"
    'sql_cabeza
    Exit Sub

End Sub

Sub visualizar_zoom()

    Dim found As Integer

    On Error GoTo cmd278_err

    'If Frame4.Visible = True Then Exit Sub

    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub

    'If "" & Data2.Recordset.Fields("estado") <> "0" Then
    '   MsgBox "Estado debe estar =0", 48, "Aviso"
    '   Exit Sub
    'End If
    tfactura.importacion = importacion

    'If importacion = "GASTOS" Then
    '   tfactura.Label51.Visible = True
    '   tfactura.cgasto.Visible = True
    '   'Exit Sub
    'End If

    found = copiar_temporal()

    If found = 0 Then
        MsgBox "Error al copiar archivo temporal", 24, "Aviso"
        Exit Sub

    End If

    borratempdao

    tfactura.zlocal = "" & rexplorap.Fields("local")
    tfactura.ztipo = "" & rexplorap.Fields("tipo")
    tfactura.zserie = "" & rexplorap.Fields("serie")
    tfactura.znumero = "" & rexplorap.Fields("numero")

    If ACU = "Z" Then
        tfactura.Label14.Visible = True
        tfactura.Label38.Visible = True
        tfactura.localf.Visible = True
        tfactura.bodegaf.Visible = True
        tfactura.Label2.Caption = "Cod.Int."
        tfactura.localf.Enabled = True
        tfactura.bodegaf.Enabled = True
        tfactura.Label14.Caption = "LocalDestino"
        tfactura.Label38.Caption = "AlmacenDestino"
 
    End If

    If ACU = "Q" Then
        tfactura.codigo = "01"
        tfactura.Label2.Caption = "Codigo"
        tfactura.Label38.Visible = True
        'tfactura.label13.Caption = "AlmacenDestino"
   
        tfactura.bodegaf.Visible = True
        tfactura.Label14.Visible = True
        tfactura.localf.Visible = True
   
        tfactura.localf.Enabled = True
        tfactura.bodegaf.Enabled = True
   
        'tfactura.Label14.Caption = "LocalOrigen"
        'tfactura.Label38.Caption = "AlmacenOrigen"
 
    End If

    If ACU = "V" Then
        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Facturacion x Ventas"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.bandera = "Ver"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.ACU = "V"
        tfactura.tipoclie = tipoclie
        tfactura.importacion = importacion
        tfactura.Show 1
        'sql_cabeza

    End If

    If ACU = "H" Then
        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Cotizacion x Ventas"
        cgusuario = "ccotizav"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "dcotizav"
        tfactura.bandera = "Ver"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.ACU = "H"
        tfactura.tipoclie = tipoclie
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "I" Then
        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Pedidos x Ventas"
        cgusuario = "cpedidov"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "dpedidov"
        tfactura.bandera = "Ver"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.ACU = "I"
        tfactura.tipoclie = tipoclie
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "3" Then

        found = copiar_servicio()

        If found = 0 Then
            MsgBox "Error al copiar servicio", 24, "Aviso"
            Exit Sub

        End If

        'tfactura.Command14.Enabled = False
        'tfactura.Command15.Enabled = False

        'tfactura.Command16.Enabled = False
        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Servicio Tecnico"
        cgusuario = "cservicio"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        sgusuario = "_s" & gusuario  'servicio tecnico
        dgusuariog = "dservicio"
        tfactura.bandera = "Ver"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.ACU = "3"
        tfactura.tipoclie = tipoclie
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "T" Then
        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Guia Remision x Ventas"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.bandera = "Ver"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.ACU = "T"
        tfactura.tipoclie = tipoclie
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "E" Then
        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Nota Credito x Ventas"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.bandera = "Ver"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.ACU = "E"
        tfactura.tipoclie = tipoclie
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "R" Then
        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Orden Compra"
        cgusuario = "CORDENC"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DORDENC"
        tfactura.bandera = "Ver"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.ACU = "R"
        tfactura.tipoclie = tipoclie
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "F" Then
        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Nota Debito x Ventas"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"

        tfactura.ACU = "F"
        tfactura.tipoclie = tipoclie
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.bandera = "Ver"
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "S" Then
        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Guia Remision Entrada"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.tipoclie = tipoclie
        tfactura.ACU = "S"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.bandera = "Ver"
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "C" Then
        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Factura de Compra"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.ACU = "C"
        tfactura.tipoclie = tipoclie
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.bandera = "Ver"
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "N" Then
        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Nota Credito Compra"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.ACU = "N"
        tfactura.tipoclie = tipoclie
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.bandera = "Ver"
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "O" Then
        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Nota debito de Compra"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.ACU = "O"
        tfactura.tipoclie = tipoclie
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.bandera = "Ver"
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "Q" Then
        tfactura.Label2 = "Codigo"
        'tfactura.Caption = "Requerimiento Almacen"
        cgusuario = "CREQUISA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DREQUISA"
        tfactura.ACU = "Q"
        tfactura.Label38.Visible = True
        tfactura.bodegaf.Visible = True
        'tfactura.label13.Caption = "AlmacenDestino"
        'tfactura.Label38.Caption = "DesdAlm."

        tfactura.localf.Enabled = True
        tfactura.bodegaf.Enabled = True

        'tfactura.Label14.Caption = "LocalOrigen"
        '   tfactura.Label38.Caption = "AlmacenOrigen"

        tfactura.tipoclie = tipoclie
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.bandera = "Ver"
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    If ACU = "Z" Then
        tfactura.Caption = "Traslado entre almacen de un mismo establecimiento"
        cgusuario = "CTRASLAD"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DTRASLAD"
        tfactura.ACU = "Z"
        tfactura.tipoclie = tipoclie
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.bandera = "Ver"
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.importacion = importacion
        tfactura.Show 1

    End If

    Exit Sub
cmd278_err:
    MsgBox "Aviso en visualizar zoon " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub modi343_Click()

End Sub

Sub xxdesmarca()

    On Error GoTo cmd117_err

    Dim found As Integer

    Dim buf   As String

    'If Frame4.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    buf = "" & rexplorap.Fields("estado") 'Data2.Recordset.Fields("estado")

    If Trim("" & rexplorap.Fields("estado")) <> "2" Then
        MsgBox "Debe encontrarse en estado 2 para desmarcar ", 48, "Aviso"
        dbgrid2.SetFocus
        Exit Sub

    End If

    If Trim("" & rexplorap.Fields(0)) = "A" Then
        MsgBox "Modo atendido,no se puede modificar ", 48, "Aviso"
        dbgrid2.SetFocus
        Exit Sub

    End If

    Frame2.Visible = True
    Frame2.Caption = "DESMARCA"
    clave = ""
    clave.SetFocus
    Exit Sub
cmd117_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Function verificar_recibo(buf As String, _
                          xlocal As String, _
                          xtipo As String, _
                          xserie As String, _
                          xnumero As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM " & buf & " where  local1='" & xlocal & "' and tipo1='" & xtipo & "' and serie1='" & xserie & "' and numero1='" & xnumero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        verificar_recibo = 1

    End If

    mytablex.Close

End Function

Function desgraba_cuentac()

    '---------- validando si es cuenta corriente

    If valida_flag("" & rexplorap.Fields("acu")) = 2 Then   'compras
        cn.Execute ("delete from cuentap where local='" & rexplorap.Fields("local") & "' and tipo='" & rexplorap.Fields("tipo") & "' and serie='" & rexplorap.Fields("serie") & "' and numero='" & rexplorap.Fields("numero") & "'")

    End If

    If valida_flag("" & rexplorap.Fields("acu")) = 1 Then   'ventas
        cn.Execute ("delete from cuentac where local='" & rexplorap.Fields("local") & "' and tipo='" & rexplorap.Fields("tipo") & "' and serie='" & rexplorap.Fields("serie") & "' and numero='" & rexplorap.Fields("numero") & "'")

    End If
 
End Function

Sub menu_excell1()

    If ACU = "V" Or ACU = "T" Or ACU = "E" Or ACU = "F" Or ACU = "S" Or ACU = "C" Or ACU = "N" Or ACU = "O" Then
        cgusuario = "FACTURA"
        dgusuariog = "DETALLE"

    End If

    If ACU = "H" Then
        cgusuario = "CCOTIZAV"
        dgusuariog = "DCOTIZAV"

    End If

    If ACU = "I" Then
        cgusuario = "CPEDIDOV"
        dgusuariog = "DPEDIDOV"

    End If

    If ACU = "R" Then
        cgusuario = "CORDENC"
        dgusuariog = "DORDENC"

    End If

    If ACU = "Q" Then
        cgusuario = "CREQUISA"
        dgusuariog = "DREQUISA"

    End If

    If ACU = "Z" Then
        cgusuario = "CTRASLAD"
        dgusuariog = "DTRASLAD"

    End If

    excel_paso1

End Sub

Sub menu_excell()

    If ACU = "V" Or ACU = "T" Or ACU = "E" Or ACU = "F" Or ACU = "S" Or ACU = "C" Or ACU = "N" Or ACU = "O" Then
        cgusuario = "FACTURA"
        dgusuariog = "DETALLE"

    End If

    If ACU = "H" Then
        cgusuario = "CCOTIZAV"
        dgusuariog = "DCOTIZAV"

    End If

    If ACU = "I" Then
        cgusuario = "CPEDIDOV"
        dgusuariog = "DPEDIDOV"

    End If

    If ACU = "R" Then
        cgusuario = "CORDENC"
        dgusuariog = "DORDENC"

    End If

    If ACU = "Q" Then
        cgusuario = "CREQUISA"
        dgusuariog = "DREQUISA"

    End If

    If ACU = "Z" Then
        cgusuario = "CTRASLAD"
        dgusuariog = "DTRASLAD"

    End If

    excel_paso

End Sub

Sub excel_paso1()

    Dim sdx As String

    On Error GoTo cmd813_err

    sdx = "" & rexplorap.Fields("numero")
    conteo_excell1
    Exit Sub
cmd813_err:
    MsgBox "Elegir un dato ", 48, "Aviso"
    Exit Sub

End Sub

Sub excel_paso()

    Dim sdx As String

    On Error GoTo cmd81_err

    sdx = "" & rexplorap.Fields("numero")
    conteo_excell
    Exit Sub
cmd81_err:
    MsgBox "Elegir un dato ", 48, "Aviso"
    Exit Sub

End Sub

Sub conteo_excell1()

    Dim v, h As Integer

    Dim R            As Long

    Dim found        As Integer

    Dim I            As Integer

    Dim sdx          As Double

    Dim sdx1         As Double

    Dim sdx2         As Double

    Dim vprecios(12) As String

    Dim Heading(13)  As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    On Error GoTo cmd56124_err

    Heading(1) = "Codigo"
    Heading(2) = "Nombre"
    Heading(3) = "Local"
    Heading(4) = "Tipo"
    Heading(5) = "Serie"
    Heading(6) = "Numero"
    Heading(7) = "M"
    Heading(8) = "Fecha"
    Heading(9) = "Total"
    Heading(10) = "E"
    Heading(11) = ""
   
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    
    '''25/10/2017 Agregar fecha en Excel de Excel Impresión total y seleccionado
    Call Formato_Excel2(11, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    '''25/10/2017 Agregar fecha en Excel de Excel Impresión total y seleccionado
    
    v = 4
    h = 0
    sdx = 0
    sdx1 = 0
    sdx2 = 0

    Do

        If rexplorap.EOF Then Exit Do

        objExcel.ActiveSheet.Cells(v, h + 1) = "'" & rexplorap.Fields("Codigo")
        objExcel.ActiveSheet.Cells(v, h + 2) = "'" & rexplorap.Fields("nombre")
        objExcel.ActiveSheet.Cells(v, h + 3) = "'" & rexplorap.Fields("Local")
        objExcel.ActiveSheet.Cells(v, h + 4) = "'" & rexplorap.Fields("tipo")
        objExcel.ActiveSheet.Cells(v, h + 5) = "'" & rexplorap.Fields("serie")
        objExcel.ActiveSheet.Cells(v, h + 6) = "'" & rexplorap.Fields("numero")
        objExcel.ActiveSheet.Cells(v, h + 7) = "'" & rexplorap.Fields("Moneda")
        objExcel.ActiveSheet.Cells(v, h + 8) = "'" & rexplorap.Fields("fecha")
        objExcel.ActiveSheet.Cells(v, h + 9) = "" & rexplorap.Fields("total")
        objExcel.ActiveSheet.Cells(v, h + 10) = "'" & rexplorap.Fields("estado")
        'objExcel.ActiveSheet.Cells(v, h + 11) = "" & rexplorap.fields("subtotal")
    
        If "" & rexplorap.Fields("Moneda") = "S" Then
            sdx1 = sdx1 + Val("" & rexplorap.Fields("total"))

        End If

        If "" & rexplorap.Fields("Moneda") = "D" Then
            sdx2 = sdx2 + Val("" & rexplorap.Fields("total"))

        End If

        v = v + 1
        rexplorap.MoveNext
    Loop

    objExcel.ActiveSheet.Cells(v, h + 1) = ""
    objExcel.ActiveSheet.Cells(v, h + 2) = ""
    objExcel.ActiveSheet.Cells(v, h + 3) = ""
    objExcel.ActiveSheet.Cells(v, h + 4) = ""
    objExcel.ActiveSheet.Cells(v, h + 5) = ""
    objExcel.ActiveSheet.Cells(v, h + 6) = dicmoneda
    objExcel.ActiveSheet.Cells(v, h + 7) = Format(sdx1, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 8) = "Dolar"
    objExcel.ActiveSheet.Cells(v, h + 9) = Format(sdx2, "0.00")

    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    Exit Sub
cmd56124_err:
    MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub conteo_excell()

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim v, h As Long

    Dim found       As Integer

    Dim I           As Integer

    Dim R           As Long

    Dim sdx         As Double

    Dim sdx1        As Double

    Dim sdx2        As Double

    Dim sdx3        As Double

    Dim isdx        As Double

    Dim vprecios(7) As String

    Dim Heading(8)  As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    On Error GoTo cmd1561212_err

    'Data1.Refresh
   
    Heading(1) = "Producto"
    Heading(2) = "Descripcio"
    Heading(3) = "Und"
    Heading(4) = "Factor"
    Heading(5) = "cantidad"
    Heading(6) = "Precio"
    Heading(7) = "Total"
   
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    
    '''09/10/2017 kenyo Testing Reportes
    'Call Formato_Excel(7, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    Call Formato_Excel2(7, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    '''09/10/2017 kenyo Testing Reportes
    
    v = 5
    h = 1

    Do

        If rexplorap.EOF Then Exit Do
        sdx = 0
        sdx1 = 0

        ''25/10/2017 Agregar fecha en Excel de Excel Impresión total y seleccionado

        objExcel.ActiveSheet.Cells(v, h + 1) = "Tipo:" & rexplorap.Fields("tipo") & " Serie:" & rexplorap.Fields("serie") & "Numero:" & rexplorap.Fields("numero") & " Fecha:" & rexplorap.Fields("fecha")
        objExcel.ActiveSheet.Cells(v, h + 1).Font.bold = True
        v = v + 1
        objExcel.ActiveSheet.Cells(v, h + 1) = "Cliente:" & rexplorap.Fields("codigo") & " Nombre:" & rexplorap.Fields("nombre") & " Vendedor:" & rexplorap.Fields("vendedor") & " Moneda:" & rexplorap.Fields("moneda")
        objExcel.ActiveSheet.Cells(v, h + 1).Font.bold = True
        v = v + 1
        ''25/10/2017 Agregar fecha en Excel de Excel Impresión total y seleccionado

        mytablex.Open "SELECT * FROM " & dgusuariog & " where  local='" & Trim("" & rexplorap.Fields("local")) & "' and tipo='" & Trim("" & rexplorap.Fields("tipo")) & "' and serie='" & Trim("" & rexplorap.Fields("serie")) & "' and numero='" & Trim("" & rexplorap.Fields("numero")) & "'", cn, adOpenKeyset, adLockOptimistic

        If mytablex.RecordCount > 0 Then  'si existe
            Do

                If mytablex.EOF Then Exit Do
                sdx = sdx + Val("" & mytablex.Fields("cantidad"))
                sdx1 = sdx1 + Val("" & mytablex.Fields("total"))
                objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("producto")
                objExcel.ActiveSheet.Cells(v, h + 1) = "" & mytablex.Fields("descripcio")
                objExcel.ActiveSheet.Cells(v, h + 2) = "" & mytablex.Fields("unidad")
                objExcel.ActiveSheet.Cells(v, h + 3) = Val("" & mytablex.Fields("factor"))
                objExcel.ActiveSheet.Cells(v, h + 4) = Val("" & mytablex.Fields("cantidad"))
                objExcel.ActiveSheet.Cells(v, h + 5) = Val("" & mytablex.Fields("precio"))
                objExcel.ActiveSheet.Cells(v, h + 6) = Val("" & mytablex.Fields("total"))
            
                ''25/10/2017 Agregar fecha en Excel de Excel Impresión total y seleccionado
                '
                '            objExcel.ActiveSheet.Cells(v, h + 6) = Val("" & mytablex.Fields("total"))
                '
                '            If mytabley.State = 1 Then mytabley.Close
                '            mytabley.Open "SELECT * FROM precios where  producto='" & mytablex.Fields("producto") & "' and local='01'", cn, adOpenKeyset, adLockOptimistic
                '            sdx2 = 0
                '            If mytabley.RecordCount > 0 Then  'si existe
                '               isdx = Val("" & mytabley.Fields("factor1"))
                '               If isdx = 0 Then
                '                  isdx = 1
                '               End If
                '               sdx2 = Val("" & mytabley.Fields("pventa1")) / isdx
                '            End If
                '            mytabley.Close
                '            isdx = Val("" & mytablex.Fields("total"))
                '            If isdx = 0 Then
                '               isdx = 1
                '            End If
                '            sdx2 = sdx2 * Val("" & mytablex.Fields("factor")) * Val("" & mytablex.Fields("cantidad"))
                '            sdx3 = (sdx2 - Val("" & mytablex.Fields("total"))) * 100 / isdx
                '            objExcel.ActiveSheet.Cells(v, h + 7) = "'" & Format(sdx3, "0.00") & "%"
                ''25/10/2017 Agregar fecha en Excel de Excel Impresión total y seleccionado

                v = v + 1
            
                mytablex.MoveNext
            Loop

        End If

        mytablex.Close

        objExcel.ActiveSheet.Cells(v, h) = ""
        objExcel.ActiveSheet.Cells(v, h + 1) = ""
        objExcel.ActiveSheet.Cells(v, h + 2) = ""
        objExcel.ActiveSheet.Cells(v, h + 3) = ""
        objExcel.ActiveSheet.Cells(v, h + 4) = "" & sdx
        objExcel.ActiveSheet.Cells(v, h + 5) = ""
        objExcel.ActiveSheet.Cells(v, h + 6) = "" & rexplorap.Fields("total")
            
        v = v + 1

        rexplorap.MoveNext
    Loop

    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    Exit Sub
cmd1561212_err:
    MsgBox "Error en exportacion excell " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub conteo_excell_uno()

    Dim mytablex As New ADODB.Recordset

    Dim v, h As Integer

    Dim found       As Integer

    Dim I           As Integer

    Dim sdx         As Double

    Dim sdx1        As Double

    Dim vprecios(8) As String

    Dim Heading(9)  As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    On Error GoTo cmd561212_err

    'Data1.Refresh
   
    Heading(1) = "Producto"
    Heading(2) = "Descripcio"
    Heading(3) = "Und"
    Heading(4) = "Factor"
    Heading(5) = "cantidad"
    Heading(6) = "Precio"
    Heading(7) = "Total"
   
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    
    '''09/10/2017 kenyo Testing Reportes
    'Call Formato_Excel(8, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    Call Formato_Excel2(8, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    '''09/10/2017 kenyo Testing Reportes

    ''25/10/2017 Agregar fecha en Excel de Excel Impresión total y seleccionado

    objExcel.ActiveSheet.Cells(1, 1) = "Tipo:" & rexplorap.Fields("tipo") & " Serie:" & rexplorap.Fields("serie") & "Numero:" & rexplorap.Fields("numero") & " Fecha:" & rexplorap.Fields("Fecha")
    objExcel.ActiveSheet.Cells(1, 1).Font.bold = True
    objExcel.ActiveSheet.Cells(2, 1) = "Cliente:" & rexplorap.Fields("codigo") & " Nombre:" & rexplorap.Fields("nombre") & " Vendedor:" & rexplorap.Fields("vendedor") & " Moneda:" & rexplorap.Fields("moneda")
    objExcel.ActiveSheet.Cells(2, 1).Font.bold = True
    ''25/10/2017 Agregar fecha en Excel de Excel Impresión total y seleccionado
    
    v = 5
    h = 1
    sdx = 0
    sdx1 = 0

    mytablex.Open "SELECT * FROM " & dgusuariog & " where  local='" & Trim("" & rexplorap.Fields("local")) & "' and tipo='" & Trim("" & rexplorap.Fields("tipo")) & "' and serie='" & Trim("" & rexplorap.Fields("serie")) & "' and numero='" & Trim("" & rexplorap.Fields("numero")) & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        Do

            If mytablex.EOF Then Exit Do
            sdx = sdx + Val("" & mytablex.Fields("cantidad"))
            sdx1 = sdx1 + Val("" & mytablex.Fields("total"))
            objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("producto")
            objExcel.ActiveSheet.Cells(v, h + 1) = "" & mytablex.Fields("descripcio")
            objExcel.ActiveSheet.Cells(v, h + 2) = "" & mytablex.Fields("unidad")
            objExcel.ActiveSheet.Cells(v, h + 3) = Val("" & mytablex.Fields("factor"))
            objExcel.ActiveSheet.Cells(v, h + 4) = Val("" & mytablex.Fields("cantidad"))
            objExcel.ActiveSheet.Cells(v, h + 5) = Val("" & mytablex.Fields("precio"))
            objExcel.ActiveSheet.Cells(v, h + 6) = Val("" & mytablex.Fields("total"))
            v = v + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

    mytablex.Open "SELECT * FROM " & cgusuario & " where  local='" & Trim("" & rexplorap.Fields("local")) & "' and tipo='" & Trim("" & rexplorap.Fields("tipo")) & "' and serie='" & Trim("" & rexplorap.Fields("serie")) & "' and numero='" & Trim("" & rexplorap.Fields("numero")) & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        objExcel.ActiveSheet.Cells(v, h) = ""
        objExcel.ActiveSheet.Cells(v, h + 1) = ""
        objExcel.ActiveSheet.Cells(v, h + 2) = ""
        objExcel.ActiveSheet.Cells(v, h + 3) = ""
        objExcel.ActiveSheet.Cells(v, h + 4) = "" & sdx
        objExcel.ActiveSheet.Cells(v, h + 5) = ""
        objExcel.ActiveSheet.Cells(v, h + 6) = "" & mytablex.Fields("total")
        v = v + 1

    End If

    mytablex.Close
    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    Exit Sub
cmd561212_err:
    MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
    Exit Sub

End Sub

'AQUI VAMOS A GENERAR LE documento automaticamente
Function valida_clave(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM vendedor where  clave='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        If "" & mytablex.Fields("modificacompra") <> "N" Then
            valida_clave = 1

        End If

    End If

    mytablex.Close

End Function

'25/06/2018 Testing Almacen General
'Function ve_descarga(buf As String)
'Dim mytablex As New ADODB.Recordset
'mytablex.Open "SELECT * FROM tipo where  tipo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic
'   If mytablex.RecordCount > 0 Then  'si existe
'      Select Case "" & mytablex.Fields("tipodoc")
'             Case "A", "B", "C", "D", "G", "E", "F", "J", "K", "L", "M", "P", "N", "O", "S", "T"
'                  ve_descarga = 1
'      End Select
'End If
'mytablex.Close
'
'End Function
Function ve_descarga(buf As String, buf1 As String)

    Dim mytablex   As New ADODB.Recordset

    Dim mytablexyz As New ADODB.Recordset

    Dim ACU        As String

    Dim acu1       As String

    ACU = ""
    acu1 = ""

    mytablex.Open "SELECT * FROM tipo where  tipo='" & buf1 & "'", cn, adOpenKeyset, adLockOptimistic
    mytablexyz.Open "SELECT * FROM tipo where  tipo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        ACU = "" & mytablexyz.Fields("tipodoc")
        acu1 = "" & mytablex.Fields("tipodoc")
    
        Select Case "" & ACU

            Case "A", "B", "C", "D", "G", "E", "F", "J", "K", "L", "M", "P", "N", "O", "S", "T"

                If (acu1 = "A" Or acu1 = "B" Or acu1 = "C" Or acu1 = "D") And ACU = "T" Then
                    ve_descarga = 1
                Else
                    ve_descarga = 0

                End If
                  
        End Select

    End If

    mytablex.Close

End Function

'25/06/2018 Testing Almacen General

Sub borratempdao()

    On Error GoTo cmdn78_err

    fuerza_borrar
    cn.Execute (" select * into _i" & gusuario & " from gastofactura where 1=2")
    Exit Sub
cmdn78_err:
    MsgBox "Aviso en borratempdao " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub fuerza_borrar()

    On Error GoTo cmd900_err

    cn.Execute ("drop table _i" & gusuario)
    Exit Sub
cmd900_err:
    Exit Sub

End Sub

Private Sub nmur41_Click()

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd9090_err

    If Frame8.Visible = True Then Exit Sub

    If Frame7.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub

    Frame8.Visible = True
    buf = " select Local,Tipo,Serie,Numero,Fecha,Fechae,Hora,Codigo,Nombre,Moneda as M,Total,Estado as E,Tipo1,Serie1,Numero1,Usuario,Caja,Turno from factura where local='" & rexplorap.Fields("local") & "'"
    buf = buf & " and tipo1 = '" & rexplorap.Fields("tipo") & "'"
    buf = buf & " and serie1 = '" & rexplorap.Fields("serie") & "'"
    buf = buf & " and numero1 = '" & rexplorap.Fields("numero") & "'"
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    Set dbgrid5.DataSource = mytablex
    dbgrid5.refresh
    Exit Sub
cmd9090_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Sub sihue()

    Dim vr

    Dim mytablex As New ADODB.Recordset

    Do

        If rexplorap.EOF Then Exit Do
        vr = DoEvents()
        mytablex.Open "select * from proveedo where codigo='" & Trim("" & rexplorap.Fields("codigo")) & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.AddNew
            mytablex.Fields("codigo") = Trim("" & rexplorap.Fields("codigo"))
            mytablex.Fields("nombre") = Trim("" & rexplorap.Fields("nombre"))

            If Len(Trim("" & rexplorap.Fields("codigo"))) = 11 Then
                mytablex.Fields("tipo") = "J"

            End If

            If Len(Trim("" & rexplorap.Fields("codigo"))) = 8 Then
                mytablex.Fields("tipo") = "D"

            End If

            If Len(Trim("" & rexplorap.Fields("codigo"))) < 8 Then
                mytablex.Fields("tipo") = "O"

            End If

            mytablex.Fields("moneda") = "S"
            mytablex.Update

        End If

        mytablex.Close
        rexplorap.MoveNext
    Loop

End Sub

' Testing Proyecto Facturacion Electronica 05/04/2018
Private Sub txtNumero_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        FrmConsultaSunat.Visible = False
        Exit Sub

    End If

End Sub

' Testing Proyecto Facturacion Electronica 05/04/2018

' Testing Proyecto Facturacion Electronica 05/04/2018
Function verificar_existencia(cgusuario As String, _
                              xlocal As String, _
                              xtipo As String, _
                              xserie As String, _
                              xnumero As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM  " & cgusuario & "   where  local='" & xlocal & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        verificar_existencia = 1
    Else
        verificar_existencia = 0

    End If

    mytablex.Close

End Function

' Testing Proyecto Facturacion Electronica 05/04/2018
