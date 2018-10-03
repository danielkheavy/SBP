VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tcheckin 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reservas"
   ClientHeight    =   9210
   ClientLeft      =   90
   ClientTop       =   -135
   ClientWidth     =   15060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   15060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "CheckOut"
      Height          =   5775
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   8175
      Begin VB.TextBox fechasalida 
         Height          =   735
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   46
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox horasalida 
         Height          =   735
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   45
         Top             =   1080
         Width           =   2775
      End
      Begin VB.CommandButton CmdGra 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Grabar"
         DisabledPicture =   "tcheckin.frx":0000
         Height          =   735
         Left            =   7200
         Picture         =   "tcheckin.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdCan 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cancelar"
         DisabledPicture =   "tcheckin.frx":0684
         Height          =   735
         Left            =   7200
         Picture         =   "tcheckin.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Salida"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   48
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HoraSalida"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   47
         Top             =   1080
         Width           =   2895
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Consulta"
      Height          =   8415
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   15015
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
         TabIndex        =   40
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Filtrar"
         Height          =   495
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   240
         Width           =   2175
      End
      Begin MSDataGridLib.DataGrid dbgrid13 
         Height          =   6870
         Left            =   120
         TabIndex        =   41
         Top             =   840
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   12118
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
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame2"
      Height          =   8865
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   14895
      Begin VB.TextBox hotelcuadre 
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
         MaxLength       =   5
         TabIndex        =   104
         Top             =   8160
         Width           =   735
      End
      Begin VB.TextBox tipocodigoh 
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   100
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox tipocodigo 
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   98
         Top             =   2760
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Habilitado"
         Height          =   495
         Left            =   6720
         TabIndex        =   97
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox brecibe 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   7920
         MaxLength       =   10
         TabIndex        =   96
         Top             =   5760
         Width           =   2055
      End
      Begin VB.ComboBox nfpago 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8520
         Style           =   2  'Dropdown List
         TabIndex        =   94
         Top             =   5400
         Width           =   1455
      End
      Begin VB.ComboBox ntipo 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8520
         Style           =   2  'Dropdown List
         TabIndex        =   93
         Top             =   4560
         Width           =   1455
      End
      Begin VB.TextBox bfpago 
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
         Left            =   7920
         MaxLength       =   4
         TabIndex        =   92
         Top             =   5400
         Width           =   615
      End
      Begin VB.TextBox bnumero 
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
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   90
         Top             =   4920
         Width           =   1215
      End
      Begin VB.TextBox bserie 
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
         Left            =   7920
         MaxLength       =   4
         TabIndex        =   89
         Top             =   4920
         Width           =   615
      End
      Begin VB.TextBox btipo 
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
         Left            =   7920
         MaxLength       =   6
         TabIndex        =   87
         Top             =   4560
         Width           =   615
      End
      Begin VB.TextBox estado 
         BackColor       =   &H00C0FFFF&
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
         MaxLength       =   7
         TabIndex        =   84
         Top             =   7800
         Width           =   1215
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   7800
         Width           =   1815
      End
      Begin VB.TextBox arribohoraf 
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
         TabIndex        =   78
         Top             =   5640
         Width           =   1935
      End
      Begin VB.ComboBox Combo4 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox categoria 
         Height          =   375
         Left            =   2280
         MaxLength       =   6
         TabIndex        =   76
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Refresca"
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
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   7080
         Width           =   855
      End
      Begin VB.TextBox precio 
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
         TabIndex        =   73
         Top             =   7080
         Width           =   1215
      End
      Begin VB.TextBox habitacion 
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
         MaxLength       =   6
         TabIndex        =   72
         Top             =   6720
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Refresca"
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
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   6720
         Width           =   855
      End
      Begin VB.ComboBox disponibles 
         Height          =   315
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   6720
         Width           =   1815
      End
      Begin VB.TextBox tipopension 
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
         TabIndex        =   68
         Top             =   6360
         Width           =   1935
      End
      Begin VB.ComboBox ntipopension 
         Height          =   315
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   67
         Top             =   6360
         Width           =   1815
      End
      Begin VB.TextBox tipotarifa 
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
         TabIndex        =   65
         Top             =   6000
         Width           =   1935
      End
      Begin VB.ComboBox ntipotarifa 
         Height          =   315
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   6000
         Width           =   1815
      End
      Begin VB.TextBox personas 
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
         MaxLength       =   2
         TabIndex        =   61
         Top             =   7440
         Width           =   735
      End
      Begin VB.ComboBox ntiporeserva 
         Height          =   315
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   3480
         Width           =   1815
      End
      Begin VB.TextBox tiporeserva 
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
         TabIndex        =   58
         Top             =   3480
         Width           =   1935
      End
      Begin VB.TextBox horareserva 
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
         MaxLength       =   10
         TabIndex        =   56
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox fechareserva 
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
         MaxLength       =   10
         TabIndex        =   54
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox hnombre 
         BackColor       =   &H00C0FFFF&
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
         MaxLength       =   100
         TabIndex        =   50
         Top             =   2040
         Width           =   6375
      End
      Begin VB.TextBox huesped 
         BackColor       =   &H00C0FFFF&
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
         MaxLength       =   15
         TabIndex        =   49
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox noches 
         Height          =   375
         Left            =   2280
         MaxLength       =   3
         TabIndex        =   36
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox codigo 
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   29
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox agente 
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
         TabIndex        =   27
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox arribohora 
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
         Top             =   5280
         Width           =   1935
      End
      Begin VB.TextBox operador 
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
         Left            =   6000
         MaxLength       =   11
         TabIndex        =   21
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox nombre 
         BackColor       =   &H00C0FFFF&
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
         MaxLength       =   80
         TabIndex        =   19
         Top             =   3120
         Width           =   6375
      End
      Begin VB.TextBox arribofechaf 
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
         Top             =   4920
         Width           =   1935
      End
      Begin VB.TextBox arribofecha 
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
         Top             =   4560
         Width           =   1935
      End
      Begin VB.TextBox checkin 
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
         Left            =   6360
         MaxLength       =   6
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cerrar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   10560
         Picture         =   "tcheckin.frx":0D08
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Imprimir todo"
         Top             =   1320
         Width           =   1470
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&GuardarReserva"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   10560
         Picture         =   "tcheckin.frx":15D2
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1470
      End
      Begin VB.Label total 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   4200
         TabIndex        =   106
         Top             =   7080
         Width           =   1815
      End
      Begin VB.Label Label33 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Turno"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   105
         Top             =   8160
         Width           =   2175
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   " (J)uridica  (D)ni (P)asaporte (O)tros"
         Height          =   195
         Left            =   2880
         TabIndex        =   103
         Top             =   2880
         Width           =   2460
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   " (J)uridica  (D)ni (P)asaporte (O)tros"
         Height          =   195
         Left            =   2880
         TabIndex        =   102
         Top             =   1680
         Width           =   2460
      End
      Begin VB.Label Label25 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoDocumento"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   101
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoDocumento"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   99
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Entrega "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   95
         Top             =   5760
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FormaPago"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   91
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie/Numero"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   88
         Top             =   4920
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoDocumento"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   86
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   82
         Top             =   7800
         Width           =   2175
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Habitacion"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   81
         Top             =   6720
         Width           =   2175
      End
      Begin VB.Label Label30 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Categoria"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   80
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label Label29 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HoraSalida(HH:MM:SS)"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   79
         Top             =   5640
         Width           =   2175
      End
      Begin VB.Label Label26 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   74
         Top             =   7080
         Width           =   2175
      End
      Begin VB.Image Image10 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4560
         Picture         =   "tcheckin.frx":1E9C
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   375
      End
      Begin VB.Image Image9 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4560
         Picture         =   "tcheckin.frx":254A
         Stretch         =   -1  'True
         Top             =   960
         Width           =   375
      End
      Begin VB.Image Image8 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5040
         Picture         =   "tcheckin.frx":2BF8
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Pension"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   69
         Top             =   6360
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo de Tarifa"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   66
         Top             =   6000
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaSalida"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   63
         Top             =   4920
         Width           =   2175
      End
      Begin VB.Label Label38 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NroPersonas"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   62
         Top             =   7440
         Width           =   2175
      End
      Begin VB.Label Label37 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo de reserva"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   59
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label Label39 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HoraReserva"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   57
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaReserva"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo (El que se aloja)"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label22 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombres"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Image Image6 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4200
         Picture         =   "tcheckin.frx":32A6
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Copia"
         Height          =   375
         Left            =   4560
         TabIndex        =   51
         Top             =   2400
         Width           =   495
      End
      Begin VB.Image Image3 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4200
         Picture         =   "tcheckin.frx":35B0
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   375
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4200
         Picture         =   "tcheckin.frx":38BA
         Stretch         =   -1  'True
         Top             =   960
         Width           =   375
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   7920
         Picture         =   "tcheckin.frx":3BC4
         Stretch         =   -1  'True
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label23 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NroDias"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Label Label28 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo(QuienPaga)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label27 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "QuienReserva"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HoraLLegada(HH:MM:SS)"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         Top             =   5280
         Width           =   2175
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Operador"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   22
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombres"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaEntrada"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         Top             =   4560
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CheckIn Id"
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
         Left            =   4920
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFC0&
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   15000
      TabIndex        =   2
      Top             =   0
      Width           =   15060
      Begin VB.ComboBox estado1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   0
         Width           =   1815
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
         Picture         =   "tcheckin.frx":3ECE
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Grabar registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.TextBox buffer 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         MaxLength       =   10
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   360
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000D&
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9600
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
         Left            =   2160
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tcheckin.frx":50E0
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
         Left            =   1440
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tcheckin.frx":62F2
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
         Left            =   2880
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tcheckin.frx":7504
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
         Picture         =   "tcheckin.frx":8716
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label vestado 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   14520
         TabIndex        =   107
         Top             =   360
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label vflag 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   11280
         TabIndex        =   85
         Top             =   120
         Width           =   45
      End
      Begin VB.Label xhabitacion 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   14520
         TabIndex        =   34
         Top             =   0
         Width           =   105
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Estado"
         Height          =   375
         Left            =   4080
         TabIndex        =   33
         Top             =   0
         Width           =   2295
      End
      Begin VB.Label xsw 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   15840
         TabIndex        =   31
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Consulta"
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   14895
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   6855
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   12091
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
         ColumnCount     =   15
         BeginProperty Column00 
            DataField       =   "Habitacion"
            Caption         =   "Habitacion"
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
            DataField       =   "Estado"
            Caption         =   "Estado"
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
            DataField       =   "tipocodigo"
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
         BeginProperty Column03 
            DataField       =   "Codigo"
            Caption         =   "Codigo"
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
            DataField       =   "nombre"
            Caption         =   "Nombre"
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
            DataField       =   "ArriboFecha"
            Caption         =   "ArriboFecha"
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
            DataField       =   "ArriboHora"
            Caption         =   "ArriboHora"
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
            DataField       =   "ARRIBOFechaf"
            Caption         =   "Salida"
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
            DataField       =   "arriboHoraf"
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
         BeginProperty Column09 
            DataField       =   "Precio"
            Caption         =   "Precio"
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
            DataField       =   "Categoria"
            Caption         =   "Categoria"
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
            DataField       =   "Tiporeserva"
            Caption         =   "TipoReserva"
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
            DataField       =   "Tipotarifa"
            Caption         =   "TipoTarifa"
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
         BeginProperty Column13 
            DataField       =   "TipoPension"
            Caption         =   "TipoPension"
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
         BeginProperty Column14 
            DataField       =   "CheckIn"
            Caption         =   "CheckIn"
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
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   434.835
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   2865.26
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   2954.835
            EndProperty
            BeginProperty Column14 
            EndProperty
         EndProperty
      End
      Begin VB.Label totalreserva 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   11760
         TabIndex        =   26
         Top             =   7200
         Width           =   2895
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         Height          =   495
         Left            =   10080
         TabIndex        =   25
         Top             =   7200
         Width           =   1695
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
   Begin VB.Menu dk88343 
      Caption         =   "An&ticipo"
   End
   Begin VB.Menu dki8834 
      Caption         =   "&Consumos"
   End
   Begin VB.Menu fkichek 
      Caption         =   "&CheckOut"
   End
   Begin VB.Menu jdfu7834 
      Caption         =   "EstadoC&uenta"
   End
   Begin VB.Menu Enviatpv 
      Caption         =   "&EnviaTPV"
   End
   Begin VB.Menu djuer1 
      Caption         =   "&Reporte"
      Begin VB.Menu fkir84 
         Caption         =   "&1.Reporte"
      End
      Begin VB.Menu dk9893 
         Caption         =   "&2.GENERAL"
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
Attribute VB_Name = "tcheckin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim txcheckinx As New ADODB.Recordset

Dim mytablexx  As New ADODB.Recordset

Dim mytableyy  As New ADODB.Recordset

Dim mytablezz  As New ADODB.Recordset

Private Sub agente_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_agente

    End If

End Sub

Private Sub ajdu1_Click()
    'If Frame5.Visible = True Then Exit Sub

    If Frame3.Visible = True Then Exit Sub
    If Frame4.Visible = True Then Exit Sub

    If Frame2.Visible = True Then
        'dbGrid1.SetFocus
        Exit Sub

    End If

    If txcheckinx.RecordCount > 20 Then

        'MsgBox "Favor llamar al Proveedor para ampliar Licencia ", 48, "Aviso"
        'Exit Sub
    End If

    inicializa
    Frame2.Visible = True
    Frame2.Caption = "Nuevo"
    cmdGuardar.Enabled = True
    habilita 1
    checkin.Enabled = False
    checkin = ""
    operador.SetFocus

End Sub

Private Sub arribofecha_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Not IsDate(arribofecha) Then
        arribofecha = Format(Now, "dd/mm/yyyy")

    End If

    valida_horas

End Sub

Private Sub arribofechaf_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Not IsDate(arribofecha) Then
        arribofecha = Format(Now, "dd/mm/yyyy")

    End If

    valida_horas

End Sub

Private Sub bo712_Click()

    Dim buf As String

    On Error GoTo cmd656_err

    'If Frame5.Visible = True Then Exit Sub

    If Frame3.Visible = True Then Exit Sub
    If Frame4.Visible = True Then Exit Sub

    buf = "" & txcheckinx.Fields("checkin")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    If MsgBox("Desea Borra " + "" & txcheckinx.Fields("checkin"), 1, "Aviso") <> 1 Then
        Exit Sub

    End If

    txcheckinx.Delete
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

Private Sub Check1_Click()
    bserie = ""
    bnumero = ""
    btipo = ""
    bfpago = ""
    brecibe = ""

    'Check1.Value = True
End Sub

Private Sub cmdAddEntry_Click()
    ajdu1_Click

End Sub

Private Sub CmdCan_Click()
    Frame4.Visible = False

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

Private Sub Codigo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_codigo1

    End If

    If KeyCode = &H76 Then  'f1
        tnclie.DBPROV = "clientes"
        tnclie.fdlo893.Visible = True
        tnclie.Show 1

    End If

End Sub

Private Sub Combo3_Click()
    estado = Trim(Combo3)

End Sub

Private Sub Combo4_Click()
    categoria = Trim("" & Combo4.Text)
    Label23 = categoria
    valida_horas

End Sub

Private Sub Command2_Click()
    busca_habitacionlibre

End Sub

Private Sub Command3_Click()
    carga_precio Trim("" & habitacion)
    suma_total

End Sub

Private Sub Command4_Click()
    filtro

End Sub

Sub filtro()

    Dim mytablex As New ADODB.Recordset

    Dim cad      As String

    If opcion1 = "1" Then  'producto
        If Len(Text1) = 0 Then
            cad = "select Habitacion,Descripcio,Estado,TipoHabitacion,Capacidad,precio from habitacion "

        End If

        If Len(Text1) > 0 Then
            cad = "select Habitacion,Descripcio,Estado,TipoHabitacion,Capacidad,precio from habitacion where  " & Combo2 & " like '" & Text1.Text & "%'"

        End If

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbgrid13.DataSource = mytablex
        dbgrid13.columns(0).Width = 800
        dbgrid13.columns(1).Width = 800
        dbgrid13.columns(2).Width = 1900
        dbgrid13.columns(3).Width = 900
        'dbgrid13.columns(4).Width = 900
        'dbgrid13.columns(2).Width = 1000
        'dbgrid13.columns(3).Width = 1000

        If mytablex.RecordCount > 0 Then
            dbgrid13.SetFocus

        End If

    End If

    If opcion1 = "2" Then  'producto
        If Len(Text1) = 0 Then
            cad = "select Nombre,Codigo,Direccion,Tipo,Correo from clientes "

        End If

        If Len(Text1) > 0 Then
            cad = "select Nombre,Codigo,Direccion,Tipo,Correo from clientes where  " & Combo2 & " like '" & Text1.Text & "%'"

        End If

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbgrid13.DataSource = mytablex
        dbgrid13.columns(0).Width = 3000
        dbgrid13.columns(1).Width = 1000
              
        If mytablex.RecordCount > 0 Then
            dbgrid13.SetFocus

        End If

    End If

    If opcion1 = "6" Then  'producto
        If Len(Text1) = 0 Then
            cad = "select Nombre,Codigo,Tipo,Direccion,Correo from clientes "

        End If

        If Len(Text1) > 0 Then
            cad = "select Nombre,Codigo,Tipo,Direccion,Correo from clientes where  " & Combo2 & " like '" & Text1.Text & "%'"

        End If

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbgrid13.DataSource = mytablex
        dbgrid13.columns(0).Width = 3000
        dbgrid13.columns(1).Width = 1000
              
        If mytablex.RecordCount > 0 Then
            dbgrid13.SetFocus

        End If

    End If

    If opcion1 = "7" Then  'producto
        If Len(Text1) = 0 Then
            cad = "select Nombre,Codigo,Tipo,Direccion,Correo from clientes "

        End If

        If Len(Text1) > 0 Then
            cad = "select Nombre,Codigo,Tipo,Direccion,Correo from clientes where  " & Combo2 & " like '" & Text1.Text & "%'"

        End If

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbgrid13.DataSource = mytablex
        dbgrid13.columns(0).Width = 3000
        dbgrid13.columns(1).Width = 1000
              
        If mytablex.RecordCount > 0 Then
            dbgrid13.SetFocus

        End If

    End If

    If opcion1 = "3" Then  'producto
        If Len(Text1) = 0 Then
            cad = "select Nombre,Codigo from Vendedor "

        End If

        If Len(Text1) > 0 Then
            cad = "select Nombre,Codigo from Vendedor where  " & Combo2 & " like '" & Text1.Text & "%'"

        End If

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbgrid13.DataSource = mytablex
        dbgrid13.columns(0).Width = 3000
        dbgrid13.columns(1).Width = 1000
              
        If mytablex.RecordCount > 0 Then
            dbgrid13.SetFocus

        End If

    End If

    If opcion1 = "43" Then  'reserva
        If Len(Text1) = 0 Then
            cad = "select Reserva,Nombre,Arribofecha,arribohora,Procedencia from reserva "

        End If

        If Len(Text1) > 0 Then
            cad = "select Reserva,Nombre,Arribofecha,arribohora,Procedencia from reserva where  " & Combo2 & " like '" & Text1.Text & "%'"

        End If

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbgrid13.DataSource = mytablex
        dbgrid13.columns(0).Width = 1000
        dbgrid13.columns(1).Width = 3000
              
        If mytablex.RecordCount > 0 Then
            dbgrid13.SetFocus

        End If

    End If

    If opcion1 = "4" Then  'producto
        If Len(Text1) = 0 Then
            cad = "select Nombre,Codigo,Tipo from clientes "

        End If

        If Len(Text1) > 0 Then
            cad = "select Nombre,Codigo,Tipo from clientes where  " & Combo2 & " like '" & Text1.Text & "%'"

        End If

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbgrid13.DataSource = mytablex
        dbgrid13.columns(0).Width = 3000
        dbgrid13.columns(1).Width = 1000
              
        If mytablex.RecordCount > 0 Then
            dbgrid13.SetFocus

        End If

    End If

    If opcion1 = "5" Then  'producto
        If Len(Text1) = 0 Then
            cad = "select producto.Descripcio,producto.producto,precios.Unidad1,precios.Factor1,precios.pventa1 from producto inner join precios on producto.producto=precios.producto "

        End If

        If Len(Text1) > 0 Then
            cad = "select producto.Descripcio,producto.producto,precios.Unidad1,precios.Factor1,precios.pventa1 from producto inner join precios on producto.producto=precios.producto and   " & Combo2 & " like '" & Text1.Text & "%'"

        End If

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbgrid13.DataSource = mytablex
        dbgrid13.columns(0).Width = 3000
        dbgrid13.columns(1).Width = 1000
              
        If mytablex.RecordCount > 0 Then
            dbgrid13.SetFocus

        End If

    End If
      
    Exit Sub

End Sub

Private Sub DataGrid2_AfterColEdit(ByVal ColIndex As Integer)

    Select Case ColIndex

        Case 2
            
    End Select

End Sub

Private Sub Command5_Click()

End Sub

Private Sub dbgrid13_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim mytablex As New ADODB.Recordset

    Dim found    As Integer

    If KeyCode = 27 Then
        Text1.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = "1" Then
            'vHabitacion = "" & Trim("" & dbgrid13.columns("habitacion"))

            'Frame3.Visible = False
        End If

        If opcion1 = "2" Then
            tipocodigoh = Trim("" & dbgrid13.columns("tipo"))
   
            huesped = Trim("" & dbgrid13.columns("codigo"))
            hnombre = Trim("" & dbgrid13.columns("nombre"))
            'hdireccion = Trim("" & dbgrid13.columns("direccion"))
            'correo = Trim("" & dbgrid13.columns("correo"))
            hnombre.SetFocus
            Frame3.Visible = False
   
        End If

        If opcion1 = "6" Then
            tipocodigo = Trim("" & dbgrid13.columns("tipo"))
            codigo = Trim("" & dbgrid13.columns("codigo"))
            nombre = Trim("" & dbgrid13.columns("nombre"))
            'direccion = Trim("" & dbgrid13.columns("direccion"))
            'correo = Trim("" & dbgrid13.columns("correo"))
            nombre.SetFocus
            Frame3.Visible = False
   
        End If

        If opcion1 = "7" Then
            'huesped1 = Trim("" & dbgrid13.columns("codigo"))
            'hnombre1 = Trim("" & dbgrid13.columns("nombre"))
            'hdireccion1 = Trim("" & dbgrid13.columns("direccion"))
            'correo = Trim("" & dbgrid13.columns("correo"))
            'hnombre1.SetFocus
            'Frame3.Visible = False
   
        End If

        If opcion1 = "3" Then
            operador = Trim("" & dbgrid13.columns("codigo"))
            Frame3.Visible = False

        End If

        If opcion1 = "43" Then
            'idreserva = Trim("" & dbgrid13.columns("reserva"))
            'carga_reserva "" & idreserva
            Frame3.Visible = False

        End If

        If opcion1 = "4" Then
            agente = Trim("" & dbgrid13.columns("codigo"))
            Frame3.Visible = False

        End If

        If opcion1 = "5" Then
            mytablex.Open "select * from hotelcheckin where checkin=" & Val(checkin) & " and producto='" & Trim("" & dbgrid13.columns("producto")) & "'", cn, adOpenStatic, adLockOptimistic

            If mytablex.RecordCount = 0 Then
                mytablex.AddNew
                mytablex.Fields("checkin") = Val(checkin)
                mytablex.Fields("producto") = Trim("" & dbgrid13.columns("producto"))
                mytablex.Fields("descripcio") = Trim("" & dbgrid13.columns("descripcio"))
                mytablex.Fields("unidad") = Trim("" & dbgrid13.columns("unidad1"))
                mytablex.Fields("factor") = Val("" & dbgrid13.columns("factor1"))
                mytablex.Fields("cantidad") = 1
                mytablex.Fields("precio") = Val("" & dbgrid13.columns("pventa1"))
                mytablex.Fields("Total") = Val("" & dbgrid13.columns("pventa1"))
                mytablex.Update
            Else
                MsgBox "Ya Existe ", 48, "Aviso"
                Exit Sub

                'mytablex.Fields("checkin") = Val(checkin)
                'mytablex.Fields("producto") = Trim("" & dbgrid13.columns("producto"))
                'mytablex.Fields("descripcio") = Trim("" & dbgrid13.columns("descripcio"))
                'mytablex.Fields("unidad") = Trim("" & dbgrid13.columns("unidad1"))
                'mytablex.Fields("factor") = Val("" & dbgrid13.columns("factor1"))
                'mytablex.Fields("cantidad") = 1
                'mytablex.Fields("precio") = Val("" & dbgrid13.columns("pventa1"))
                'mytablex.Fields("Total") = Val("" & dbgrid13.columns("pventa1"))
                'mytablex.Update
            End If

            mytablex.Close
            Frame3.Visible = False

        End If

    End If

End Sub

Private Sub disponibles_Click()
    vHabitacion = Trim(extra_loquesea(disponibles))

End Sub

Private Sub dk88343_Click()

    Dim buf As String

    On Error GoTo cmd86712_err

    'If Frame5.Visible = True Then Exit Sub

    If Frame2.Visible = True Then Exit Sub

    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If Frame4.Visible = True Then Exit Sub

    buf = "" & txcheckinx.Fields("checkin")
    thotelan.idreserva = Trim(buf)
    'thotelan.idhabitacion = Trim("" & txcheckinx.Fields("habitacion"))
    thotelan.Show 1
    Exit Sub
cmd86712_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub dk9893_Click()

    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If Frame4.Visible = True Then Exit Sub

    If Frame2.Visible = True Then Exit Sub
    reporgen.NAMETABLA = "hotelcheckin"
    reporgen.Show 1

End Sub

Sub prueba_reporte()

    'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\checkinesproducto.rpt", "")
End Sub

Private Sub dki8834_Click()

    Dim buf As String

    On Error GoTo cmd186712_err

    'If Frame5.Visible = True Then Exit Sub

    If Frame2.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If Frame4.Visible = True Then Exit Sub

    buf = "" & txcheckinx.Fields("checkin")
    thotelco.idcheckin = Trim(buf)
    thotelco.idhabitacion = Trim("" & txcheckinx.Fields("habitacion"))
    thotelco.Show 1
    Exit Sub
cmd186712_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Enviatpv_Click()

    On Error GoTo cmd89123_err

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    Dim sdx      As Double

    Dim found    As Integer

    'If Frame5.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If Frame4.Visible = True Then Exit Sub

    If MsgBox("Desea enviar a la facturacion los Consumos ", 1, "Aviso") <> 1 Then Exit Sub
    mytablex.Open "select * from hotelconsumo where idcheckin=" & Val("" & txcheckinx.Fields("checkin")), cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    Do

        If mytablex.EOF Then Exit Do
        tptovta.Data2.Recordset.AddNew
        tptovta.Data2.Recordset.Fields("zona") = ""
        tptovta.Data2.Recordset.Fields("nroprecio") = "1"
        tptovta.Data2.Recordset.Fields("hora") = Format(Now, "hh:mm:ss")
        tptovta.Data2.Recordset.Fields("fecha") = "" & Format(Now, "dd/mm/yyyy")
        tptovta.Data2.Recordset.Fields("producto") = Trim("" & mytablex.Fields("producto"))
        tptovta.Data2.Recordset.Fields("vendedor") = ""
        tptovta.Data2.Recordset.Fields("descripcio") = Mid$(Trim("" & mytablex.Fields("descripcio")), 1, 80)

        tptovta.Data2.Recordset.Fields("cantidad") = Val("" & mytablex.Fields("cantidad"))
        tptovta.Data2.Recordset.Fields("unidad") = Trim("" & mytablex.Fields("unidad"))
        tptovta.Data2.Recordset.Fields("factor") = Val("" & mytablex.Fields("factor"))
        tptovta.Data2.Recordset.Fields("precio") = Val("" & mytablex.Fields("precio"))
        tptovta.Data2.Recordset.Fields("total") = Val("" & mytablex.Fields("total"))
        tptovta.Data2.Recordset.Fields("deslipo") = 0
        tptovta.Data2.Recordset.Fields("impuesto") = 0

        tptovta.Data2.Recordset.Fields("igv") = 18 'Val("" & mytablex.Fields("igv"))
        tptovta.Data2.Recordset.Fields("serviciopo") = 0 'Val("" & mytablex.Fields("serviciomesa"))
        tptovta.Data2.Recordset.Fields("descuento") = 0

        tptovta.Data2.Recordset.Fields("neto") = Val("" & mytablex.Fields("total"))
        tptovta.Data2.Recordset.Fields("FAMILIA") = Trim("" & mytablex.Fields("tipo"))
        tptovta.Data2.Recordset.Fields("ivap") = 0 'Val("" & mytablex.Fields("ivap"))
        tptovta.Data2.Recordset.Update
        mytablex.MoveNext
    Loop
    mytablex.Close
    vHabitacion = Trim("" & txcheckinx.Fields("checkin"))
    MsgBox "Proceso Realizado ", 48, "Aviso"
    dlo132_Click
    Exit Sub
cmd89123_err:
    MsgBox "seleccione un dato " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub fkichek_Click()

    'If Frame5.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If Frame4.Visible = True Then Exit Sub
    Frame4.Visible = True
    fechasalida = Format(Now, "dd/mm/yyyy")
    horasalida = Format(Now, "hh:mm:ss")

End Sub

Private Sub fkir84_Click()

    'If Frame5.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If Frame4.Visible = True Then Exit Sub
    reporte_excell txcheckinx

    'reporte
End Sub

Private Sub huesped_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_codigo

    End If

    If KeyCode = &H76 Then  'f1
        tnclie.DBPROV = "clientes"
        tnclie.fdlo893.Visible = True
        tnclie.Show 1

    End If

End Sub

Private Sub Image1_Click()
    consulta_vendedor

End Sub

Private Sub Image10_Click()
    tnclie.FLAG = "NUEVO"
    tnclie.DBPROV = "clientes"
    tnclie.fdlo893.Visible = True
    tnclie.Show 1

End Sub

Private Sub Image2_Click()
    consulta_agente

End Sub

Private Sub Image3_Click()
    consulta_codigo1

End Sub

Private Sub Image4_Click()

End Sub

Private Sub Image6_Click()
    consulta_codigo

End Sub

Private Sub Image8_Click()
    tnclie.FLAG = "NUEVO"
    tnclie.DBPROV = "clientes"
    tnclie.fdlo893.Visible = True
    tnclie.Show 1

End Sub

Private Sub Image9_Click()
    tnclie.FLAG = "NUEVO"
    tnclie.DBPROV = "clientes"
    tnclie.fdlo893.Visible = True
    tnclie.Show 1

End Sub

Private Sub jdfu7834_Click()

    On Error GoTo cmd9088_err

    'If Frame5.Visible = True Then Exit Sub

    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If Frame4.Visible = True Then Exit Sub

    Dim buf As String

    'buf = Trim("" & xcheckin)
    thotelpr.idcheckin = Trim("" & txcheckinx.Fields("checkin"))
    thotelpr.idhabitacion = ""
    thotelpr.Show 1

    'thotelct.idcheckin = Trim("" & txcheckinx.Fields("checkin"))
    'thotelct.idreserva = Trim("" & txcheckinx.Fields("idreserva"))
    'thotelct.habitacion = "" & txcheckinx.Fields("habitacion")
    'thotelct.Show 1
    Exit Sub
cmd9088_err:
    MsgBox "Seleccione un Datos", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Label2_Click()
    tipocodigo = tipocodigoh
    codigo = huesped
    nombre = hnombre

    'direccion = hdireccion
End Sub

Private Sub Label26_Click()

    If Frame2.Caption <> "Modifica" Then Exit Sub

    'Frame5.Caption = "NUEVO"
    'Frame5.Visible = True
    inicializa_huesped

    'huesped1.SetFocus
End Sub

Private Sub Label29_Click()

    On Error GoTo cmd89123_err

    If Frame2.Caption <> "Modifica" Then Exit Sub

    'huesped1 = Trim("" & mytablezz.Fields("huesped"))
    'Frame5.Caption = "MODIFICA"
    'Frame5.Visible = True
    inicializa_huesped
    'IDE = Trim("" & mytablezz.Fields("ide"))
    'huesped1 = Trim("" & mytablezz.Fields("huesped"))
    'hnombre1 = Trim("" & mytablezz.Fields("nombre"))
    'hdireccion1 = Trim("" & mytablezz.Fields("direccion"))
    'procedencia1 = Trim("" & mytablezz.Fields("procedencia"))
    'tipoviaje1 = Trim("" & mytablezz.Fields("tipoviaje"))
    'tipopersona1 = Trim("" & mytablezz.Fields("tipopersona"))
    'huesped1.SetFocus
    Exit Sub
cmd89123_err:
    MsgBox "Seleccione un Dato", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Label30_Click()

    On Error GoTo cmd289123_err

    If Frame2.Caption <> "Modifica" Then Exit Sub

    'huesped1 = Trim("" & mytablezz.Fields("huesped"))
    'IDE = Trim("" & mytablezz.Fields("ide"))
    'Frame5.Caption = "BORRA"
    'grabar_huespedes
    Exit Sub
cmd289123_err:
    MsgBox "Seleccione un Dato", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Label31_Click()

    'Frame5.Visible = True
End Sub

Private Sub Label6_Click()
    consulta_producto

End Sub

Private Sub Label7_Click()
    consulta_codigo

End Sub

Private Sub mesa_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_producto

    End If

End Sub

Private Sub Label9_Click()

    'huesped1 = codigo
    'hnombre1 = nombre
    'hdireccion1 = direccion
End Sub

Private Sub nfpago_Click()

    If nfpago <> "%" Then
        bfpago = extra_loquesea1(nfpago)

    End If

End Sub

Private Sub noches_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    valida_horas
    suma_total

End Sub

Sub valida_horas()

    Dim hoyi As Date

    Dim hoyf As Date

    Dim xvar

    If Not IsDate(arribofecha) Then Exit Sub
    If Val(noches) = 0 Then
        noches = "1"

    End If

    If categoria = "DIAS" Then
        hoyi = CVDate(arribofecha)
        hoyf = DateAdd("D", Int(noches), hoyi)
        arribofechaf = Format(hoyf, "dd/mm/yyyy")
        arribohoraf = "13:00:00"

    End If

    If categoria = "HORAS" Then
        arribofechaf = Format(arribofecha, "dd/mm/yyyy")
        suma_lashoras

    End If

End Sub

Private Sub nombre_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_codigo

    End If

End Sub

Private Sub checkin_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        dlo132_Click
        Exit Sub

    End If

    If Len(checkin) = 0 Then Exit Sub

    'descripcio.SetFocus
End Sub

Private Sub Command1_Click()
    Frame1.Visible = True
    Frame1.Enabled = True
    'buffer = ""
    opcion1 = "1"
    ejecuta 1

End Sub

Sub ejecuta(sw As Integer)

    Dim cad As String

    Dim sdx As Double

    cad = "SELECT * from hotelcheckin  "
    cad = cad & " where estado like '" & estado1 & "'"

    If Len(xhabitacion) > 0 Then
        cad = cad & " and habitacion='" & xhabitacion & "' "

    End If

    If Len(buffer) > 0 Then
        cad = cad & " and " & Combo1 & " like '" & buffer & "%'"

    End If

    cad = cad & "order by habitacion ,arribofecha"
   
    If txcheckinx.State = 1 Then txcheckinx.Close
    txcheckinx.Open cad, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = txcheckinx

    'dbGrid1.columns(0).Width = 4000
    'dbGrid1.columns(1).Width = 2000
    If txcheckinx.RecordCount > 0 Then

        'dbGrid1.SetFocus
    End If

    sumar_todo txcheckinx

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then

        'checkin = dbGrid1.Columns(1)
        'Frame1.Visible = False
        'Frame1.Enabled = False
        'checkin.SetFocus
        'checkin_KeyPress 13
    End If

End Sub

Private Sub dlo132_Click()

    If Frame3.Visible = True Then
        Frame3.Visible = False
        Exit Sub

    End If

    If vflag = "NUEVO" Then
        tcheckin.Hide
        Unload tcheckin
        Exit Sub

    End If

    'If Frame5.Visible = True Then
    '   Frame5.Visible = False
    '   Exit Sub
    'End If

    If Frame2.Visible = True Then
        habilita 0
        Frame2.Visible = False
        dbGrid1.Enabled = True
        ejecuta 1
   
        Exit Sub

    End If

    If Frame4.Visible = True Then
        Frame4.Visible = False
        Exit Sub

    End If

    tcheckin.Hide
    Unload tcheckin

End Sub

Private Sub f8443_Click()

    Dim buf As String

    On Error GoTo cmd456_err

    'If Frame5.Visible = True Then Exit Sub

    buf = "" & txcheckinx.Fields("checkin")

    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If Frame4.Visible = True Then Exit Sub

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
    'refresca_huesped
    checkin.Enabled = False
    operador.SetFocus
    Exit Sub
cmd456_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fjh433_Click()

    Dim buf As String

    On Error GoTo cmd556_err

    'If Frame5.Visible = True Then Exit Sub

    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If Frame4.Visible = True Then Exit Sub

    buf = "" & txcheckinx.Fields("checkin")

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
    'refresca_huesped
    checkin.Enabled = False
    operador.SetFocus
    Exit Sub
cmd556_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Form_Activate()
    'agregar_menus
    Enviatpv.Visible = False
    jdfu7834.Visible = False 'estado cuenta
    fkichek.Visible = False
    dki8834.Visible = False
    dk88343.Visible = False

    If xsw = "CONSUMO" Then
        dki8834.Visible = True
        habilita 1
        dbGrid1.Enabled = True

    End If

    If xsw = "TPV" Then
        Enviatpv.Visible = True
        dki8834.Visible = True
        habilita 1
        dbGrid1.Enabled = True

    End If

    If xsw = "ANTICIPO" Then
        dk88343.Visible = True
        habilita 1
        dbGrid1.Enabled = True

    End If

    If xsw = "SALIDA" Then
        fkichek.Visible = True
        habilita 1
        dbGrid1.Enabled = True

    End If

    If xsw = "PRECUENTA" Then
        jdfu7834.Visible = True
        habilita 1
        dbGrid1.Enabled = True

    End If

    Command1_Click

    If vflag = "NUEVO" Then
        ajdu1_Click

    End If

End Sub

Sub consulta_mesas()

    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.AddItem "Habitacion"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "1"
    Text1.SetFocus
    Command4_Click

End Sub

Sub consulta_codigo()
    Combo2.Clear
    Combo2.AddItem "Nombre"
    Combo2.AddItem "Codigo"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "2"
    Text1.SetFocus
    Command4_Click

End Sub

Sub consulta_codigo1()
    Combo2.Clear
    Combo2.AddItem "Nombre"
    Combo2.AddItem "Codigo"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "6"
    Text1.SetFocus
    Command4_Click

End Sub

Sub consulta_codigo2()
    Combo2.Clear
    Combo2.AddItem "Nombre"
    Combo2.AddItem "Codigo"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "7"
    Text1.SetFocus
    Command4_Click

End Sub

Sub consulta_producto()
    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.AddItem "Producto"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "5"
    Text1.SetFocus
    Command4_Click

End Sub

Sub consulta_reserva()
    Combo2.Clear
    Combo2.AddItem "Nombre"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "43"
    Text1.SetFocus
    Command4_Click

End Sub

Sub consulta_vendedor()
    Combo2.Clear
    Combo2.AddItem "Nombre"
    Combo2.AddItem "Codigo"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "3"
    Text1.SetFocus
    Command4_Click

End Sub

Sub consulta_agente()
    Combo2.Clear
    Combo2.AddItem "Nombre"
    Combo2.AddItem "Codigo"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "4"
    Text1.SetFocus
    Command4_Click

End Sub

Private Sub Form_Load()

    Dim mytablex As New ADODB.Recordset

    carga_tipopension
    carga_tipotarifa
    carga_tiporeserva
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.ListIndex = 0

    Combo3.Clear
    Combo3.AddItem ""
    Combo3.AddItem "RESERVA"
    Combo3.AddItem "ENTRADA"
    Combo3.AddItem "CERRADO"
    Combo3.ListIndex = 0

    Combo4.Clear
    Combo4.AddItem ""
    Combo4.AddItem "DIAS"
    Combo4.AddItem "HORAS"
    Combo4.ListIndex = 0

    estado1.Clear
    estado1.AddItem "%"
    estado1.AddItem "RESERVA"
    estado1.AddItem "ENTRADA"
    estado1.AddItem "CERRADO"
    estado1.ListIndex = 0

    'icategoria.Clear
    'icategoria.AddItem "DIAS"
    'icategoria.AddItem "HORAS"
    'icategoria.ListIndex = 0
    ntipo.Clear
    ntipo.AddItem "%"
    mytablex.Open "select * from tipo ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        If Trim("" & mytablex.Fields("tipodoc")) = "A" Or Trim("" & mytablex.Fields("tipodoc")) = "B" Or Trim("" & mytablex.Fields("tipodoc")) = "C" Or Trim("" & mytablex.Fields("tipodoc")) = "D" Or Trim("" & mytablex.Fields("tipodoc")) = "G" Then
            ntipo.AddItem Trim("" & mytablex.Fields("descripcio")) & "|" & Trim("" & mytablex.Fields("tipo"))

        End If

        mytablex.MoveNext
    Loop
    mytablex.Close
    ntipo.ListIndex = 0

    nfpago.Clear
    nfpago.AddItem "%"
    mytablex.Open "select * from fpago ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        nfpago.AddItem Trim("" & mytablex.Fields("descripcio")) & "|" & Trim("" & mytablex.Fields("fpago"))
        mytablex.MoveNext
    Loop
    mytablex.Close
    nfpago.ListIndex = 0

End Sub

Sub inicializa()
    hotelcuadre = Trim("" & treevho.turno)
    tipocodigo = ""
    tipocodigoh = ""
    bserie = ""
    bnumero = ""
    btipo = ""
    bfpago = ""
    brecibe = ""

    'horas = ""
    precio = ""
    estado = "" & vestado
    Label23 = "DIAS"
    'reservadas.Clear
    tipotarifa = ""
    tiporeserva = ""
    tipopension = ""
    personas = "1"
    tiporeserva = ""
    horareserva = Format(Now, "HH:MM:SS")
    fechareserva = Format(Now, "dd/mm/yyyy")
    Combo4.ListIndex = 0
    categoria = "DIAS"
    'idreserva = ""
    arribofecha = Format(Now, "dd/mm/yyyy")
    arribofechaf = Format(Now, "dd/mm/yyyy")
    arribohora = Format(Now, "hh:mm:ss")
    arribohoraf = "13:00:00"  'Format(Now, "hh:mm")
    vHabitacion = ""
    nombre = ""
    codigo = ""
    'direccion = ""
    huesped = ""
    hnombre = ""
    'hdireccion = ""
    'tipoviaje = ""
    'procedencia = ""
    agente = ""
    operador = Trim(gusuario)
    'adulto = ""
    'nino = ""
    noches = "1"
    valida_horas
    vHabitacion = Trim(xhabitacion)
    carga_precio Trim(xhabitacion)
    suma_total

    'carga_precio Trim("" & vHabitacion)
    'tipoviaje = ""
End Sub

Sub inicializa_huesped()

    'huesped1 = ""
    'hnombre1 = ""
    'hdireccion1 = ""
    'tipopersona1 = "ADULTO"
    'tipoviaje1 = ""
    'procedencia1 = "PERU"
End Sub

Sub pone_registro()
    tipocodigoh = Trim("" & txcheckinx.Fields("tipocodigoh"))
    tipocodigo = Trim("" & txcheckinx.Fields("tipocodigo"))
    tipotarifa = Trim("" & txcheckinx.Fields("tipotarifa"))
    tiporeserva = Trim("" & txcheckinx.Fields("tiporeserva"))
    tipopension = Trim("" & txcheckinx.Fields("tipopension"))

    personas = Trim("" & txcheckinx.Fields("personas"))
    tiporeserva = Trim("" & txcheckinx.Fields("tiporeserva"))
    horareserva = Trim("" & txcheckinx.Fields("horareserva"))
    fechareserva = Trim("" & txcheckinx.Fields("fechareserva"))
    categoria = Trim("" & txcheckinx.Fields("categoria"))
    Label23 = categoria
    checkin = Trim("" & txcheckinx.Fields("checkin"))
    arribofecha = Trim("" & txcheckinx.Fields("arribofecha"))
    arribofechaf = Trim("" & txcheckinx.Fields("arribofechaf"))
    arribohora = Trim("" & txcheckinx.Fields("arribohora"))
    arribohoraf = Trim("" & txcheckinx.Fields("arribohoraf"))
    noches = Trim("" & txcheckinx.Fields("noches"))

    codigo = Trim("" & txcheckinx.Fields("codigo"))
    nombre = Trim("" & txcheckinx.Fields("nombre"))
    'direccion = Trim("" & txcheckinx.Fields("direccion"))

    huesped = Trim("" & txcheckinx.Fields("huesped"))
    hnombre = Trim("" & txcheckinx.Fields("hnombre"))
    'hdireccion = Trim("" & txcheckinx.Fields("hdireccion"))

    'procedencia = Trim("" & txcheckinx.Fields("procedencia"))
    operador = Trim("" & txcheckinx.Fields("operador"))
    agente = Trim("" & txcheckinx.Fields("agente"))
    precio = Trim("" & txcheckinx.Fields("precio"))

    'adulto = Trim("" & txcheckinx.Fields("adulto"))
    'nino = Trim("" & txcheckinx.Fields("nino"))
    vHabitacion = Trim("" & txcheckinx.Fields("habitacion"))
    estado = Trim("" & txcheckinx.Fields("estado"))
    hotelcuadre = Trim("" & txcheckinx.Fields("hotelcuadre"))

End Sub

Sub grabando()

    Dim X As Integer

    If Val(hotelcuadre) <= 0 Then
        hotelcuadre = Trim("" & treevho.turno)

    End If

    txcheckinx.Fields("hotelcuadre") = Val(hotelcuadre)
    txcheckinx.Fields("tipocodigo") = Trim(tipocodigo)
    txcheckinx.Fields("tipocodigoh") = Trim(tipocodigoh)
    txcheckinx.Fields("tipopension") = Trim(tipopension)
    txcheckinx.Fields("tipotarifa") = Trim(tipotarifa)
    txcheckinx.Fields("tiporeserva") = Trim(tiporeserva)

    txcheckinx.Fields("personas") = Val(personas)
    txcheckinx.Fields("tiporeserva") = Trim(tiporeserva)
    txcheckinx.Fields("fechareserva") = Trim(fechareserva)
    txcheckinx.Fields("horareserva") = Trim(horareserva)
    'txcheckinx.Fields("precio") = Val(precio)
    txcheckinx.Fields("categoria") = Trim(categoria)
    'txcheckinx.Fields("idreserva") = Val(idreserva)
    txcheckinx.Fields("arribofecha") = Trim(arribofecha)
    txcheckinx.Fields("arribohora") = Trim(arribohora)
    txcheckinx.Fields("arribofechaf") = Trim(arribofechaf)
    txcheckinx.Fields("arribohoraf") = Trim(arribohoraf)

    txcheckinx.Fields("codigo") = Trim(codigo)
    txcheckinx.Fields("nombre") = Trim(nombre)
    'txcheckinx.Fields("direccion") = Trim(direccion)

    txcheckinx.Fields("huesped") = Trim(huesped)
    txcheckinx.Fields("hnombre") = Trim(hnombre)
    'txcheckinx.Fields("hdireccion") = Trim(hdireccion)

    'txcheckinx.Fields("procedencia") = Trim(procedencia)
    txcheckinx.Fields("agente") = Trim(agente)
    txcheckinx.Fields("operador") = Trim(operador)
    txcheckinx.Fields("estado") = Trim(estado)

    txcheckinx.Fields("noches") = Val(noches)
    'txcheckinx.Fields("adulto") = Val(adulto)
    txcheckinx.Fields("precio") = Val(precio)
    'txcheckinx.Fields("habitacion") = Trim(vHabitacion)
    txcheckinx.Fields("habitacion") = Trim(vHabitacion)

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
        'If Len(checkin) = 0 Then
        '   checkin.SetFocus
        '   Exit Function
        'End If
        'rbusca.Open "select checkin from checkin where checkin='" & checkin & "'", cn, adOpenStatic, adLockOptimistic
        'If rbusca.RecordCount > 0 Then
        '   rbusca.Close
        '   MsgBox "Ya existe checkin ", 48, "Aviso"
        '   Exit Function
        'End If
        txcheckinx.AddNew
        'txcheckinx.Fields("checkin") = checkin
        grabando
        'if estado="RESER"
        'txcheckinx.Fields("estado") = "0"
        'txcheckinx.Fields("estado1") = "0"
        'If estado = "RESERVA" Then
        'txcheckinx.Fields("estado") = "R"
        'End If
        'If estado = "ENTRADA" Then
        'txcheckinx.Fields("estado") = "e"
        'End If
        txcheckinx.Update

        If estado = "ENTRADA" Then
            actualiza_habitacion "" & txcheckinx.Fields("habitacion")

        End If

        If Check1.Value = 1 Then
            graba_factura Trim("" & txcheckinx.Fields("checkin"))

        End If

        dlo132_Click
        Exit Function

    End If

    If Frame2.Caption = "Modifica" Then
        'txcheckinx.Fields("checkin") = checkin
        grabando
        'If estado = "RESERVA" Then
        'txcheckinx.Fields("estado") = "R"
        'End If
        'If estado = "ENTRADA" Then
        'txcheckinx.Fields("estado") = "E"
        'End If
        txcheckinx.Update
        dlo132_Click
        Exit Function

    End If

End Function

Function valida()

    If Not IsDate(fechareserva) Then
        fechareserva = Format(Now, "dd/mm/yyyy")

    End If

    If Len(Trim(horareserva)) = 0 Then
        horareserva = Format(Now, "HH:MM:SS")

    End If

    If Len(Trim(arribofecha)) < 10 Or Not IsDate(Trim(arribofecha)) Then
        arribofecha.SetFocus
        Exit Function

    End If

    If Len(Trim(arribofecha)) < 10 Or Not IsDate(Trim(arribofecha)) Then
        arribofecha.SetFocus
        Exit Function

    End If

    If Len(Trim(arribohora)) <> 8 Then
        arribohora.SetFocus
        Exit Function

    End If

    If Len(Trim(arribohoraf)) <> 8 Then
        arribohora.SetFocus
        Exit Function

    End If

    If tipocodigo <> "J" And tipocodigo <> "X" And tipocodigo <> "N" And tipocodigo <> "D" And tipocodigo <> "O" And tipocodigo <> "P" Then
        tipocodigo.SetFocus
        Exit Function

    End If

    If Len(Trim(codigo)) = 0 Then
        codigo.SetFocus
        Exit Function

    End If

    If Len(Trim(nombre)) = 0 Then
        nombre.SetFocus
        Exit Function

    End If

    valida_horas

    If tipocodigoh <> "J" And tipocodigoh <> "X" And tipocodigoh <> "N" And tipocodigoh <> "D" And tipocodigoh <> "O" And tipocodigoh <> "P" Then
        tipocodigoh.SetFocus
        Exit Function

    End If

    If Len(Trim(huesped)) = 0 Then
        huesped.SetFocus
        Exit Function

    End If

    If Len(Trim(hnombre)) = 0 Then
        hnombre.SetFocus
        Exit Function

    End If

    If Len(Trim(operador)) = 0 Then
        operador.SetFocus
        Exit Function

    End If

    If Len(Trim(estado)) = 0 Then
        MsgBox "Ingresar Estado del Documento ", 48, "Aviso"
        Exit Function

    End If

    If Val(precio) <= 0 Then
        If estado = "ENTRADA" Then
            MsgBox "Debe Seleccionar un Precio", 48, "Aviso"
            precio.SetFocus
            Exit Function

        End If

    End If

    If Check1.Value = 1 Then

        'If estado = "ENTRADA" Then
        If Len(Trim(btipo)) = 0 Then
            MsgBox "Seleccione Tipo Documento ", 48, "Aviso"
            Exit Function

        End If

        If Len(Trim(bserie)) = 0 Then
            MsgBox "Seleccione Serie Documento ", 48, "Aviso"
            'bserie.SetFocus
            Exit Function

        End If

        If Len(Trim(bnumero)) = 0 Then
            MsgBox "Seleccione numero Documento ", 48, "Aviso"
            Exit Function

        End If

        If Len(Trim(bfpago)) = 0 Then
            MsgBox "Seleccione Forma Pago ", 48, "Aviso"
            Exit Function

        End If

        If Val(brecibe) <= 0 Then
            MsgBox "Seleccione Valor Pagado ", 48, "Aviso"
            Exit Function

        End If
   
        ' End If
   
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

    mytablex.Open "select * from archivo where menu='checkin' and   estado='S'", cn, adOpenStatic, adLockOptimistic

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
    mytablex.Open "select * from archivo where menu='checkin' and descripcio='" & Trim(buf) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close

    End If

    'busca el reporte
    buf = mytablex.Fields("archivo")
    mytablex.Close

    'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\" & buf, "")
End Sub

Private Sub reserva_Change()

End Sub

Private Sub vendedor_Change()

End Sub

Private Sub ntipo_Click()

    Dim sdx As Double

    If ntipo <> "%" Then
        btipo = Trim(extra_loquesea1(ntipo))
        busca_parameca

    End If

End Sub

Private Sub ntipopension_Click()

    If Len(Trim(ntipopension.Text)) = 0 Then
        tipopension = ""
        Exit Sub

    End If

    tipopension = Trim(extra_loquesea1(ntipopension.Text))

End Sub

Private Sub ntiporeserva_Click()

    If Len(Trim(ntiporeserva.Text)) = 0 Then
        tiporeserva = ""
        Exit Sub

    End If

    tiporeserva = Trim(extra_loquesea1(ntiporeserva.Text))

End Sub

Private Sub ntipotarifa_Click()

    If Len(Trim(ntipotarifa.Text)) = 0 Then
        tipotarifa = ""
        Exit Sub

    End If

    tipotarifa = Trim(extra_loquesea1(ntipotarifa.Text))

End Sub

Private Sub operador_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_vendedor

    End If

End Sub

Sub carga_reserva(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "Select * from reserva where reserva=" & buf & "", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        'idreserva = Trim("" & mytablex.Fields("reserva"))
        arribofecha = Trim("" & mytablex.Fields("arribofecha"))
        arribofechaf = Trim("" & mytablex.Fields("arribofechaf"))
        arribohora = Trim("" & mytablex.Fields("arribohora"))
        arribohoraf = Trim("" & mytablex.Fields("arribohoraf"))
        nombre = Trim("" & mytablex.Fields("nombre"))
        'procedencia = Trim("" & mytablex.Fields("procedencia"))
        operador = Trim("" & mytablex.Fields("operador"))
        agente = Trim("" & mytablex.Fields("agente"))
        'adulto = Trim("" & mytablex.Fields("adulto"))
        'nino = Trim("" & mytablex.Fields("nino"))

    End If

    mytablex.Close

End Sub

Sub Reporte()

    Dim found As Integer

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    borrar_archivo FileName
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_documento1
    cuerpo_programa_documento1
    '------------------------------------
    Close #1
    cerrar_archivo
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1

End Sub

Sub cabecera_documento1()

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
    buf = "Reporte de Habitaciones  "
    found = formateaa(buf, 90, 2, 0)
    
    found = formateaa("CheckIn", 8, 0, 0)
    found = formateaa("Reserva", 8, 0, 0)
    found = formateaa("Nombre", 51, 0, 0)
    found = formateaa("Habi", 7, 0, 0)
    found = formateaa("Entrada ", 22, 0, 0)
    found = formateaa("Salida ", 22, 0, 0)
    found = formateaa("Dias ", 5, 0, 0)
    found = formateaa("Total ", 7, 0, 0)
    found = formateaa("Abono ", 7, 0, 0)
    found = formateaa("Saldo ", 7, 2, 0)
        
    found = formateaa("", 8, 0, 0)
    found = formateaa("", 60, 0, 0)
    found = formateaa("", 7, 0, 0)
    found = formateaa("Fecha ", 11, 0, 0)
    found = formateaa("Hora ", 11, 0, 0)
    found = formateaa("Fecha ", 11, 0, 0)
    found = formateaa("Hora ", 11, 2, 0)
    
    found = formateaa("", 10, 0, 0)
    buf = "Tipo"
    found = formateaa(buf, 2, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Fecha"
    found = formateaa(buf, 10, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Producto"
    found = formateaa(buf, 7, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Descripcio"
    found = formateaa(buf, 20, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Und"
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Fac"
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Cant"
    found = formateaa(buf, 7, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Precio"
    found = formateaa(buf, 7, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Total"
    found = formateaa(buf, 7, 0, 0)
    found = formateaa("", 1, 2, 0)
    
    buf = String(150, "-")
    found = formateaa(buf, 150, 2, 0)

End Sub

Sub cuerpo_programa_documento1()

    Dim buf   As String

    Dim found As Integer

    Dim sdx   As Double

    On Error GoTo cmd78812_err

    Do

        If txcheckinx.EOF Then Exit Do
        buf = "+" & txcheckinx.Fields("checkin")
        found = formateaa(buf, 7, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & txcheckinx.Fields("idreserva")
        found = formateaa(buf, 7, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & txcheckinx.Fields("nombre")
        found = formateaa(buf, 50, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & txcheckinx.Fields("habitacion")
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & txcheckinx.Fields("arribofecha")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & txcheckinx.Fields("arribohora")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & txcheckinx.Fields("arribofechaf")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & txcheckinx.Fields("arribohoraf")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
      
        buf = "" & txcheckinx.Fields("noches")
        found = formateaa(buf, 4, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & suma_consumos("" & txcheckinx.Fields("checkin"))
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & suma_abonos("" & txcheckinx.Fields("checkin"), "" & txcheckinx.Fields("idreserva"))
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        sdx = -suma_abonos("" & txcheckinx.Fields("checkin"), "" & txcheckinx.Fields("idreserva")) + suma_consumos("" & txcheckinx.Fields("checkin"))
        buf = "" & sdx
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 2, 0)
             
        nlineas
        imprime_consumos "" & txcheckinx.Fields("checkin")
        txcheckinx.MoveNext
    Loop
    Exit Sub
cmd78812_err:
    MsgBox "Aviso en cuerpo " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub nlineas()
    contlin = contlin + 1

    If contlin > 45 Then
        cabecera_documento1

    End If

End Sub

Sub imprime_consumos(buf1 As String)

    Dim buf      As String

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from hotelconsumo where idcheckin=" & buf1, cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        found = formateaa("", 10, 0, 0)
        buf = "-" & mytablex.Fields("tipo")
        found = formateaa(buf, 2, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("fecha")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("Producto")
        found = formateaa(buf, 7, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("descripcio")
        found = formateaa(buf, 20, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("unidad")
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("factor")
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("cantidad")
        found = formateaa(buf, 7, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("precio")
        found = formateaa(buf, 7, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("total")
        found = formateaa(buf, 7, 0, 0)
        found = formateaa("", 1, 2, 0)
        nlineas
        mytablex.MoveNext
    Loop
    mytablex.Close

End Sub

Function suma_consumos(buf1 As String) As Double

    Dim buf   As String

    Dim found As Integer

    Dim sdx   As Double

    sdx = 0

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from hotelconsumo where idcheckin=" & buf1, cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        sdx = sdx + Val("" & mytablex.Fields("total"))
        mytablex.MoveNext
    Loop
    mytablex.Close
    suma_consumos = sdx

End Function

Function suma_abonos(buf1 As String, buf2 As String) As Double

    Dim buf   As String

    Dim found As Integer

    Dim sdx   As Double

    Dim sdx1  As Double

    sdx = 0
    sdx1 = 0

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from hotelanticipo where idcheckin=" & buf1, cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        sdx = sdx + Val("" & mytablex.Fields("monto"))
        mytablex.MoveNext
    Loop
    mytablex.Close
 
    mytablex.Open "select * from hotelanticipo where idreserva=" & buf2, cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        sdx1 = sdx1 + Val("" & mytablex.Fields("monto"))
        mytablex.MoveNext
    Loop
    mytablex.Close
    suma_abonos = sdx + sdx1
 
End Function

Sub carga_precio(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from habitacion where habitacion='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        precio = Trim("" & mytablex.Fields("precio"))
        brecibe = precio

    End If

    mytablex.Close
            
End Sub

Sub carga_tiporeserva()

    Dim mytablex As New ADODB.Recordset

    ntiporeserva.Clear
    ntiporeserva.AddItem ""
    mytablex.Open "select * from tiporeserva", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        ntiporeserva.AddItem Trim("" & mytablex.Fields("descripcio")) & "|" & Trim("" & mytablex.Fields("tiporeserva"))
        mytablex.MoveNext
    Loop
    mytablex.Close
    ntiporeserva.ListIndex = 0

End Sub

Sub carga_tipotarifa()

    Dim mytablex As New ADODB.Recordset

    ntipotarifa.Clear
    ntipotarifa.AddItem ""
    mytablex.Open "select * from tipotarifa", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        ntipotarifa.AddItem Trim("" & mytablex.Fields("descripcio")) & "|" & Trim("" & mytablex.Fields("tipotarifa"))
        mytablex.MoveNext
    Loop
    mytablex.Close
    ntipotarifa.ListIndex = 0

End Sub

Sub carga_tipopension()

    Dim mytablex As New ADODB.Recordset

    ntipopension.Clear
    ntipopension.AddItem ""
    mytablex.Open "select * from tipopension", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        ntipopension.AddItem Trim("" & mytablex.Fields("descripcio")) & "|" & Trim("" & mytablex.Fields("tipopension"))
        mytablex.MoveNext
    Loop
    mytablex.Close
    ntipopension.ListIndex = 0

End Sub

Sub busca_habitacionlibre()

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    If Not IsDate(arribofecha) Then
        MsgBox "Verifica arribo fecha"
        Exit Sub

    End If

    If Not IsDate(arribofechaf) Then
        MsgBox "Verifica arribo fecha"
        Exit Sub

    End If

    disponibles.Clear
    mytablex.Open "select * from habitacion ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        buf = "SELECT     dbo.hotelcheckin.checkin, dbo.hotelcheckin.habitacion"
        buf = buf & " FROM         dbo.hotelcheckin"
        buf = buf & " where  (dbo.hotelcheckin.arribofecha>='" & Format(arribofecha, "YYYYMMDD") & "'"
        buf = buf & " and dbo.hotelcheckin.arribofecha<='" & Format(arribofechaf, "YYYYMMDD") & "') "
        buf = buf & " and dbo.hotelcheckin.habitacion='" & Trim("" & mytablex.Fields("habitacion")) & "'"
        mytabley.Open buf, cn, adOpenStatic, adLockOptimistic

        If mytabley.RecordCount = 0 Then
            disponibles.AddItem Trim("" & mytablex.Fields("habitacion")) & "|" & Trim("" & mytablex.Fields("descripcio"))

        End If

        mytabley.Close
        mytablex.MoveNext
    Loop
    mytablex.Close

End Sub

Private Function CalculateTime(Time As Double) As String

    Dim TimeHour    As Double

    Dim TimeMin     As Double

    Dim TimeSec     As Byte

    Dim CalcTime    As String

    Dim strtimemin  As String

    Dim strtimesec  As String

    Dim strtimehour As String

    'Calculate the actual times
    TimeHour = Int((Time / 60) / 60)
    TimeMin = Int(Time / 60)
    TimeMin = TimeMin Mod 60
    TimeSec = Int(Time Mod 60)
  
    'Change times to appropriate formats
    strtimemin = "" & TimeMin

    If Len(strtimemin) = 1 Then
        strtimemin = "0" & strtimemin

    End If
  
    strtimesec = "" & TimeSec

    If Len(strtimesec) = 1 Then
        strtimesec = "0" & strtimesec

    End If
  
    strtimehour = "" & TimeHour

    If Len(strtimehour) = 1 Then
        strtimehour = "0" & strtimehour

    End If
  
    'MsgBox strtimehour & ":" & strtimemin & ":" & strtimesec
  
    'Assign the appropriate values to the function
    CalculateTime = strtimehour & ":" & strtimesec & ":" & strtimemin
  
End Function

Sub suma_lashoras()

    Dim FirstTime  As Double

    Dim SecondTime As Double

    Dim vTotal     As Double

    Dim txthours1  As Double

    Dim txtmin1    As Double

    Dim txtsec1    As Double

    Dim txthours2  As Double

    Dim txtmin2    As Double

    Dim txtsec2    As Double

    'MsgBox arribohora
    txthours1 = Val(Mid$("" & arribohora, 1, 2))
    txtmin1 = Val(Mid$("" & arribohora, 4, 2))
    txtsec1 = Val(Mid$("" & arribohora, 7, 2))

    txthours2 = Val("" & noches)
    txtmin2 = 0
    txtsec2 = 0
    'Convert the Hours and minutes to seconds, and add them up
    FirstTime = ((txthours1 * 60) * 60) + (txtmin1) + txtsec1
    SecondTime = ((txthours2 * 60) * 60) + (txtmin2) + txtsec2
  
    'SecondTime = ((txthours2 * 60)) + (txtmin2) + txtsec2
  
    'Add the two times
    vTotal = FirstTime + SecondTime
    'MsgBox vTotal
    'MsgBox FirstTime & " " & SecondTime
  
    'Assign the appropriate values to the correct textboxes
    'txtTotalSec.Text = vTotal

    arribohoraf = CalculateTime(vTotal)

End Sub

Sub reporte_excell(mytablex As ADODB.Recordset)

    Dim found       As Integer

    Dim buf         As String

    Dim buf1        As String

    Dim sdx         As Double

    Dim sdx1        As Double

    Dim Heading(15) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    Dim v           As Long

    Dim h           As Long

    Command1.Visible = True

    On Error GoTo cmd6561245_err
    
    Heading(1) = "Habitacion"
    Heading(2) = "Estado"
    Heading(3) = "FechaEnt."
    Heading(4) = "FechaSal."
    Heading(5) = "H.Ingreso"
    Heading(6) = "H.Salida"
    Heading(7) = "Apellidos y Nombres"
    Heading(8) = "Doc.Ident"
    Heading(9) = "Categ."
    Heading(10) = "Precio"
    Heading(11) = "Perman."
    Heading(12) = "Total"
    Heading(13) = "Id"
    
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_orden(15, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook

    objExcel.ActiveSheet.Cells(1, 1) = "HABITACION CON DATOS"
    objExcel.ActiveSheet.Cells(2, 1) = "FECHA HOY  " + Format(Now, "dd/mm/yyyy") & "- HORA HOY  " + Format(Now, "HH:MM:SS")

    v = 4
    h = 1
    sdx1 = 0
    
    Do

        If mytablex.EOF Then Exit Do
        objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("habitacion")
        objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("estado")
        objExcel.ActiveSheet.Cells(v, h + 2) = "'" & mytablex.Fields("arribofecha")
        objExcel.ActiveSheet.Cells(v, h + 3) = "'" & mytablex.Fields("arribofechaf")
        objExcel.ActiveSheet.Cells(v, h + 4) = "'" & mytablex.Fields("arribohora")
        objExcel.ActiveSheet.Cells(v, h + 5) = "'" & mytablex.Fields("arribohoraf")
        objExcel.ActiveSheet.Cells(v, h + 6) = "'" & mytablex.Fields("hnombre")
        objExcel.ActiveSheet.Cells(v, h + 7) = "'" & mytablex.Fields("huesped")
        objExcel.ActiveSheet.Cells(v, h + 8) = "'" & mytablex.Fields("categoria")
        objExcel.ActiveSheet.Cells(v, h + 9) = "" & mytablex.Fields("Precio")

        Dim Fecha1 As Date

        Dim Fecha2 As Date

        Dim meses  As Integer

        If Trim("" & mytablex.Fields("categoria")) = "DIAS" Then
            Fecha1 = Format(Now, "yyyy/mm/dd")
            Fecha2 = Format("" & mytablex.Fields("arribofecha"), "yyyy/mm/dd")
            meses = DateDiff("d", Fecha2, Fecha1)
            buf1 = Format(meses, "###0")

            If Val(buf1) = 0 Then
                buf1 = "1"

            End If

            objExcel.ActiveSheet.Cells(v, h + 10) = buf1
            sdx = Val(buf1) * Val("" & mytablex.Fields("Precio"))
            objExcel.ActiveSheet.Cells(v, h + 11) = "" & sdx
            sdx1 = sdx1 + sdx

        End If

        If Trim("" & mytablex.Fields("categoria")) = "HORAS" Then
            objExcel.ActiveSheet.Cells(v, h + 10) = "1"
            sdx = Val("" & mytablex.Fields("Precio"))
            objExcel.ActiveSheet.Cells(v, h + 11) = "" & sdx
            sdx1 = sdx1 + sdx

        End If

        objExcel.ActiveSheet.Cells(v, h + 12) = "'" & mytablex.Fields("CHECKIN")
        v = v + 1
        mytablex.MoveNext
    Loop
    objExcel.ActiveSheet.Cells(v, h + 11) = "" & sdx1
    Set objExcel = Nothing
    Exit Sub
cmd6561245_err:
    MsgBox "Aviso en reporte orden " + error$, 48, "Aviso"
    Exit Sub

End Sub

Function Formato_orden(Num_Campos As Integer, Nombre_Campos() As String) As Boolean

    Dim I As Integer

    With objExcel.ActiveSheet
        'MsgBox Num_Campos
        'Formato de las celdas de los titulos
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 8)).Font.bold = True
        
        For I = 1 To Num_Campos Step 1
            .Cells(3, I) = Nombre_Campos(I)
        Next I

        'hasta aki pa colocar los titulos
        
        'a partir de aki ta claro que es pa darle el ancho a las celdas ;-)
       
        .columns("A").ColumnWidth = 10
        .columns("B").ColumnWidth = 10
        .columns("C").ColumnWidth = 10
        .columns("D").ColumnWidth = 10
        .columns("E").ColumnWidth = 10
        .columns("F").ColumnWidth = 10
        .columns("G").ColumnWidth = 30
        .columns("H").ColumnWidth = 10
        .columns("i").ColumnWidth = 10
        .columns("j").ColumnWidth = 7
        .columns("k").ColumnWidth = 7
        .columns("l").ColumnWidth = 7
    
    End With

End Function

Sub sumar_todo(mytablex As ADODB.Recordset)

    Dim Fecha1 As Date

    Dim Fecha2 As Date

    Dim meses  As Integer

    Dim sdx    As Double

    Dim sdx1   As Double

    Dim buf1   As String

    sdx = 0
    sdx1 = 0
    Do

        If mytablex.EOF Then Exit Do
        If Trim("" & mytablex.Fields("categoria")) = "DIAS" Then
            Fecha1 = Format(Now, "yyyy/mm/dd")
            Fecha2 = Format("" & mytablex.Fields("arribofecha"), "yyyy/mm/dd")
            meses = DateDiff("d", Fecha2, Fecha1)
            buf1 = Format(meses, "###0")

            If Val(buf1) = 0 Then
                buf1 = "1"

            End If

            sdx = Val(buf1) * Val("" & mytablex.Fields("Precio"))
            sdx1 = sdx1 + sdx

        End If

        If Trim("" & mytablex.Fields("categoria")) = "HORAS" Then
            sdx = Val("" & mytablex.Fields("Precio"))
            sdx1 = sdx1 + sdx

        End If

        mytablex.MoveNext
    Loop
    totalreserva = Format(sdx1, "0.00")

End Sub

Private Sub precio_KeyPress(KeyAscii As Integer)

    Dim sdx As Double

    If KeyAscii <> 13 Then Exit Sub
    sdx = Val(noches) * Val(precio)
    total = Format(sdx, "0.00")

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    filtro

End Sub

Function busca_parameca()

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Double

    mytablex.Open "select * from tipo where tipo='" & btipo & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        bserie = Trim("" & mytablex.Fields("serie"))
        sdx = Val("" & mytablex.Fields("numero")) + 1
        bnumero = "" & sdx

    End If

    If Len(Trim(bserie)) = 0 Then
        bserie = "001"

    End If

    mytablex.Close

End Function

Sub graba_factura(buf As String)

    Dim buf1     As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from hotelfactura ", cn, adOpenStatic, adLockOptimistic
    mytablex.AddNew
    mytablex.Fields("idcheckin") = Val(buf)
    mytablex.Fields("fecha") = Trim(arribofecha)
    mytablex.Fields("hora") = Format(Now, "hh:mm:ss")
    mytablex.Fields("tipo") = Trim(btipo)
    mytablex.Fields("serie") = Trim(bserie)
    mytablex.Fields("numero") = Trim(bnumero)
    mytablex.Fields("codigo") = Trim(codigo)
    mytablex.Fields("nombre") = Trim(nombre)
    'mytablex.Fields("direccion") = Trim(direccion)
    mytablex.Fields("operador") = Trim(operador)
    'mytablex.Fields("subtotal") = Val(subtotal)
    'mytablex.Fields("impuesto") = Val(impuesto)
    mytablex.Fields("total") = Val(brecibe)
    mytablex.Fields("moneda") = "S"
    mytablex.Update
    buf1 = Trim("" & mytablex.Fields("idfactura"))
    mytablex.Close

    '-----------------detalle
    mytablex.Open "select * from hoteldetalle where idfactura=" & Val(buf1), cn, adOpenStatic, adLockOptimistic
    mytablex.AddNew
    mytablex.Fields("idfactura") = Trim(buf1)
    mytablex.Fields("idecheckin") = Trim(buf)
    mytablex.Fields("tipo") = "H"
    mytablex.Fields("producto") = Trim(vHabitacion)
    mytablex.Fields("descripcio") = "HABITACION " + Trim(arribofecha)
    mytablex.Fields("unidad") = "UND"
    mytablex.Fields("factor") = 1
    mytablex.Fields("precio") = Val(brecibe)
    mytablex.Fields("cantidad") = 1
    mytablex.Fields("total") = Val(brecibe)
    mytablex.Fields("fecha") = Trim(arribofecha)
    mytablex.Update
    mytablex.Close

    '------------- Forma de Pago
    mytablex.Open "select * from hotelanticipo where idfactura=" & Val(buf1), cn, adOpenStatic, adLockOptimistic
    mytablex.AddNew
    mytablex.Fields("idfactura") = Trim(buf1)
    mytablex.Fields("idcheckin") = Trim(buf)
    mytablex.Fields("fpago") = Trim(bfpago)
    mytablex.Fields("monto") = Val(brecibe)
    mytablex.Fields("banco") = ""
    mytablex.Fields("numero") = ""
    mytablex.Fields("observa") = ""
    mytablex.Fields("fecha") = Trim(arribofecha)
    mytablex.Fields("habitacion") = ""
    mytablex.Update
    mytablex.Close

End Sub

Sub suma_total()

    Dim sdx As Double

    sdx = Val(noches) * Val(precio)
    total = Format(sdx, "0.00")

End Sub

Function actualiza_habitacion(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from habitacion where habitacion='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        mytablex.Fields("checkin") = Val("" & txcheckinx.Fields("checkin"))
        mytablex.Update

    End If

    mytablex.Close

End Function

