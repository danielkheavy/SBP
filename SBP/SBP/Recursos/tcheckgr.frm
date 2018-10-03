VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form tcheckgr 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reservas En Grupo"
   ClientHeight    =   10065
   ClientLeft      =   90
   ClientTop       =   -135
   ClientWidth     =   15090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10065
   ScaleWidth      =   15090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "CheckOut"
      Height          =   5775
      Left            =   13080
      TabIndex        =   44
      Top             =   1800
      Visible         =   0   'False
      Width           =   8175
      Begin VB.TextBox fechasalida 
         Height          =   735
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   48
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox horasalida 
         Height          =   735
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   47
         Top             =   1080
         Width           =   2775
      End
      Begin VB.CommandButton CmdGra 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Grabar"
         DisabledPicture =   "tcheckgr.frx":0000
         Height          =   735
         Left            =   7200
         Picture         =   "tcheckgr.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdCan 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cancelar"
         DisabledPicture =   "tcheckgr.frx":0684
         Height          =   735
         Left            =   7200
         Picture         =   "tcheckgr.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   45
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
         TabIndex        =   50
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
         TabIndex        =   49
         Top             =   1080
         Width           =   2895
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consulta"
      Height          =   9975
      Left            =   12720
      TabIndex        =   39
      Top             =   1200
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
         Top             =   240
         Width           =   2175
      End
      Begin MSDataGridLib.DataGrid dbgrid13 
         Height          =   8895
         Left            =   120
         TabIndex        =   43
         Top             =   840
         Width           =   14655
         _ExtentX        =   25850
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
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame2"
      Height          =   9975
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   14895
      Begin VB.Frame Frame5 
         Caption         =   "Estado"
         Height          =   3495
         Left            =   6240
         TabIndex        =   92
         Top             =   4200
         Width           =   2415
         Begin VB.TextBox estado 
            BackColor       =   &H00E0E0E0&
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
            Left            =   120
            MaxLength       =   7
            TabIndex        =   94
            Top             =   1200
            Width           =   2175
         End
         Begin VB.ComboBox Combo3 
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   93
            Top             =   840
            Width           =   2175
         End
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
         TabIndex        =   89
         Top             =   5640
         Width           =   1935
      End
      Begin VB.ComboBox Combo4 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   88
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox categoria 
         Height          =   375
         Left            =   2280
         MaxLength       =   6
         TabIndex        =   87
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
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
         Left            =   4560
         TabIndex        =   86
         Top             =   9120
         Width           =   1215
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
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   84
         Top             =   9120
         Width           =   2175
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
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   82
         Top             =   8400
         Width           =   6255
      End
      Begin VB.CommandButton Command2 
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
         Left            =   4680
         TabIndex        =   81
         Top             =   7320
         Width           =   1335
      End
      Begin VB.ComboBox disponibles 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   7680
         Width           =   2175
      End
      Begin VB.ListBox reservadas 
         BackColor       =   &H00FFFFFF&
         Height          =   1620
         Left            =   120
         TabIndex        =   76
         Top             =   7680
         Width           =   2175
      End
      Begin VB.TextBox Text4 
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
         TabIndex        =   74
         Top             =   6720
         Width           =   855
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
         TabIndex        =   72
         Top             =   6360
         Width           =   1935
      End
      Begin VB.ComboBox ntipopension 
         Height          =   315
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   71
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
         TabIndex        =   69
         Top             =   6000
         Width           =   1935
      End
      Begin VB.ComboBox ntipotarifa 
         Height          =   315
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   68
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
         Left            =   7920
         MaxLength       =   2
         TabIndex        =   65
         Top             =   3480
         Width           =   735
      End
      Begin VB.ComboBox ntiporeserva 
         Height          =   315
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   64
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
         TabIndex        =   62
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
         TabIndex        =   60
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
         TabIndex        =   58
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox hnombre 
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
         TabIndex        =   53
         Top             =   1680
         Width           =   6375
      End
      Begin VB.TextBox hdireccion 
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
         MaxLength       =   60
         TabIndex        =   52
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
         TabIndex        =   51
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox noches 
         Height          =   375
         Left            =   2280
         MaxLength       =   3
         TabIndex        =   38
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox direccion 
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
         TabIndex        =   31
         Top             =   3120
         Width           =   6375
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
         Top             =   2760
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
         Picture         =   "tcheckgr.frx":0D08
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
         Picture         =   "tcheckgr.frx":15D2
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1470
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
         TabIndex        =   91
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
         TabIndex        =   90
         Top             =   5640
         Width           =   2175
      End
      Begin VB.Label Label26 
         BackColor       =   &H00C0FFC0&
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
         Height          =   375
         Left            =   2400
         TabIndex        =   85
         Top             =   8760
         Width           =   2175
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Habitaciones Seleccionadas"
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
         Left            =   2400
         TabIndex        =   83
         Top             =   8040
         Width           =   2175
      End
      Begin VB.Image Image10 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4560
         Picture         =   "tcheckgr.frx":1E9C
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   375
      End
      Begin VB.Image Image9 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4560
         Picture         =   "tcheckgr.frx":254A
         Stretch         =   -1  'True
         Top             =   960
         Width           =   375
      End
      Begin VB.Image Image8 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5040
         Picture         =   "tcheckgr.frx":2BF8
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Borra"
         Height          =   375
         Left            =   120
         TabIndex        =   80
         Top             =   9480
         Width           =   735
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Habitaciones Disponibles"
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
         Left            =   2400
         TabIndex        =   78
         Top             =   7320
         Width           =   2175
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Habitaciones Reservadas"
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
         TabIndex        =   77
         Top             =   7320
         Width           =   2175
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NroHabitaciones"
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
         TabIndex        =   75
         Top             =   6720
         Width           =   2175
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
         TabIndex        =   73
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
         TabIndex        =   70
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
         TabIndex        =   67
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
         Left            =   6360
         TabIndex        =   66
         Top             =   3480
         Width           =   1575
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
         TabIndex        =   63
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
         TabIndex        =   61
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
         TabIndex        =   59
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cliente (El que se aloja)"
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
         Height          =   375
         Left            =   120
         TabIndex        =   56
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Direccion"
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
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Image Image6 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4200
         Picture         =   "tcheckgr.frx":32A6
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
         TabIndex        =   54
         Top             =   2400
         Width           =   495
      End
      Begin VB.Image Image3 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4200
         Picture         =   "tcheckgr.frx":35B0
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   375
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4200
         Picture         =   "tcheckgr.frx":38BA
         Stretch         =   -1  'True
         Top             =   960
         Width           =   375
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   7920
         Picture         =   "tcheckgr.frx":3BC4
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
         TabIndex        =   37
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Direccion"
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
         TabIndex        =   32
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label28 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pagador(QuienPaga)"
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
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   2760
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
      ScaleWidth      =   15030
      TabIndex        =   2
      Top             =   0
      Width           =   15090
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
         TabIndex        =   34
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
         Picture         =   "tcheckgr.frx":3ECE
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
         Picture         =   "tcheckgr.frx":50E0
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
         Picture         =   "tcheckgr.frx":62F2
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
         Picture         =   "tcheckgr.frx":7504
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
         Picture         =   "tcheckgr.frx":8716
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label xhabitacion 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   12840
         TabIndex        =   36
         Top             =   480
         Width           =   105
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Estado"
         Height          =   375
         Left            =   4080
         TabIndex        =   35
         Top             =   0
         Width           =   2295
      End
      Begin VB.Label xsw 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   15840
         TabIndex        =   33
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
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
         ColumnCount     =   14
         BeginProperty Column00 
            DataField       =   "estado"
            Caption         =   "E"
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
            DataField       =   "CheckIn"
            Caption         =   "CheckIn"
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
         BeginProperty Column03 
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
         BeginProperty Column04 
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
         BeginProperty Column05 
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
         BeginProperty Column06 
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
         BeginProperty Column07 
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
         BeginProperty Column08 
            DataField       =   "Habitacion"
            Caption         =   "Habitaciones"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2865.26
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   2055.118
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
Attribute VB_Name = "tcheckgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim txcheckinx As New ADODB.Recordset
Dim mytablexx As New ADODB.Recordset
Dim mytableyy As New ADODB.Recordset
Dim mytablezz As New ADODB.Recordset

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
   dbGrid1.SetFocus
   Exit Sub
End If
If txcheckinx.RecordCount > 20 Then
   MsgBox "Favor llamar al Proveedor para ampliar Licencia ", 48, "Aviso"
   Exit Sub
End If
inicializa
Frame2.Visible = True
Frame2.Caption = "Nuevo"
cmdGuardar.Enabled = True
habilita 1
refresca_huesped
checkin.Enabled = False
checkin = ""
operador.SetFocus
End Sub

Private Sub bo712_Click()
Dim buf As String

On Error GoTo cmd656_err
If Frame5.Visible = True Then Exit Sub

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

Private Sub CmdGra_Click()
If Not IsDate(fechasalida) Then
   fechasalida.SetFocus
   Exit Sub
End If
If Len(Trim(horasalida)) = 0 Then
   horasalida.SetFocus
   Exit Sub
End If
If "" & txcheckinx.Fields("estado1") <> "2" Then
   MsgBox "Habitacion todavia no cancelado ", 48, "Aviso"
   Exit Sub
End If
txcheckinx.Fields("arribofechaf") = fechasalida
txcheckinx.Fields("arribohoraf") = horasalida
txcheckinx.Fields("estado") = "2"
txcheckinx.Update

MsgBox "Proceso Realizado ", 48, "Aviso"
   ejecuta 1

Frame4.Visible = False
Exit Sub
End Sub

Private Sub cmdGuardar_Click()
Dim found As Integer
If estado = "ENTRADA" Then
If reservadas.ListCount = 0 Then
   MsgBox "No existen habitaciones seleccionados ", 48, "Aviso"
   Exit Sub
End If
End If
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
sumar_precios
End Sub

Private Sub Command4_Click()
filtro
End Sub
Sub filtro()
Dim mytablex As New ADODB.Recordset
Dim cad As String
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
      cad = "select Nombre,Codigo,Direccion,Correo from clientes "
   End If
   If Len(Text1) > 0 Then
      cad = "select Nombre,Codigo,Direccion,Correo from clientes where  " & Combo2 & " like '" & Text1.Text & "%'"
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
      cad = "select Nombre,Codigo,Direccion,Correo from clientes "
   End If
   If Len(Text1) > 0 Then
      cad = "select Nombre,Codigo,Direccion,Correo from clientes where  " & Combo2 & " like '" & Text1.Text & "%'"
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
      cad = "select Nombre,Codigo,Direccion,Correo from clientes "
   End If
   If Len(Text1) > 0 Then
      cad = "select Nombre,Codigo,Direccion,Correo from clientes where  " & Combo2 & " like '" & Text1.Text & "%'"
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
Frame5.Visible = False
End Sub

Private Sub dbgrid13_KeyDown(KeyCode As Integer, Shift As Integer)
Dim mytablex As New ADODB.Recordset
Dim found As Integer
If KeyCode = 27 Then
   Text1.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
If opcion1 = "1" Then
   'habitacion = "" & Trim("" & dbgrid13.columns("habitacion"))
   'habitacion.SetFocus
   'Frame3.Visible = False
End If
If opcion1 = "2" Then
   huesped = Trim("" & dbgrid13.columns("codigo"))
   hnombre = Trim("" & dbgrid13.columns("nombre"))
   hdireccion = Trim("" & dbgrid13.columns("direccion"))
   'correo = Trim("" & dbgrid13.columns("correo"))
   hnombre.SetFocus
   Frame3.Visible = False
   
End If
If opcion1 = "6" Then
   codigo = Trim("" & dbgrid13.columns("codigo"))
   nombre = Trim("" & dbgrid13.columns("nombre"))
   direccion = Trim("" & dbgrid13.columns("direccion"))
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
If ya_existe(Trim("" & disponibles)) <> 1 Then
   reservadas.AddItem Trim(disponibles)
   sumar_precios
End If
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
thotelco.Idcheckin = Trim(buf)
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
Dim buf As String
Dim sdx As Double
Dim found As Integer
If Frame5.Visible = True Then Exit Sub
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
tptovta.habitacion.Caption = Trim("" & txcheckinx.Fields("checkin"))
MsgBox "Proceso Realizado ", 48, "Aviso"
dlo132_Click
Exit Sub
cmd89123_err:
MsgBox "seleccione un dato " + error$, 48, "Aviso"
Exit Sub
End Sub

Private Sub fkichek_Click()
If Frame5.Visible = True Then Exit Sub
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

reporte
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
tnclie.DBPROV = "clientes"
tnclie.fdlo893.Visible = True
tnclie.Show 1

End Sub

Private Sub image2_Click()
consulta_agente
End Sub

Private Sub image3_Click()
consulta_codigo1
End Sub

Private Sub Image4_Click()
End Sub


Private Sub Image6_Click()
consulta_codigo
End Sub


Private Sub Image8_Click()
tnclie.DBPROV = "clientes"
tnclie.fdlo893.Visible = True
tnclie.Show 1

End Sub

Private Sub Image9_Click()
tnclie.DBPROV = "clientes"
tnclie.fdlo893.Visible = True
tnclie.Show 1

End Sub

Private Sub jdfu7834_Click()
If Frame5.Visible = True Then Exit Sub

If Frame2.Visible = True Then Exit Sub
If Frame3.Visible = True Then Exit Sub
If Frame4.Visible = True Then Exit Sub
thotelct.Idcheckin = Trim("" & txcheckinx.Fields("checkin"))
'thotelct.idreserva = Trim("" & txcheckinx.Fields("idreserva"))
'thotelct.habitacion = "" & txcheckinx.Fields("habitacion")
thotelct.Show 1
End Sub

Private Sub Label14_Click()
If reservadas.ListIndex <> -1 Then
   'Eliminamos el elemento que se encuentra seleccionado
   reservadas.RemoveItem reservadas.ListIndex
   sumar_precios
End If

End Sub

Private Sub Label2_Click()
codigo = huesped
nombre = hnombre
direccion = hdireccion
End Sub


Private Sub Label26_Click()
If Frame2.Caption <> "Modifica" Then Exit Sub

Frame5.Caption = "NUEVO"
Frame5.Visible = True
inicializa_huesped
'huesped1.SetFocus
End Sub

Private Sub Label29_Click()
On Error GoTo cmd89123_err
If Frame2.Caption <> "Modifica" Then Exit Sub

'huesped1 = Trim("" & mytablezz.Fields("huesped"))
Frame5.Caption = "MODIFICA"
Frame5.Visible = True
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
Frame5.Caption = "BORRA"
'grabar_huespedes
Exit Sub
cmd289123_err:
MsgBox "Seleccione un Dato", 48, "Aviso"
Exit Sub

End Sub

Private Sub Label31_Click()
Frame5.Visible = True
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

Private Sub noches_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
valida_horas
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
buffer = ""
opcion1 = "1"
ejecuta 1

End Sub
Sub ejecuta(sw As Integer)
Dim cad As String
Dim sdx As Double

      cad = "SELECT * from hotelcheckin  "
      cad = cad & " where estado like '" & estado1 & "'"
   If Len(buffer) > 0 Then
      cad = cad & " and " & Combo1 & " like '" & buffer & "%'"
   End If

   
   If txcheckinx.State = 1 Then txcheckinx.Close
   txcheckinx.Open cad, cn, adOpenStatic, adLockOptimistic
   Set dbGrid1.DataSource = txcheckinx
   'dbGrid1.columns(0).Width = 4000
   'dbGrid1.columns(1).Width = 2000
   If txcheckinx.RecordCount > 0 Then
     'dbGrid1.SetFocus
  End If
  
  

End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
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
refresca_huesped
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
refresca_huesped
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





End Sub
Sub inicializa()
'horas = ""
precio = ""
estado = ""
Label23 = "DIAS"
reservadas.Clear
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
habitacion = ""
nombre = ""
codigo = ""
direccion = ""
huesped = ""
hnombre = ""
hdireccion = ""
'tipoviaje = ""
'procedencia = ""
agente = ""
operador = Trim(gusuario)
'adulto = ""
'nino = ""
noches = "1"
valida_horas
'carga_precio Trim("" & habitacion)
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
direccion = Trim("" & txcheckinx.Fields("direccion"))

huesped = Trim("" & txcheckinx.Fields("huesped"))
hnombre = Trim("" & txcheckinx.Fields("hnombre"))
hdireccion = Trim("" & txcheckinx.Fields("hdireccion"))


'procedencia = Trim("" & txcheckinx.Fields("procedencia"))
operador = Trim("" & txcheckinx.Fields("operador"))
agente = Trim("" & txcheckinx.Fields("agente"))
precio = Trim("" & txcheckinx.Fields("precio"))

'adulto = Trim("" & txcheckinx.Fields("adulto"))
'nino = Trim("" & txcheckinx.Fields("nino"))
habitacion = Trim("" & txcheckinx.Fields("habitacion"))
estado = Trim("" & txcheckinx.Fields("estado"))
carga_reservadas Trim("" & txcheckinx.Fields("checkin"))


End Sub
Sub grabando()
Dim X As Integer
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
txcheckinx.Fields("direccion") = Trim(direccion)

txcheckinx.Fields("huesped") = Trim(huesped)
txcheckinx.Fields("hnombre") = Trim(hnombre)
txcheckinx.Fields("hdireccion") = Trim(hdireccion)

'txcheckinx.Fields("procedencia") = Trim(procedencia)
txcheckinx.Fields("agente") = Trim(agente)
txcheckinx.Fields("operador") = Trim(operador)
txcheckinx.Fields("estado") = Trim(estado)

txcheckinx.Fields("noches") = Val(noches)
'txcheckinx.Fields("adulto") = Val(adulto)
txcheckinx.Fields("precio") = Val(precio)
'txcheckinx.Fields("habitacion") = Trim(habitacion)
habitacion = ""
For X = 0 To reservadas.ListCount - 1
If Len(Trim("" & reservadas.List(X))) > 0 Then
   habitacion = habitacion & Trim("" & reservadas.List(X)) & ","
End If
Next X
txcheckinx.Fields("habitacion") = Trim(habitacion)

End Sub

Private Sub grba1_Click()

End Sub

Function grabar()
Dim found As Integer
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
   txcheckinx.Update
   'actualiza_habitacion "" & txcheckinx.Fields("habitacion")
   If estado = "RESERVA" Then
   graba_reservadas Trim("" & txcheckinx.Fields("checkin"))
   End If
   If estado = "ENTRADA" Then
   graba_reservadas1 Trim("" & txcheckinx.Fields("checkin"))
   End If
   dlo132_Click
   Exit Function
End If
If Frame2.Caption = "Modifica" Then
   'txcheckinx.Fields("checkin") = checkin
   grabando
   txcheckinx.Update
   If estado = "RESERVA" Then
   graba_reservadas Trim("" & txcheckinx.Fields("checkin"))
   End If
   If estado = "ENTRADA" Then
   graba_reservadas1 Trim("" & txcheckinx.Fields("checkin"))
   End If
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
If Len(Trim(codigo)) = 0 Then
   codigo.SetFocus
   Exit Function
End If
If Len(Trim(nombre)) = 0 Then
   nombre.SetFocus
   Exit Function
End If
valida_horas

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
   'habitacion.SetFocus
   Exit Function
End If
If Val(precio) <= 0 Then
   If estado = "ENTRADA" Then
      MsgBox "Debe Seleccionar un Precio", 48, "Aviso"
      precio.SetFocus
      Exit Function
   End If
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
Dim buf As String
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

Sub reporte()
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
Dim buf As String
Dim found As Integer
Dim sdx As Double
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
Dim buf As String
Dim found As Integer
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
Dim buf As String
Dim found As Integer
Dim sdx As Double
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
Dim buf As String
Dim found As Integer
Dim sdx As Double
Dim sdx1 As Double
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
Sub actualiza_habitacion(buf1 As String)
Dim mytablex As New ADODB.Recordset
 mytablex.Open "select * from habitacion where habitacion='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic
 If mytablex.RecordCount > 0 Then
 mytablex.Fields("estado") = "1"
 mytablex.Update
 End If
 mytablex.Close

End Sub
Sub refresca_huesped()
   'If mytablezz.State = 1 Then mytablezz.Close
   'mytablezz.Open "select * from habitacionhuesped where checkin=" & Val(checkin), cn, adOpenStatic, adLockOptimistic
   'Set DataGrid1.DataSource = mytablezz
            

End Sub
Sub carga_precio(buf As String)
Dim mytablex As New ADODB.Recordset

   mytablex.Open "select * from habitacion where habitacion='" & buf & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount > 0 Then
   'precio = Trim("" & mytablex.Fields("precio"))
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
Dim buf As String
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
buf = "SELECT     dbo.hotelcheckinh.checkin, dbo.hotelcheckinh.habitacion"
buf = buf & " FROM         dbo.hotelcheckinh INNER JOIN"
buf = buf & "  dbo.hotelcheckin ON dbo.hotelcheckinh.checkin = dbo.hotelcheckin.checkin"
buf = buf & " and   (dbo.hotelcheckin.arribofecha>='" & Format(arribofecha, "YYYYMMDD") & "'"
buf = buf & " and dbo.hotelcheckin.arribofecha<='" & Format(arribofechaf, "YYYYMMDD") & "') "
buf = buf & " and dbo.hotelcheckinh.habitacion='" & Trim("" & mytablex.Fields("habitacion")) & "'"
mytabley.Open buf, cn, adOpenStatic, adLockOptimistic
If mytabley.RecordCount = 0 Then
   disponibles.AddItem "" & mytablex.Fields("habitacion")
End If
mytabley.Close
mytablex.MoveNext
Loop
mytablex.Close
End Sub
Sub graba_reservadas(buf As String)
Dim X As Integer
Dim mytablex As New ADODB.Recordset
cn.Execute ("delete from hotelcheckinh where checkin=" & Val(buf))
mytablex.Open "select * from hotelcheckinh where checkin=" & Val(buf), cn, adOpenStatic, adLockOptimistic
For X = 0 To reservadas.ListCount - 1
If Len(Trim("" & reservadas.List(X))) > 0 Then
   mytablex.AddNew
   mytablex.Fields("estado") = "R"
   mytablex.Fields("checkin") = Val(buf)
   mytablex.Fields("habitacion") = Trim("" & reservadas.List(X))
   mytablex.Update
End If
Next X
mytablex.Close
End Sub
Sub graba_reservadas1(buf As String)
Dim X As Integer
Dim mytablex As New ADODB.Recordset
cn.Execute ("delete from hotelcheckinh where checkin=" & Val(buf))
mytablex.Open "select * from hotelcheckinh where checkin=" & Val(buf), cn, adOpenStatic, adLockOptimistic
For X = 0 To reservadas.ListCount - 1
If Len(Trim("" & reservadas.List(X))) > 0 Then
   mytablex.AddNew
   mytablex.Fields("estado") = "E"
   mytablex.Fields("checkin") = Val(buf)
   mytablex.Fields("habitacion") = Trim("" & reservadas.List(X))
   mytablex.Update
End If
Next X
mytablex.Close
End Sub

Sub carga_reservadas(buf As String)
Dim mytablex As New ADODB.Recordset
reservadas.Clear
'reservadas.AddItem ""
mytablex.Open "select * from hotelcheckinh where checkin=" & Val(buf), cn, adOpenStatic, adLockOptimistic
   Do
   If mytablex.EOF Then Exit Do
   reservadas.AddItem "" & Trim("" & mytablex.Fields("habitacion"))
   mytablex.MoveNext
   Loop
   mytablex.Close
   'reservadas.ListIndex = 0
End Sub
Function ya_existe(buf As String)
Dim X As Integer

For X = 0 To reservadas.ListCount - 1
If Len(Trim("" & reservadas.List(X))) > 0 Then
   If Trim(buf) = Trim("" & reservadas.List(X)) Then
   ya_existe = 1
   End If
End If
Next X

End Function
Sub sumar_precios()
Dim mytablex As New ADODB.Recordset
Dim sdx As Double
Dim X As Integer
sdx = 0
For X = 0 To reservadas.ListCount - 1
If Len(Trim("" & reservadas.List(X))) > 0 Then
   mytablex.Open "select * from habitacion where habitacion='" & Trim("" & reservadas.List(X)) & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount > 0 Then
   sdx = sdx + Val("" & mytablex.Fields("precio"))
   End If
   mytablex.Close
End If
Next X
precio = Format(sdx, "0.00")
End Sub
Private Function CalculateTime(Time As Double) As String
Dim TimeHour As Double
Dim TimeMin As Double
Dim TimeSec As Byte
Dim CalcTime As String

Dim strtimemin As String
Dim strtimesec As String
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
Dim FirstTime As Double
Dim SecondTime As Double
Dim vTotal As Double
Dim txthours1 As Double
Dim txtmin1 As Double
Dim txtsec1 As Double

Dim txthours2 As Double
Dim txtmin2 As Double
Dim txtsec2 As Double

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


