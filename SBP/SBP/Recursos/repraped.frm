VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form repraped 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte por Producto"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   13125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   13125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox conigv 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   91
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Frame Frame2 
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
      Height          =   9015
      Left            =   -240
      TabIndex        =   77
      Top             =   8400
      Visible         =   0   'False
      Width           =   14775
      Begin VB.ComboBox Combo5 
         Appearance      =   0  'Flat
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
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   615
         Visible         =   0   'False
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
         Left            =   5400
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox Combo4 
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
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox buffer2 
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
         Left            =   2385
         MaxLength       =   10
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   225
         Width           =   2895
      End
      Begin VB.TextBox buffer1 
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
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   600
         Visible         =   0   'False
         Width           =   2895
      End
      Begin MSDataGridLib.DataGrid dbgrid3 
         Height          =   5865
         Left            =   120
         TabIndex        =   83
         Top             =   1095
         Width           =   11565
         _ExtentX        =   20399
         _ExtentY        =   10345
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12648447
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
      Height          =   8175
      Left            =   -240
      TabIndex        =   72
      Top             =   8400
      Visible         =   0   'False
      Width           =   12975
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
         Left            =   6120
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
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
         TabIndex        =   74
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
         Left            =   2430
         MaxLength       =   10
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   6975
         Left            =   120
         TabIndex        =   76
         Top             =   960
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   12303
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
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLICK PARA CANCELAR...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   1200
      TabIndex        =   71
      Top             =   3840
      Visible         =   0   'False
      Width           =   7935
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00808080&
      Caption         =   "CostoDelDiaVenta"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7800
      TabIndex        =   69
      Top             =   2160
      Width           =   2655
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00808080&
      Caption         =   "CostoUltimoTablaProductos"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7800
      TabIndex        =   68
      Top             =   1560
      Value           =   -1  'True
      Width           =   2655
   End
   Begin VB.ComboBox servicio 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   65
      Top             =   1560
      Width           =   1575
   End
   Begin VB.ComboBox PROVEEDOR 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8760
      Style           =   2  'Dropdown List
      TabIndex        =   64
      Top             =   4080
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox horaf 
      BackColor       =   &H00FFFFFF&
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
      Left            =   6120
      MaxLength       =   5
      TabIndex        =   60
      Text            =   "%"
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox horai 
      BackColor       =   &H00FFFFFF&
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
      Left            =   6120
      MaxLength       =   5
      TabIndex        =   59
      Text            =   "%"
      Top             =   5160
      Width           =   1575
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   57
      Top             =   3480
      Width           =   1575
   End
   Begin VB.ComboBox ttflujo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6600
      Style           =   2  'Dropdown List
      TabIndex        =   55
      Top             =   120
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.ComboBox bodega 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   52
      Top             =   5880
      Width           =   1575
   End
   Begin VB.ComboBox local1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   50
      Top             =   360
      Width           =   1575
   End
   Begin VB.ComboBox ccosto 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   48
      Top             =   5880
      Width           =   1575
   End
   Begin VB.ComboBox marca 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   46
      Top             =   5520
      Width           =   1575
   End
   Begin VB.ComboBox Subfamilia 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   44
      Top             =   5160
      Width           =   1575
   End
   Begin VB.ComboBox familia 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   42
      Top             =   4800
      Width           =   1575
   End
   Begin VB.ComboBox estado 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   40
      Top             =   1920
      Width           =   1575
   End
   Begin VB.ComboBox usuario 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox unidad 
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   36
      Text            =   "%"
      Top             =   6240
      Width           =   1575
   End
   Begin VB.ComboBox caja 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   2760
      Width           =   1575
   End
   Begin VB.ComboBox turno 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   3120
      Width           =   1575
   End
   Begin VB.ComboBox orden 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6960
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   7440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox tallas 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   7800
      Width           =   3855
   End
   Begin VB.TextBox descripcio 
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   25
      Text            =   "%"
      Top             =   4440
      Width           =   3855
   End
   Begin VB.TextBox producto 
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   23
      Text            =   "%"
      Top             =   4080
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   7440
      Width           =   3855
   End
   Begin VB.ComboBox moneda 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox nrolineas 
      BackColor       =   &H00FFFFFF&
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
      MaxLength       =   30
      TabIndex        =   8
      Text            =   "45"
      Top             =   7080
      Width           =   1575
   End
   Begin VB.TextBox titulo 
      BackColor       =   &H00FFFFFF&
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
      MaxLength       =   30
      TabIndex        =   7
      Top             =   6720
      Width           =   3855
   End
   Begin VB.TextBox fechaf 
      BackColor       =   &H0080FFFF&
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
      TabIndex        =   6
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox fechai 
      BackColor       =   &H0080FFFF&
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
      TabIndex        =   5
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox serie 
      BackColor       =   &H00FFFFFF&
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
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "%"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ComboBox tipo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox numero 
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   2
      Text            =   "%"
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox codigo 
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   1
      Text            =   "%"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox vendedor 
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   0
      Text            =   "%"
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox TxtCodProv 
      Height          =   405
      Left            =   6150
      TabIndex        =   84
      Text            =   "%"
      Top             =   4050
      Width           =   1500
   End
   Begin VB.ComboBox CboRedondeo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   90
      Top             =   480
      Width           =   1575
   End
   Begin VB.ComboBox quecosto 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   86
      Top             =   780
      Width           =   1575
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Con Igv"
      Height          =   375
      Left            =   3960
      TabIndex        =   92
      Top             =   1080
      Width           =   2140
   End
   Begin VB.Label lblRedondeo 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Redondeo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3960
      TabIndex        =   89
      Top             =   450
      Width           =   2145
   End
   Begin VB.Label lbltipocosto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lbltipocosto"
      Height          =   195
      Left            =   7680
      TabIndex        =   88
      Top             =   1320
      Width           =   795
   End
   Begin VB.Label lblCostoPrecioVta 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Costo/PrecioVta"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3960
      TabIndex        =   87
      Top             =   795
      Width           =   2145
   End
   Begin VB.Label TxtRazonSocialProv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   195
      Left            =   7875
      TabIndex        =   85
      Top             =   4155
      Width           =   120
   End
   Begin VB.Label Label33 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2280
      TabIndex        =   70
      Top             =   8160
      Width           =   2655
   End
   Begin VB.Label donde 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Height          =   195
      Left            =   8280
      TabIndex        =   67
      Top             =   480
      Width           =   45
   End
   Begin VB.Label Label32 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Servicio"
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
      Left            =   3960
      TabIndex        =   66
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label31 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Proveedor"
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
      Left            =   3840
      TabIndex        =   63
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Label30 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HoraFinal HH:MM"
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
      Left            =   3960
      TabIndex        =   62
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Label29 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HoraInicio HH:MM"
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
      Left            =   3960
      TabIndex        =   61
      Top             =   5160
      Width           =   2175
   End
   Begin VB.Label Label28 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo Impresion"
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
      Left            =   3960
      TabIndex        =   58
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label27 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Flujo de Recaudacion"
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
      Left            =   4440
      TabIndex        =   56
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label esactivo 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3960
      TabIndex        =   54
      Top             =   360
      Width           =   105
   End
   Begin VB.Label Label25 
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
      Left            =   3960
      TabIndex        =   53
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Local"
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
      TabIndex        =   51
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label26 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ccosto"
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
      TabIndex        =   49
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label Label24 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Marca"
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
      TabIndex        =   47
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Label23 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SubFamilia"
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
      TabIndex        =   45
      Top             =   5160
      Width           =   2175
   End
   Begin VB.Label Label22 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Familia"
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
      TabIndex        =   43
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label Label18 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estado"
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
      Left            =   3960
      TabIndex        =   41
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cajero"
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
      Left            =   3960
      TabIndex        =   39
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unidad"
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
      TabIndex        =   37
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Caja"
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
      Left            =   3960
      TabIndex        =   35
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label20 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Turno"
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
      Left            =   3960
      TabIndex        =   34
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Orden"
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
      Left            =   6240
      TabIndex        =   31
      Top             =   7440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label xdata 
      AutoSize        =   -1  'True
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
      Height          =   300
      Left            =   8520
      TabIndex        =   29
      Top             =   600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VisualizaTallas"
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
      TabIndex        =   28
      Top             =   7800
      Width           =   2175
   End
   Begin VB.Label Label7 
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
      TabIndex        =   26
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Producto"
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
      TabIndex        =   24
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Agrupacion"
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
      TabIndex        =   22
      Top             =   7440
      Width           =   2175
   End
   Begin VB.Label acu 
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
      Left            =   8280
      TabIndex        =   20
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lineas x Pagina"
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
      Top             =   7080
      Width           =   2175
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Titulo reporte"
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
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label Label17 
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
      Left            =   120
      TabIndex        =   17
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label5 
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
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label1 
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
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Serie"
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
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label10 
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
      TabIndex        =   13
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label11 
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
      TabIndex        =   12
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label21 
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
      TabIndex        =   11
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vendedor"
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
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Menu eki 
      Caption         =   "&Ejecuta"
   End
   Begin VB.Menu lso3232 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "repraped"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xsaa           As String

Dim flujo(11)      As Double

Dim tflujo(11)     As Double

Dim recaudo(11)    As Double

Dim trecaudo(11)   As Double

Dim flujo_ejes(11) As Double

Dim xmeses(32)     As Double

Dim xxmeses(32)    As Double

Private Sub buffer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        Frame1.Visible = False
        codigo.SetFocus
        Exit Sub

    End If

    Command1_Click

End Sub

Private Sub Codigo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        Combo2.Clear
        Combo2.AddItem "nombre"
        Combo2.ListIndex = 0
        Frame1.Visible = True
        buffer = ""
        buffer.SetFocus
        Command1_Click

    End If

End Sub

Private Sub Command1_Click()
    ejecuta 1

End Sub

Sub ejecuta(sw As Integer)

    Dim buf       As String

    Dim rconsulta As New ADODB.Recordset

    If acu = "C" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo from proveedo "
        Else
            buf = "select Nombre,Codigo from proveedo where " & Combo2 & " like '" & buffer & "%'"

        End If

    End If

    If acu = "V" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo from clientes "
        Else
            buf = "select Nombre,Codigo from clientes where " & Combo2 & " like '" & buffer & "%'"

        End If

    End If

    'MsgBox buf

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open buf, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = rconsulta

    If rconsulta.EOF = True And rconsulta.BOF = True Then
        rconsulta.Close
        buffer.SetFocus
        Exit Sub

    End If

    dbGrid1.columns(0).Width = 4000
    dbGrid1.columns(1).Width = 2000

    If sw = 1 Then
        dbGrid1.SetFocus

    End If

End Sub

Private Sub Command2_Click()
    Command2.Visible = False

End Sub

Private Sub Command4_Click()

    'Frame2.Visible = False
End Sub

Private Sub Command5_Click()

End Sub

Sub ejecuta2(sw As Integer)

    Dim tipoclie  As String

    Dim buf       As String

    Dim buf1      As String

    Dim buf2      As String

    Dim xbuf      As String

    Dim xbuf2     As String

    Dim sfound    As String

    Dim rconsulta As New ADODB.Recordset

    buf2 = ""
    tipoclie = "P"

    opcion1 = "2"

    If opcion1 = "2" Then
        If Len(buffer2) = 0 Then
            buf = "select Nombre,Codigo,Codigo1 from proveedo "
        Else
            buf = "select Nombre,Codigo,Codigo1 from proveedo where " & Combo4 & " like '" & buffer2 & "%'"

        End If

    End If

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open buf, cn, adOpenStatic, adLockOptimistic

    If rconsulta.RecordCount = 0 Then
        rconsulta.Close
   
        buffer.SelStart = Len(buffer.Text)
        buffer.SetFocus
        Exit Sub

    End If
   
    Set dbgrid3.DataSource = rconsulta
    'refresca_precios
    sw_consulta = 1
   
    If opcion1 = "444" Or opcion1 = "443" Or opcion1 = "21" Or opcion1 = "1" Or opcion1 = "2" Or opcion1 = "3" Or opcion1 = "4" Or opcion1 = "5" Or opcion1 = "6" Or opcion1 = "7" Then
        dbgrid3.columns(0).Width = 4000
        dbgrid3.columns(1).Width = 2000

    End If

    If opcion1 = "22" Or opcion1 = "23" Or opcion1 = "24" Then
        dbgrid3.columns(0).Width = 1000
        dbgrid3.columns(1).Width = 1500
        dbgrid3.columns(2).Width = 1500
        dbgrid3.columns(3).Width = 1500
        dbgrid3.columns(4).Width = 700

    End If
               
    If opcion1 = "8" Or opcion1 = "888" Or opcion1 = "50" Or opcion1 = "45" Then
        dbgrid3.columns(0).Width = 5000
        dbgrid3.columns(1).Width = 1300
        dbgrid3.columns(2).Width = 1000
        dbgrid3.columns(3).Width = 900
        dbgrid3.columns(4).Width = 500
        dbgrid3.columns(5).Width = 900
        dbgrid3.columns(6).Width = 500
        dbgrid3.columns(7).Width = 800
        dbgrid3.columns(8).Width = 800
        dbgrid3.columns(9).Width = 1700

        'dbgrid3.Columns(10).Width = 500
    End If
             
    If sw = 1 Then

        '         dbgrid3.SetFocus
    End If

End Sub

Private Sub Command3_Click()

    ejecuta2 1

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        codigo = dbGrid1.columns(1)
        Frame1.Visible = False
        codigo.SetFocus

    End If

End Sub

Private Sub DBGrid3_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim found As Integer

    Dim buf   As String

    Dim xtemp As Variant

    If KeyCode = &H70 Then  'f1
        If Len(dbGrid1.columns(0)) > 0 Then
            If opcion1 = "20" Then

                ' consulta_detalles
            End If

            Exit Sub

        End If

    End If

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then

        If opcion1 = "2" Then
            TxtCodProv = Trim(dbgrid3.columns(1))
            TxtRazonSocialProv = Trim(dbgrid3.columns(0))
            Frame2.Visible = False
            Frame2.Enabled = False
            TxtCodProv.SetFocus
  
        End If

    End If

End Sub

Private Sub eki_Click()

    Dim found As Integer

    If Frame1.Visible = True Then Exit Sub

    'If Frame2.Visible = True Then Exit Sub
    If Command2.Visible = True Then Exit Sub

    'costo kenyo 12/04/2017
    If quecosto.ListIndex = 0 Then 'COSTOULTIMO
        lbltipocosto.Caption = "costou"

    End If

    If quecosto.ListIndex = 1 Then 'COSTOPROMEDIO
        lbltipocosto.Caption = "costop"

    End If

    'MsgBox opcion2
    'OPCION2 REPORTE RANKING
    If horai <> "%" And horaf <> "%" Then
        If Not IsDate(horai) Then Exit Sub
        If Not IsDate(horaf) Then Exit Sub

    End If

    If Frame1.Visible = True Then Exit Sub
    
    If ttflujo.Visible = True Then
        If ttflujo = "FLUJO RECAUDACION DIAS" Then '73
            opcion2 = 73
            proceso_impresion73

        End If

        If ttflujo = "FLUJO RECAUDACON HORARIA" Then '74
            opcion2 = 74
            proceso_impresion74

        End If

        If ttflujo = "FLUJO RECAUDACION MENSUAL" Then '75
            opcion2 = 75
            proceso_impresion75

        End If

        If ttflujo = "FLUJO RECAUDACION CASETA" Then '80
            opcion2 = 80
            proceso_impresion80

        End If

        If ttflujo = "FLUJO RECAUDACION FAMILIA" Then '90
            opcion2 = 90
            proceso_impresion90

        End If

        If ttflujo = "FLUJO RECAUDACION PRODUCTO" Then '91
            opcion2 = 91
            proceso_impresion91

        End If

        If ttflujo = "FLUJO RECAUDACION CAJERO" Then '92
            opcion2 = 92
            proceso_impresion92

        End If

        If ttflujo = "FLUJO RECAUDACION TIPO" Then '93
            opcion2 = 93
            proceso_impresion93

        End If

        If ttflujo = "RECAUDACION CAJERO" Then '67
            opcion2 = 67
            proceso_impresion67

        End If

        If ttflujo = "RECAUDACION CAJERO CAJA TIPO SENTIDO" Then '94
            opcion2 = 94
            proceso_impresion94

        End If

        Command2.Visible = False
        Exit Sub

    End If

    If opcion2 = "2" Then
        If Option3.Value = True Then
            Command2.Visible = True
            xsaa = "1"

            If Combo3 = "NORMAL" Then
                imprime_opcion2

            End If

            If Combo3 = "EXCELL" Then
                imprime_opcion2_excel

            End If

            Command2.Visible = False

        End If

        If Option4.Value = True Then
            Command2.Visible = True
            xsaa = "2"

            If Combo3 = "NORMAL" Then
                imprime_opcion2

            End If

            If Combo3 = "EXCELL" Then
                imprime_opcion2_excel

            End If

            Command2.Visible = False

        End If

        'Frame2.Visible = True
   
        Exit Sub

    End If

    Command2.Visible = True

    If opcion2 = "n" Then
   
        '''21/10/2017 Reporte de Productos diarios En Excel
        'imprime_opcionn
        If Combo3 = "NORMAL" Then
            imprime_opcionn

        End If

        If Combo3 = "EXCELL" Then
            imprime_productosdiarios_excel

        End If

        '''21/10/2017 Reporte de Productos diarios En Excel

    End If

    If opcion2 = "1" Then
        If Combo3 = "NORMAL" Then
            imprime_opcion1

        End If
   
        If Combo3 = "EXCELL" Then
            ''''13/09/2017 kenyo Reporte Comisiones Productos en Excel
            'imprime_opcion2_excel
            imprime_opcion2_excelcomision
      
            ''''13/09/2017 kenyo Reporte Comisiones Productos en Excel
   
        End If

    End If

    If opcion2 = "3" Then
        imprime_opcion3

    End If

    Command2.Visible = False

End Sub

Sub imprime_opcion1()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    contpag = 0
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0

    found = sql_producto(mytablex)

    If found = 0 Then
        mytablex.Close
    
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_producto
    cuerpo_programa_producto mytablex
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
    found = valida_wordpad(FileName)
     
    'genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1

End Sub

Private Sub Form_Activate()
    Frame1.Top = 10: Frame1.Left = 10

    '--------------------------------
    Dim mytablex As New ADODB.Recordset

    If donde = "VENDEDOR" Then
        Combo1.ListIndex = 3
   
    End If

    If donde = "Comisiones" Then
        Combo1.ListIndex = 3
   
    End If

    If donde = "CLIENTES" Then
        Combo1.ListIndex = 1
   
    End If

    If donde = "Producto" Then
        Combo1.ListIndex = 2

    End If

    If acu = "Q" Then
        Combo1.ListIndex = 10

    End If

    If Len(esactivo) = 0 Then
        esactivo = "1"
        tipo.Clear
        tipo.AddItem "%"
        mytablex.Open "select * from tipo ", cn, adOpenStatic, adLockOptimistic
        Do

            If mytablex.EOF Then Exit Do
            If acu = "V" Then
                If "" & mytablex.Fields("tipodoc") = "1" Or "" & mytablex.Fields("tipodoc") = "A" Or "" & mytablex.Fields("tipodoc") = "B" Or "" & mytablex.Fields("tipodoc") = "C" Or "" & mytablex.Fields("tipodoc") = "D" Or "" & mytablex.Fields("tipodoc") = "E" Or "" & mytablex.Fields("tipodoc") = "G" Or "" & mytablex.Fields("tipodoc") = "F" Then
                    tipo.AddItem "" & mytablex.Fields("tipo")

                End If

            End If

            If acu = "C" Then
                If "" & mytablex.Fields("tipodoc") = "J" Or "" & mytablex.Fields("tipodoc") = "K" Or "" & mytablex.Fields("tipodoc") = "L" Or "" & mytablex.Fields("tipodoc") = "M" Or "" & mytablex.Fields("tipodoc") = "P" Or "" & mytablex.Fields("tipodoc") = "N" Or "" & mytablex.Fields("tipodoc") = "O" Then
                    tipo.AddItem "" & mytablex.Fields("tipo")

                End If

            End If

            If acu <> "C" And acu <> "V" Then
                tipo.AddItem "" & mytablex.Fields("tipo")

            End If

            mytablex.MoveNext
        Loop
        mytablex.Close
        tipo.ListIndex = 0
        caja.Clear
        caja.AddItem "%"
        mytablex.Open "select * from parameca ", cn, adOpenStatic, adLockOptimistic

        Do

            If mytablex.EOF Then Exit Do
            caja.AddItem "" & mytablex.Fields("caja")
            mytablex.MoveNext
        Loop
        mytablex.Close
        caja.ListIndex = 0

        turno.Clear
        turno.AddItem "%"
        mytablex.Open "select * from turno ", cn, adOpenStatic, adLockOptimistic

        Do

            If mytablex.EOF Then Exit Do
            turno.AddItem "" & mytablex.Fields("turno")
            mytablex.MoveNext
        Loop
        mytablex.Close
        turno.ListIndex = 0

        usuario.Clear
        usuario.AddItem "%"
        mytablex.Open "select * from vendedor ", cn, adOpenStatic, adLockOptimistic

        Do

            If mytablex.EOF Then Exit Do
            usuario.AddItem "" & mytablex.Fields("codigo")
            mytablex.MoveNext
        Loop
        mytablex.Close
        usuario.ListIndex = 0

        familia.Clear
        familia.AddItem "%"
        mytablex.Open "select * from familia ", cn, adOpenStatic, adLockOptimistic
        Do

            If mytablex.EOF Then Exit Do
            familia.AddItem "" & mytablex.Fields("familia")
            mytablex.MoveNext
        Loop
        mytablex.Close
        familia.ListIndex = 0

        subfamilia.Clear
        subfamilia.AddItem "%"
        subfamilia.ListIndex = 0

        marca.Clear
        marca.AddItem "%"
        mytablex.Open "select * from marca ", cn, adOpenStatic, adLockOptimistic

        Do

            If mytablex.EOF Then Exit Do
            marca.AddItem "" & mytablex.Fields("marca")
            mytablex.MoveNext
        Loop
        mytablex.Close
        marca.ListIndex = 0

        ccosto.Clear
        ccosto.AddItem "%"
        mytablex.Open "select * from Vendedor ", cn, adOpenStatic, adLockOptimistic

        Do

            If mytablex.EOF Then Exit Do
            If "" & mytablex.Fields("dueno") = "S" Then
                ccosto.AddItem "" & mytablex.Fields("codigo")

            End If

            mytablex.MoveNext
        Loop
        mytablex.Close
        ccosto.ListIndex = 0
        marca.ListIndex = 0

        local1.Clear
        local1.AddItem "%"
        mytablex.Open "select * from tlocal ", cn, adOpenStatic, adLockOptimistic
        Do

            If mytablex.EOF Then Exit Do
            local1.AddItem "" & mytablex.Fields("CODIGO") & "|" & "" & mytablex.Fields("nombre")
            mytablex.MoveNext
        Loop
        mytablex.Close
        local1.ListIndex = 0

        If local1.ListCount = 2 Then
            local1.ListIndex = 1

        End If

        bodega.Clear
        bodega.AddItem "%"
        mytablex.Open "select * from bodega ", cn, adOpenStatic, adLockOptimistic
        Do

            If mytablex.EOF Then Exit Do
            bodega.AddItem "" & mytablex.Fields("CODIGO") & "|" & "" & mytablex.Fields("nombre")
            mytablex.MoveNext
        Loop
        mytablex.Close
        bodega.ListIndex = 0

        proveedor.Clear
        proveedor.AddItem "%"
        mytablex.Open "select * from proveedo ", cn, adOpenStatic, adLockOptimistic

        Do

            If mytablex.EOF Then Exit Do
            proveedor.AddItem Trim("" & mytablex.Fields("CODIGO")) & "|" & "" & mytablex.Fields("nombre")
            mytablex.MoveNext
        Loop
        mytablex.Close
        proveedor.ListIndex = 0

    End If

End Sub

Private Sub Form_Load()

    Dim mytablex As New ADODB.Recordset

    servicio.Clear
    servicio.AddItem "%"
    'servicio.AddItem "Autoservicio"
    'servicio.AddItem "Comanda"
    'servicio.AddItem "Delivery"

    mytablex.Open "SELECT * FROM servicio ", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        servicio.AddItem "" & mytablex.Fields("servicio") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    mytablex.Close
    servicio.ListIndex = 0

    Combo3.Clear
    Combo3.AddItem "NORMAL"
    Combo3.AddItem "EXCELL"

    '19/06/2017 kenyo REPORTE DEFECTO BLOCK DE NOTA
    Combo3.ListIndex = 0
    '19/06/2017 kenyo  REPORTE DEFECTO BLOCK DE NOTA

    ''08/07/2017 KENYO redondeo reporte productos vendidos
    CboRedondeo.Clear
    CboRedondeo.AddItem "S"
    CboRedondeo.AddItem "N"
    CboRedondeo.ListIndex = 0
    ''08/07/2017 KENYO redondeo reporte productos vendidos

    quecosto.AddItem "COSTOULTIMO"
    quecosto.AddItem "COSTOPROMEDIO"
    'quecosto.AddItem "PRECIOVENTA"
    quecosto.ListIndex = 0

    estado.AddItem "2"
    estado.AddItem "1"
    estado.AddItem "0"
    estado.AddItem "%"
    estado.ListIndex = 0

    ttflujo.Clear
    ttflujo.AddItem "FLUJO RECAUDACION DIAS" '73
    ttflujo.AddItem "FLUJO RECAUDACON HORARIA" '74
    ttflujo.AddItem "FLUJO RECAUDACION MENSUAL" '75
    ttflujo.AddItem "FLUJO RECAUDACION CASETA"  '80
    ttflujo.AddItem "FLUJO RECAUDACION FAMILIA"  '90
    ttflujo.AddItem "FLUJO RECAUDACION PRODUCTO"  '91
    ttflujo.AddItem "FLUJO RECAUDACION CAJERO"  '92
    ttflujo.AddItem "FLUJO RECAUDACION TIPO"  '93
    ttflujo.AddItem "RECAUDACION CAJERO"  '67
    ttflujo.AddItem "RECAUDACION CAJERO CAJA TIPO SENTIDO"  '94
    ttflujo.ListIndex = 0

    orden.AddItem "CANT"
    orden.AddItem "MONTO"
    orden.AddItem "GANANCIA"
    'orden.AddItem "PERDIDA"
    orden.ListIndex = 0

    tallas.AddItem "N"
    tallas.AddItem "S"
    tallas.ListIndex = 0

    Combo1.AddItem "Familia"
    Combo1.AddItem "Codigo"
    Combo1.AddItem "Producto"
    Combo1.AddItem "Vendedor"
    Combo1.AddItem "Zona"
    Combo1.AddItem "Cajero"
    Combo1.AddItem "Caja"
    Combo1.AddItem "Turno"
    Combo1.AddItem "Subfamilia"
    Combo1.AddItem "Marca"
    Combo1.AddItem "Ccosto"
    Combo1.AddItem "Almacen"
    Combo1.AddItem "Comisiones"

    ''' 27/12/2017 Ranking de ventas de productos por Hora
    If opcion2 = "2" Then
        Combo1.AddItem "Hora"

    End If

    ''' 27/12/2017 Ranking de ventas de productos por Hora

    Combo1.ListIndex = 0

    If acu = "Q" Then

    End If

    moneda.AddItem "%"
    moneda.AddItem "S"
    moneda.AddItem "D"
    moneda.ListIndex = 0
    fechai = Format(Now, "dd/mm/yyyy") '"01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    fechaf = Format(Now, "dd/mm/yyyy")

    Frame2.Top = 60
    Frame2.Left = 0

    '' 31/01/2018 Productos unidades vendidos con o sin igv
    conigv.Clear
    conigv.AddItem ""
    conigv.AddItem "S"
    conigv.AddItem "N"
    conigv.ListIndex = 0
    '' 31/01/2018 Productos unidades vendidos con o sin igv

End Sub

Sub consulta_codigo()

    sw_consulta = 0
    Combo5.Clear
    Combo5.AddItem "%"
    Combo5.AddItem "Nombre"
    Combo5.AddItem "Codigo"
    Combo5.ListIndex = 0

    Combo4.Clear
    Combo4.AddItem "Nombre"
    Combo4.AddItem "Codigo"
    Combo4.AddItem "Codigo1"
    Combo4.ListIndex = 0

    Frame2.Visible = True
    Frame2.Enabled = True
    opcion1 = "2"

    If Len(Trim(buffer2)) > 0 Then
        Command3_Click
        Exit Sub

    End If

    Set dbgrid3.DataSource = Nothing

End Sub

Private Sub lso3232_Click()

    If Command2.Visible = True Then
        Command2.Visible = False
        Exit Sub

    End If

    If Frame1.Visible = True Then
        Frame1.Visible = False
        Exit Sub

    End If

    'If Frame2.Visible = True Then
    '   Frame2.Visible = False
    '   Exit Sub
    'End If

    repraped.Hide
    Unload repraped

End Sub

Sub cabecera_producto()

    Dim buf   As String

    Dim I     As Integer

    Dim found As Integer

    'MsgBox "xxx"
    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Fechai : " & fechai, 25, 2, 0)
    found = formateaa("Fechaf : " & fechaf, 25, 2, 0)
    
    buf = String(152, "-")
    found = formateaa(buf, 152, 2, 0)
    buf = "Tipo"
    found = formateaa(buf, 3, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Serie"
    found = formateaa(buf, 3, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Numero"
    found = formateaa(buf, 11, 0, 0)
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
    found = formateaa(buf, 3, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Fact"
    found = formateaa(buf, 4, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Cant"
    found = formateaa(buf, 7, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = "Precio"
    found = formateaa(buf, 7, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = "Total"
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = "M"
    found = formateaa(buf, 1, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Vend"
    found = formateaa(buf, 7, 0, 0)
    buf = "Cajero"
    found = formateaa(buf, 7, 0, 0)
    buf = "Ca"
    found = formateaa(buf, 3, 0, 0)
    buf = "T"
    found = formateaa(buf, 2, 0, 0)
    buf = "E"
    found = formateaa(buf, 2, 0, 0)
    buf = "HORA"
    found = formateaa(buf, 6, 0, 0)
      
    buf = "%Comis"
    found = formateaa(buf, 8, 0, 0)
      
    buf = "TotComi"
    found = formateaa(buf, 8, 2, 0)
      
    buf = String(152, "-")
    found = formateaa(buf, 152, 2, 0)

End Sub

Sub cuerpo_programa_producto(mytablex As ADODB.Recordset)

    Dim sw1      As Integer

    Dim Tmp      As String

    Dim tmp1     As String

    Dim sw       As Integer

    Dim buf      As String

    Dim found    As Integer

    Dim sdx      As Double

    Dim sdxcomi  As Double

    Dim tsdxcomi As Double

    Dim vr

    sdx = 0
    sw = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0

    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    ssuma4 = 0

    suma5 = 0
    suma6 = 0
    suma7 = 0
    suma8 = 0

    ssuma5 = 0
    ssuma6 = 0
    ssuma7 = 0
    ssuma8 = 0
    tmp1 = ""
    sdxcomi = 0
    tsdxcomi = 0
    Do

        If mytablex.EOF Then Exit Do
        vr = DoEvents()

        If Command2.Visible = False Then Exit Do
        If TxtCodProv.Text <> "%" Then
            found = ver_proveedor("" & mytablex.Fields("producto"))

            If found = 0 Then GoTo seguy15

        End If

        If Combo1 = "Producto" Then
            tmp1 = "" & mytablex.Fields("producto")

        End If

        If Combo1 = "Codigo" Then
            tmp1 = "" & mytablex.Fields("codigo")

        End If

        If Combo1 = "Vendedor" Then
            tmp1 = "" & mytablex.Fields("vendedor")

        End If

        If Combo1 = "Comisiones" Then
            tmp1 = "" & mytablex.Fields("vendedor")

        End If

        If Combo1 = "Zona" Then
            tmp1 = "" & mytablex.Fields("Zona")

        End If

        If Combo1 = "Cajero" Then
            tmp1 = "" & mytablex.Fields("Usuario")

        End If

        If Combo1 = "Caja" Then
            tmp1 = "" & mytablex.Fields("Caja")

        End If

        If Combo1 = "Turno" Then
            tmp1 = "" & mytablex.Fields("Turno")

        End If

        If Combo1 = "Familia" Then
   
            tmp1 = "" & mytablex.Fields("familia")

        End If

        If Combo1 = "Subfamilia" Then
   
            tmp1 = "" & mytablex.Fields("subfamilia")

        End If

        If Combo1 = "Ccosto" Then
            tmp1 = "" & mytablex.Fields("ccosto")

        End If

        If Combo1 = "Marca" Then
            tmp1 = "" & mytablex.Fields("marca")

        End If

        If Combo1 = "Almacen" Then
            tmp1 = "" & mytablex.Fields("bodega")

        End If

        If sw = 0 Then
            If Combo1 = "Ccosto" Then
                buf = "" & mytablex.Fields("Ccosto")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Ccosto")

            End If

            If Combo1 = "Almacen" Then
                buf = "" & mytablex.Fields("bodega")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_bodega(buf)
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("bodega")

            End If

            If Combo1 = "Familia" Then
                buf = "" & mytablex.Fields("Familia")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Familia")

            End If

            If Combo1 = "Subfamilia" Then
                buf = "" & mytablex.Fields("Subfamilia")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Subfamilia")

            End If

            If Combo1 = "Marca" Then
                buf = "" & mytablex.Fields("Marca")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Marca")

            End If

            If Combo1 = "Producto" Then
                buf = "" & mytablex.Fields("producto")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = ""
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("producto")

            End If

            If Combo1 = "Codigo" Then
                buf = "" & mytablex.Fields("codigo")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_nombre(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("codigo")

            End If

            If Combo1 = "Vendedor" Then
                buf = "" & mytablex.Fields("vendedor")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Comisiones" Then
                buf = "" & mytablex.Fields("vendedor")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_zona(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("zona")

            End If

            If Combo1 = "Cajero" Then
                buf = "" & mytablex.Fields("usuario")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("usuario")

            End If

            If Combo1 = "Caja" Then
                buf = "" & mytablex.Fields("Caja")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Caja")

            End If

            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("Turno")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Turno")

            End If
   
            sw = 1
            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            suma6 = 0
            suma7 = 0
            suma8 = 0

        End If

        If Tmp <> tmp1 Then
            found = formateaa("", 66, 0, 0)
            sw1 = 0

            If suma2 > 0 Then
                buf = Format(suma1, "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                found = formateaa("", 7, 0, 1)
                found = formateaa("", 1, 0, 0)
                buf = Format(suma2, "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 35, 0, 0)
   
                buf = Format(sdxcomi, "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                found = formateaa("", 1, 2, 0)
   
                nlineas
                sw1 = 1

            End If
   
            If suma4 > 0 Then
                buf = Format(suma3, "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                found = formateaa("", 7, 0, 1)
                found = formateaa("", 1, 0, 0)
                buf = Format(suma4, "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 1, 2, 0)
                nlineas
                sw1 = 1

            End If
   
            If sw1 = 0 Then
                found = formateaa("", 1, 2, 0)
                nlineas

            End If

            If Combo1 = "Ccosto" Then
                buf = "" & mytablex.Fields("Ccosto")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Ccosto")

            End If

            If Combo1 = "Almacen" Then
                buf = "" & mytablex.Fields("bodega")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_bodega(buf)
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("bodega")

            End If

            If Combo1 = "Familia" Then
                buf = "" & mytablex.Fields("Familia")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Familia")

            End If

            If Combo1 = "Subfamilia" Then
                buf = "" & mytablex.Fields("Subfamilia")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Subfamilia")

            End If

            If Combo1 = "Marca" Then
                buf = "" & mytablex.Fields("Marca")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Marca")

            End If

            If Combo1 = "Producto" Then
                buf = "" & mytablex.Fields("producto")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = ""
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("producto")

            End If
   
            If Combo1 = "Codigo" Then
                buf = "" & mytablex.Fields("codigo")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_nombre(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("codigo")

            End If

            If Combo1 = "Vendedor" Then
                buf = "" & mytablex.Fields("vendedor")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Comisiones" Then
                buf = "" & mytablex.Fields("vendedor")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_zona(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("zona")

            End If

            If Combo1 = "Cajero" Then
                buf = "" & mytablex.Fields("usuario")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("usuario")

            End If

            If Combo1 = "Caja" Then
                buf = "" & mytablex.Fields("Caja")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Caja")

            End If

            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("Turno")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Turno")

            End If

            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            suma6 = 0
            suma7 = 0
            suma8 = 0
            sdxcomi = 0

        End If

        buf = "" & mytablex.Fields("tipo")
        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("serie")
        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("numero")
        found = formateaa(buf, 11, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("fecha")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("producto")
        found = formateaa(buf, 7, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("descripcio")
        found = formateaa(buf, 20, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("unidad")
        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("factor")
        found = formateaa(buf, 4, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("cantidad")
        found = formateaa(buf, 7, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("precio")
        found = formateaa(buf, 7, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("Total")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("moneda")
        found = formateaa(buf, 1, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("vendedor")
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("usuario")
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("caja")
        found = formateaa(buf, 2, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("turno")
        found = formateaa(buf, 1, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("estado")
        found = formateaa(buf, 1, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("hora")
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
   
        buf = "" & mytablex.Fields("comision")
        found = formateaa(buf, 8, 0, 0)
        found = formateaa("", 1, 0, 0)
   
        sdx = Val("" & mytablex.Fields("comision")) * Val("" & mytablex.Fields("total")) / 100
   
        buf = "" & sdx
        found = formateaa(buf, 8, 0, 0)
        found = formateaa("", 1, 2, 0)
   
        sdxcomi = sdxcomi + sdx
        tsdxcomi = tsdxcomi + sdx
   
        nlineas

        If tallas = "S" Then
            busca_linea mytablex

        End If

        If "" & mytablex.Fields("estado") = "2" Then
            If "" & mytablex.Fields("moneda") = "S" Then
                suma1 = suma1 + Val("" & mytablex.Fields("cantidad"))
                suma2 = suma2 + Val("" & mytablex.Fields("total"))

            End If

            If "" & mytablex.Fields("moneda") = "D" Then
                suma3 = suma3 + Val("" & mytablex.Fields("cantidad"))
                suma4 = suma4 + Val("" & mytablex.Fields("total"))

            End If

        End If

seguy15:
        mytablex.MoveNext
    Loop
    found = formateaa("", 66, 0, 0)
    sw1 = 0

    If suma2 > 0 Then
        buf = Format(suma1, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        found = formateaa("", 7, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma2, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 35, 0, 0)
   
        buf = Format(sdxcomi, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 2, 0)
   
        nlineas
        sw1 = 1

    End If

    If suma4 > 0 Then
        buf = Format(suma3, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        found = formateaa("", 7, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma4, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 2, 0)
        nlineas
        sw1 = 1

    End If
   
    found = formateaa("Total --> ", 130, 0, 1)
    buf = Format(tsdxcomi, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)
   
    If sw1 = 0 Then
        found = formateaa("", 1, 2, 0)
        nlineas

    End If
   
End Sub

Sub nlineas()
    contlin = contlin + 1

    If contlin > Val(nrolineas) Then
        If opcion2 = "1" Then
            cabecera_producto

        End If

        If opcion2 = "2" Then
            cabecera_producto1

        End If

        If opcion2 = "3" Then
            cabecera_producto1

        End If

    End If

End Sub

Function busca_nombre(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    Dim buf1     As String

    If acu = "V" Then
        buf1 = "clientes"

    End If

    If acu = "C" Then
        buf1 = "proveedo"

    End If

    If acu <> "C" And acu <> "V" Then
        buf1 = "clientes"

    End If

    mytablex.Open "select * from " & buf1 & " where codigo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_nombre = "" & mytablex.Fields("nombre")

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_vendedor(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from vendedor where codigo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_vendedor = "" & mytablex.Fields("nombre")

    End If

    mytablex.Close
 
End Function

Function busca_producto(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select descripcio from producto where [producto]='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_producto = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close
 
End Function

' ''10/10/2017 Reporte de Seguimiento de facturas En Excel
Function busca_familia(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select descripcio from familia where familia='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_familia = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close
 
End Function

Function busca_subfamilia(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select descripcio from subfamilia where subfamilia='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_subfamilia = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close
 
End Function

Function busca_marca(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select descripcio from marca where marca='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_marca = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close
 
End Function

' ''10/10/2017 Reporte de Seguimiento de facturas En Excel

Function busca_zona(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from zona where zona='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_zona = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close
 
End Function

Function busca_bodega(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from bodega where codigo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_bodega = "" & mytablex.Fields("nombre")

    End If

    mytablex.Close
 
End Function

Function sql_producto(mytablex As ADODB.Recordset)

    Dim buf As String

    If Len(fechai) <> 10 Then Exit Function
    If Len(fechaf) <> 10 Then Exit Function
    If Not IsDate(fechai) Then Exit Function
    If Not IsDate(fechaf) Then Exit Function
    buf = "select * from " & xdata & " where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If local1 <> "%" Then
        buf = buf & " and local='" & extra_loquesea(local1) & "'"

    End If

    If tipo <> "%" Then
        buf = buf & " and tipo like '" & tipo & "'"

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

    If usuario <> "%" Then
        buf = buf & " and usuario like '" & usuario & "'"

    End If

    If bodega <> "%" Then
        buf = buf & " and bodega like '" & extra_loquesea(bodega) & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & caja & "'"

    End If

    If servicio <> "%" Then
        'If servicio = "Autoservicio" Then
        buf = buf & " and servicio='" & extra_loquesea(servicio) & "'"
        'End If
        'If servicio = "Comanda" Then
        '   buf = buf & " and servicio='C'"
        'End If
        'If servicio = "Delivery" Then
        '   buf = buf & " and servicio='D'"
        'End If

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & turno & "'"

    End If

    If producto <> "%" Then
        buf = buf & " and [producto] like '" & producto & "'"

    End If

    If descripcio <> "%" Then
        buf = buf & " and [descripcio] like '" & descripcio & "'"

    End If

    If moneda <> "%" Then
        buf = buf & " and moneda like '" & moneda & "'"

    End If

    If ccosto <> "%" Then
        buf = buf & " and ccosto like '" & ccosto & "'"

    End If

    If familia <> "%" Then
        buf = buf & " and familia like '" & familia & "'"

    End If

    If subfamilia <> "%" Then
        buf = buf & " and subfamilia like '" & subfamilia & "'"

    End If

    If marca <> "%" Then
        buf = buf & " and marca like '" & marca & "'"

    End If

    If horai <> "%" And horaf <> "%" Then
        buf = buf & " and HORA BETWEEN '" & horai & "' AND '" & horaf & "'"

    End If

    If unidad <> "%" Then
        buf = buf & " and unidad like '" & unidad & "'"

    End If

    If vendedor <> "%" Then
        buf = buf & " and vendedor like '" & vendedor & "'"

    End If

    If acu = "V" Then
        buf = buf & " and (acu='1' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') "

    End If

    If acu = "C" Then
        buf = buf & " and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') "

    End If

    If acu <> "C" And acu <> "V" Then
        buf = buf & " and acu='" & acu & "'"

    End If

    If estado <> "%" Then
        buf = buf & " and estado='" & estado & "'"

    End If

    If Combo1 = "Producto" Then
        buf = buf & " order by Producto,fecha "

    End If

    If Combo1 = "Codigo" Then
        buf = buf & " order by codigo,fecha "

    End If

    If Combo1 = "Vendedor" Then
        buf = buf & " order by Vendedor,fecha "

    End If

    If Combo1 = "Comisiones" Then
        buf = buf & " order by Vendedor,fecha "

    End If

    If Combo1 = "Zona" Then
        buf = buf & " order by Zona,fecha "

    End If

    If Combo1 = "Cajero" Then
        buf = buf & " order by Usuario,fecha "

    End If

    If Combo1 = "Caja" Then
        buf = buf & " order by Caja,fecha "

    End If

    If Combo1 = "Turno" Then
        buf = buf & " order by Turno,fecha "

    End If

    If Combo1 = "Familia" Then
        buf = buf & " order by Familia,fecha "

    End If

    If Combo1 = "Subfamilia" Then
        buf = buf & " order by Subfamilia,fecha "

    End If

    If Combo1 = "Ccosto" Then
        buf = buf & " order by Ccosto,fecha "

    End If

    If Combo1 = "Almacen" Then
        buf = buf & " order by Bodega,fecha "

    End If

    If Combo1 = "Marca" Then
        buf = buf & " order by Marca,fecha "

    End If

    'MsgBox buf
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    sql_producto = 1

End Function

Sub busca_linea(mytabley As Table)

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    mytablex.Open "select * from linea where linea='" & "" & mytabley.Fields("linea") & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
   
        found = formateaa("" & mytablex.Fields("descripcio"), 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("t1"), 3, 0, 0)
        found = formateaa(":", 1, 0, 0)
        found = formateaa("" & mytabley.Fields("t1"), 3, 0, 0)
        found = formateaa("", 2, 0, 0)
        found = formateaa("" & mytablex.Fields("t2"), 3, 0, 0)
        found = formateaa(":", 1, 0, 0)
        found = formateaa("" & mytabley.Fields("t2"), 3, 0, 0)
        found = formateaa("", 2, 0, 0)
        found = formateaa("" & mytablex.Fields("t3"), 3, 0, 0)
        found = formateaa(":", 1, 0, 0)
        found = formateaa("" & mytabley.Fields("t3"), 3, 0, 0)
        found = formateaa("", 2, 0, 0)
        found = formateaa("" & mytablex.Fields("t4"), 3, 0, 0)
        found = formateaa(":", 1, 0, 0)
        found = formateaa("" & mytabley.Fields("t4"), 3, 0, 0)
        found = formateaa("", 2, 0, 0)
        found = formateaa("" & mytablex.Fields("t5"), 3, 0, 0)
        found = formateaa(":", 1, 0, 0)
        found = formateaa("" & mytabley.Fields("t5"), 3, 0, 0)
        found = formateaa("", 2, 0, 0)
        found = formateaa("" & mytablex.Fields("t6"), 3, 0, 0)
        found = formateaa(":", 1, 0, 0)
        found = formateaa("" & mytabley.Fields("t6"), 3, 0, 0)
        found = formateaa("", 2, 0, 0)
        found = formateaa("" & mytablex.Fields("t7"), 3, 0, 0)
        found = formateaa(":", 1, 0, 0)
        found = formateaa("" & mytabley.Fields("t7"), 3, 0, 0)
        found = formateaa("", 2, 0, 0)
        found = formateaa("" & mytablex.Fields("t8"), 3, 0, 0)
        found = formateaa(":", 1, 0, 0)
        found = formateaa("" & mytabley.Fields("t8"), 3, 0, 0)
        found = formateaa("", 2, 0, 0)
        found = formateaa("" & mytablex.Fields("t9"), 3, 0, 0)
        found = formateaa(":", 1, 0, 0)
        found = formateaa("" & mytabley.Fields("t9"), 3, 0, 0)
        found = formateaa("", 2, 0, 0)
        found = formateaa("" & mytablex.Fields("t10"), 3, 0, 0)
        found = formateaa(":", 1, 0, 0)
        found = formateaa("" & mytabley.Fields("t10"), 3, 0, 0)
        found = formateaa("", 2, 0, 0)
        found = formateaa("" & mytablex.Fields("t11"), 3, 0, 0)
        found = formateaa(":", 1, 0, 0)
        found = formateaa("" & mytabley.Fields("t11"), 3, 0, 0)
        found = formateaa("", 2, 0, 0)
        found = formateaa("" & mytablex.Fields("t12"), 3, 0, 0)
        found = formateaa(":", 1, 0, 0)
        found = formateaa("" & mytabley.Fields("t12"), 3, 0, 0)
        found = formateaa("", 2, 0, 0)
        found = formateaa("" & mytablex.Fields("t13"), 3, 0, 0)
        found = formateaa(":", 1, 0, 0)
        found = formateaa("" & mytabley.Fields("t13"), 3, 0, 0)
        found = formateaa("", 2, 0, 0)
        found = formateaa("" & mytablex.Fields("t14"), 3, 0, 0)
        found = formateaa(":", 1, 0, 0)
        found = formateaa("" & mytabley.Fields("t14"), 3, 0, 0)
        found = formateaa("", 2, 0, 0)
        found = formateaa("" & mytablex.Fields("t15"), 3, 0, 0)
        found = formateaa(":", 1, 0, 0)
        found = formateaa("" & mytabley.Fields("t15"), 3, 0, 0)
        found = formateaa("", 2, 0, 0)
        found = formateaa("" & mytablex.Fields("t16"), 3, 0, 0)
        found = formateaa(":", 1, 0, 0)
        found = formateaa("" & mytabley.Fields("t16"), 3, 0, 0)
        found = formateaa("", 1, 2, 0)
        nlineas
   
    End If

    '------------------------------------- ------------
    mytablex.Close
 
End Sub

Sub imprime_opcion2()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    Dim vr

    contpag = 0
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0

    found = sql_producto1(mytablex)

    If found = 0 Then
        mytablex.Close
        Exit Sub

    End If

    'MsgBox "BC"

    Label33 = "" & mytablex.RecordCount
    vr = DoEvents

    'MsgBox "Presente una tecla", 48, "Aviso"
    'generar_temporal mytablex

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_producto1
    cuerpo_programa_producto1 mytablex
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
     
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1

End Sub

Sub imprime_opcion2_excel()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    contpag = 0
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0

    found = sql_producto1(mytablex)

    If found = 0 Then
        mytablex.Close
    
        Exit Sub

    End If

    cuerpo_programa_producto1xx mytablex
    mytablex.Close

End Sub

Sub imprime_productosdiarios_excel()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    contpag = 0
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0

    found = sql_producton(mytablex)

    If found = 0 Then
        mytablex.Close
    
        Exit Sub

    End If

    cuerpo_programa_productosdiariosexcel mytablex
    mytablex.Close

End Sub

Sub imprime_opcion2_excelcomision()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    contpag = 0
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0

    ''''13/09/2017 kenyo Reporte Comisiones Productos en Excel
    'found = sql_producto1comision(mytablex)
    found = sql_producto(mytablex)
    ''''13/09/2017 kenyo Reporte Comisiones Productos en Excel

    If found = 0 Then
        mytablex.Close
    
        Exit Sub

    End If

    cuerpo_programa_producto1xxcomision mytablex
    mytablex.Close

End Sub

Function sql_producto1(mytablex As ADODB.Recordset)

    Dim buf   As String

    Dim ybuf  As String

    Dim xbuf  As String

    Dim found As Integer

    If Len(fechai) <> 10 Then Exit Function
    If Len(fechaf) <> 10 Then Exit Function
    If Not IsDate(fechai) Then Exit Function
    If Not IsDate(fechaf) Then Exit Function
    If horai <> "%" And horaf <> "%" Then
        found = valida_hora(horai)

        If found = 0 Then Exit Function
        found = valida_hora(horaf)

        If found = 0 Then Exit Function

    End If

    xbuf = ""

    ''' 27/12/2017 Ranking de ventas de productos por Hora
    If Combo1 = "Hora" Then
        xbuf = "left(hora,2) as Hora"

    End If

    ''' 27/12/2017 Ranking de ventas de productos por Hora

    If Combo1 = "Producto" Then
        xbuf = "Producto"

    End If

    If Combo1 = "Codigo" Then
        xbuf = "Codigo"

    End If

    If Combo1 = "Vendedor" Then
        xbuf = "vendedor"

    End If

    If Combo1 = "Comisiones" Then
        xbuf = "vendedor"

    End If

    If Combo1 = "Zona" Then
        xbuf = "Zona"

    End If

    If Combo1 = "Cajero" Then
        xbuf = "Usuario"

    End If

    If Combo1 = "Caja" Then
        xbuf = "Caja"

    End If

    If Combo1 = "Turno" Then
        xbuf = "Turno"

    End If

    If Combo1 = "Familia" Then
        xbuf = "Familia"

    End If

    If Combo1 = "Subfamilia" Then
        xbuf = "Subfamilia"

    End If

    If Combo1 = "Ccosto" Then
        xbuf = "Ccosto"

    End If

    If Combo1 = "Almacen" Then
        xbuf = "bodega"

    End If

    If Combo1 = "Marca" Then
        xbuf = "Marca"

    End If

    buf = "select " & xbuf & ",Producto,Descripcio,moneda as m,sum(cantidad*factor) as xcanti,sum(total) as xtotal,sum(tcosto*cantidad*factor) as xcosto,(sum(total)-sum(tcosto*cantidad*factor)) as xmargen from " & xdata & " where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If local1 <> "%" Then
        buf = buf & " and local='" & extra_loquesea(local1) & "'"

    End If

    If tipo <> "%" Then
        buf = buf & " and tipo like '" & tipo & "'"

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

    If usuario <> "%" Then
        buf = buf & " and usuario like '" & usuario & "'"

    End If

    If bodega <> "%" Then
        buf = buf & " and bodega like '" & extra_loquesea(bodega) & "'"

    End If

    If servicio <> "%" Then
        'If servicio = "Autoservicio" Then
        buf = buf & " and servicio='" & extra_loquesea(servicio) & "'"
        'End If
        'If servicio = "Comanda" Then
        '   buf = buf & " and servicio='C'"
        'End If
        'If servicio = "Delivery" Then
        '   buf = buf & " and servicio='D'"
        'End If

    End If

    If unidad <> "%" Then
        buf = buf & " and unidad like '" & unidad & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & caja & "'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & turno & "'"

    End If

    If producto <> "%" Then
        buf = buf & " and [producto] like '" & producto & "'"

    End If

    If ccosto <> "%" Then
        buf = buf & " and ccosto like '" & ccosto & "'"

    End If

    If horai <> "%" And horaf <> "%" Then
        buf = buf & " and HORA BETWEEN '" & horai & "' AND '" & horaf & "'"

    End If

    If familia <> "%" Then
        buf = buf & " and familia like '" & familia & "'"

    End If

    If subfamilia <> "%" Then
        buf = buf & " and subfamilia like '" & subfamilia & "'"

    End If

    If marca <> "%" Then
        buf = buf & " and marca like '" & marca & "'"

    End If

    If descripcio <> "%" Then
        buf = buf & " and [descripcio] like '" & descripcio & "'"

    End If

    If moneda <> "%" Then
        buf = buf & " and moneda like '" & moneda & "'"

    End If

    If vendedor <> "%" Then
        buf = buf & " and vendedor like '" & vendedor & "'"

    End If

    If estado <> "%" Then
        buf = buf & " and estado='" & estado & "'"

    End If

    If acu = "V" Then
        buf = buf & " and (acu='1' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') "

    End If

    If acu = "C" Then
        buf = buf & " and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') "

    End If

    If acu <> "C" And acu <> "V" Then

        'buf = buf & " and acu='" & acu & "'"
    End If

    If horai <> "%" And horaf <> "%" Then
        buf = buf & " and HORA BETWEEN '" & horai & "' AND '" & horaf & "'"

    End If

    If orden = "CANT" Then
        ybuf = " SUM(cantidad*factor) "

    End If

    If orden = "MONTO" Then
        ybuf = " SUM(total) "

    End If

    If orden = "GANANCIA" Then
        ybuf = " SUM(total)-SUM(cantidad*factor*tcosto) "

    End If

    ''' 27/12/2017 Ranking de ventas de productos por Hora
    If Combo1 = "Hora" Then
        xbuf = "left(hora,2)"

    End If

    ''' 27/12/2017 Ranking de ventas de productos por Hora

    buf = buf & "  group by " & xbuf & ", producto,Descripcio,moneda  order  by " & xbuf & " ," & ybuf & " DESC "
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    sql_producto1 = 1

End Function

Function sql_producto1comision(mytablex As ADODB.Recordset)

    Dim buf   As String

    Dim ybuf  As String

    Dim xbuf  As String

    Dim found As Integer

    If Len(fechai) <> 10 Then Exit Function
    If Len(fechaf) <> 10 Then Exit Function
    If Not IsDate(fechai) Then Exit Function
    If Not IsDate(fechaf) Then Exit Function
    If horai <> "%" And horaf <> "%" Then
        found = valida_hora(horai)

        If found = 0 Then Exit Function
        found = valida_hora(horaf)

        If found = 0 Then Exit Function

    End If

    xbuf = ""

    If Combo1 = "Producto" Then
        xbuf = "Producto"

    End If

    If Combo1 = "Codigo" Then
        xbuf = "Codigo"

    End If

    If Combo1 = "Vendedor" Then
        xbuf = "vendedor"

    End If

    If Combo1 = "Comisiones" Then
        xbuf = "vendedor"

    End If

    If Combo1 = "Zona" Then
        xbuf = "Zona"

    End If

    If Combo1 = "Cajero" Then
        xbuf = "Usuario"

    End If

    If Combo1 = "Caja" Then
        xbuf = "Caja"

    End If

    If Combo1 = "Turno" Then
        xbuf = "Turno"

    End If

    If Combo1 = "Familia" Then
        xbuf = "Familia"

    End If

    If Combo1 = "Subfamilia" Then
        xbuf = "Subfamilia"

    End If

    If Combo1 = "Ccosto" Then
        xbuf = "Ccosto"

    End If

    If Combo1 = "Almacen" Then
        xbuf = "bodega"

    End If

    If Combo1 = "Marca" Then
        xbuf = "Marca"

    End If

    buf = "select " & xbuf & ",tipo,serie,numero,Producto,Descripcio,moneda as m,(cantidad) as xcanti,(total) as xtotal,(tcosto) as xcosto from " & xdata & " where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If local1 <> "%" Then
        buf = buf & " and local='" & extra_loquesea(local1) & "'"

    End If

    If tipo <> "%" Then
        buf = buf & " and tipo like '" & tipo & "'"

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

    If usuario <> "%" Then
        buf = buf & " and usuario like '" & usuario & "'"

    End If

    If bodega <> "%" Then
        buf = buf & " and bodega like '" & extra_loquesea(bodega) & "'"

    End If

    If servicio <> "%" Then
        buf = buf & " and servicio='" & extra_loquesea(servicio) & "'"

    End If

    If unidad <> "%" Then
        buf = buf & " and unidad like '" & unidad & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & caja & "'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & turno & "'"

    End If

    If producto <> "%" Then
        buf = buf & " and [producto] like '" & producto & "'"

    End If

    If ccosto <> "%" Then
        buf = buf & " and ccosto like '" & ccosto & "'"

    End If

    If horai <> "%" And horaf <> "%" Then
        buf = buf & " and HORA BETWEEN '" & horai & "' AND '" & horaf & "'"

    End If

    If familia <> "%" Then
        buf = buf & " and familia like '" & familia & "'"

    End If

    If subfamilia <> "%" Then
        buf = buf & " and subfamilia like '" & subfamilia & "'"

    End If

    If marca <> "%" Then
        buf = buf & " and marca like '" & marca & "'"

    End If

    If descripcio <> "%" Then
        buf = buf & " and [descripcio] like '" & descripcio & "'"

    End If

    If moneda <> "%" Then
        buf = buf & " and moneda like '" & moneda & "'"

    End If

    If vendedor <> "%" Then
        buf = buf & " and vendedor like '" & vendedor & "'"

    End If

    If estado <> "%" Then
        buf = buf & " and estado='" & estado & "'"

    End If

    If acu = "V" Then
        buf = buf & " and (acu='1' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') "

    End If

    If acu = "C" Then
        buf = buf & " and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') "

    End If

    If acu <> "C" And acu <> "V" Then

    End If

    If horai <> "%" And horaf <> "%" Then
        buf = buf & " and HORA BETWEEN '" & horai & "' AND '" & horaf & "'"

    End If

    If orden = "CANT" Then
        ybuf = " SUM(cantidad*factor) "

    End If

    If orden = "MONTO" Then
        ybuf = " SUM(total) "

    End If

    buf = buf & " order  by " & xbuf & " ," & ybuf & " DESC "
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    sql_producto1comision = 1

End Function

Sub cuerpo_programa_producto1(mytablex As ADODB.Recordset)

    Dim vr

    Dim sw1        As Integer

    Dim Tmp        As String

    Dim tmp1       As String

    Dim sw         As Integer

    Dim buf        As String

    Dim found      As Integer

    Dim sdx        As Double

    Dim comision   As Double

    Dim mytabley   As New ADODB.Recordset

    Dim mytablez   As New ADODB.Recordset

    Dim xunidad    As String

    Dim xfactor    As String

    Dim sdxindx    As Double

    Dim xcosto     As Double

    Dim sdx1       As Double

    Dim sdxtmp     As Double

    Dim sstock     As Double

    Dim xcomision1 As Double

    Dim xcomision2 As Double

    On Error GoTo cmd666_err

    sdx1 = 0
    sstock = 0
    sdx = 0
    sw = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0
    suma5 = 0
    suma6 = 0

    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    ssuma4 = 0
    ssuma5 = 0
    ssuma6 = 0

    suma5 = 0
    suma6 = 0
    suma7 = 0
    suma8 = 0

    ssuma5 = 0
    ssuma6 = 0
    ssuma7 = 0
    ssuma8 = 0

    xcomision1 = 0
    xcomision2 = 0
    tmp1 = ""
    sdxindx = 0
    Do

        If mytablex.EOF Then Exit Do
        vr = DoEvents()
        sdxindx = sdxindx + 1
        Command2.Caption = "" & sdxindx

        If Command2.Visible = False Then Exit Do

        'MsgBox "" & mytablex.Fields("FAMILIA")
        If TxtCodProv.Text <> "%" Then
            found = ver_proveedor("" & mytablex.Fields("producto"))

            If found = 0 Then GoTo seguy12

        End If

        '' 27/12/2017 Ranking de ventas de productos por Hora

        If Combo1 = "Hora" Then
            tmp1 = "" & mytablex.Fields("Hora")

        End If

        '' 27/12/2017 Ranking de ventas de productos por Hora

        If Combo1 = "Producto" Then
            tmp1 = "" & mytablex.Fields("Producto")

        End If

        If Combo1 = "Codigo" Then
            tmp1 = "" & mytablex.Fields("codigo")

        End If

        If Combo1 = "Vendedor" Then
            tmp1 = "" & mytablex.Fields("vendedor")

        End If

        If Combo1 = "Comisiones" Then
            tmp1 = "" & mytablex.Fields("vendedor")

        End If

        If Combo1 = "Zona" Then
            tmp1 = "" & mytablex.Fields("Zona")

        End If

        If Combo1 = "Cajero" Then
            tmp1 = "" & mytablex.Fields("Usuario")

        End If

        If Combo1 = "Caja" Then
            tmp1 = "" & mytablex.Fields("Caja")

        End If

        If Combo1 = "Turno" Then
            tmp1 = "" & mytablex.Fields("Turno")

        End If

        If Combo1 = "Familia" Then
            tmp1 = "" & mytablex.Fields("Familia")

        End If

        If Combo1 = "Subfamilia" Then
            tmp1 = "" & mytablex.Fields("Subfamilia")

        End If

        If Combo1 = "Marca" Then
            tmp1 = "" & mytablex.Fields("Marca")

        End If

        If Combo1 = "Ccosto" Then
            tmp1 = "" & mytablex.Fields("Ccosto")

        End If

        If Combo1 = "Almacen" Then
            tmp1 = "" & mytablex.Fields("bodega")

        End If

        If sw = 0 Then

            '' 27/12/2017 Ranking de ventas de productos por Hora
            If Combo1 = "Hora" Then
                buf = "" & mytablex.Fields("Hora")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Hora")

            End If

            '' 27/12/2017 Ranking de ventas de productos por Hora
   
            If Combo1 = "Ccosto" Then
                buf = "" & mytablex.Fields("Ccosto")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Ccosto")

            End If
   
            If Combo1 = "Almacen" Then
                buf = "" & mytablex.Fields("bodega")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_bodega(buf)
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("bodega")

            End If

            If Combo1 = "Familia" Then
                buf = "" & mytablex.Fields("Familia")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Familia")

            End If

            If Combo1 = "Subfamilia" Then
                buf = "" & mytablex.Fields("Subfamilia")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Subfamilia")

            End If

            If Combo1 = "Marca" Then
                buf = "" & mytablex.Fields("Marca")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Marca")

            End If

            If Combo1 = "Producto" Then
                buf = "" & mytablex.Fields("producto")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = ""
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("producto")

            End If

            If Combo1 = "Codigo" Then
                buf = "" & mytablex.Fields("codigo")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_nombre(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("codigo")

            End If

            If Combo1 = "Vendedor" Then
                buf = "" & mytablex.Fields("vendedor")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Comisiones" Then
                buf = "" & mytablex.Fields("vendedor")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_zona(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("zona")

            End If

            If Combo1 = "Cajero" Then
                buf = "" & mytablex.Fields("usuario")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("usuario")

            End If

            If Combo1 = "Caja" Then
                buf = "" & mytablex.Fields("Caja")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Caja")

            End If

            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("Turno")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Turno")

            End If
   
            sw = 1
            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            suma6 = 0
            suma7 = 0
            suma8 = 0
            xcomision1 = 0

        End If

        If Tmp <> tmp1 Then
            '''24/08/2017  Kenyo descripcion larga en reportes ticket
            'found = formateaa("", 80, 0, 0)
            found = formateaa("", 40, 0, 0)
            '''24/08/2017  Kenyo descripcion larga en reportes ticket
   
            sw1 = 0

            If suma2 > 0 Then
                buf = Format(suma1, "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                buf = Format(suma2, "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                buf = Format(suma5, "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                buf = Format(suma6, "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
   
                buf = ""
                found = formateaa(buf, 10, 0, 0)
                found = formateaa("", 1, 0, 0)
   
                '''24/08/2017  Kenyo descripcion larga en reportes ticket
                'buf = "" & xcomision1
                'found = formateaa(buf, 10, 0, 1)
                '''24/08/2017  Kenyo descripcion larga en reportes ticket
   
                found = formateaa("", 1, 2, 0)
   
                nlineas
                sw1 = 1

            End If

            If suma4 > 0 Then
                buf = Format(suma3, "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
   
                buf = Format(suma4, "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 1, 2, 0)
                nlineas
                sw1 = 1

            End If

            If sw1 = 0 Then
                found = formateaa("", 1, 2, 0)
                nlineas

            End If
   
            '' 27/12/2017 Ranking de ventas de productos por Hora
            If Combo1 = "Hora" Then
                buf = "" & mytablex.Fields("Hora")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Hora")

            End If

            '' 27/12/2017 Ranking de ventas de productos por Hora
   
            If Combo1 = "Ccosto" Then
                buf = "" & mytablex.Fields("Ccosto")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Ccosto")

            End If
   
            If Combo1 = "Almacen" Then
                buf = "" & mytablex.Fields("bodega")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_bodega(buf)
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("bodega")

            End If

            If Combo1 = "Familia" Then
                buf = "" & mytablex.Fields("Familia")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Familia")

            End If

            If Combo1 = "Subfamilia" Then
                buf = "" & mytablex.Fields("Subfamilia")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Subfamilia")

            End If

            If Combo1 = "Marca" Then
                buf = "" & mytablex.Fields("Marca")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Marca")

            End If
   
            If Combo1 = "Producto" Then
                buf = "" & mytablex.Fields("producto")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = ""
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("producto")

            End If
   
            If Combo1 = "Codigo" Then
                buf = "" & mytablex.Fields("codigo")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_nombre(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("codigo")

            End If

            If Combo1 = "Vendedor" Then
                buf = "" & mytablex.Fields("vendedor")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Comisiones" Then
                buf = "" & mytablex.Fields("vendedor")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_zona(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("zona")

            End If

            If Combo1 = "Cajero" Then
                buf = "" & mytablex.Fields("usuario")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("usuario")

            End If

            If Combo1 = "Caja" Then
                buf = "" & mytablex.Fields("Caja")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Caja")

            End If

            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("Turno")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Turno")

            End If

            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            suma6 = 0
            suma7 = 0
            suma8 = 0
            xcomision1 = 0

        End If

        buf = "" & mytablex.Fields("producto")
        found = formateaa(buf, 11, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("descripcio")
   
        '''24/08/2017  Kenyo descripcion larga en reportes ticket
        'found = formateaa(buf, 59, 0, 0)
        found = formateaa(buf, 30, 0, 0)
        '''24/08/2017  Kenyo descripcion larga en reportes ticke
   
        found = formateaa("", 1, 0, 0)
        xunidad = "UND"
        xfactor = "1"
        xcosto = 0
        sdxtmp = 0
        procesa_formatop mytablex, xunidad, xfactor, sdxtmp, xcosto

        If Val(xfactor) <= 0 Then
            xfactor = "1"

        End If
   
        '''24/08/2017  Kenyo descripcion larga en reportes ticket
        'buf = xunidad
        'found = formateaa(buf, 5, 0, 0)
        'found = formateaa("", 1, 0, 0)
            
        'buf = xfactor
        'Found = formateaa(buf, 4, 0, 0)
        'found = formateaa("", 1, 0, 0)
  
        buf = calcula_saldo(Val("" & mytablex.Fields("xcanti")), Val(xfactor))
        found = formateaa(buf, 7, 0, 1)
        found = formateaa("", 1, 0, 0)
   
        '''24/08/2017  Kenyo descripcion larga en reportes ticket

        buf = "" & mytablex.Fields("xtotal")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
   
        If xsaa = "2" Then
            sdx = Val("" & mytablex.Fields("xcosto"))
            buf = Format(sdx, "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
   
            sdx1 = Val("" & mytablex.Fields("xmargen"))
            buf = Format(sdx1, "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
      
            buf = "" & mytablex.Fields("m")
            found = formateaa(buf, 1, 0, 0)
            found = formateaa("", 1, 0, 0)
   
            sstock = 0
            mytablez.Open "select saldo from almacen where producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "'", cn, adOpenStatic, adLockOptimistic

            If mytablez.RecordCount > 0 Then
                sstock = Val("" & mytablez.Fields("saldo"))

            End If

            mytablez.Close
   
            buf = "" & sstock
            found = formateaa(buf, 10, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas

        End If

        If xsaa = "1" Then
   
            sdx = xcosto * Val("" & mytablex.Fields("xcanti"))
            buf = Format(sdx, "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
   
            sdx1 = Val("" & mytablex.Fields("xtotal")) - xcosto * Val("" & mytablex.Fields("xcanti"))
            buf = Format(sdx1, "0.00")
            found = formateaa(buf, 10, 0, 1)
   
            '''24/08/2017  Kenyo descripcion larga en reportes ticket
            'found = formateaa("", 1, 0, 0)
            'buf = "" & mytablex.Fields("m")
            'found = formateaa(buf, 1, 0, 0)
            'found = formateaa("", 1, 0, 0)
            'sstock = 0
            'mytablez.Open "select saldo from almacen where producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "'", cn, adOpenStatic, adLockOptimistic
            'If mytablez.RecordCount > 0 Then
            'sstock = Val("" & mytablez.Fields("saldo"))
            'End If
            'mytablez.Close
            'buf = "" & sstock
            'found = formateaa(buf, 10, 0, 0)
            'found = formateaa("", 1, 0, 0)

            '''24/08/2017  Kenyo descripcion larga en reportes ticket
   
            '-------------COMISIONES-------------------------------
            comision = pone_comisiones(mytablex)
            buf = "" & comision
            '''24/08/2017  Kenyo descripcion larga en reportes ticket
            'found = formateaa(buf, 10, 0, 0)
            '''24/08/2017  Kenyo descripcion larga en reportes ticket
   
            found = formateaa("", 1, 2, 0)
            nlineas
            xcomision1 = xcomision1 + comision
   
        End If
   
        If "" & mytablex.Fields("m") = "S" Then
            suma1 = suma1 + Val("" & mytablex.Fields("xcanti"))
            suma2 = suma2 + Val("" & mytablex.Fields("xtotal"))
            suma5 = suma5 + sdx
            suma6 = suma6 + sdx1
      
            ssuma1 = ssuma1 + Val("" & mytablex.Fields("xcanti"))
            ssuma2 = ssuma2 + Val("" & mytablex.Fields("xtotal"))
            ssuma5 = ssuma5 + sdx
            ssuma6 = ssuma6 + sdx1

        End If

        If "" & mytablex.Fields("m") = "D" Then
            suma3 = suma3 + Val("" & mytablex.Fields("xcanti"))
            suma4 = suma4 + Val("" & mytablex.Fields("xtotal"))
      
            ssuma3 = ssuma3 + Val("" & mytablex.Fields("xcanti"))
            ssuma4 = ssuma4 + Val("" & mytablex.Fields("xtotal"))

        End If
   
seguy12:
        mytablex.MoveNext
    Loop

    '''24/08/2017  Kenyo descripcion larga en reportes ticket
    'found = formateaa("", 80, 0, 0)
    found = formateaa("", 40, 0, 0)
    '''24/08/2017  Kenyo descripcion larga en reportes ticket
   
    sw1 = 0

    If suma2 > 0 Then
        buf = Format(suma1, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma2, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma5, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma6, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
   
        buf = ""
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
    
        '''24/08/2017  Kenyo descripcion larga en reportes ticket
        ' buf = "" & xcomision1
        ' found = formateaa(buf, 10, 0, 1)
        '''24/08/2017  Kenyo descripcion larga en reportes ticket

        found = formateaa("", 1, 2, 0)
        nlineas
        sw1 = 1

    End If

    If suma4 > 0 Then
        buf = Format(suma3, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma4, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 2, 0)
        nlineas
        sw1 = 1

    End If

    If sw1 = 0 Then
        found = formateaa("", 1, 2, 0)
        nlineas

    End If
   
    '''24/08/2017  Kenyo descripcion larga en reportes ticket
    ' found = formateaa("Gran Total ", 50, 0, 1)
    found = formateaa("GRAN TOTAL ", 40, 0, 1)
    '''24/08/2017  Kenyo descripcion larga en reportes ticket
   
    If ssuma2 > 0 Then
        buf = Format(ssuma1, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(ssuma2, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(ssuma5, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(ssuma6, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 2, 0)
        nlineas
        sw1 = 1

    End If

    If ssuma4 > 0 Then
        buf = Format(ssuma3, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(ssuma4, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 2, 0)
        nlineas
        sw1 = 1

    End If

    If sw1 = 0 Then
        found = formateaa("", 1, 2, 0)
        nlineas

    End If

    Exit Sub
cmd666_err:
    MsgBox "Aviso en cuerpo programa producto 1 " + error$, 48, "Aviso"
    Exit Sub
   
End Sub

Sub procesa_formatop(mytablex As ADODB.Recordset, _
                     xunidad As String, _
                     xfactor As String, _
                     sdxtmp As Double, _
                     xcosto As Double)

    Dim sdx      As Double

    Dim mytabley As New ADODB.Recordset

    On Error GoTo cmd9090_err

    mytabley.Open "select * from producto where [producto]='" & "" & mytablex.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        xunidad = "" & mytabley.Fields("unidad")
        xfactor = "" & mytabley.Fields("factor")
        sdxtmp = 0

        If Val(xfactor) <= 0 Then
            xfactor = "1"

        End If

        If "" & mytablex.Fields("m") = "S" Then
            If "" & mytabley.Fields("monedac") = "S" Then
                sdxtmp = Val("" & mytabley.Fields(lbltipocosto.Caption))

            End If

            If "" & mytabley.Fields("monedac") = "D" Then
                sdxtmp = Val("" & mytabley.Fields(lbltipocosto.Caption)) * 3

            End If

        End If

        If "" & mytablex.Fields("m") = "D" Then
            If "" & mytabley.Fields("monedac") = "D" Then
                sdxtmp = Val("" & mytabley.Fields(lbltipocosto.Caption))

            End If

            If "" & mytabley.Fields("monedac") = "S" Then
                sdxtmp = Val("" & mytabley.Fields(lbltipocosto.Caption)) / 3

            End If

        End If

        xcosto = Val(Format(sdxtmp, "0.00"))
        sdx = xcosto '/ Val(xfactor)
        xcosto = Val(Format(sdx, "0.00"))
        xcosto = xcosto '* Val("" & mytablex.Fields("xcanti"))

    End If

    mytabley.Close
    Exit Sub
cmd9090_err:
    Exit Sub

End Sub

Sub cuerpo_programa_producto1xx(mytablex As ADODB.Recordset)

    ''''13/09/2017 kenyo Mejora Reportes Familias Producto
    'Ajustar h
    ''''13/09/2017 kenyo Mejora Reportes Familias Producto
    Dim vr

    Dim sw1          As Integer

    Dim Tmp          As String

    Dim tmp1         As String

    Dim sw           As Integer

    Dim buf          As String

    Dim found        As Integer

    Dim sdx          As Double

    Dim mytabley     As New ADODB.Recordset

    Dim xunidad      As String

    Dim xfactor      As String

    Dim xcosto       As String

    Dim sdx1         As Double

    Dim sdxtmp       As Double

    Dim v            As Long

    Dim h            As Integer

    Dim vprecios(10) As String

    Dim Heading(11)  As String

    h = 1
    sdx1 = 0
    sdx = 0
    sw = 0

    ''''13/09/2017 kenyo Mejora Reportes Familias Producto
    v = 4
    ''''13/09/2017 kenyo Mejora Reportes Familias Producto

    h = 1
    suma1 = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0
    suma5 = 0
    suma6 = 0

    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    ssuma4 = 0
    ssuma5 = 0
    ssuma6 = 0

    suma5 = 0
    suma6 = 0
    suma7 = 0
    suma8 = 0

    ssuma5 = 0
    ssuma6 = 0
    ssuma7 = 0
    ssuma8 = 0

    ''''13/09/2017 kenyo Mejora Reportes Familias Producto
    Heading(1) = Combo1
    Heading(2) = "Cod.Producto"
    Heading(3) = "Descripcion"
    Heading(4) = "Und"
    Heading(5) = "Factor"
    Heading(6) = "cantidad"
    Heading(7) = "Total"
    Heading(8) = "Costo"
    Heading(9) = "Ganancia"
    Heading(10) = "Moneda"
    
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_Excel(10, Heading())
    
    ''''09/10/2017 kenyo Testing Reportes

    If donde = "VENDEDOR" Then
        objExcel.ActiveSheet.Cells(1, 3) = "                                       VENTAS POR VENDEDOR                          "
    ElseIf donde = "CLIENTES" Then
        objExcel.ActiveSheet.Cells(1, 3) = "                                        VENTAS POR CLIENTE                      "
    Else
        objExcel.ActiveSheet.Cells(1, 3) = "                                  RANKING DE PRODUCTOS VENDIDOS                       "

    End If
    
    objExcel.ActiveSheet.Cells(1, 3).Font.bold = True
    objExcel.ActiveSheet.Cells(1, 3).Font.Size = 14
    objExcel.ActiveSheet.Cells(1, 3).Font.color = RGB(0, 112, 184)
    
    objExcel.ActiveSheet.Cells(2, 1) = "FECHA INICIO  " + fechai
    objExcel.ActiveSheet.Cells(2, 3) = "FECHA FIN  " + fechaf
    ''''09/10/2017 kenyo Testing Reportes
     
    tmp1 = ""

    Do

        If mytablex.EOF Then Exit Do
        vr = DoEvents()

        If Command2.Visible = False Then Exit Do
        If TxtCodProv.Text <> "%" Then
            found = ver_proveedor("" & mytablex.Fields("producto"))

            If found = 0 Then GoTo seguy13

        End If

        ''' 27/12/2017 Ranking de ventas de productos por Hora
        If Combo1 = "Hora" Then
            tmp1 = "" & mytablex.Fields("hora")

        End If

        ''' 27/12/2017 Ranking de ventas de productos por Hora

        If Combo1 = "Producto" Then
            tmp1 = "" & mytablex.Fields("Producto")

        End If

        If Combo1 = "Codigo" Then
            tmp1 = "" & mytablex.Fields("codigo")

        End If

        If Combo1 = "Vendedor" Then
            tmp1 = "" & mytablex.Fields("vendedor")

        End If

        If Combo1 = "Comisiones" Then
            tmp1 = "" & mytablex.Fields("vendedor")

        End If

        If Combo1 = "Zona" Then
            tmp1 = "" & mytablex.Fields("Zona")

        End If

        If Combo1 = "Cajero" Then
            tmp1 = "" & mytablex.Fields("Usuario")

        End If

        If Combo1 = "Caja" Then
            tmp1 = "" & mytablex.Fields("Caja")

        End If

        If Combo1 = "Turno" Then
            tmp1 = "" & mytablex.Fields("Turno")

        End If

        If Combo1 = "Familia" Then
            tmp1 = "" & mytablex.Fields("Familia")

        End If

        If Combo1 = "Subfamilia" Then
            tmp1 = "" & mytablex.Fields("Subfamilia")

        End If

        If Combo1 = "Marca" Then
            tmp1 = "" & mytablex.Fields("Marca")

        End If

        If Combo1 = "Ccosto" Then
            tmp1 = "" & mytablex.Fields("Ccosto")

        End If

        If Combo1 = "Almacen" Then
            tmp1 = "" & mytablex.Fields("bodega")

        End If

        If sw = 0 Then
    
            ''' 27/12/2017 Ranking de ventas de productos por Hora
            If Combo1 = "Hora" Then
                buf = "" & mytablex.Fields("Hora")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf
                v = v + 1
                Tmp = "" & mytablex.Fields("Hora")

            End If

            ''' 27/12/2017 Ranking de ventas de productos por Hora
   
            If Combo1 = "Ccosto" Then
                buf = "" & mytablex.Fields("Ccosto")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf
                v = v + 1
                Tmp = "" & mytablex.Fields("Ccosto")

            End If
   
            If Combo1 = "Almacen" Then
                buf = "" & mytablex.Fields("bodega")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf & " " & busca_bodega(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("bodega")

            End If

            If Combo1 = "Familia" Then
                buf = "" & mytablex.Fields("Familia")
                v = v + 1
                objExcel.ActiveSheet.Cells(v - 1, h) = " "
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_familia(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("Familia")

            End If
   
            If Combo1 = "Subfamilia" Then
                buf = "" & mytablex.Fields("Subfamilia")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf
    
                v = v + 1
                Tmp = "" & mytablex.Fields("Subfamilia")

            End If
   
            If Combo1 = "Marca" Then
                buf = "" & mytablex.Fields("Marca")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf
                v = v + 1
                Tmp = "" & mytablex.Fields("Marca")

            End If
   
            If Combo1 = "Producto" Then
                buf = "" & mytablex.Fields("producto")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf
                v = v + 1
                Tmp = "" & mytablex.Fields("producto")

            End If

            If Combo1 = "Codigo" Then
                buf = "" & mytablex.Fields("codigo")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf
                objExcel.ActiveSheet.Cells(v, h + 1) = "" & busca_nombre(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("codigo")

            End If
   
            If Combo1 = "Vendedor" Then
                buf = "" & mytablex.Fields("vendedor")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf & ""
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_vendedor(buf)

                v = v + 1
                Tmp = "" & mytablex.Fields("vendedor")

            End If
   
            If Combo1 = "Comisiones" Then
                buf = "" & mytablex.Fields("vendedor")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf & ""
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_vendedor(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("vendedor")

            End If
   
            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf & " " & busca_zona(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("zona")

            End If
   
            If Combo1 = "Cajero" Then
                buf = "" & mytablex.Fields("usuario")
                objExcel.ActiveSheet.Cells(v, h) = "'" & busca_vendedor(buf)
                objExcel.ActiveSheet.Cells(v, h + 1) = "" & busca_vendedor(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("usuario")

            End If
   
            If Combo1 = "Caja" Then
                buf = "" & mytablex.Fields("Caja")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf
                v = v + 1
                nlineas
                Tmp = "" & mytablex.Fields("Caja")

            End If
   
            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("Turno")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf
                v = v + 1
                Tmp = "" & mytablex.Fields("Turno")

            End If
   
            objExcel.ActiveSheet.Cells(v - 1, h).Font.bold = True
            objExcel.ActiveSheet.Cells(v - 1, h).Font.color = RGB(62, 95, 138)
   
            objExcel.ActiveSheet.Cells(v - 1, h + 1).Font.bold = True
            objExcel.ActiveSheet.Cells(v - 1, h + 1).Font.color = RGB(62, 95, 138)
   
            sw = 1
            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            suma6 = 0
            suma7 = 0
            suma8 = 0

        End If

        If Tmp <> tmp1 Then

            buf = Format(suma1, "0.00")
            objExcel.ActiveSheet.Cells(v, h + 5) = "" & buf
            objExcel.ActiveSheet.Cells(v, h + 5).Font.bold = True
            buf = Format(suma2, "0.00")
            objExcel.ActiveSheet.Cells(v, h + 6) = "" & buf
            objExcel.ActiveSheet.Cells(v, h + 6).Font.bold = True
            buf = Format(suma5, "0.00")
            objExcel.ActiveSheet.Cells(v, h + 7) = "" & buf
            objExcel.ActiveSheet.Cells(v, h + 7).Font.bold = True
            buf = Format(suma6, "0.00")
            objExcel.ActiveSheet.Cells(v, h + 8) = "" & buf
            objExcel.ActiveSheet.Cells(v, h + 8).Font.bold = True
            buf = Format(suma3, "0.00")
   
            ''''13/09/2017 kenyo Mejora Reportes Familias Producto
            'objExcel.ActiveSheet.Cells(v, h + 9) = ""
            'buf = Format(suma4, "0.00")
            'objExcel.ActiveSheet.Cells(v, h + 10) = ""
   
            objExcel.ActiveSheet.Cells(v, h + 9) = ""
            buf = Format(suma4, "0.00")
            objExcel.ActiveSheet.Cells(v, h + 10) = ""
            ''''13/09/2017 kenyo Mejora Reportes Familias Producto
   
            v = v + 1
   
            ''' 27/12/2017 Ranking de ventas de productos por Hora
            If Combo1 = "Hora" Then
                buf = "" & mytablex.Fields("Hora")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf
                v = v + 1
                Tmp = "" & mytablex.Fields("Hora")

            End If

            ''' 27/12/2017 Ranking de ventas de productos por Hora
   
            If Combo1 = "Ccosto" Then
                buf = "" & mytablex.Fields("Ccosto")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf
                v = v + 1
                Tmp = "" & mytablex.Fields("Ccosto")

            End If
   
            If Combo1 = "Almacen" Then
                buf = "" & mytablex.Fields("bodega")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf & "" & busca_bodega(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("bodega")

            End If

            If Combo1 = "Familia" Then
                buf = "" & mytablex.Fields("Familia")
                v = v + 1
                objExcel.ActiveSheet.Cells(v - 1, h) = ""
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_familia(buf)
                v = v + 1
                nlineas
                Tmp = "" & mytablex.Fields("Familia")

            End If
   
            If Combo1 = "Subfamilia" Then
                buf = "" & mytablex.Fields("Subfamilia")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf
                v = v + 1
                Tmp = "" & mytablex.Fields("Subfamilia")

            End If
   
            If Combo1 = "Marca" Then
                buf = "" & mytablex.Fields("Marca")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf
                v = v + 1
                Tmp = "" & mytablex.Fields("Marca")

            End If
   
            If Combo1 = "Producto" Then
                buf = "" & mytablex.Fields("producto")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf
                objExcel.ActiveSheet.Cells(v, h + 1) = "" & busca_producto(buf)
   
                v = v + 1
                Tmp = "" & mytablex.Fields("producto")

            End If
   
            If Combo1 = "Codigo" Then
                buf = "" & mytablex.Fields("codigo")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & " "
                objExcel.ActiveSheet.Cells(v, h + 1) = "" & busca_nombre(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("codigo")

            End If
   
            If Combo1 = "Vendedor" Then
                buf = "" & mytablex.Fields("vendedor")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf
                objExcel.ActiveSheet.Cells(v, h + 1) = "" & busca_vendedor(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("vendedor")

            End If
   
            If Combo1 = "Comisiones" Then
                buf = "" & mytablex.Fields("vendedor")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf & " " & busca_vendedor(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("vendedor")

            End If
   
            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf & " " & busca_zona(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("zona")

            End If
   
            If Combo1 = "Cajero" Then
                buf = "" & mytablex.Fields("usuario")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf & " " & busca_vendedor(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("usuario")

            End If
   
            If Combo1 = "Caja" Then
                buf = "" & mytablex.Fields("Caja")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf
                v = v + 1
                Tmp = "" & mytablex.Fields("Caja")

            End If
   
            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("Turno")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf
                v = v + 1
                Tmp = "" & mytablex.Fields("Turno")

            End If
   
            objExcel.ActiveSheet.Cells(v - 2, h) = " "
            objExcel.ActiveSheet.Cells(v - 1, h).Font.bold = True
            objExcel.ActiveSheet.Cells(v - 1, h).Font.color = RGB(62, 95, 138)
    
            objExcel.ActiveSheet.Cells(v - 1, h + 1).Font.bold = True
            objExcel.ActiveSheet.Cells(v - 1, h + 1).Font.color = RGB(62, 95, 138)
   
            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            suma6 = 0
            suma7 = 0
            suma8 = 0

        End If

        ''''13/09/2017 kenyo Mejora Reportes Familias Producto
   
        ''' 27/12/2017 Ranking de ventas de productos por Hora
        If Combo1 = "Hora" Then
            buf = "'" & mytablex.Fields("producto")
            objExcel.ActiveSheet.Cells(v, h + 1) = "" & buf
            buf = "" & mytablex.Fields("descripcio")
            objExcel.ActiveSheet.Cells(v, h + 2) = "" & buf
            buf = "" & mytablex.Fields("Hora")
            objExcel.ActiveSheet.Cells(v, h) = "'" & buf

        End If

        ''' 27/12/2017 Ranking de ventas de productos por Hora
   
        If Combo1 = "Familia" Then
            buf = "'" & mytablex.Fields("producto")
            objExcel.ActiveSheet.Cells(v, h + 1) = "" & buf
            buf = "" & mytablex.Fields("descripcio")
            objExcel.ActiveSheet.Cells(v, h + 2) = "" & buf
            buf = "" & mytablex.Fields("familia")
            objExcel.ActiveSheet.Cells(v, h) = "'" & buf
        
        ElseIf Combo1 = "Subfamilia" Then
            buf = "'" & mytablex.Fields("producto")
            objExcel.ActiveSheet.Cells(v, h + 1) = "" & buf
            buf = "" & mytablex.Fields("descripcio")
            objExcel.ActiveSheet.Cells(v, h + 2) = "" & buf
            buf = "" & mytablex.Fields("subfamilia")
            objExcel.ActiveSheet.Cells(v, h) = "'" & buf
         
        ElseIf Combo1 = "Producto" Then
            buf = "'" & mytablex.Fields("producto")
            objExcel.ActiveSheet.Cells(v, h + 1) = "" & buf
            buf = "" & mytablex.Fields("descripcio")
            objExcel.ActiveSheet.Cells(v, h + 2) = "" & buf
            buf = "" & mytablex.Fields("Producto")
            objExcel.ActiveSheet.Cells(v, h) = "'" & buf
        
        ElseIf Combo1 = "Marca" Then
            buf = "'" & mytablex.Fields("producto")
            objExcel.ActiveSheet.Cells(v, h + 1) = "" & buf
            buf = "" & mytablex.Fields("descripcio")
            objExcel.ActiveSheet.Cells(v, h + 2) = "" & buf
            buf = "" & mytablex.Fields("marca")
            objExcel.ActiveSheet.Cells(v, h) = "" & buf
        
        ElseIf Combo1 = "Vendedor" Then
            buf = "'" & mytablex.Fields("producto")
            objExcel.ActiveSheet.Cells(v, h + 1) = "" & buf
            buf = "" & mytablex.Fields("descripcio")
            objExcel.ActiveSheet.Cells(v, h + 2) = "" & buf
            buf = "" & mytablex.Fields("vendedor")
            objExcel.ActiveSheet.Cells(v, h) = "" & buf
        ElseIf Combo1 = "Codigo" Then
            buf = "'" & mytablex.Fields("producto")
            objExcel.ActiveSheet.Cells(v, h + 1) = "" & buf
            buf = "" & mytablex.Fields("descripcio")
            objExcel.ActiveSheet.Cells(v, h + 2) = "" & buf
            buf = "'" & mytablex.Fields("codigo")
            objExcel.ActiveSheet.Cells(v, h) = "" & buf
        Else
            buf = "'" & mytablex.Fields("producto")
            objExcel.ActiveSheet.Cells(v, h + 1) = "" & buf
            buf = "" & mytablex.Fields("descripcio")
            objExcel.ActiveSheet.Cells(v, h + 2) = "" & buf

        End If

        ''''13/09/2017 kenyo Mejora Reportes Familias Producto
        xunidad = "UND"
        xfactor = "1"
        xcosto = "0"
        sdxtmp = 0
        mytabley.Open "select * from producto where [producto]='" & "" & mytablex.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic

        If mytabley.RecordCount > 0 Then
            xunidad = "" & mytabley.Fields("unidad")
            xfactor = "" & mytabley.Fields("factor")
            sdxtmp = 0

            If Val(xfactor) <= 0 Then
                xfactor = "1"

            End If

            If "" & mytablex.Fields("m") = "S" Then
                If "" & mytabley.Fields("monedac") = "S" Then
                    sdxtmp = Val("" & mytabley.Fields(lbltipocosto.Caption))

                End If

                If "" & mytabley.Fields("monedac") = "D" Then
                    sdxtmp = Val("" & mytabley.Fields(lbltipocosto.Caption)) * 3

                End If

            End If

            If "" & mytablex.Fields("m") = "D" Then
                If "" & mytabley.Fields("monedac") = "D" Then
                    sdxtmp = Val("" & mytabley.Fields(lbltipocosto.Caption))

                End If

                If "" & mytabley.Fields("monedac") = "S" Then
                    sdxtmp = Val("" & mytabley.Fields(lbltipocosto.Caption)) / 3

                End If

            End If
       
            ''''09/10/2017 kenyo Familias productos Correcion Costos
            'xcosto = Format(sdxtmp, "0.00000")
            'sdx = Val(xcosto) / Val(xfactor)
            'xcosto = Format(sdx, "0.00000")
            
            xcosto = Format(sdxtmp, "0.0000000")
            sdx = Val(xcosto) / Val(xfactor)
            xcosto = Format(sdx, "0.0000000")
            ''''09/10/2017 kenyo Familias productos Correcion Costos
       
        End If

        mytabley.Close

        If Val(xfactor) <= 0 Then
            xfactor = "1"

        End If

        buf = xunidad
        objExcel.ActiveSheet.Cells(v, h + 3) = "" & buf
        buf = xfactor
        objExcel.ActiveSheet.Cells(v, h + 4) = "" & buf
        buf = calcula_saldo(Val("" & mytablex.Fields("xcanti")), Val(xfactor))
        'buf = "" & mytablex.Fields("xcanti")
        objExcel.ActiveSheet.Cells(v, h + 5) = "" & buf
   
        '' 31/01/2018 Productos unidades vendidos con o sin igv
        'buf = "" & mytablex.Fields("xtotal")
        If conigv = "" Or conigv = "S" Then
            buf = "" & mytablex.Fields("xtotal")

        End If
    
        If conigv = "N" Then

            Dim valorigv As String

            valorigv = 0
            valorigv = obtiene_igv_producto(mytablex.Fields("producto"))
              
            If Val("" & valorigv) > 0 Then
                buf = mytablex.Fields("xtotal") / (1 + (Val("" & valorigv) / 100))
                buf = Val(Format(buf, "0.00000"))
            Else
                buf = "" & mytablex.Fields("xtotal")

            End If

        End If

        '' 31/01/2018 Productos unidades vendidos con o sin igv
   
        ''08/07/2017 KENYO redondeo reporte productos vendidos
        If CboRedondeo = "S" Then
            buf = Format(buf, "0.00")
        Else
            buf = Format(buf, "0.00000")

        End If

        ''08/07/2017 KENYO redondeo reporte productos vendidos
  
        objExcel.ActiveSheet.Cells(v, h + 6) = "" & buf
   
        '' 31/01/2018 Productos unidades vendidos con o sin igv
        'buf = "" & mytablex.Fields("xtotal")
        If conigv = "" Or conigv = "S" Then
            xcosto = "" & xcosto

        End If
    
        If conigv = "N" Then
            valorigv = obtiene_igv_producto(mytablex.Fields("producto"))
              
            If Val("" & valorigv) > 0 Then
                xcosto = xcosto / (1 + (Val("" & valorigv) / 100))
                xcosto = Val(Format(xcosto, "0.00000"))
            Else
                xcosto = "" & xcosto

            End If

        End If

        '' 31/01/2018 Productos unidades vendidos con o sin igv
   
        ''''09/10/2017 kenyo Familias productos Correcion Costos
        'sdx = Val(xcosto) * Val("" & mytablex.Fields("xcanti"))
        sdx = (Val(xcosto) * Val("" & mytablex.Fields("xcanti"))) * Val(xfactor)
        ''''09/10/2017 kenyo Familias productos Correcion Costos
   
        ''08/07/2017 KENYO redondeo reporte productos vendidos
        If CboRedondeo = "S" Then
            buf = Format(sdx, "0.00")
        Else
            buf = Format(sdx, "0.00000")

        End If

        ''08/07/2017 KENYO redondeo reporte productos vendidos
        
        objExcel.ActiveSheet.Cells(v, h + 7) = "" & buf
      
        ''''09/10/2017 kenyo Familias productos Correcion Costos
        'sdx1 = Val("" & mytablex.Fields("xtotal")) - (Val(xcosto) * Val(xfactor)) * Val("" & mytablex.Fields("xcanti"))
   
        '' 31/01/2018 Productos unidades vendidos con o sin igv
        'sdx1 = Val("" & mytablex.Fields("xtotal")) - (Val(xcosto) * Val(xfactor)) * Val("" & mytablex.Fields("xcanti"))
  
        If conigv = "" Or conigv = "S" Then
            sdx1 = Val("" & mytablex.Fields("xtotal")) - (Val(xcosto) * Val(xfactor)) * Val("" & mytablex.Fields("xcanti"))

        End If
    
        If conigv = "N" Then
            valorigv = obtiene_igv_producto(mytablex.Fields("producto"))

            Dim total As Double

            total = 0
            total = mytablex.Fields("xtotal")

            If Val("" & valorigv) > 0 Then
                total = total / (1 + (Val("" & valorigv) / 100))
                sdx1 = total - (Val(xcosto) * Val(xfactor)) * Val("" & mytablex.Fields("xcanti"))
            Else
                sdx1 = Val("" & mytablex.Fields("xtotal")) - (Val(xcosto) * Val(xfactor)) * Val("" & mytablex.Fields("xcanti"))

            End If

        End If

        '' 31/01/2018 Productos unidades vendidos con o sin igv
        '' 31/01/2018 Productos unidades vendidos con o sin igv
   
        ''''09/10/2017 kenyo Familias productos Correcion Costos
   
        ''08/07/2017 KENYO redondeo reporte productos vendidos
        If CboRedondeo = "S" Then
            buf = Format(sdx1, "0.00")
        Else
            buf = Format(sdx1, "0.00000")

        End If

        ''08/07/2017 KENYO redondeo reporte productos vendidos
      
        objExcel.ActiveSheet.Cells(v, h + 8) = "" & buf
        buf = "" & mytablex.Fields("m")
        objExcel.ActiveSheet.Cells(v, h + 9) = "" & buf
     
        v = v + 1
   
        If "" & mytablex.Fields("m") = "S" Then
       
            '' 31/01/2018 Productos unidades vendidos con o sin igv
            ' suma2 = suma2 + Val("" & mytablex.Fields("xtotal"))
            total = mytablex.Fields("xtotal")

            If conigv = "" Or conigv = "S" Then
                total = Val("" & mytablex.Fields("xtotal"))

            End If
        
            If conigv = "N" Then
                valorigv = obtiene_igv_producto(mytablex.Fields("producto"))

                If Val("" & valorigv) > 0 Then
                    total = total / (1 + (Val("" & valorigv) / 100))
                Else
                    total = total

                End If

            End If
        
            suma2 = suma2 + total
      
            '' 31/01/2018 Productos unidades vendidos con o sin igv
   
            suma1 = suma1 + Val("" & mytablex.Fields("xcanti"))
 
            suma5 = suma5 + sdx
            suma6 = suma6 + sdx1
      
            ssuma1 = ssuma1 + Val("" & mytablex.Fields("xcanti"))
      
            '' 31/01/2018 Productos unidades vendidos con o sin igv
            'ssuma2 = ssuma2 + Val("" & mytablex.Fields("xtotal"))
            ssuma2 = ssuma2 + total
            '' 31/01/2018 Productos unidades vendidos con o sin igv
      
            ssuma5 = ssuma5 + sdx
            ssuma6 = ssuma6 + sdx1

        End If
   
        If "" & mytablex.Fields("m") = "D" Then
   
            '' 31/01/2018 Productos unidades vendidos con o sin igv
            total = mytablex.Fields("xtotal")

            If conigv = "" Or conigv = "S" Then
                total = Val("" & mytablex.Fields("xtotal"))

            End If
        
            If conigv = "N" Then
                valorigv = obtiene_igv_producto(mytablex.Fields("producto"))

                If Val("" & valorigv) > 0 Then
                    total = total / (1 + (Val("" & valorigv) / 100))
                Else
                    total = total

                End If

            End If

            '' 31/01/2018 Productos unidades vendidos con o sin igv
   
            suma3 = suma3 + Val("" & mytablex.Fields("xcanti"))
            suma4 = suma4 + total
      
            ssuma3 = ssuma3 + Val("" & mytablex.Fields("xcanti"))
            ssuma4 = ssuma4 + total

            '' 31/01/2018 Productos unidades vendidos con o sin igv
        End If
   
seguy13:
        mytablex.MoveNext
    Loop

    sw1 = 0
   
    ''08/07/2017 KENYO redondeo reporte productos vendidos
    If CboRedondeo = "S" Then
      
        buf = Format(suma1, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 5) = "" & buf
        objExcel.ActiveSheet.Cells(v, h + 5).Font.bold = True
        buf = Format(suma2, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 6) = "" & buf
        objExcel.ActiveSheet.Cells(v, h + 6).Font.bold = True
        buf = Format(suma5, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 7) = "" & buf
        objExcel.ActiveSheet.Cells(v, h + 7).Font.bold = True
        buf = Format(suma6, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 8) = "" & buf
        objExcel.ActiveSheet.Cells(v, h + 8).Font.bold = True
        buf = Format(suma3, "0.00")
        
        ''''13/09/2017 kenyo Mejora Reportes Familias Producto - quita buf
        objExcel.ActiveSheet.Cells(v, h + 9) = ""
        buf = Format(suma4, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 10) = ""
        ''''13/09/2017 kenyo Mejora Reportes Familias Producto - quita buf
        v = v + 1
        objExcel.ActiveSheet.Cells(v, h + 3) = "Gran Total"
           
        buf = Format(ssuma1, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 5) = "" & buf
        buf = Format(ssuma2, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 6) = "" & buf
        buf = Format(ssuma5, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 7) = "" & buf
        buf = Format(ssuma6, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 8) = "" & buf
        buf = Format(ssuma3, "0.00")
        ''''13/09/2017 kenyo Mejora Reportes Familias Producto - quita buf
        objExcel.ActiveSheet.Cells(v, h + 9) = ""
        buf = Format(ssuma4, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 10) = ""
        ''''13/09/2017 kenyo Mejora Reportes Familias Producto
    Else
       
        buf = Format(suma1, "0.00000")
        objExcel.ActiveSheet.Cells(v, h + 5) = "" & buf
        buf = Format(suma2, "0.00000")
        objExcel.ActiveSheet.Cells(v, h + 6) = "" & buf
        buf = Format(suma5, "0.00000")
        objExcel.ActiveSheet.Cells(v, h + 7) = "" & buf
        buf = Format(suma6, "0.00000")
        objExcel.ActiveSheet.Cells(v, h + 8) = "" & buf
        buf = Format(suma3, "0.00000")
        ''''13/09/2017 kenyo Mejora Reportes Familias Producto - quita buf
        objExcel.ActiveSheet.Cells(v, h + 9) = ""
        buf = Format(suma4, "0.00000")
        objExcel.ActiveSheet.Cells(v, h + 10) = ""
        ''''13/09/2017 kenyo Mejora Reportes Familias Producto

        v = v + 1
        buf = Format(ssuma1, "0.00000")
        objExcel.ActiveSheet.Cells(v, h + 5) = "" & buf
        buf = Format(ssuma2, "0.00000")
        objExcel.ActiveSheet.Cells(v, h + 6) = "" & buf
        buf = Format(ssuma5, "0.00000")
        objExcel.ActiveSheet.Cells(v, h + 7) = "" & buf
        buf = Format(ssuma6, "0.00000")
        objExcel.ActiveSheet.Cells(v, h + 8) = "" & buf
        buf = Format(ssuma3, "0.00000")
        ''''13/09/2017 kenyo Mejora Reportes Familias Producto - quita buf
        objExcel.ActiveSheet.Cells(v, h + 9) = ""
        buf = Format(ssuma4, "0.00000")
        objExcel.ActiveSheet.Cells(v, h + 10) = ""
        ''''13/09/2017 kenyo Mejora Reportes Familias Producto

    End If

    ''08/07/2017 KENYO redondeo reporte productos vendidos
   
    ''''13/09/2017 kenyo Mejora Reportes Familias Producto
    Dim k As Integer

    For k = 4 To 9
        objExcel.ActiveSheet.Cells(v, k).Font.bold = True
        objExcel.ActiveSheet.Cells(v, k).Interior.color = RGB(248, 243, 53)
    Next
    ''''13/09/2017 kenyo Mejora Reportes Familias Producto
  
    v = v + 1
    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto

End Sub

'' 31/01/2018 Productos unidades vendidos con o sin igv
Function obtiene_igv_producto(producto As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "Select igv from producto where producto='" & producto & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        obtiene_igv_producto = "" & mytablex.Fields("igv")

    End If

    mytablex.Close

End Function

'' 31/01/2018 Productos unidades vendidos con o sin igv

Sub cuerpo_programa_producto1xxcomision(mytablex As ADODB.Recordset)

    Dim vr

    Dim sw1          As Integer

    Dim Tmp          As String

    Dim tmp1         As String

    Dim sw           As Integer

    Dim buf          As String

    Dim found        As Integer

    Dim sdx          As Double

    Dim mytabley     As New ADODB.Recordset

    Dim xunidad      As String

    Dim xfactor      As String

    Dim xcosto       As String

    Dim sdx1         As Double

    Dim sdxtmp       As Double

    Dim v            As Long

    Dim h            As Integer

    Dim vprecios(10) As String

    Dim Heading(20)  As String

    h = 1
    sdx1 = 0
    sdx = 0
    sw = 0

    v = 4

    h = 1
    suma1 = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0
    suma5 = 0
    suma6 = 0

    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    ssuma4 = 0
    ssuma5 = 0
    ssuma6 = 0

    suma5 = 0
    suma6 = 0
    suma7 = 0
    suma8 = 0

    ssuma5 = 0
    ssuma6 = 0
    ssuma7 = 0
    ssuma8 = 0
    
    Heading(1) = Combo1
    Heading(2) = "Tipo"
    Heading(3) = "Serie"
    Heading(4) = "Numero"
    Heading(5) = "Fecha"
    Heading(6) = "Producto"
    Heading(7) = "Descripcion"
    Heading(8) = "Und"
    Heading(9) = "Fact"
    Heading(10) = "Cant"
    Heading(11) = "Precio"
    Heading(12) = "Total"
    Heading(13) = "M"
    Heading(14) = "Cajero"
    Heading(15) = "Caja"
    Heading(16) = "Turno"
    Heading(17) = "Estado"
    Heading(18) = "Hora"
    Heading(19) = "%Comis"
    Heading(20) = "TotComision"
  
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_ExcelComision(20, Heading())
    
    If Combo1 = "Producto" Then
        objExcel.ActiveSheet.Cells(1, 6) = "SEGUIMIENTO DE PRODUCTOS POR COMPROBANTES"
    ElseIf Combo1 = "Familia" Then
        objExcel.ActiveSheet.Cells(1, 6) = "SEGUIMIENTO DE COMPROBANTES POR FAMILIA"
    ElseIf Combo1 = "Vendedor" Then
        objExcel.ActiveSheet.Cells(1, 6) = "SEGUIMIENTO DE COMISIONES POR VENDEDOR"
    Else
        objExcel.ActiveSheet.Cells(1, 6) = "     REPORTE DE SEGUIMIENTO"

    End If
     
    objExcel.ActiveSheet.Cells(1, 6).Font.bold = True
    objExcel.ActiveSheet.Cells(1, 6).Font.Size = 14
    objExcel.ActiveSheet.Cells(1, 6).Font.color = RGB(0, 112, 184)
    
    objExcel.ActiveSheet.Cells(2, 1) = "FECHA INICIO  " + fechai
    objExcel.ActiveSheet.Cells(2, 4) = "FECHA FIN  " + fechaf

    tmp1 = ""

    Do

        If mytablex.EOF Then Exit Do
        vr = DoEvents()

        If Command2.Visible = False Then Exit Do
        If TxtCodProv.Text <> "%" Then
            found = ver_proveedor("" & mytablex.Fields("producto"))

            If found = 0 Then GoTo seguy13

        End If

        If Combo1 = "Producto" Then
            tmp1 = "" & mytablex.Fields("Producto")

        End If

        If Combo1 = "Codigo" Then
            tmp1 = "" & mytablex.Fields("codigo")

        End If

        If Combo1 = "Vendedor" Then
            tmp1 = "" & mytablex.Fields("vendedor")

        End If

        If Combo1 = "Comisiones" Then
            tmp1 = "" & mytablex.Fields("vendedor")

        End If

        If Combo1 = "Zona" Then
            tmp1 = "" & mytablex.Fields("Zona")

        End If

        If Combo1 = "Cajero" Then
            tmp1 = "" & mytablex.Fields("Usuario")

        End If

        If Combo1 = "Caja" Then
            tmp1 = "" & mytablex.Fields("Caja")

        End If

        If Combo1 = "Turno" Then
            tmp1 = "" & mytablex.Fields("Turno")

        End If

        If Combo1 = "Familia" Then
            tmp1 = "" & mytablex.Fields("Familia")

        End If

        If Combo1 = "Subfamilia" Then
            tmp1 = "" & mytablex.Fields("Subfamilia")

        End If

        If Combo1 = "Marca" Then
            tmp1 = "" & mytablex.Fields("Marca")

        End If

        If Combo1 = "Ccosto" Then
            tmp1 = "" & mytablex.Fields("Ccosto")

        End If

        If Combo1 = "Almacen" Then
            tmp1 = "" & mytablex.Fields("bodega")

        End If

        If sw = 0 Then
   
            If Combo1 = "Hora" Then
                buf = "" & mytablex.Fields("Hora")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf
                v = v + 1
                Tmp = "" & mytablex.Fields("Hora")

            End If
     
            If Combo1 = "Ccosto" Then
                buf = "" & mytablex.Fields("Ccosto")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf
                v = v + 1
                Tmp = "" & mytablex.Fields("Ccosto")

            End If
   
            If Combo1 = "Almacen" Then
                buf = "" & mytablex.Fields("bodega")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & " "
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_bodega(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("bodega")

            End If

            If Combo1 = "Familia" Then
                buf = "" & mytablex.Fields("Familia")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf & " " & busca_familia(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("Familia")

            End If
   
            If Combo1 = "Subfamilia" Then
                buf = "" & mytablex.Fields("Subfamilia")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_subfamilia(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("Subfamilia")

            End If
   
            If Combo1 = "Marca" Then
                buf = "" & mytablex.Fields("Marca")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_marca(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("Marca")

            End If
   
            If Combo1 = "Producto" Then
                buf = "" & mytablex.Fields("producto")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_producto(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("producto")

            End If

            If Combo1 = "Codigo" Then
                buf = "" & mytablex.Fields("codigo")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & ""
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_nombre(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("codigo")

            End If
   
            If Combo1 = "Vendedor" Then
                buf = "" & mytablex.Fields("vendedor")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & " "
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_vendedor(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("vendedor")

            End If
   
            If Combo1 = "Comisiones" Then
                buf = "" & mytablex.Fields("vendedor")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf & " " & busca_vendedor(buf)
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_vendedor(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("vendedor")

            End If
   
            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf & " " & busca_zona(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("zona")

            End If
   
            If Combo1 = "Cajero" Then
                buf = "" & mytablex.Fields("usuario")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_vendedor(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("usuario")

            End If
   
            If Combo1 = "Caja" Then
                buf = "" & mytablex.Fields("Caja")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf
                v = v + 1
                nlineas
                Tmp = "" & mytablex.Fields("Caja")

            End If
   
            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("Turno")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf
                v = v + 1
                Tmp = "" & mytablex.Fields("Turno")

            End If
   
            objExcel.ActiveSheet.Cells(v - 1, h).Font.bold = True
            objExcel.ActiveSheet.Cells(v - 1, h).Font.color = RGB(62, 95, 138)
            objExcel.ActiveSheet.Cells(v - 1, h + 1).Font.bold = True
            objExcel.ActiveSheet.Cells(v - 1, h + 1).Font.color = RGB(62, 95, 138)
   
            sw = 1
            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            suma6 = 0
            suma7 = 0
            suma8 = 0

        End If

        If Tmp <> tmp1 Then

            buf = Format(suma1, "0.00")
            objExcel.ActiveSheet.Cells(v, h + 9) = "" & buf
            objExcel.ActiveSheet.Cells(v, h + 9).Font.bold = True
            buf = Format(suma2, "0.00")
            objExcel.ActiveSheet.Cells(v, h + 11) = "" & buf
            objExcel.ActiveSheet.Cells(v, h + 11).Font.bold = True
            buf = Format(suma5, "0.00")
            objExcel.ActiveSheet.Cells(v, h + 19) = "" & buf
            objExcel.ActiveSheet.Cells(v, h + 19).Font.bold = True
       
            v = v + 1
   
            If Combo1 = "Ccosto" Then
                buf = "" & mytablex.Fields("Ccosto")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf
                v = v + 1
                Tmp = "" & mytablex.Fields("Ccosto")

            End If
   
            If Combo1 = "Almacen" Then
                buf = "" & mytablex.Fields("bodega")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & " "
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_bodega(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("bodega")

            End If

            If Combo1 = "Familia" Then
                buf = "" & mytablex.Fields("Familia")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf & " " & busca_familia(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("Familia")

            End If
   
            If Combo1 = "Subfamilia" Then
                buf = "" & mytablex.Fields("Subfamilia")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_subfamilia(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("Subfamilia")

            End If
   
            If Combo1 = "Marca" Then
                buf = "" & mytablex.Fields("Marca")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_marca(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("Marca")

            End If
   
            If Combo1 = "Producto" Then
                buf = "" & mytablex.Fields("producto")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_producto(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("producto")

            End If

            If Combo1 = "Codigo" Then
                buf = "" & mytablex.Fields("codigo")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & ""
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_nombre(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("codigo")

            End If
   
            If Combo1 = "Vendedor" Then
                buf = "" & mytablex.Fields("vendedor")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & " "
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_vendedor(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("vendedor")

            End If
   
            If Combo1 = "Comisiones" Then
                buf = "" & mytablex.Fields("vendedor")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf & " " & busca_vendedor(buf)
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_vendedor(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("vendedor")

            End If
   
            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf & " " & busca_zona(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("zona")

            End If
   
            If Combo1 = "Cajero" Then
                buf = "" & mytablex.Fields("usuario")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_vendedor(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("usuario")

            End If
   
            If Combo1 = "Caja" Then
                buf = "" & mytablex.Fields("Caja")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf
                v = v + 1
                nlineas
                Tmp = "" & mytablex.Fields("Caja")

            End If
   
            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("Turno")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf
                v = v + 1
                Tmp = "" & mytablex.Fields("Turno")

            End If
   
            objExcel.ActiveSheet.Cells(v - 1, h).Font.bold = True
            objExcel.ActiveSheet.Cells(v - 1, h).Font.color = RGB(62, 95, 138)
     
            objExcel.ActiveSheet.Cells(v - 1, h + 1).Font.bold = True
            objExcel.ActiveSheet.Cells(v - 1, h + 1).Font.color = RGB(62, 95, 138)
   
            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            suma6 = 0
            suma7 = 0
            suma8 = 0

        End If
    
        If Combo1 = "Familia" Then
            buf = "" & mytablex.Fields("familia")
            objExcel.ActiveSheet.Cells(v, h) = "'" & buf
            objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("tipo")
        
        ElseIf Combo1 = "Subfamilia" Then
            buf = "" & mytablex.Fields("producto")
            objExcel.ActiveSheet.Cells(v, h + 2) = "'" & buf
            buf = "" & mytablex.Fields("descripcio")
            objExcel.ActiveSheet.Cells(v, h + 1) = "'" & buf
            buf = "" & mytablex.Fields("subfamilia")
         
        ElseIf Combo1 = "Producto" Then
            buf = "" & mytablex.Fields("producto")
            objExcel.ActiveSheet.Cells(v, h + 2) = "'" & buf
            buf = "" & mytablex.Fields("descripcio")
            objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("tipo")
         
        ElseIf Combo1 = "Cajero" Then
            buf = "" & mytablex.Fields("usuario")
            objExcel.ActiveSheet.Cells(v, h + 2) = "'" & buf
            buf = "" & mytablex.Fields("descripcio")
            objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("tipo")
         
        ElseIf Combo1 = "Marca" Then
            buf = "" & mytablex.Fields("producto")
            objExcel.ActiveSheet.Cells(v, h + 2) = "'" & buf
            buf = "" & mytablex.Fields("descripcio")
            objExcel.ActiveSheet.Cells(v, h + 1) = "'" & buf
            buf = "" & mytablex.Fields("marca")
            objExcel.ActiveSheet.Cells(v, h) = "'" & buf
        
        ElseIf Combo1 = "Vendedor" Then
            buf = "" & mytablex.Fields("vendedor")
            objExcel.ActiveSheet.Cells(v, h) = "'" & buf
            buf = "" & mytablex.Fields("descripcio")
            objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("tipo")
   
        ElseIf Combo1 = "Codigo" Then
            buf = "" & mytablex.Fields("codigo")
            objExcel.ActiveSheet.Cells(v, h) = "'" & buf
   
        ElseIf Combo1 = "Almacen" Then
            buf = "" & mytablex.Fields("bodega")
            objExcel.ActiveSheet.Cells(v, h) = "'" & buf
            buf = "" & mytablex.Fields("descripcio")
            objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("tipo")
   
        Else
            buf = "" & mytablex.Fields("producto")
            objExcel.ActiveSheet.Cells(v, h + 2) = "'" & buf
            buf = "" & mytablex.Fields("descripcio")
            objExcel.ActiveSheet.Cells(v, h + 1) = "'" & buf
        
        End If

        buf = "" & mytablex.Fields("serie")
        objExcel.ActiveSheet.Cells(v, h + 2) = "" & buf

        buf = "" & mytablex.Fields("numero")
        objExcel.ActiveSheet.Cells(v, h + 3) = "" & buf

        buf = "" & mytablex.Fields("fecha")
        objExcel.ActiveSheet.Cells(v, h + 4) = "" & buf

        buf = "'" & mytablex.Fields("producto")
        objExcel.ActiveSheet.Cells(v, h + 5) = "" & buf

        buf = "" & mytablex.Fields("Descripcio")
        objExcel.ActiveSheet.Cells(v, h + 6) = "" & buf

        buf = "" & mytablex.Fields("unidad")
        objExcel.ActiveSheet.Cells(v, h + 7) = "" & buf

        buf = "" & mytablex.Fields("factor")
        objExcel.ActiveSheet.Cells(v, h + 8) = "" & buf

        buf = "" & mytablex.Fields("cantidad")
        suma1 = suma1 + Val("" & mytablex.Fields("cantidad"))
        ssuma1 = ssuma1 + Val("" & mytablex.Fields("cantidad"))
        objExcel.ActiveSheet.Cells(v, h + 9) = "" & buf

        buf = "" & mytablex.Fields("precio")
        objExcel.ActiveSheet.Cells(v, h + 10) = "" & buf

        buf = "" & mytablex.Fields("total")
        suma2 = suma2 + Val("" & mytablex.Fields("total"))
        ssuma2 = ssuma2 + Val("" & mytablex.Fields("total"))
        objExcel.ActiveSheet.Cells(v, h + 11) = "" & buf

        buf = "" & mytablex.Fields("moneda")
        objExcel.ActiveSheet.Cells(v, h + 12) = "" & buf

        buf = "'" & mytablex.Fields("usuario")
        objExcel.ActiveSheet.Cells(v, h + 13) = "" & buf

        buf = "'" & mytablex.Fields("caja")
        objExcel.ActiveSheet.Cells(v, h + 14) = "" & buf

        buf = "'" & mytablex.Fields("turno")
        objExcel.ActiveSheet.Cells(v, h + 15) = "" & buf

        buf = "" & mytablex.Fields("estado")
        objExcel.ActiveSheet.Cells(v, h + 16) = "" & buf

        buf = "" & mytablex.Fields("hora")
        objExcel.ActiveSheet.Cells(v, h + 17) = "" & buf

        buf = "" & mytablex.Fields("comision")
        objExcel.ActiveSheet.Cells(v, h + 18) = "" & buf

        buf = Format(mytablex.Fields("comision") * mytablex.Fields("total") / 100, "0.00")
        suma5 = suma5 + ((mytablex.Fields("comision") * mytablex.Fields("total") / 100))
        suma5 = Format(suma5, "0.00")

        ssuma5 = ssuma5 + mytablex.Fields("comision") * mytablex.Fields("total") / 100
        ssuma5 = Format(ssuma5, "0.00")

        objExcel.ActiveSheet.Cells(v, h + 19) = "" & buf
       
        xunidad = "UND"
        xfactor = "1"
        xcosto = "0"
        sdxtmp = 0

        If Val(xfactor) <= 0 Then
            xfactor = "1"

        End If
    
        v = v + 1

seguy13:
        mytablex.MoveNext
    Loop

    sw1 = 0
   
    If CboRedondeo = "S" Then
        buf = Format(suma1, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 9) = "" & buf
        objExcel.ActiveSheet.Cells(v, h + 9).Font.bold = True
        buf = Format(suma2, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 11) = "" & buf
        objExcel.ActiveSheet.Cells(v, h + 11).Font.bold = True

        buf = Format(suma5, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 19) = "" & buf
        objExcel.ActiveSheet.Cells(v, h + 19).Font.bold = True
        v = v + 1

        objExcel.ActiveSheet.Cells(v, h + 7) = "Gran Total"

        buf = Format(ssuma1, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 9) = "" & buf
        buf = Format(ssuma2, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 11) = "" & buf
        buf = Format(ssuma5, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 19) = "" & buf

    Else

        buf = Format(suma1, "0.00000")
        objExcel.ActiveSheet.Cells(v, h + 9) = "" & buf
        objExcel.ActiveSheet.Cells(v, h + 9).Font.bold = True
        buf = Format(suma2, "0.00000")
        objExcel.ActiveSheet.Cells(v, h + 11) = "" & buf
        objExcel.ActiveSheet.Cells(v, h + 11).Font.bold = True
        buf = Format(suma5, "0.00000")
        objExcel.ActiveSheet.Cells(v, h + 19) = "" & buf
        objExcel.ActiveSheet.Cells(v, h + 19).Font.bold = True

        v = v + 1
        objExcel.ActiveSheet.Cells(v, h + 7) = "Gran Total"
        buf = Format(ssuma1, "0.00000")
        objExcel.ActiveSheet.Cells(v, h + 9) = "" & buf
        buf = Format(ssuma2, "0.00000")
        objExcel.ActiveSheet.Cells(v, h + 11) = "" & buf
        buf = Format(ssuma5, "0.00000")
        objExcel.ActiveSheet.Cells(v, h + 19) = "" & buf

    End If

    Dim k As Integer

    For k = 8 To 20
        objExcel.ActiveSheet.Cells(v, k).Font.bold = True
        objExcel.ActiveSheet.Cells(v, k).Interior.color = RGB(248, 243, 53)
    Next
  
    v = v + 1
    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto

End Sub

Sub cuerpo_programa_productosdiariosexcel(mytablex As ADODB.Recordset)

    Dim vr

    Dim sw1         As Integer

    Dim Tmp         As String

    Dim tmp1        As String

    Dim sw          As Integer

    Dim buf         As String

    Dim found       As Integer

    Dim sdx         As Double

    Dim mytabley    As New ADODB.Recordset

    Dim xunidad     As String

    Dim xfactor     As String

    Dim xcosto      As String

    Dim sdx1        As Double

    Dim sdxtmp      As Double

    Dim v           As Long

    Dim h           As Integer

    Dim Heading(33) As String

    h = 1
    sdx1 = 0
    sdx = 0
    sw = 0

    v = 4

    h = 1
    suma1 = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0
    suma5 = 0
    suma6 = 0

    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    ssuma4 = 0
    ssuma5 = 0
    ssuma6 = 0

    suma5 = 0
    suma6 = 0
    suma7 = 0
    suma8 = 0

    ssuma5 = 0
    ssuma6 = 0
    ssuma7 = 0
    ssuma8 = 0

    Dim I As Integer

    For I = 1 To 31
        xmeses(I) = 0
        xxmeses(I) = 0
    Next I

    Heading(1) = Combo1
    
    Dim ij As Long

    For ij = 1 To 31
        Heading(ij + 1) = ij
    Next ij
  
    Heading(33) = "TOTAL"
  
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_Productosdiarios(33, Heading())
    
    If Combo1 = "Producto" Then
        objExcel.ActiveSheet.Cells(1, 6) = "SEGUIMIENTO DE COMPROBANTES POR PRODUCTOS"
    ElseIf Combo1 = "Familia" Then
        objExcel.ActiveSheet.Cells(1, 6) = "SEGUIMIENTO DE COMPROBANTES POR FAMILIA"
    ElseIf Combo1 = "Vendedor" Then
        objExcel.ActiveSheet.Cells(1, 6) = "SEGUIMIENTO DE COMISIONES POR VENDEDOR"
    Else
        objExcel.ActiveSheet.Cells(1, 6) = "     REPORTE DE SEGUIMIENTO"

    End If
     
    objExcel.ActiveSheet.Cells(1, 6).Font.bold = True
    objExcel.ActiveSheet.Cells(1, 6).Font.Size = 14
    objExcel.ActiveSheet.Cells(1, 6).Font.color = RGB(0, 112, 184)
    
    objExcel.ActiveSheet.Cells(2, 1) = "FECHA INICIO  " + fechai
    objExcel.ActiveSheet.Cells(2, 4) = "FECHA FIN  " + fechaf

    tmp1 = ""

    Do

        If mytablex.EOF Then Exit Do
        vr = DoEvents()

        If Command2.Visible = False Then Exit Do
        If TxtCodProv.Text <> "%" Then
            found = ver_proveedor("" & mytablex.Fields("producto"))

            If found = 0 Then GoTo seguy13

        End If

        If Combo1 = "Producto" Then
            tmp1 = "" & mytablex.Fields("Producto")

        End If

        If Combo1 = "Codigo" Then
            tmp1 = "" & mytablex.Fields("codigo")

        End If

        If Combo1 = "Vendedor" Then
            tmp1 = "" & mytablex.Fields("vendedor")

        End If

        If Combo1 = "Comisiones" Then
            tmp1 = "" & mytablex.Fields("vendedor")

        End If

        If Combo1 = "Zona" Then
            tmp1 = "" & mytablex.Fields("Zona")

        End If

        If Combo1 = "Cajero" Then
            tmp1 = "" & mytablex.Fields("Usuario")

        End If

        If Combo1 = "Caja" Then
            tmp1 = "" & mytablex.Fields("Caja")

        End If

        If Combo1 = "Turno" Then
            tmp1 = "" & mytablex.Fields("Turno")

        End If

        If Combo1 = "Familia" Then
            tmp1 = "" & mytablex.Fields("Familia")

        End If

        If Combo1 = "Subfamilia" Then
            tmp1 = "" & mytablex.Fields("Subfamilia")

        End If

        If Combo1 = "Marca" Then
            tmp1 = "" & mytablex.Fields("Marca")

        End If

        If Combo1 = "Ccosto" Then
            tmp1 = "" & mytablex.Fields("Ccosto")

        End If

        If Combo1 = "Almacen" Then
            tmp1 = "" & mytablex.Fields("bodega")

        End If

        If sw = 0 Then
            v = v + 1
   
            If Combo1 = "Ccosto" Then
                buf = "" & mytablex.Fields("Ccosto")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf
                v = v + 1
                Tmp = "" & mytablex.Fields("Ccosto")

            End If
   
            If Combo1 = "Almacen" Then
                buf = "" & mytablex.Fields("bodega")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & " "
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_bodega(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("bodega")

            End If

            If Combo1 = "Familia" Then
                '   buf = "" & mytablex.Fields("Familia")
                '   objExcel.ActiveSheet.Cells(v, h) = "" & buf & " " & busca_familia(buf)
                '   v = v + 1
                Tmp = "" & mytablex.Fields("Familia")

            End If
   
            If Combo1 = "Subfamilia" Then
                buf = "" & mytablex.Fields("Subfamilia")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_subfamilia(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("Subfamilia")

            End If
   
            If Combo1 = "Marca" Then
                buf = "" & mytablex.Fields("Marca")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_marca(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("Marca")

            End If
   
            If Combo1 = "Producto" Then
                buf = "" & mytablex.Fields("producto")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_producto(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("producto")

            End If

            If Combo1 = "Codigo" Then
                buf = "" & mytablex.Fields("codigo")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & ""
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_nombre(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("codigo")

            End If
   
            If Combo1 = "Vendedor" Then
                buf = "" & mytablex.Fields("vendedor")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & " "
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_vendedor(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("vendedor")

            End If
   
            If Combo1 = "Comisiones" Then
                buf = "" & mytablex.Fields("vendedor")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf & " " & busca_vendedor(buf)
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_vendedor(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("vendedor")

            End If
   
            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf & " " & busca_zona(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("zona")

            End If
   
            If Combo1 = "Cajero" Then
                buf = "" & mytablex.Fields("usuario")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_vendedor(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("usuario")

            End If
   
            If Combo1 = "Caja" Then
                buf = "" & mytablex.Fields("Caja")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf
                v = v + 1
                nlineas
                Tmp = "" & mytablex.Fields("Caja")

            End If
   
            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("Turno")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf
                v = v + 1
                Tmp = "" & mytablex.Fields("Turno")

            End If
   
            '   objExcel.ActiveSheet.Cells(v - 1, h).Font.bold = True
            '   objExcel.ActiveSheet.Cells(v - 1, h).Font.color = RGB(62, 95, 138)
            '   objExcel.ActiveSheet.Cells(v - 1, h + 1).Font.bold = True
            '   objExcel.ActiveSheet.Cells(v - 1, h + 1).Font.color = RGB(62, 95, 138)
            '
            If orden = "CANT" Then
                xmeses(Val("" & mytablex.Fields("xmes"))) = xmeses(Val("" & mytablex.Fields("xmes"))) + Val("" & mytablex.Fields("xcanti"))
                xxmeses(Val("" & mytablex.Fields("xmes"))) = xxmeses(Val("" & mytablex.Fields("xmes"))) + Val("" & mytablex.Fields("xcanti"))

            End If
   
            '   For i = 1 To 31
            '        sdx = sdx + xmeses(i)
            '        objExcel.ActiveSheet.Cells(v, i + 1) = xmeses(i)
            '   Next i
   
            sw = 1
   
            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            suma6 = 0
            suma7 = 0
            suma8 = 0

        End If

        If Tmp <> tmp1 Then
   
            sdx = 0

            For I = 1 To 31
                sdx = sdx + xmeses(I)
                objExcel.ActiveSheet.Cells(v, I + 1) = xmeses(I)
            Next I

            v = v + 1
   
            If Combo1 = "Ccosto" Then
                buf = "" & mytablex.Fields("Ccosto")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf
                v = v + 1
                Tmp = "" & mytablex.Fields("Ccosto")

            End If
   
            If Combo1 = "Almacen" Then
                buf = "" & mytablex.Fields("bodega")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & " "
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_bodega(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("bodega")

            End If

            If Combo1 = "Familia" Then
                '   buf = "" & mytablex.Fields("Familia")
                '   objExcel.ActiveSheet.Cells(v, h) = "" & buf & " " & busca_familia(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("Familia")
        
            End If
   
            If Combo1 = "Subfamilia" Then
                buf = "" & mytablex.Fields("Subfamilia")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_subfamilia(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("Subfamilia")

            End If
   
            If Combo1 = "Marca" Then
                buf = "" & mytablex.Fields("Marca")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_marca(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("Marca")

            End If
   
            If Combo1 = "Producto" Then
                buf = "" & mytablex.Fields("producto")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_producto(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("producto")

            End If

            If Combo1 = "Codigo" Then
                buf = "" & mytablex.Fields("codigo")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & ""
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_nombre(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("codigo")

            End If
   
            If Combo1 = "Vendedor" Then
                buf = "" & mytablex.Fields("vendedor")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & " "
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_vendedor(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("vendedor")

            End If
   
            If Combo1 = "Comisiones" Then
                buf = "" & mytablex.Fields("vendedor")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf & " " & busca_vendedor(buf)
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_vendedor(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("vendedor")

            End If
   
            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                objExcel.ActiveSheet.Cells(v, h) = "" & buf & " " & busca_zona(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("zona")

            End If
   
            If Combo1 = "Cajero" Then
                buf = "" & mytablex.Fields("usuario")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_vendedor(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("usuario")

            End If
   
            If Combo1 = "Caja" Then
                buf = "" & mytablex.Fields("Caja")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf
                v = v + 1
                nlineas
                Tmp = "" & mytablex.Fields("Caja")

            End If
   
            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("Turno")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf
                v = v + 1
                Tmp = "" & mytablex.Fields("Turno")

            End If
   
            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            suma6 = 0
            suma7 = 0
            suma8 = 0

            If orden = "MONTO" Then
                xmeses(Val("" & mytablex.Fields("xmes"))) = xmeses(Val("" & mytablex.Fields("xmes"))) + Val("" & mytablex.Fields("xtotal"))
                xxmeses(Val("" & mytablex.Fields("xmes"))) = xxmeses(Val("" & mytablex.Fields("xmes"))) + Val("" & mytablex.Fields("xtotal"))

            End If
   
            If orden = "CANT" Then
                xmeses(Val("" & mytablex.Fields("xmes"))) = xmeses(Val("" & mytablex.Fields("xmes"))) + Val("" & mytablex.Fields("xcanti"))
                xxmeses(Val("" & mytablex.Fields("xmes"))) = xxmeses(Val("" & mytablex.Fields("xmes"))) + Val("" & mytablex.Fields("xcanti"))

            End If

        End If
    
        If Combo1 = "Familia" Then
            buf = "" & mytablex.Fields("familia")
            objExcel.ActiveSheet.Cells(v, h) = "'" & buf
            Tmp = "" & mytablex.Fields("familia")
        ElseIf Combo1 = "Subfamilia" Then
            buf = "" & mytablex.Fields("producto")
            objExcel.ActiveSheet.Cells(v, h + 2) = "'" & buf
            buf = "" & mytablex.Fields("descripcio")
            objExcel.ActiveSheet.Cells(v, h + 1) = "'" & buf
            buf = "" & mytablex.Fields("subfamilia")
         
        ElseIf Combo1 = "Producto" Then
            buf = "" & mytablex.Fields("producto")
            objExcel.ActiveSheet.Cells(v, h + 2) = "'" & buf
        ElseIf Combo1 = "Marca" Then
            buf = "" & mytablex.Fields("producto")
            objExcel.ActiveSheet.Cells(v, h + 2) = "'" & buf
            buf = "" & mytablex.Fields("descripcio")
            objExcel.ActiveSheet.Cells(v, h + 1) = "'" & buf
            buf = "" & mytablex.Fields("marca")
            objExcel.ActiveSheet.Cells(v, h) = "'" & buf
        
        ElseIf Combo1 = "Vendedor" Then
            buf = "" & mytablex.Fields("vendedor")
            objExcel.ActiveSheet.Cells(v, h) = "'" & buf
   
        ElseIf Combo1 = "Codigo" Then
            buf = "" & mytablex.Fields("codigo")
            objExcel.ActiveSheet.Cells(v, h) = "'" & buf
        Else
            buf = "" & mytablex.Fields("producto")
            objExcel.ActiveSheet.Cells(v, h + 2) = "'" & buf
            buf = "" & mytablex.Fields("descripcio")
            objExcel.ActiveSheet.Cells(v, h + 1) = "'" & buf

        End If
     
        ' For i = 1 To 31
        ' xmeses(Val("" & mytablex.Fields("xmes"))) = xmeses(Val("" & mytablex.Fields("xmes"))) + Val("" & mytablex.Fields("xcanti"))
        xmeses(Val("" & mytablex.Fields("xmes"))) = Val("" & mytablex.Fields("xcanti"))
        objExcel.ActiveSheet.Cells(v, h + 1) = xmeses(I)
        ' Next

seguy13:
        mytablex.MoveNext
    Loop

    'sw1 = 0
    '   Dim k As Integer
    '   For k = 8 To 20
    '      objExcel.ActiveSheet.Cells(v, k).Font.bold = True
    '      objExcel.ActiveSheet.Cells(v, k).Interior.color = RGB(248, 243, 53)
    '   Next
  
    v = v + 1
    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
   
End Sub

Sub cabecera_producto1()

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
    buf = titulo
    
    '''24/08/2017  Kenyo descripcion larga en reportes ticket
    'found = formateaa(buf, 90, 2, 0)
    found = formateaa(buf, 90, 2, 0)
    '''24/08/2017  Kenyo descripcion larga en reportes ticket
    
    found = formateaa("Fechai : " & fechai, 25, 2, 0)
    found = formateaa("Fechaf : " & fechaf, 25, 2, 0)
    
    '''24/08/2017  Kenyo descripcion larga en reportes ticket
    ' buf = String(152, "-")
    ' found = formateaa(buf, 152, 2, 0)
    buf = String(83, "-")
    found = formateaa(buf, 83, 2, 0)
    '''24/08/2017  Kenyo descripcion larga en reportes ticket
    
    buf = "Producto"
    found = formateaa(buf, 11, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Descripcio"
   
    '''24/08/2017  Kenyo descripcion larga en reportes ticket
    'found = formateaa(buf, 59, 0, 0)
    found = formateaa(buf, 30, 0, 0)
    '''24/08/2017  Kenyo descripcion larga en reportes ticket
   
    found = formateaa("", 1, 0, 0)
   
    '''24/08/2017  Kenyo descripcion larga en reportes ticket
    'buf = "Unida"
    'found = formateaa(buf, 5, 0, 0)
    'found = formateaa("", 1, 0, 0)
        
    ' buf = "fact"
    ' found = formateaa(buf, 4, 0, 0)
    ' found = formateaa("", 1, 0, 0)
    '''24/08/2017  Kenyo descripcion larga en reportes ticket
   
    buf = "Canti "
    found = formateaa(buf, 7, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = "Total "
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = "TCosto "
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = "Ganancia "
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)

    '''24/08/2017  Kenyo descripcion larga en reportes ticket
    'buf = "M"
    'found = formateaa(buf, 1, 0, 0)
    found = formateaa("", 1, 0, 0)
   
    'buf = "Saldo"
    'found = formateaa(buf, 10, 0, 0)
    found = formateaa("", 1, 2, 0)
    '''24/08/2017  Kenyo descripcion larga en reportes ticket
   
    'buf = "Comision"
    'found = formateaa(buf, 10, 0, 0)
    'found = formateaa("", 1, 2, 0)
   
    '''24/08/2017  Kenyo descripcion larga en reportes ticket
    'buf = String(152, "-")
    'found = formateaa(buf, 152, 2, 0)
    buf = String(83, "-")
    found = formateaa(buf, 83, 2, 0)
    '''24/08/2017  Kenyo descripcion larga en reportes ticket
   
End Sub

Sub imprime_opcion3()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset

    Dim buf      As String

    contpag = 0
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0

    found = sql_producto3(mytablex)

    If found = 0 Then
        mytablex.Close
    
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_producto3
    cuerpo_programa_producto3 mytablex
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
     
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1

End Sub

Function sql_producto3(mytablex As ADODB.Recordset)

    Dim buf  As String

    Dim ybuf As String

    Dim xbuf As String

    If Len(fechai) <> 10 Then Exit Function
    If Len(fechaf) <> 10 Then Exit Function
    If Not IsDate(fechai) Then Exit Function
    If Not IsDate(fechaf) Then Exit Function

    xbuf = ""

    If Combo1 = "Producto" Then
        xbuf = "Producto"

    End If

    If Combo1 = "Codigo" Then
        xbuf = "Codigo"

    End If

    If Combo1 = "Vendedor" Then
        xbuf = "vendedor"

    End If

    If Combo1 = "Comisiones" Then
        xbuf = "vendedor"

    End If

    If Combo1 = "Zona" Then
        xbuf = "Zona"

    End If

    If Combo1 = "Cajero" Then
        xbuf = "Usuario"

    End If

    If Combo1 = "Caja" Then
        xbuf = "Caja"

    End If

    If Combo1 = "Turno" Then
        xbuf = "Turno"

    End If

    If Combo1 = "Familia" Then
        xbuf = "Familia"

    End If

    If Combo1 = "Subfamilia" Then
        xbuf = "Subfamilia"

    End If

    If Combo1 = "Ccosto" Then
        xbuf = "Ccosto"

    End If

    If Combo1 = "Almacen" Then
        xbuf = "bodega"

    End If

    If Combo1 = "Marca" Then
        xbuf = "Marca"

    End If

    buf = "select " & xbuf & ",Producto,Descripcio,moneda as m,sum(cantidad*factor) as xcanti,sum(cantdev*factor) as Despacho,sum(total) as xtotal from " & xdata & " where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If local1 <> "%" Then
        buf = buf & " and local='" & extra_loquesea(local1) & "'"

    End If

    If tipo <> "%" Then
        buf = buf & " and tipo like '" & tipo & "'"

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

    If usuario <> "%" Then
        buf = buf & " and usuario like '" & usuario & "'"

    End If

    If bodega <> "%" Then
        buf = buf & " and bodega like '" & extra_loquesea(bodega) & "'"

    End If

    If horai <> "%" And horaf <> "%" Then
        buf = buf & " and HORA BETWEEN '" & horai & "' AND '" & horaf & "'"

    End If

    If unidad <> "%" Then
        buf = buf & " and unidad like '" & unidad & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & caja & "'"

    End If

    If servicio <> "%" Then
        'If servicio = "Autoservicio" Then
        buf = buf & " and servicio='" & extra_loquesea(servicio) & "'"
        'End If
        'If servicio = "Comanda" Then
        '   buf = buf & " and servicio='C'"
        'End If
        'If servicio = "Delivery" Then
        '   buf = buf & " and servicio='D'"
        'End If

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & turno & "'"

    End If

    If producto <> "%" Then
        buf = buf & " and [producto] like '" & producto & "'"

    End If

    If ccosto <> "%" Then
        buf = buf & " and ccosto like '" & ccosto & "'"

    End If

    If familia <> "%" Then
        buf = buf & " and familia like '" & familia & "'"

    End If

    If subfamilia <> "%" Then
        buf = buf & " and subfamilia like '" & subfamilia & "'"

    End If

    If marca <> "%" Then
        buf = buf & " and marca like '" & marca & "'"

    End If

    If descripcio <> "%" Then
        buf = buf & " and [descripcio] like '" & descripcio & "'"

    End If

    If moneda <> "%" Then
        buf = buf & " and moneda like '" & moneda & "'"

    End If

    If vendedor <> "%" Then
        buf = buf & " and vendedor like '" & vendedor & "'"

    End If

    If estado <> "%" Then
        buf = buf & " and estado='" & estado & "'"

    End If

    If acu = "V" Then
        buf = buf & " and (acu='1' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') "

    End If

    If acu = "C" Then
        buf = buf & " and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') "

    End If

    If acu <> "C" And acu <> "V" Then

        'buf = buf & " and acu='" & acu & "'"
    End If

    If orden = "CANT" Then
        ybuf = " SUM(cantidad*factor) "

    End If

    If orden = "MONTO" Then
        ybuf = " SUM(total) "

    End If

    'If orden = "GANANCIA" Then
    '   ybuf = " SUM(total)-SUM(cantidad*factor*tcosto) "
    'End If

    buf = buf & "  group by " & xbuf & ", producto,Descripcio,moneda  order  by " & xbuf & " ," & ybuf & " DESC "
    'MsgBox buf
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    sql_producto3 = 1

End Function

Sub cuerpo_programa_producto3(mytablex As ADODB.Recordset)

    Dim vr

    Dim xsaldo   As Double

    Dim sw1      As Integer

    Dim Tmp      As String

    Dim tmp1     As String

    Dim sw       As Integer

    Dim buf      As String

    Dim found    As Integer

    Dim sdx      As Double

    Dim mytabley As New ADODB.Recordset

    Dim xunidad  As String

    Dim xfactor  As String

    Dim xcosto   As String

    Dim sdx1     As Double

    Dim sdxtmp   As Double

    Dim psdx     As Double

    sdx1 = 0
    sdx = 0
    sw = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0
    suma5 = 0
    suma6 = 0

    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    ssuma4 = 0
    ssuma5 = 0
    ssuma6 = 0

    suma5 = 0
    suma6 = 0
    suma7 = 0
    suma8 = 0

    ssuma5 = 0
    ssuma6 = 0
    ssuma7 = 0
    ssuma8 = 0
    tmp1 = ""
    Do

        If mytablex.EOF Then Exit Do
        vr = DoEvents()

        If Command2.Visible = False Then Exit Do
        If TxtCodProv.Text <> "%" Then
            found = ver_proveedor("" & mytablex.Fields("producto"))

            If found = 0 Then GoTo seguy14

        End If

        If Combo1 = "Producto" Then
            tmp1 = "" & mytablex.Fields("Producto")

        End If

        If Combo1 = "Codigo" Then
            tmp1 = "" & mytablex.Fields("codigo")

        End If

        If Combo1 = "Vendedor" Then
            tmp1 = "" & mytablex.Fields("vendedor")

        End If

        If Combo1 = "Comisiones" Then
            tmp1 = "" & mytablex.Fields("vendedor")

        End If

        If Combo1 = "Zona" Then
            tmp1 = "" & mytablex.Fields("Zona")

        End If

        If Combo1 = "Cajero" Then
            tmp1 = "" & mytablex.Fields("Usuario")

        End If

        If Combo1 = "Caja" Then
            tmp1 = "" & mytablex.Fields("Caja")

        End If

        If Combo1 = "Turno" Then
            tmp1 = "" & mytablex.Fields("Turno")

        End If

        If Combo1 = "Familia" Then
            tmp1 = "" & mytablex.Fields("Familia")

        End If

        If Combo1 = "Subfamilia" Then
            tmp1 = "" & mytablex.Fields("Subfamilia")

        End If

        If Combo1 = "Marca" Then
            tmp1 = "" & mytablex.Fields("Marca")

        End If

        If Combo1 = "Ccosto" Then
            tmp1 = "" & mytablex.Fields("Ccosto")

        End If

        If Combo1 = "Almacen" Then
            tmp1 = "" & mytablex.Fields("bodega")

        End If

        If sw = 0 Then
            If Combo1 = "Ccosto" Then
                buf = "" & mytablex.Fields("Ccosto")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Ccosto")

            End If

            If Combo1 = "Almacen" Then
                buf = "" & mytablex.Fields("bodega")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_bodega(buf)
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("bodega")

            End If

            If Combo1 = "Familia" Then
                buf = "" & mytablex.Fields("Familia")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Familia")

            End If

            If Combo1 = "Subfamilia" Then
                buf = "" & mytablex.Fields("Subfamilia")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Subfamilia")

            End If

            If Combo1 = "Marca" Then
                buf = "" & mytablex.Fields("Marca")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Marca")

            End If

            If Combo1 = "Producto" Then
                buf = "" & mytablex.Fields("Producto")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = ""
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("producto")

            End If

            If Combo1 = "Codigo" Then
                buf = "" & mytablex.Fields("codigo")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_nombre(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("codigo")

            End If

            If Combo1 = "Vendedor" Then
                buf = "" & mytablex.Fields("vendedor")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Comisiones" Then
                buf = "" & mytablex.Fields("vendedor")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_zona(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("zona")

            End If

            If Combo1 = "Cajero" Then
                buf = "" & mytablex.Fields("usuario")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("usuario")

            End If

            If Combo1 = "Caja" Then
                buf = "" & mytablex.Fields("Caja")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Caja")

            End If

            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("Turno")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Turno")

            End If
   
            sw = 1
            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            suma6 = 0
            suma7 = 0
            suma8 = 0

        End If

        If Tmp <> tmp1 Then
            found = formateaa("", 83, 0, 0)
   
            buf = Format(suma1, "0.00")
            found = formateaa(buf, 7, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = Format(suma2, "0.00")
            found = formateaa(buf, 7, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = Format(suma3, "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 2, 0)
            nlineas
   
            If Combo1 = "Ccosto" Then
                buf = "" & mytablex.Fields("Ccosto")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Ccosto")

            End If

            If Combo1 = "Almacen" Then
                buf = "" & mytablex.Fields("bodega")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_bodega(buf)
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("bodega")

            End If

            If Combo1 = "Familia" Then
                buf = "" & mytablex.Fields("Familia")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Familia")

            End If

            If Combo1 = "Subfamilia" Then
                buf = "" & mytablex.Fields("Subfamilia")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Subfamilia")

            End If

            If Combo1 = "Marca" Then
                buf = "" & mytablex.Fields("Marca")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Marca")

            End If

            If Combo1 = "Producto" Then
                buf = "" & mytablex.Fields("Producto")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = ""
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("producto")

            End If
   
            If Combo1 = "Codigo" Then
                buf = "" & mytablex.Fields("codigo")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_nombre(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("codigo")

            End If

            If Combo1 = "Vendedor" Then
                buf = "" & mytablex.Fields("vendedor")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Comisiones" Then
                buf = "" & mytablex.Fields("vendedor")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_zona(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("zona")

            End If

            If Combo1 = "Cajero" Then
                buf = "" & mytablex.Fields("usuario")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("usuario")

            End If

            If Combo1 = "Caja" Then
                buf = "" & mytablex.Fields("Caja")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Caja")

            End If

            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("Turno")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Turno")

            End If

            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            suma6 = 0
            suma7 = 0
            suma8 = 0

        End If

        buf = "" & mytablex.Fields("producto")
        found = formateaa(buf, 11, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("descripcio")
        found = formateaa(buf, 59, 0, 0)
        found = formateaa("", 1, 0, 0)
        xunidad = "UND"
        xfactor = "1"
        xcosto = "0"
        sdxtmp = 0
        mytabley.Open "select * from producto where [producto]='" & "" & mytablex.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic

        If mytabley.RecordCount > 0 Then
            xunidad = "" & mytabley.Fields("unidad")
            xfactor = "" & mytabley.Fields("factor")
            sdxtmp = 0

            If Val(xfactor) <= 0 Then
                xfactor = "1"

            End If
       
        End If

        mytabley.Close
        xsaldo = 0
        mytabley.Open "select producto,sum(saldo) as xsaldo from almacen where [producto]='" & "" & mytablex.Fields("producto") & "' group by producto ", cn, adOpenStatic, adLockOptimistic

        If mytabley.RecordCount > 0 Then
            xsaldo = Val("" & mytabley.Fields("xsaldo"))

        End If

        mytabley.Close
   
        If Val(xfactor) <= 0 Then
            xfactor = "1"

        End If

        buf = xunidad
        found = formateaa(buf, 5, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = xfactor
        found = formateaa(buf, 4, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = calcula_saldo(Val("" & mytablex.Fields("xcanti")), Val(xfactor))
        found = formateaa(buf, 7, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = calcula_saldo(xsaldo, Val(xfactor))
        found = formateaa(buf, 7, 0, 1)
        found = formateaa("", 1, 0, 0)
   
        psdx = 0

        If xsaldo > Val("" & mytablex.Fields("xcanti")) * Val(xfactor) Then
            psdx = 0

        End If

        If xsaldo < Val("" & mytablex.Fields("xcanti")) * Val(xfactor) Then
            psdx = Val("" & mytablex.Fields("xcanti")) * Val(xfactor) - xsaldo

        End If
   
        buf = calcula_saldo(psdx, Val(xfactor))
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 2, 0)
        nlineas
        suma1 = suma1 + Val("" & mytablex.Fields("xcanti")) * Val(xfactor)
        suma2 = suma2 + xsaldo
        suma3 = suma3 + psdx
      
        suma4 = suma4 + Val("" & mytablex.Fields("xcanti")) * Val(xfactor)
        suma5 = suma5 + xsaldo
        suma6 = suma6 + psdx
      
seguy14:
        mytablex.MoveNext
    Loop

    found = formateaa("", 83, 0, 0)
    buf = Format(suma1, "0.00")
    found = formateaa(buf, 7, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(suma2, "0.00")
    found = formateaa(buf, 7, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(suma3, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)
    nlineas
    found = formateaa("", 83, 0, 0)
    buf = Format(suma4, "0.00")
    found = formateaa(buf, 7, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(suma5, "0.00")
    found = formateaa(buf, 7, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(suma6, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)
   
End Sub

Sub cabecera_producto3()

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
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Fechai : " & fechai, 25, 2, 0)
    found = formateaa("Fechaf : " & fechaf, 25, 2, 0)
    
    buf = String(152, "-")
    found = formateaa(buf, 152, 2, 0)
    buf = "Producto"
    found = formateaa(buf, 11, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Descripcio"
    found = formateaa(buf, 59, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Unida"
    found = formateaa(buf, 5, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "fact"
    found = formateaa(buf, 4, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Pedido "
    found = formateaa(buf, 7, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = "Stock "
    found = formateaa(buf, 7, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = "Falta "
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)
      
    buf = String(152, "-")
    found = formateaa(buf, 152, 2, 0)

End Sub

Sub proceso_impresion73()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset

    Dim buf      As String

    contpag = 0
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    found = proceso_sql()

    If found = 0 Then
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera73
    cuerpo_programa73
    '------------------------------------
    Close #1
    cerrar_archivo
    mysnap.Close
     
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1
    '---------------------------------

End Sub

Function proceso_sql()

    Dim buf1 As String

    Dim buf  As String

    Dim buf2 As String

    On Error GoTo cmd299_err

    buf2 = " sum(cantidad) as TCANTIDAD "

    If opcion2 = 90 Then  'ventas por fecha,categoria+subfamilia
        buf = "select familia,categoria,subfamilia,moneda," & buf2 & ",sum(total) as TTOTAL from detalle where "

    End If

    If opcion2 = 91 Then  'ventas por fecha,categoria+subfamilia
        buf = "select producto,categoria,subfamilia,moneda," & buf2 & ",sum(total) as TTOTAL from detalle where "

    End If

    If opcion2 = 92 Then  'ventas por fecha,categoria+subfamilia
        buf = "select usuario,categoria,subfamilia,moneda," & buf2 & ",sum(total) as TTOTAL from detalle where "

    End If

    If opcion2 = 93 Then  'ventas por fecha,categoria+subfamilia
        buf = "select tipo,categoria,subfamilia,moneda," & buf2 & ",sum(total) as TTOTAL from detalle where "

    End If

    If opcion2 = 94 Then  'ventas por fecha,categoria+subfamilia
        buf = "select turno,caja,usuario,sentido,categoria,subfamilia,moneda," & buf2 & ",sum(total) as TTOTAL from detalle where "

    End If

    If opcion2 = 73 Then  'ventas por fecha,categoria+subfamilia
        buf = "select fecha,categoria,subfamilia,moneda," & buf2 & ",sum(total) as TTOTAL from detalle where "

    End If

    If opcion2 = 80 Then  'ventas por caseta,categoria+subfamilia
        buf = "select caja,categoria,subfamilia,moneda," & buf2 & ",sum(total) as TTOTAL from detalle where "

    End If

    If opcion2 = 74 Then  'ventas por hora,categoria+subfamilia
        buf = "select MID$(hora,1,2) as horax,categoria,subfamilia,moneda," & buf2 & ",sum(total) as TTOTAL from detalle WHERE "

    End If

    If opcion2 = 75 Then  'ventas por hora,categoria+subfamilia
        buf = "select month(fecha) as mes,categoria,subfamilia,moneda," & buf2 & ",sum(total) as TTOTAL from detalle where "

    End If

    If opcion2 = 67 Then  'ventas por cajero,subfamilias
        buf = "select turno,caja,usuario,categoria,subfamilia,moneda," & buf2 & ",sum(total) as TTOTAL from detalle where "

    End If
    
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If local1 <> "%" Then
        buf = buf & " and local='" & extra_loquesea(local1) & "'"

    End If

    If tipo <> "%" Then
        buf = buf & " and tipo like '" & tipo & "'"

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

    If usuario <> "%" Then
        buf = buf & " and usuario like '" & usuario & "'"

    End If

    If bodega <> "%" Then
        buf = buf & " and bodega like '" & extra_loquesea(bodega) & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & caja & "'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & turno & "'"

    End If

    If producto <> "%" Then
        buf = buf & " and [producto] like '" & producto & "'"

    End If

    If descripcio <> "%" Then
        buf = buf & " and [descripcio] like '" & descripcio & "'"

    End If

    If moneda <> "%" Then
        buf = buf & " and moneda like '" & moneda & "'"

    End If

    If ccosto <> "%" Then
        buf = buf & " and ccosto like '" & ccosto & "'"

    End If

    If familia <> "%" Then
        buf = buf & " and familia like '" & familia & "'"

    End If

    If subfamilia <> "%" Then
        buf = buf & " and subfamilia like '" & subfamilia & "'"

    End If

    If marca <> "%" Then
        buf = buf & " and marca like '" & marca & "'"

    End If

    If unidad <> "%" Then
        buf = buf & " and unidad like '" & unidad & "'"

    End If

    If vendedor <> "%" Then
        buf = buf & " and vendedor like '" & vendedor & "'"

    End If

    If acu = "V" Then
        buf = buf & " and (acu='1' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') "

    End If

    If acu = "C" Then
        buf = buf & " and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') "

    End If

    If acu <> "C" And acu <> "V" Then
        buf = buf & " and acu='" & acu & "'"

    End If

    If estado <> "%" Then
        buf = buf & " and estado='" & estado & "'"

    End If

    If opcion2 = 73 Then  'ventas por fecha+subfamilias+categoria
        buf = buf & " group by fecha,categoria,subfamilia,moneda order by fecha"

    End If

    If opcion2 = 90 Then  'ventas por fecha+subfamilias+categoria
        buf = buf & " group by FAMILIA,categoria,subfamilia,moneda order by familia"

    End If

    If opcion2 = 91 Then  'ventas por fecha+subfamilias+categoria
        buf = buf & " group by producto,categoria,subfamilia,moneda order by producto"

    End If

    If opcion2 = 92 Then  'ventas por fecha+subfamilias+categoria
        buf = buf & " group by usuario,categoria,subfamilia,moneda order by usuario"

    End If

    If opcion2 = 93 Then  'ventas por fecha+subfamilias+categoria
        buf = buf & " group by tipo,categoria,subfamilia,moneda order by tipo"

    End If

    If opcion2 = 94 Then  'ventas por fecha+subfamilias+categoria
        buf = buf & " group by turno,caja,usuario,sentido,categoria,subfamilia,moneda order by caja,turno,usuario,sentido"

    End If

    If opcion2 = 80 Then  'ventas por caseta+subfamilias+categoria
        buf = buf & " group by caja,categoria,subfamilia,moneda order by caja"

    End If

    If opcion2 = 74 Then  'ventas por fecha+subfamilias+categoria
        buf = buf & " group by MID$(hora,1,2),categoria,subfamilia,moneda order by MID$(hora,1,2)"

    End If

    If opcion2 = 75 Then  'ventas por fecha+subfamilias+categoria
        buf = buf & " group by month(fecha),categoria,subfamilia,moneda order by month(fecha)"

    End If

    If opcion2 = 67 Then  'ventas por turno,caja,usuario+subfamilias+
        buf = buf & " group by turno,caja,usuario,categoria,subfamilia,moneda order by turno,caja,usuario "

    End If

    'MsgBox buf
    If mysnap.State = 1 Then mysnap.Close
    mysnap.Open buf, cn, adOpenStatic, adLockOptimistic
    
    proceso_sql = 1
    Exit Function
cmd299_err:
    MsgBox "Mensaje, Error en Sql .." & error$
    Exit Function

End Function

Sub cabecera73()

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
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Fechai : " & fechai, 25, 2, 0)
    found = formateaa("Fechaf : " & fechaf, 25, 2, 0)
    
    buf = String(197, "-")
    found = formateaa(buf, 197, 2, 0)
    found = formateaa("FECHA ", 11, 0, 0)
    found = formateaa("    LIVIANOS    ", 18, 0, 0)
    found = formateaa("                                 VEHICULOS PESADOS ", 112, 2, 0)

    found = formateaa("", 11, 0, 0)
    found = formateaa("                ", 18, 0, 0)
    found = formateaa("----2 EJES-----", 18, 0, 0)
    found = formateaa("----3 EJES-----", 18, 0, 0)
    found = formateaa("----4 EJES-----", 18, 0, 0)
    found = formateaa("----5 EJES-----", 18, 0, 0)
    found = formateaa("----6 EJES-----", 18, 0, 0)
    found = formateaa("----7 EJES-----", 18, 0, 0)

    found = formateaa("-TOTAL PESADOS-", 18, 0, 0)
    found = formateaa("-TOTAL GENERAL-", 18, 2, 0)

    found = formateaa("", 11, 0, 0)
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1)  'LIVIANOS
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '3 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '4EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '5 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '6 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '7 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) ' 8 EJES

    found = formateaa("RECAUDA ", 9, 0, 1) 'TOTAL PESADO
    found = formateaa("FLUJO ", 9, 0, 1) ' TOTAL PESADO

    found = formateaa("TOT/REC ", 9, 0, 1)
    found = formateaa("TOT/FLU ", 9, 2, 1)

    buf = String(197, "-")
    found = formateaa(buf, 197, 2, 0)
    
End Sub

Sub cuerpo_programa73()

    Dim Tmp   As String

    Dim tmp1  As String

    Dim buf1  As String

    Dim vr    As Integer

    Dim sw    As Integer

    Dim I     As Integer

    Dim j     As Integer

    Dim found As Integer

    On Error GoTo cmd99165_err

    sw = 0
    inicializa_recaudo
    Do

        If mysnap.EOF Then Exit Do
        'vr = DoEvents()
        tmp1 = "" & mysnap.Fields("fecha")

        If sw = 0 Then
            buf1 = "" & mysnap.Fields("fecha")
            found = formateaa(buf1, 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            Tmp = "" & mysnap.Fields(0)
            sw = 1

        End If

        If Tmp <> tmp1 Then

            '----------------------
            For I = 1 To 8
                buf1 = Format(recaudo(I), "0.00")

                If Val(buf1) = 0 Then buf1 = ""
                found = formateaa(buf1, 8, 0, 1)
                found = formateaa("", 1, 0, 0)
                buf1 = Format(flujo(I), "0")

                If Val(buf1) = 0 Then buf1 = ""
                found = formateaa(buf1, 8, 0, 1)
                found = formateaa("", 1, 0, 0)
            Next I

            buf1 = Format(recaudo(10), "0.00")

            If Val(buf1) = 0 Then buf1 = ""
            found = formateaa(buf1, 8, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf1 = Format(flujo(11), "0")

            If Val(buf1) = 0 Then buf1 = ""
            found = formateaa(buf1, 8, 0, 1)
            found = formateaa("", 1, 0, 0)
                 
            buf1 = Format(tflujo(11), "0")

            If Val(buf1) = 0 Then buf1 = ""
            found = formateaa(buf1, 8, 0, 1)
            found = formateaa("", 1, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas

            For I = 1 To 11
                flujo(I) = 0
                recaudo(I) = 0
            Next I

            buf1 = "" & mysnap.Fields("fecha")
            found = formateaa(buf1, 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            Tmp = "" & mysnap.Fields("fecha")

        End If

        '---------------
        imprime_detalle73
        '--------------
        mysnap.MoveNext
    Loop
             
    For I = 1 To 8
        buf1 = Format(recaudo(I), "0.00")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf1 = Format(flujo(I), "0")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
    Next I

    buf1 = Format(recaudo(10), "0.00")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf1 = Format(flujo(11), "0")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)

    found = formateaa("", 1, 2, 0)
    nlineas
    found = formateaa("", 11, 0, 0)

    For I = 1 To 8
        buf1 = Format(trecaudo(I), "0.00")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf1 = Format(tflujo(I), "0")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
    Next I

    buf1 = Format(trecaudo(10), "0.00")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf1 = Format(tflujo(11), "0")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)

    found = formateaa("", 1, 2, 0)
    Exit Sub
cmd99165_err:
    MsgBox "Error en Cuerpo Programa-30 " & error$, 24, "Aviso"
    Exit Sub

End Sub

Sub imprime_detalle73()

    If Val("" & mysnap.Fields("CATEGORIA")) = 1 Then  'LIVIANO
        recaudo(10) = recaudo(10) + Val("" & mysnap.Fields("ttotal"))
        trecaudo(10) = trecaudo(10) + Val(mysnap.Fields("ttotal"))
        recaudo(1) = recaudo(1) + Val(mysnap.Fields("ttotal"))

        trecaudo(1) = trecaudo(1) + Val(mysnap.Fields("ttotal"))
        flujo(1) = flujo(1) + Val("" & mysnap.Fields("tcantidad"))
        tflujo(1) = tflujo(1) + Val("" & mysnap.Fields("tcantidad"))

        flujo(11) = flujo(11) + Val("" & mysnap.Fields("tcantidad"))
        tflujo(11) = tflujo(11) + Val("" & mysnap.Fields("tcantidad"))

    End If

    If Val("" & mysnap.Fields("CATEGORIA")) = 2 Then   'PESADO
        If Val("" & mysnap.Fields("subfamilia")) = 2 Then
            recaudo(10) = recaudo(10) + Val("" & mysnap.Fields("ttotal"))
            trecaudo(10) = trecaudo(10) + Val("" & mysnap.Fields("ttotal"))
            recaudo(2) = recaudo(2) + Val("" & mysnap.Fields("ttotal"))
            trecaudo(2) = trecaudo(2) + Val("" & mysnap.Fields("ttotal"))

            recaudo(8) = recaudo(8) + Val("" & mysnap.Fields("ttotal"))
            trecaudo(8) = trecaudo(8) + Val("" & mysnap.Fields("ttotal"))
            tflujo(8) = tflujo(8) + Val("" & mysnap.Fields("tcantidad"))
            flujo(8) = flujo(8) + Val("" & mysnap.Fields("tcantidad"))

            tflujo(2) = tflujo(2) + Val("" & mysnap.Fields("tcantidad"))
            flujo(2) = flujo(2) + Val("" & mysnap.Fields("tcantidad"))
            flujo(11) = flujo(11) + Val("" & mysnap.Fields("tcantidad"))
            tflujo(11) = tflujo(11) + Val("" & mysnap.Fields("tcantidad"))
            flujo_ejes(11) = flujo_ejes(11) + Val("" & mysnap.Fields("tcantidad")) * 2

        End If

        If Val("" & mysnap.Fields("subfamilia")) = 3 Then
            recaudo(10) = recaudo(10) + Val("" & mysnap.Fields("ttotal"))
            trecaudo(10) = trecaudo(10) + Val("" & mysnap.Fields("ttotal"))
            recaudo(3) = recaudo(3) + Val("" & mysnap.Fields("ttotal"))
            trecaudo(3) = trecaudo(3) + Val("" & mysnap.Fields("ttotal"))
            recaudo(8) = recaudo(8) + Val("" & mysnap.Fields("ttotal"))
            trecaudo(8) = trecaudo(8) + Val("" & mysnap.Fields("ttotal"))
            tflujo(8) = tflujo(8) + Val("" & mysnap.Fields("tcantidad"))
            flujo(8) = flujo(8) + Val("" & mysnap.Fields("tcantidad"))
            tflujo(3) = tflujo(3) + Val("" & mysnap.Fields("tcantidad"))
            flujo(3) = flujo(3) + Val("" & mysnap.Fields("tcantidad"))
            flujo(11) = flujo(11) + Val("" & mysnap.Fields("tcantidad"))
            tflujo(11) = tflujo(11) + Val("" & mysnap.Fields("tcantidad"))
            flujo_ejes(11) = flujo_ejes(11) + Val("" & mysnap.Fields("tcantidad")) * 3

        End If

        If Val("" & mysnap.Fields("subfamilia")) = 4 Then
            recaudo(10) = recaudo(10) + Val("" & mysnap.Fields("ttotal"))
            trecaudo(10) = trecaudo(10) + Val("" & mysnap.Fields("ttotal"))
            recaudo(4) = recaudo(4) + Val("" & mysnap.Fields("ttotal"))
            trecaudo(4) = trecaudo(4) + Val("" & mysnap.Fields("ttotal"))
            recaudo(8) = recaudo(8) + Val("" & mysnap.Fields("ttotal"))
            trecaudo(8) = trecaudo(8) + Val("" & mysnap.Fields("ttotal"))
            tflujo(8) = tflujo(8) + Val("" & mysnap.Fields("tcantidad"))
            flujo(8) = flujo(8) + Val("" & mysnap.Fields("tcantidad"))
            tflujo(4) = tflujo(4) + Val("" & mysnap.Fields("tcantidad"))
            flujo(4) = flujo(4) + Val("" & mysnap.Fields("tcantidad"))
            flujo(11) = flujo(11) + Val("" & mysnap.Fields("tcantidad"))
            tflujo(11) = tflujo(11) + Val("" & mysnap.Fields("tcantidad"))
            flujo_ejes(11) = flujo_ejes(11) + Val("" & mysnap.Fields("tcantidad")) * 4

        End If

        If Val("" & mysnap.Fields("subfamilia")) = 5 Then
            flujo(11) = flujo(11) + Val("" & mysnap.Fields("tcantidad"))
            tflujo(11) = tflujo(11) + Val("" & mysnap.Fields("tcantidad"))
            flujo_ejes(11) = flujo_ejes(11) + Val("" & mysnap.Fields("tcantidad")) * 5
            recaudo(10) = recaudo(10) + Val("" & mysnap.Fields("ttotal"))
            trecaudo(10) = trecaudo(10) + Val("" & mysnap.Fields("ttotal"))
            recaudo(5) = recaudo(5) + Val("" & mysnap.Fields("ttotal"))
            trecaudo(5) = trecaudo(5) + Val("" & mysnap.Fields("ttotal"))
            recaudo(8) = recaudo(8) + Val("" & mysnap.Fields("ttotal"))
            trecaudo(8) = trecaudo(8) + Val("" & mysnap.Fields("ttotal"))
            tflujo(8) = tflujo(8) + Val("" & mysnap.Fields("tcantidad"))
            flujo(8) = flujo(8) + Val("" & mysnap.Fields("tcantidad"))
            tflujo(5) = tflujo(5) + Val("" & mysnap.Fields("tcantidad"))
            flujo(5) = flujo(5) + Val("" & mysnap.Fields("tcantidad"))

        End If

        If Val("" & mysnap.Fields("subfamilia")) = 6 Then
            flujo(11) = flujo(11) + Val("" & mysnap.Fields("tcantidad"))
            tflujo(11) = tflujo(11) + Val("" & mysnap.Fields("tcantidad"))
            flujo_ejes(11) = flujo_ejes(11) + Val("" & mysnap.Fields("tcantidad")) * 6
            recaudo(10) = recaudo(10) + Val("" & mysnap.Fields("ttotal"))
            trecaudo(10) = trecaudo(10) + Val("" & mysnap.Fields("ttotal"))
            recaudo(6) = recaudo(6) + Val("" & mysnap.Fields("ttotal"))
            trecaudo(6) = trecaudo(6) + Val("" & mysnap.Fields("ttotal"))
            recaudo(8) = recaudo(8) + Val("" & mysnap.Fields("ttotal"))
            trecaudo(8) = trecaudo(8) + Val("" & mysnap.Fields("ttotal"))
            tflujo(8) = tflujo(8) + Val("" & mysnap.Fields("tcantidad"))
            flujo(8) = flujo(8) + Val("" & mysnap.Fields("tcantidad"))
            tflujo(6) = tflujo(6) + Val("" & mysnap.Fields("tcantidad"))
            flujo(6) = flujo(6) + Val("" & mysnap.Fields("tcantidad"))

        End If

        If Val("" & mysnap.Fields("subfamilia")) = 7 Then
            flujo(11) = flujo(11) + Val("" & mysnap.Fields("tcantidad"))
            tflujo(11) = tflujo(11) + Val("" & mysnap.Fields("tcantidad"))
            flujo_ejes(11) = flujo_ejes(11) + Val("" & mysnap.Fields("tcantidad")) * 7
            recaudo(10) = recaudo(10) + Val("" & mysnap.Fields("ttotal"))
            trecaudo(10) = trecaudo(10) + Val("" & mysnap.Fields("ttotal"))
            recaudo(7) = recaudo(7) + Val("" & mysnap.Fields("ttotal"))
            trecaudo(7) = trecaudo(7) + Val("" & mysnap.Fields("ttotal"))
            recaudo(8) = recaudo(8) + Val("" & mysnap.Fields("ttotal"))
            trecaudo(8) = trecaudo(8) + Val("" & mysnap.Fields("ttotal"))
            tflujo(8) = tflujo(8) + Val("" & mysnap.Fields("tcantidad"))
            flujo(8) = flujo(8) + Val("" & mysnap.Fields("tcantidad"))
            tflujo(7) = tflujo(7) + Val("" & mysnap.Fields("tcantidad"))
            flujo(7) = flujo(7) + Val("" & mysnap.Fields("tcantidad"))

        End If

    End If

End Sub

Sub inicializa_recaudo()

    Dim I As Integer

    For I = 1 To 11
        flujo(I) = 0
        tflujo(I) = 0
        recaudo(I) = 0
        trecaudo(I) = 0
        flujo_ejes(I) = 0
    Next I

End Sub

Sub proceso_impresion74()

    Dim found As Integer

    Dim buf   As String

    contpag = 0
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    found = proceso_sql()

    If found = 0 Then
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera74
    CUERPO_PROGRAMA74
    '------------------------------------
    Close #1
    cerrar_archivo
    mysnap.Close
     
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1
    '---------------------------------

End Sub

Sub cabecera74()

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
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Fechai : " & fechai, 25, 2, 0)
    found = formateaa("Fechaf : " & fechaf, 25, 2, 0)
    buf = String(197, "-")
    found = formateaa(buf, 197, 2, 0)
    found = formateaa("HORA  ", 11, 0, 0)
    found = formateaa("    LIVIANOS    ", 18, 0, 0)
    found = formateaa("                                 VEHICULOS PESADOS ", 112, 2, 0)

    found = formateaa("", 11, 0, 0)
    found = formateaa("                ", 18, 0, 0)
    found = formateaa("----2 EJES-----", 18, 0, 0)
    found = formateaa("----3 EJES-----", 18, 0, 0)
    found = formateaa("----4 EJES-----", 18, 0, 0)
    found = formateaa("----5 EJES-----", 18, 0, 0)
    found = formateaa("----6 EJES-----", 18, 0, 0)
    found = formateaa("----7 EJES-----", 18, 0, 0)
    found = formateaa("-TOTAL PESADOS-", 18, 0, 0)
    found = formateaa("-TOTAL GENERAL-", 18, 2, 0)

    found = formateaa("", 11, 0, 0)
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1)  'LIVIANOS
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '3 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '4EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '5 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '6 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '7 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) ' 8 EJES

    found = formateaa("RECAUDA ", 9, 0, 1) 'TOTAL PESADO
    found = formateaa("FLUJO ", 9, 0, 1) ' TOTAL PESADO

    found = formateaa("TOT/REC ", 9, 0, 1)
    found = formateaa("TOT/FLU ", 9, 2, 1)

    buf = String(197, "-")
    found = formateaa(buf, 197, 2, 0)
    
End Sub

Sub CUERPO_PROGRAMA74()

    Dim Tmp   As String

    Dim tmp1  As String

    Dim buf1  As String

    Dim vr    As Integer

    Dim sw    As Integer

    Dim I     As Integer

    Dim j     As Integer

    Dim found As Integer

    On Error GoTo cmd9916599_err

    sw = 0
    inicializa_recaudo

    Do Until mysnap.EOF
        vr = DoEvents()
          
        tmp1 = "" & mysnap.Fields("horax")

        If sw = 0 Then
            buf1 = "" & mysnap.Fields("horax")
            found = formateaa(buf1, 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            Tmp = "" & mysnap.Fields("horax")
            sw = 1

        End If

        If Tmp <> tmp1 Then

            '----------------------
            For I = 1 To 8
                buf1 = Format(recaudo(I), "0.00")

                If Val(buf1) = 0 Then buf1 = ""
                found = formateaa(buf1, 8, 0, 1)
                found = formateaa("", 1, 0, 0)
                buf1 = Format(flujo(I), "0")

                If Val(buf1) = 0 Then buf1 = ""
                found = formateaa(buf1, 8, 0, 1)
                found = formateaa("", 1, 0, 0)
            Next I

            buf1 = Format(recaudo(10), "0.00")

            If Val(buf1) = 0 Then buf1 = ""
            found = formateaa(buf1, 8, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf1 = Format(flujo(11), "0")

            If Val(buf1) = 0 Then buf1 = ""
            found = formateaa(buf1, 8, 0, 1)
            found = formateaa("", 1, 0, 0)

            found = formateaa("", 1, 2, 0)
            nlineas

            For I = 1 To 11
                flujo(I) = 0
                recaudo(I) = 0
            Next I

            buf1 = "" & mysnap.Fields("hora")
            found = formateaa(buf1, 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            Tmp = "" & mysnap.Fields("horax")

        End If

        imprime_detalle73
        '--------------
        mysnap.MoveNext
    Loop

    For I = 1 To 8
        buf1 = Format(recaudo(I), "0.00")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf1 = Format(flujo(I), "0")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
    Next I

    buf1 = Format(recaudo(10), "0.00")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf1 = Format(flujo(11), "0")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)

    found = formateaa("", 1, 2, 0)
    nlineas
    found = formateaa("", 11, 0, 0)

    For I = 1 To 8
        buf1 = Format(trecaudo(I), "0.00")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf1 = Format(tflujo(I), "0")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
    Next I

    buf1 = Format(trecaudo(10), "0.00")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf1 = Format(tflujo(11), "0")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)

    found = formateaa("", 1, 2, 0)
    Exit Sub
cmd9916599_err:
    MsgBox "Error en Cuerpo Programa-30 " & error$, 24, "Aviso"
    Exit Sub

End Sub

Sub proceso_impresion75()

    Dim found As Integer

    Dim buf   As String

    contpag = 0
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    found = proceso_sql()

    If found = 0 Then
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera75
    CUERPO_PROGRAMA75
    '------------------------------------
    Close #1
    cerrar_archivo
    mysnap.Close
     
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1
    '---------------------------------

End Sub

Sub cabecera75()

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
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Fechai : " & fechai, 25, 2, 0)
    found = formateaa("Fechaf : " & fechaf, 25, 2, 0)
    
    buf = String(197, "-")
    found = formateaa(buf, 197, 2, 0)
    found = formateaa("MENSUAL ", 11, 0, 0)
    found = formateaa("    LIVIANOS    ", 18, 0, 0)
    found = formateaa("                                 VEHICULOS PESADOS ", 112, 2, 0)

    found = formateaa("", 11, 0, 0)
    found = formateaa("                ", 18, 0, 0)
    found = formateaa("----2 EJES-----", 18, 0, 0)
    found = formateaa("----3 EJES-----", 18, 0, 0)
    found = formateaa("----4 EJES-----", 18, 0, 0)
    found = formateaa("----5 EJES-----", 18, 0, 0)
    found = formateaa("----6 EJES-----", 18, 0, 0)
    found = formateaa("----7 EJES-----", 18, 0, 0)
    found = formateaa("-TOTAL PESADOS-", 18, 0, 0)
    found = formateaa("-TOTAL GENERAL-", 18, 2, 0)

    found = formateaa("", 11, 0, 0)
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1)  'LIVIANOS
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '3 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '4EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '5 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '6 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '7 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) ' 8 EJES

    found = formateaa("RECAUDA ", 9, 0, 1) 'TOTAL PESADO
    found = formateaa("FLUJO ", 9, 0, 1) ' TOTAL PESADO

    found = formateaa("TOT/REC ", 9, 0, 1)
    found = formateaa("TOT/FLU ", 9, 2, 1)

    buf = String(197, "-")
    found = formateaa(buf, 197, 2, 0)
    
End Sub

Sub CUERPO_PROGRAMA75()

    Dim Tmp   As String

    Dim tmp1  As String

    Dim buf1  As String

    Dim vr    As Integer

    Dim sw    As Integer

    Dim I     As Integer

    Dim j     As Integer

    Dim found As Integer

    On Error GoTo cmd99165991_err

    sw = 0
    inicializa_recaudo

    Do Until mysnap.EOF
        vr = DoEvents()
        tmp1 = "" & mysnap.Fields("mes")

        If sw = 0 Then
            buf1 = "" & mysnap.Fields("mes")
            found = formateaa(buf1, 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            Tmp = "" & mysnap.Fields("mes")
            sw = 1

        End If

        If Tmp <> tmp1 Then

            '----------------------
            For I = 1 To 8
                buf1 = Format(recaudo(I), "0.00")

                If Val(buf1) = 0 Then buf1 = ""
                found = formateaa(buf1, 8, 0, 1)
                found = formateaa("", 1, 0, 0)
                buf1 = Format(flujo(I), "0")

                If Val(buf1) = 0 Then buf1 = ""
                found = formateaa(buf1, 8, 0, 1)
                found = formateaa("", 1, 0, 0)
            Next I

            buf1 = Format(recaudo(10), "0.00")

            If Val(buf1) = 0 Then buf1 = ""
            found = formateaa(buf1, 8, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf1 = Format(flujo(11), "0")

            If Val(buf1) = 0 Then buf1 = ""
            found = formateaa(buf1, 8, 0, 1)
            found = formateaa("", 1, 0, 0)

            found = formateaa("", 1, 2, 0)
            nlineas

            For I = 1 To 11
                flujo(I) = 0
                recaudo(I) = 0
            Next I

            buf1 = "" & mysnap.Fields("mes")
            found = formateaa(buf1, 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            Tmp = "" & mysnap.Fields("mes")

        End If

        imprime_detalle73
        '--------------
        mysnap.MoveNext
    Loop

    For I = 1 To 8
        buf1 = Format(recaudo(I), "0.00")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf1 = Format(flujo(I), "0")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
    Next I

    buf1 = Format(recaudo(10), "0.00")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf1 = Format(flujo(11), "0")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)

    found = formateaa("", 1, 2, 0)
    nlineas
    found = formateaa("", 11, 0, 0)

    For I = 1 To 8
        buf1 = Format(trecaudo(I), "0.00")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf1 = Format(tflujo(I), "0")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
    Next I

    buf1 = Format(trecaudo(10), "0.00")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf1 = Format(tflujo(11), "0")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)

    found = formateaa("", 1, 2, 0)
    Exit Sub
cmd99165991_err:

    MsgBox "Error en Cuerpo Programa-30 " & error$, 24, "Aviso"
    Exit Sub

End Sub

'80
Sub proceso_impresion80()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset

    Dim buf      As String

    contpag = 0
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    found = proceso_sql()

    If found = 0 Then
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera80
    CUERPO_PROGRAMA80
    '------------------------------------
    Close #1
    cerrar_archivo
    mysnap.Close
     
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1
    '---------------------------------

End Sub

Sub cabecera80()

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
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Fechai : " & fechai, 25, 2, 0)
    found = formateaa("Fechaf : " & fechaf, 25, 2, 0)
    
    buf = String(197, "-")
    found = formateaa(buf, 197, 2, 0)
    found = formateaa("CAJA  ", 11, 0, 0)
    found = formateaa("    LIVIANOS    ", 18, 0, 0)
    found = formateaa("                                 VEHICULOS PESADOS ", 112, 2, 0)

    found = formateaa("", 11, 0, 0)
    found = formateaa("                ", 18, 0, 0)
    found = formateaa("----2 EJES-----", 18, 0, 0)
    found = formateaa("----3 EJES-----", 18, 0, 0)
    found = formateaa("----4 EJES-----", 18, 0, 0)
    found = formateaa("----5 EJES-----", 18, 0, 0)
    found = formateaa("----6 EJES-----", 18, 0, 0)
    found = formateaa("----7 EJES-----", 18, 0, 0)

    found = formateaa("-TOTAL PESADOS-", 18, 0, 0)
    found = formateaa("-TOTAL GENERAL-", 18, 2, 0)

    found = formateaa("", 11, 0, 0)
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1)  'LIVIANOS
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '3 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '4EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '5 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '6 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '7 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) ' 8 EJES

    found = formateaa("RECAUDA ", 9, 0, 1) 'TOTAL PESADO
    found = formateaa("FLUJO ", 9, 0, 1) ' TOTAL PESADO

    found = formateaa("TOT/REC ", 9, 0, 1)
    found = formateaa("TOT/FLU ", 9, 2, 1)

    buf = String(197, "-")
    found = formateaa(buf, 197, 2, 0)
    
End Sub

Sub CUERPO_PROGRAMA80()

    Dim Tmp   As String

    Dim tmp1  As String

    Dim buf1  As String

    Dim vr    As Integer

    Dim sw    As Integer

    Dim I     As Integer

    Dim j     As Integer

    Dim found As Integer

    On Error GoTo cmd991651_err

    sw = 0
    inicializa_recaudo
    Do

        If mysnap.EOF Then Exit Do
        'vr = DoEvents()
        tmp1 = "" & mysnap.Fields("caja")

        If sw = 0 Then
            buf1 = "" & mysnap.Fields("caja")
            found = formateaa(buf1, 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            Tmp = "" & mysnap.Fields(0)
            sw = 1

        End If

        If Tmp <> tmp1 Then

            '----------------------
            For I = 1 To 8
                buf1 = Format(recaudo(I), "0.00")

                If Val(buf1) = 0 Then buf1 = ""
                found = formateaa(buf1, 8, 0, 1)
                found = formateaa("", 1, 0, 0)
                buf1 = Format(flujo(I), "0")

                If Val(buf1) = 0 Then buf1 = ""
                found = formateaa(buf1, 8, 0, 1)
                found = formateaa("", 1, 0, 0)
            Next I

            buf1 = Format(recaudo(10), "0.00")

            If Val(buf1) = 0 Then buf1 = ""
            found = formateaa(buf1, 8, 0, 1)
            found = formateaa("", 1, 0, 0)

            buf1 = Format(flujo(11), "0")

            If Val(buf1) = 0 Then buf1 = ""
            found = formateaa(buf1, 8, 0, 1)
            found = formateaa("", 1, 0, 0)

            found = formateaa("", 1, 2, 0)
            nlineas

            For I = 1 To 11
                flujo(I) = 0
                recaudo(I) = 0
            Next I

            buf1 = "" & mysnap.Fields("caja")
            found = formateaa(buf1, 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            Tmp = "" & mysnap.Fields("caja")

        End If

        '---------------
        imprime_detalle73
        '--------------
        mysnap.MoveNext
    Loop
             
    For I = 1 To 8
        buf1 = Format(recaudo(I), "0.00")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf1 = Format(flujo(I), "0")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
    Next I

    buf1 = Format(recaudo(10), "0.00")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)

    buf1 = Format(flujo(11), "0")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)

    found = formateaa("", 1, 2, 0)
    nlineas
    found = formateaa("", 11, 0, 0)

    For I = 1 To 8
        buf1 = Format(trecaudo(I), "0.00")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf1 = Format(tflujo(I), "0")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
    Next I

    buf1 = Format(trecaudo(10), "0.00")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf1 = Format(tflujo(11), "0")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)

    found = formateaa("", 1, 2, 0)
    Exit Sub
cmd991651_err:

    MsgBox "Error en Cuerpo Programa-30 " & error$, 24, "Aviso"
    Exit Sub

End Sub

Sub proceso_impresion90()

    Dim found As Integer

    Dim buf   As String

    contpag = 0
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    found = proceso_sql()

    If found = 0 Then
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera90
    CUERPO_PROGRAMA90
    '------------------------------------
    Close #1
    cerrar_archivo
    mysnap.Close
     
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1
    '---------------------------------

End Sub

Sub cabecera90()

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
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Fechai : " & fechai, 25, 2, 0)
    found = formateaa("Fechaf : " & fechaf, 25, 2, 0)
    
    buf = String(197, "-")
    found = formateaa(buf, 197, 2, 0)
    found = formateaa("FAMILIA ", 11, 0, 0)
    found = formateaa("    LIVIANOS    ", 18, 0, 0)
    found = formateaa("                                 VEHICULOS PESADOS ", 112, 2, 0)

    found = formateaa("", 11, 0, 0)
    found = formateaa("                ", 18, 0, 0)
    found = formateaa("----2 EJES-----", 18, 0, 0)
    found = formateaa("----3 EJES-----", 18, 0, 0)
    found = formateaa("----4 EJES-----", 18, 0, 0)
    found = formateaa("----5 EJES-----", 18, 0, 0)
    found = formateaa("----6 EJES-----", 18, 0, 0)
    found = formateaa("----7 EJES-----", 18, 0, 0)
    found = formateaa("-TOTAL PESADOS-", 18, 0, 0)
    found = formateaa("-TOTAL GENERAL-", 18, 2, 0)

    found = formateaa("", 11, 0, 0)
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1)  'LIVIANOS
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '3 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '4EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '5 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '6 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '7 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) ' 8 EJES

    found = formateaa("RECAUDA ", 9, 0, 1) 'TOTAL PESADO
    found = formateaa("FLUJO ", 9, 0, 1) ' TOTAL PESADO

    found = formateaa("TOT/REC ", 9, 0, 1)
    found = formateaa("TOT/FLU ", 9, 2, 1)

    buf = String(197, "-")
    found = formateaa(buf, 197, 2, 0)
    
End Sub

Sub CUERPO_PROGRAMA90()

    Dim Tmp   As String

    Dim tmp1  As String

    Dim buf1  As String

    Dim vr    As Integer

    Dim sw    As Integer

    Dim I     As Integer

    Dim j     As Integer

    Dim found As Integer

    On Error GoTo cmd991652_err

    sw = 0
    inicializa_recaudo
    Do

        If mysnap.EOF Then Exit Do
        tmp1 = "" & mysnap.Fields("familia")

        If sw = 0 Then
            buf1 = "" & mysnap.Fields("familia")
            found = formateaa(buf1, 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            Tmp = "" & mysnap.Fields("familia")
            sw = 1

        End If

        If Tmp <> tmp1 Then

            '----------------------
            For I = 1 To 8
                buf1 = Format(recaudo(I), "0.00")

                If Val(buf1) = 0 Then buf1 = ""
                found = formateaa(buf1, 8, 0, 1)
                found = formateaa("", 1, 0, 0)
                buf1 = Format(flujo(I), "0")

                If Val(buf1) = 0 Then buf1 = ""
                found = formateaa(buf1, 8, 0, 1)
                found = formateaa("", 1, 0, 0)
            Next I

            buf1 = Format(recaudo(10), "0.00")

            If Val(buf1) = 0 Then buf1 = ""
            found = formateaa(buf1, 8, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf1 = Format(flujo(11), "0")

            If Val(buf1) = 0 Then buf1 = ""
            found = formateaa(buf1, 8, 0, 1)
            found = formateaa("", 1, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas

            For I = 1 To 11
                flujo(I) = 0
                recaudo(I) = 0
            Next I

            buf1 = "" & mysnap.Fields("familia")
            found = formateaa(buf1, 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            Tmp = "" & mysnap.Fields("familia")

        End If

        '---------------
        imprime_detalle73
        '--------------
        mysnap.MoveNext
    Loop
             
    For I = 1 To 8
        buf1 = Format(recaudo(I), "0.00")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf1 = Format(flujo(I), "0")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
    Next I

    buf1 = Format(recaudo(10), "0.00")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf1 = Format(flujo(11), "0")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("", 1, 2, 0)
    nlineas
    found = formateaa("", 11, 0, 0)

    For I = 1 To 8
        buf1 = Format(trecaudo(I), "0.00")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf1 = Format(tflujo(I), "0")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
    Next I

    buf1 = Format(trecaudo(10), "0.00")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf1 = Format(tflujo(11), "0")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("", 1, 2, 0)
    Exit Sub
cmd991652_err:
    MsgBox "Error en Cuerpo Programa-90 " & error$, 24, "Aviso"
    Exit Sub

End Sub

'91
Sub proceso_impresion91()

    Dim found As Integer

    Dim buf   As String

    contpag = 0
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    found = proceso_sql()

    If found = 0 Then
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera91
    CUERPO_PROGRAMA91
    '------------------------------------
    Close #1
    cerrar_archivo
    mysnap.Close
     
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1
    '---------------------------------

End Sub

Sub cabecera91()

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
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Fechai : " & fechai, 25, 2, 0)
    found = formateaa("Fechaf : " & fechaf, 25, 2, 0)
    
    buf = String(197, "-")
    found = formateaa(buf, 197, 2, 0)
    found = formateaa("PRODUCTO ", 11, 0, 0)
    found = formateaa("    LIVIANOS    ", 18, 0, 0)
    found = formateaa("                                 VEHICULOS PESADOS ", 112, 2, 0)

    found = formateaa("", 11, 0, 0)
    found = formateaa("                ", 18, 0, 0)
    found = formateaa("----2 EJES-----", 18, 0, 0)
    found = formateaa("----3 EJES-----", 18, 0, 0)
    found = formateaa("----4 EJES-----", 18, 0, 0)
    found = formateaa("----5 EJES-----", 18, 0, 0)
    found = formateaa("----6 EJES-----", 18, 0, 0)
    found = formateaa("----7 EJES-----", 18, 0, 0)
    found = formateaa("-TOTAL PESADOS-", 18, 0, 0)
    found = formateaa("-TOTAL GENERAL-", 18, 2, 0)

    found = formateaa("", 11, 0, 0)
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1)  'LIVIANOS
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '3 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '4EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '5 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '6 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '7 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) ' 8 EJES

    found = formateaa("RECAUDA ", 9, 0, 1) 'TOTAL PESADO
    found = formateaa("FLUJO ", 9, 0, 1) ' TOTAL PESADO

    found = formateaa("TOT/REC ", 9, 0, 1)
    found = formateaa("TOT/FLU ", 9, 2, 1)

    buf = String(197, "-")
    found = formateaa(buf, 197, 2, 0)
    
End Sub

Sub CUERPO_PROGRAMA91()

    Dim Tmp   As String

    Dim tmp1  As String

    Dim buf1  As String

    Dim vr    As Integer

    Dim sw    As Integer

    Dim I     As Integer

    Dim j     As Integer

    Dim found As Integer

    On Error GoTo cmd991653_err

    sw = 0
    inicializa_recaudo
    Do

        If mysnap.EOF Then Exit Do
        'vr = DoEvents()
        tmp1 = "" & mysnap.Fields("producto")

        If sw = 0 Then
            buf1 = "" & mysnap.Fields("producto")
            found = formateaa(buf1, 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            Tmp = "" & mysnap.Fields("producto")
            sw = 1

        End If

        If Tmp <> tmp1 Then

            '----------------------
            For I = 1 To 8
                buf1 = Format(recaudo(I), "0.00")

                If Val(buf1) = 0 Then buf1 = ""
                found = formateaa(buf1, 8, 0, 1)
                found = formateaa("", 1, 0, 0)
                buf1 = Format(flujo(I), "0")

                If Val(buf1) = 0 Then buf1 = ""
                found = formateaa(buf1, 8, 0, 1)
                found = formateaa("", 1, 0, 0)
            Next I

            buf1 = Format(recaudo(10), "0.00")

            If Val(buf1) = 0 Then buf1 = ""
            found = formateaa(buf1, 8, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf1 = Format(flujo(11), "0")

            If Val(buf1) = 0 Then buf1 = ""
            found = formateaa(buf1, 8, 0, 1)
            found = formateaa("", 1, 0, 0)

            found = formateaa("", 1, 2, 0)
            nlineas

            For I = 1 To 11
                flujo(I) = 0
                recaudo(I) = 0
            Next I

            buf1 = "" & mysnap.Fields("producto")
            found = formateaa(buf1, 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            Tmp = "" & mysnap.Fields("producto")

        End If

        '---------------
        imprime_detalle73
        '--------------
        mysnap.MoveNext
    Loop
             
    For I = 1 To 8
        buf1 = Format(recaudo(I), "0.00")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf1 = Format(flujo(I), "0")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
    Next I

    buf1 = Format(recaudo(10), "0.00")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf1 = Format(flujo(11), "0")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)

    found = formateaa("", 1, 2, 0)
    nlineas
    found = formateaa("", 11, 0, 0)

    For I = 1 To 8
        buf1 = Format(trecaudo(I), "0.00")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf1 = Format(tflujo(I), "0")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
    Next I

    buf1 = Format(trecaudo(10), "0.00")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf1 = Format(tflujo(11), "0")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)

    found = formateaa("", 1, 2, 0)
    Exit Sub
cmd991653_err:

    MsgBox "Error en Cuerpo Programa-30 " & error$, 24, "Aviso"
    Exit Sub

End Sub

'92
Sub proceso_impresion92()

    Dim found As Integer

    Dim buf   As String

    contpag = 0
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    found = proceso_sql()

    If found = 0 Then
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera92
    'MsgBox "xx"
    cuerpo_prog92
    '------------------------------------
    Close #1
    cerrar_archivo
    mysnap.Close
     
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1
    '---------------------------------

End Sub

Sub cabecera92()

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
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Fechai : " & fechai, 25, 2, 0)
    found = formateaa("Fechaf : " & fechaf, 25, 2, 0)
    
    buf = String(197, "-")
    found = formateaa(buf, 197, 2, 0)
    found = formateaa("CAJE ", 6, 0, 0)
    found = formateaa("NOMBRE ", 16, 0, 0)

    found = formateaa("    LIVIANOS    ", 18, 0, 0)
    found = formateaa("                                 VEHICULOS PESADOS ", 112, 2, 0)

    found = formateaa("", 22, 0, 0)
    found = formateaa("                ", 18, 0, 0)
    found = formateaa("----2 EJES-----", 18, 0, 0)
    found = formateaa("----3 EJES-----", 18, 0, 0)
    found = formateaa("----4 EJES-----", 18, 0, 0)
    found = formateaa("----5 EJES-----", 18, 0, 0)
    found = formateaa("----6 EJES-----", 18, 0, 0)
    found = formateaa("----7 EJES-----", 18, 0, 0)
    found = formateaa("-TOTAL PESADOS-", 18, 0, 0)
    found = formateaa("-TOTAL GENERAL-", 18, 2, 0)

    found = formateaa("", 11, 0, 0)
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1)  'LIVIANOS
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '3 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '4EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '5 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '6 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '7 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) ' 8 EJES

    found = formateaa("RECAUDA ", 9, 0, 1) 'TOTAL PESADO
    found = formateaa("FLUJO ", 9, 0, 1) ' TOTAL PESADO

    found = formateaa("TOT/REC ", 9, 0, 1)
    found = formateaa("TOT/FLU ", 9, 2, 1)

    buf = String(197, "-")
    found = formateaa(buf, 197, 2, 0)
    
End Sub

Sub cuerpo_prog92()

    Dim Tmp   As String

    Dim tmp1  As String

    Dim buf1  As String

    Dim vr    As Integer

    Dim sw    As Integer

    Dim I     As Integer

    Dim j     As Integer

    Dim found As Integer

    On Error GoTo cmd991654_err

    sw = 0
    inicializa_recaudo
    Do

        If mysnap.EOF Then Exit Do
        'vr = DoEvents()
        tmp1 = "" & mysnap.Fields("usuario")

        If sw = 0 Then
            buf1 = "" & mysnap.Fields("usuario")
            found = formateaa(buf1, 5, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf1 = busca_cajero("" & mysnap.Fields("usuario"))
            found = formateaa(buf1, 15, 0, 0)
            found = formateaa("", 1, 0, 0)
            Tmp = "" & mysnap.Fields("usuario")
            sw = 1

        End If

        If Tmp <> tmp1 Then

            '----------------------
            For I = 1 To 8
                buf1 = Format(recaudo(I), "0.00")

                If Val(buf1) = 0 Then buf1 = ""
                found = formateaa(buf1, 8, 0, 1)
                found = formateaa("", 1, 0, 0)
                buf1 = Format(flujo(I), "0")

                If Val(buf1) = 0 Then buf1 = ""
                found = formateaa(buf1, 8, 0, 1)
                found = formateaa("", 1, 0, 0)
            Next I

            buf1 = Format(recaudo(10), "0.00")

            If Val(buf1) = 0 Then buf1 = ""
            found = formateaa(buf1, 8, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf1 = Format(flujo(11), "0")

            If Val(buf1) = 0 Then buf1 = ""
            found = formateaa(buf1, 8, 0, 1)
            found = formateaa("", 1, 0, 0)

            found = formateaa("", 1, 2, 0)
            nlineas

            For I = 1 To 11
                flujo(I) = 0
                recaudo(I) = 0
            Next I

            buf1 = "" & mysnap.Fields("usuario")
            found = formateaa(buf1, 5, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf1 = busca_cajero("" & mysnap.Fields("usuario"))
            found = formateaa(buf1, 15, 0, 0)
            found = formateaa("", 1, 0, 0)
            Tmp = "" & mysnap.Fields("usuario")

        End If

        '---------------
        imprime_detalle73
        '--------------
        mysnap.MoveNext
    Loop
             
    For I = 1 To 8
        buf1 = Format(recaudo(I), "0.00")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf1 = Format(flujo(I), "0")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
    Next I

    buf1 = Format(recaudo(10), "0.00")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf1 = Format(flujo(11), "0")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)

    found = formateaa("", 1, 2, 0)
    nlineas
    found = formateaa("", 22, 0, 0)

    For I = 1 To 8
        buf1 = Format(trecaudo(I), "0.00")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf1 = Format(tflujo(I), "0")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
    Next I

    buf1 = Format(trecaudo(10), "0.00")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf1 = Format(tflujo(11), "0")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)

    found = formateaa("", 1, 2, 0)
    Exit Sub
cmd991654_err:
    MsgBox "Error en Cuerpo Programa-30 " & error$, 24, "Aviso"
    Exit Sub

End Sub

Function busca_cajero(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from vendedor where codigo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_cajero = "" & mytablex.Fields("nombre")

    End If

    mytablex.Close

End Function

'93
Sub proceso_impresion93()

    Dim found As Integer

    Dim buf   As String

    contpag = 0
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    found = proceso_sql()

    If found = 0 Then
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera93
    cuerpo_prog93
    '------------------------------------
    Close #1
    cerrar_archivo
    mysnap.Close
     
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1
    '---------------------------------

End Sub

Sub cabecera93()

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
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Fechai : " & fechai, 25, 2, 0)
    found = formateaa("Fechaf : " & fechaf, 25, 2, 0)
    buf = String(197, "-")
    found = formateaa(buf, 197, 2, 0)
    found = formateaa("CLIENTE ", 20, 0, 0)
    found = formateaa("    LIVIANOS    ", 18, 0, 0)
    found = formateaa("                                 VEHICULOS PESADOS ", 112, 2, 0)

    found = formateaa("", 20, 0, 0)
    found = formateaa("                ", 18, 0, 0)
    found = formateaa("----2 EJES-----", 18, 0, 0)
    found = formateaa("----3 EJES-----", 18, 0, 0)
    found = formateaa("----4 EJES-----", 18, 0, 0)
    found = formateaa("----5 EJES-----", 18, 0, 0)
    found = formateaa("----6 EJES-----", 18, 0, 0)
    found = formateaa("----7 EJES-----", 18, 0, 0)

    found = formateaa("-TOTAL PESADOS-", 18, 0, 0)
    found = formateaa("-TOTAL GENERAL-", 18, 2, 0)

    found = formateaa("", 20, 0, 0)
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1)  'LIVIANOS
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '3 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '4EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '5 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '6 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '7 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) ' 8 EJES

    found = formateaa("RECAUDA ", 9, 0, 1) 'TOTAL PESADO
    found = formateaa("FLUJO ", 9, 0, 1) ' TOTAL PESADO

    found = formateaa("TOT/REC ", 9, 0, 1)
    found = formateaa("TOT/FLU ", 9, 2, 1)

    buf = String(197, "-")
    found = formateaa(buf, 197, 2, 0)
    
End Sub

Sub cuerpo_prog93()

    Dim Tmp   As String

    Dim tmp1  As String

    Dim buf1  As String

    Dim vr    As Integer

    Dim sw    As Integer

    Dim I     As Integer

    Dim j     As Integer

    Dim found As Integer

    On Error GoTo cmd9916512_err

    sw = 0
    inicializa_recaudo
    Do

        If mysnap.EOF Then Exit Do
        'vr = DoEvents()
        tmp1 = "" & mysnap.Fields("tipo")

        If sw = 0 Then
            buf1 = busca_tipo("" & mysnap.Fields("tipo"))
            found = formateaa(buf1, 19, 0, 0)
            found = formateaa("", 1, 0, 0)
            Tmp = "" & mysnap.Fields("tipo")
            sw = 1

        End If

        If Tmp <> tmp1 Then

            '----------------------
            For I = 1 To 8
                buf1 = Format(recaudo(I), "0.00")

                If Val(buf1) = 0 Then buf1 = ""
                found = formateaa(buf1, 8, 0, 1)
                found = formateaa("", 1, 0, 0)
                buf1 = Format(flujo(I), "0")

                If Val(buf1) = 0 Then buf1 = ""
                found = formateaa(buf1, 8, 0, 1)
                found = formateaa("", 1, 0, 0)
            Next I

            buf1 = Format(recaudo(10), "0.00")

            If Val(buf1) = 0 Then buf1 = ""
            found = formateaa(buf1, 8, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf1 = Format(flujo(11), "0")

            If Val(buf1) = 0 Then buf1 = ""
            found = formateaa(buf1, 8, 0, 1)
            found = formateaa("", 1, 0, 0)

            found = formateaa("", 1, 2, 0)
            nlineas

            For I = 1 To 11
                flujo(I) = 0
                recaudo(I) = 0
            Next I

            buf1 = busca_tipo("" & mysnap.Fields("tipo"))
            found = formateaa(buf1, 19, 0, 0)
            found = formateaa("", 1, 0, 0)
            Tmp = "" & mysnap.Fields("tipo")

        End If

        '---------------
        imprime_detalle73
        '--------------
        mysnap.MoveNext
    Loop
             
    For I = 1 To 8
        buf1 = Format(recaudo(I), "0.00")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf1 = Format(flujo(I), "0")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
    Next I

    buf1 = Format(recaudo(10), "0.00")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf1 = Format(flujo(11), "0")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)

    found = formateaa("", 1, 2, 0)
    nlineas
    found = formateaa("", 20, 0, 0)

    For I = 1 To 8
        buf1 = Format(trecaudo(I), "0.00")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf1 = Format(tflujo(I), "0")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
    Next I

    buf1 = Format(trecaudo(10), "0.00")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf1 = Format(tflujo(11), "0")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)

    found = formateaa("", 1, 2, 0)
    Exit Sub
cmd9916512_err:

    MsgBox "Error en Cuerpo Programa-30 " & error$, 24, "Aviso"
    Exit Sub

End Sub

Function busca_tipo(buf1 As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from tipo where tipo='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_tipo = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close
    Exit Function

End Function

'67
Sub proceso_impresion67()

    Dim found As Integer

    Dim buf   As String

    contpag = 0
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    found = proceso_sql()

    If found = 0 Then
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera67
    cuerpo_prog67
    '------------------------------------
    Close #1
    cerrar_archivo
    mysnap.Close
     
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1
    '---------------------------------

End Sub

Sub cabecera67()

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
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Fechai : " & fechai, 25, 2, 0)
    found = formateaa("Fechaf : " & fechaf, 25, 2, 0)
    
    buf = String(197, "-")
    found = formateaa(buf, 197, 2, 0)
    found = formateaa("CJ ", 3, 0, 0)
    found = formateaa("T ", 2, 0, 0)
    found = formateaa("CAJE ", 6, 0, 0)
    found = formateaa("NOMBRE ", 16, 0, 0)
    found = formateaa("    LIVIANOS    ", 18, 0, 0)
    found = formateaa("                                 VEHICULOS PESADOS ", 112, 2, 0)

    found = formateaa("", 27, 0, 0)
    found = formateaa("                ", 18, 0, 0)
    found = formateaa("----2 EJES-----", 18, 0, 0)
    found = formateaa("----3 EJES-----", 18, 0, 0)
    found = formateaa("----4 EJES-----", 18, 0, 0)
    found = formateaa("----5 EJES-----", 18, 0, 0)
    found = formateaa("----6 EJES-----", 18, 0, 0)
    found = formateaa("----7 EJES-----", 18, 0, 0)
    found = formateaa("-TOTAL PESADOS-", 18, 0, 0)
    found = formateaa("-TOTAL GENERAL-", 18, 0, 0)
    found = formateaa("PESADO ", 9, 2, 1)

    found = formateaa("", 27, 0, 0)
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1)  'LIVIANOS
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '3 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '4EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '5 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '6 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '7 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) ' 8 EJES

    found = formateaa("RECAUDA ", 9, 0, 1) 'TOTAL PESADO
    found = formateaa("FLUJO ", 9, 0, 1) ' TOTAL PESADO

    found = formateaa("TOT/REC ", 9, 0, 1)
    found = formateaa("TOT/FLU ", 9, 0, 1)

    found = formateaa("NRO/EJE ", 9, 2, 1)

    buf = String(197, "-")
    found = formateaa(buf, 197, 2, 0)
    
End Sub

Sub cuerpo_prog67()

    Dim Tmp   As String

    Dim tmp1  As String

    Dim buf1  As String

    Dim vr    As Integer

    Dim sw    As Integer

    Dim I     As Integer

    Dim j     As Integer

    Dim found As Integer

    On Error GoTo cmd9916567_err

    sw = 0
    inicializa_recaudo

    Do Until mysnap.EOF
        tmp1 = "" & mysnap.Fields("turno") & "" & mysnap.Fields("caja") & "" & mysnap.Fields("usuario")

        If sw = 0 Then
            buf1 = "" & mysnap.Fields("caja")
            found = formateaa(buf1, 2, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf1 = "" & mysnap.Fields("turno")
            found = formateaa(buf1, 1, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf1 = "" & mysnap.Fields("usuario")
            found = formateaa(buf1, 5, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf1 = busca_cajero("" & mysnap.Fields("usuario"))
            found = formateaa(buf1, 15, 0, 0)
            found = formateaa("", 1, 0, 0)
            Tmp = "" & mysnap.Fields("turno") & "" & mysnap.Fields("caja") & "" & mysnap.Fields("usuario")
            sw = 1

        End If

        If Tmp <> tmp1 Then

            '----------------------
            For I = 1 To 8
                buf1 = Format(recaudo(I), "0.00")

                If Val(buf1) = 0 Then buf1 = ""
                found = formateaa(buf1, 8, 0, 1)
                found = formateaa("", 1, 0, 0)
                buf1 = Format(flujo(I), "0")

                If Val(buf1) = 0 Then buf1 = ""
                found = formateaa(buf1, 8, 0, 1)
                found = formateaa("", 1, 0, 0)
            Next I

            buf1 = Format(recaudo(10), "0.00")

            If Val(buf1) = 0 Then buf1 = ""
            found = formateaa(buf1, 8, 0, 1)
            found = formateaa("", 1, 0, 0)

            buf1 = Format(flujo(11), "0")

            If Val(buf1) = 0 Then buf1 = ""
            found = formateaa(buf1, 8, 0, 1)
            found = formateaa("", 1, 0, 0)

            buf1 = Format(flujo_ejes(11), "0")

            If Val(buf1) = 0 Then buf1 = ""
            found = formateaa(buf1, 8, 0, 1)
            found = formateaa("", 1, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas

            For I = 1 To 11
                flujo(I) = 0
                recaudo(I) = 0
                flujo_ejes(I) = 0
            Next I

            buf1 = "" & mysnap.Fields("caja")
            found = formateaa(buf1, 2, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf1 = "" & mysnap.Fields("turno")
            found = formateaa(buf1, 1, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf1 = "" & mysnap.Fields("usuario")
            found = formateaa(buf1, 5, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf1 = busca_cajero("" & mysnap.Fields("usuario"))
            found = formateaa(buf1, 15, 0, 0)
            found = formateaa("", 1, 0, 0)
            Tmp = "" & mysnap.Fields("turno") & "" & mysnap.Fields("caja") & "" & mysnap.Fields("usuario")

        End If

        imprime_detalle73
        '--------------
        mysnap.MoveNext
    Loop

    For I = 1 To 8
        buf1 = Format(recaudo(I), "0.00")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf1 = Format(flujo(I), "0")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
    Next I

    buf1 = Format(recaudo(10), "0.00")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf1 = Format(flujo(11), "0")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf1 = Format(flujo_ejes(11), "0")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("", 1, 2, 0)
    nlineas
    found = formateaa("", 27, 0, 0)

    For I = 1 To 8
        buf1 = Format(trecaudo(I), "0.00")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf1 = Format(tflujo(I), "0")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
    Next I

    buf1 = Format(trecaudo(10), "0.00")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf1 = Format(tflujo(11), "0")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)

    found = formateaa("", 1, 2, 0)
    Exit Sub
cmd9916567_err:
    MsgBox "Error en Cuerpo Programa-30 " & error$, 24, "Aviso"
    Exit Sub

End Sub

'94
Sub proceso_impresion94()

    Dim found As Integer

    Dim buf   As String

    contpag = 0
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    found = proceso_sql()

    If found = 0 Then
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera94
    cuerpo_prog94
    '------------------------------------
    Close #1
    cerrar_archivo
    mysnap.Close
     
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1
    '---------------------------------

End Sub

Sub cabecera94()

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
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Fechai : " & fechai, 25, 2, 0)
    found = formateaa("Fechaf : " & fechaf, 25, 2, 0)
    
    buf = String(197, "-")
    found = formateaa(buf, 197, 2, 0)
    found = formateaa("CJ ", 3, 0, 0)
    found = formateaa("T ", 2, 0, 0)
    found = formateaa("CAJE ", 6, 0, 0)
    found = formateaa("NOMBRE ", 16, 0, 0)
    found = formateaa("S", 2, 0, 0)
    
    found = formateaa("    LIVIANOS    ", 18, 0, 0)
    found = formateaa("                                 VEHICULOS PESADOS ", 112, 2, 0)

    found = formateaa("", 20, 0, 0)
    found = formateaa("                ", 18, 0, 0)
    found = formateaa("----2 EJES-----", 18, 0, 0)
    found = formateaa("----3 EJES-----", 18, 0, 0)
    found = formateaa("----4 EJES-----", 18, 0, 0)
    found = formateaa("----5 EJES-----", 18, 0, 0)
    found = formateaa("----6 EJES-----", 18, 0, 0)
    found = formateaa("----7 EJES-----", 18, 0, 0)

    found = formateaa("-TOTAL PESADOS-", 18, 0, 0)
    found = formateaa("-TOTAL GENERAL-", 18, 2, 0)

    found = formateaa("", 29, 0, 0)
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1)  'LIVIANOS
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '3 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '4EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '5 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '6 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) '7 EJES
    
    found = formateaa("RECAUDA ", 9, 0, 1)
    found = formateaa("FLUJO ", 9, 0, 1) ' 8 EJES

    found = formateaa("RECAUDA ", 9, 0, 1) 'TOTAL PESADO
    found = formateaa("FLUJO ", 9, 0, 1) ' TOTAL PESADO

    found = formateaa("TOT/REC ", 9, 0, 1)
    found = formateaa("TOT/FLU ", 9, 2, 1)

    buf = String(197, "-")
    found = formateaa(buf, 197, 2, 0)
    
End Sub

Sub cuerpo_prog94()

    Dim Tmp   As String

    Dim tmp1  As String

    Dim buf1  As String

    Dim vr    As Integer

    Dim sw    As Integer

    Dim I     As Integer

    Dim j     As Integer

    Dim found As Integer

    On Error GoTo cmd99165121_err

    sw = 0
    inicializa_recaudo
    Do

        If mysnap.EOF Then Exit Do
        'vr = DoEvents()
        tmp1 = "" & mysnap.Fields("turno") & "" & mysnap.Fields("caja") & "" & mysnap.Fields("usuario") & "" & mysnap.Fields("sentido")

        If sw = 0 Then
            buf1 = "" & mysnap.Fields("caja")
            found = formateaa(buf1, 2, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf1 = "" & mysnap.Fields("turno")
            found = formateaa(buf1, 1, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf1 = "" & mysnap.Fields("usuario")
            found = formateaa(buf1, 5, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf1 = busca_cajero("" & mysnap.Fields("usuario"))
            found = formateaa(buf1, 15, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf1 = "" & mysnap.Fields("sentido")
            found = formateaa(buf1, 1, 0, 0)
            found = formateaa("", 1, 0, 0)
            Tmp = "" & mysnap.Fields("turno") & "" & mysnap.Fields("caja") & "" & mysnap.Fields("usuario") & "" & mysnap.Fields("sentido")
            sw = 1

        End If

        If Tmp <> tmp1 Then

            '----------------------
            For I = 1 To 8
                buf1 = Format(recaudo(I), "0.00")

                If Val(buf1) = 0 Then buf1 = ""
                found = formateaa(buf1, 8, 0, 1)
                found = formateaa("", 1, 0, 0)
                buf1 = Format(flujo(I), "0")

                If Val(buf1) = 0 Then buf1 = ""
                found = formateaa(buf1, 8, 0, 1)
                found = formateaa("", 1, 0, 0)
            Next I

            buf1 = Format(recaudo(10), "0.00")

            If Val(buf1) = 0 Then buf1 = ""
            found = formateaa(buf1, 8, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf1 = Format(flujo(11), "0")

            If Val(buf1) = 0 Then buf1 = ""
            found = formateaa(buf1, 8, 0, 1)
            found = formateaa("", 1, 0, 0)

            found = formateaa("", 1, 2, 0)
            nlineas

            For I = 1 To 11
                flujo(I) = 0
                recaudo(I) = 0
            Next I

            buf1 = "" & mysnap.Fields("caja")
            found = formateaa(buf1, 2, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf1 = "" & mysnap.Fields("turno")
            found = formateaa(buf1, 1, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf1 = "" & mysnap.Fields("usuario")
            found = formateaa(buf1, 5, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf1 = busca_cajero("" & mysnap.Fields("usuario"))
            found = formateaa(buf1, 15, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf1 = "" & mysnap.Fields("sentido")
            found = formateaa(buf1, 1, 0, 0)
            found = formateaa("", 1, 0, 0)
            Tmp = "" & mysnap.Fields("turno") & "" & mysnap.Fields("caja") & "" & mysnap.Fields("usuario") & "" & mysnap.Fields("sentido")

        End If

        '---------------
        imprime_detalle73
        '--------------
        mysnap.MoveNext
    Loop
             
    For I = 1 To 8
        buf1 = Format(recaudo(I), "0.00")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf1 = Format(flujo(I), "0")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
    Next I

    buf1 = Format(recaudo(10), "0.00")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf1 = Format(flujo(11), "0")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)

    found = formateaa("", 1, 2, 0)
    nlineas
    found = formateaa("", 29, 0, 0)

    For I = 1 To 8
        buf1 = Format(trecaudo(I), "0.00")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf1 = Format(tflujo(I), "0")

        If Val(buf1) = 0 Then buf1 = ""
        found = formateaa(buf1, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
    Next I

    buf1 = Format(trecaudo(10), "0.00")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf1 = Format(tflujo(11), "0")

    If Val(buf1) = 0 Then buf1 = ""
    found = formateaa(buf1, 8, 0, 1)
    found = formateaa("", 1, 0, 0)

    found = formateaa("", 1, 2, 0)
    Exit Sub
cmd99165121_err:

    MsgBox "Error en Cuerpo Programa-30 " & error$, 24, "Aviso"
    Exit Sub

End Sub

Function ver_proveedor(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "Select * from codprov where [producto]='" & buf & "' and codigo='" & "" & TxtCodProv.Text & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        ver_proveedor = 1

    End If

    mytablex.Close

End Function

Private Sub option1_Click()
    Command2.Visible = True
    xsaa = "1"

    If Combo3 = "NORMAL" Then
        imprime_opcion2

    End If

    If Combo3 = "EXCELL" Then
        imprime_opcion2_excel

    End If

    Command2.Visible = False

End Sub

Private Sub option2_Click()
    Command2.Visible = True
    xsaa = "2"

    If Combo3 = "NORMAL" Then
        imprime_opcion2

    End If

    If Combo3 = "EXCELL" Then
        imprime_opcion2_excel

    End If

    Command2.Visible = False

End Sub

Function pone_comisiones(mytabley As ADODB.Recordset) As Double

    Dim mytablex As New ADODB.Recordset

    'pone_comisiones = Val("" & mytabley.Fields("comision"))
    'Exit Function

    mytablex.Open "Select * from producto where [producto]='" & "" & mytabley.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        pone_comisiones = Val("" & mytablex.Fields("comision")) * Val("" & mytabley.Fields("xtotal")) / 100

    End If

    mytablex.Close

End Function

Function sql_producton(mytablex As ADODB.Recordset)

    Dim buf   As String

    Dim ybuf  As String

    Dim xbuf  As String

    Dim found As Integer

    If Len(fechai) <> 10 Then Exit Function
    If Len(fechaf) <> 10 Then Exit Function
    If Not IsDate(fechai) Then Exit Function
    If Not IsDate(fechaf) Then Exit Function
    If horai <> "%" And horaf <> "%" Then
        found = valida_hora(horai)

        If found = 0 Then Exit Function
        found = valida_hora(horaf)

        If found = 0 Then Exit Function

    End If

    xbuf = ""

    If Combo1 = "Producto" Then
        xbuf = "Producto"

    End If

    If Combo1 = "Codigo" Then
        xbuf = "Codigo"

    End If

    If Combo1 = "Vendedor" Then
        xbuf = "vendedor"

    End If

    If Combo1 = "Comisiones" Then
        xbuf = "vendedor"

    End If

    If Combo1 = "Zona" Then
        xbuf = "Zona"

    End If

    If Combo1 = "Cajero" Then
        xbuf = "Usuario"

    End If

    If Combo1 = "Caja" Then
        xbuf = "Caja"

    End If

    If Combo1 = "Turno" Then
        xbuf = "Turno"

    End If

    If Combo1 = "Familia" Then
        xbuf = "Familia"

    End If

    If Combo1 = "Subfamilia" Then
        xbuf = "Subfamilia"

    End If

    If Combo1 = "Ccosto" Then
        xbuf = "Ccosto"

    End If

    If Combo1 = "Almacen" Then
        xbuf = "bodega"

    End If

    If Combo1 = "Marca" Then
        xbuf = "Marca"

    End If

    buf = "select " & xbuf & ",DAY(fecha) as xmes,sum(cantidad*factor) as xcanti,sum(total) as xtotal,sum(tcosto*cantidad*factor) as xcosto,(sum(total)-sum(tcosto*cantidad*factor)) as xmargen from " & xdata & " where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If local1 <> "%" Then
        buf = buf & " and local='" & extra_loquesea(local1) & "'"

    End If

    If tipo <> "%" Then
        buf = buf & " and tipo like '" & tipo & "'"

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

    If usuario <> "%" Then
        buf = buf & " and usuario like '" & usuario & "'"

    End If

    If bodega <> "%" Then
        buf = buf & " and bodega like '" & extra_loquesea(bodega) & "'"

    End If

    If servicio <> "%" Then
        'If servicio = "Autoservicio" Then
        buf = buf & " and servicio='" & extra_loquesea(servicio) & "'"
        'End If
        'If servicio = "Comanda" Then
        '   buf = buf & " and servicio='C'"
        'End If
        'If servicio = "Delivery" Then
        '   buf = buf & " and servicio='D'"
        'End If

    End If

    If unidad <> "%" Then
        buf = buf & " and unidad like '" & unidad & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & caja & "'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & turno & "'"

    End If

    If producto <> "%" Then
        buf = buf & " and producto like '" & producto & "'"

    End If

    If ccosto <> "%" Then
        buf = buf & " and ccosto like '" & ccosto & "'"

    End If

    If horai <> "%" And horaf <> "%" Then
        buf = buf & " and HORA BETWEEN '" & horai & "' AND '" & horaf & "'"

    End If

    If familia <> "%" Then
        buf = buf & " and familia like '" & familia & "'"

    End If

    If subfamilia <> "%" Then
        buf = buf & " and subfamilia like '" & subfamilia & "'"

    End If

    If marca <> "%" Then
        buf = buf & " and marca like '" & marca & "'"

    End If

    If descripcio <> "%" Then
        buf = buf & " and [descripcio] like '" & descripcio & "'"

    End If

    If moneda <> "%" Then
        buf = buf & " and moneda like '" & moneda & "'"

    End If

    If vendedor <> "%" Then
        buf = buf & " and vendedor like '" & vendedor & "'"

    End If

    If estado <> "%" Then
        buf = buf & " and estado='" & estado & "'"

    End If

    If acu = "V" Then
        buf = buf & " and (acu='1' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') "

    End If

    If acu = "C" Then
        buf = buf & " and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') "

    End If

    If acu <> "C" And acu <> "V" Then

        'buf = buf & " and acu='" & acu & "'"
    End If

    If horai <> "%" And horaf <> "%" Then
        buf = buf & " and HORA BETWEEN '" & horai & "' AND '" & horaf & "'"

    End If

    If orden = "CANT" Then
        ybuf = " SUM(cantidad*factor) "

    End If

    If orden = "MONTO" Then
        ybuf = " SUM(total) "

    End If

    If orden = "GANANCIA" Then
        ybuf = " SUM(total)-SUM(cantidad*factor*tcosto) "

    End If

    buf = buf & "  group by " & xbuf & ", day(fecha)  order  by " & xbuf & " ," & ybuf & " DESC "
    'MsgBox buf
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    sql_producton = 1

End Function

Sub imprime_opcionn()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    Dim vr

    contpag = 0
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0

    found = sql_producton(mytablex)

    If found = 0 Then
        mytablex.Close
        Exit Sub

    End If

    Label33 = "" & mytablex.RecordCount
    vr = DoEvents
    'MsgBox "Presente una tecla", 48, "Aviso"
    'generar_temporal mytablex
    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_producton
    cuerpo_programa_producton mytablex
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
     
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1

End Sub

Sub cabecera_producton()

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
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Fechai : " & fechai, 25, 2, 0)
    found = formateaa("Fechaf : " & fechaf, 25, 2, 0)
    
    buf = String(352, "-")
    found = formateaa(buf, 352, 2, 0)
    buf = "" & Combo1
    found = formateaa(buf, 11, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("", 31, 0, 0)
   
    For I = 1 To 31
        found = formateaa("" & I, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
    Next I

    found = formateaa("", 1, 2, 0)
   
    buf = String(352, "-")
    found = formateaa(buf, 352, 2, 0)

End Sub

Sub cuerpo_programa_producton(mytablex As ADODB.Recordset)

    Dim vr

    Dim I        As Integer

    Dim sw1      As Integer

    Dim Tmp      As String

    Dim tmp1     As String

    Dim sw       As Integer

    Dim buf      As String

    Dim found    As Integer

    Dim sdx      As Double

    Dim mytabley As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset

    Dim xunidad  As String

    Dim xfactor  As String

    Dim sdxindx  As Double

    Dim xcosto   As Double

    Dim sdx1     As Double

    Dim sdxtmp   As Double

    Dim sstock   As Double

    On Error GoTo cmd2666_err

    sdx1 = 0
    sstock = 0
    sdx = 0
    sw = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0

    For I = 1 To 31
        xmeses(I) = 0
        xxmeses(I) = 0
    Next I

    tmp1 = ""
    sdxindx = 0
    Do

        If mytablex.EOF Then Exit Do
        vr = DoEvents()
        sdxindx = sdxindx + 1
        Command2.Caption = "" & sdxindx

        If Command2.Visible = False Then Exit Do
        If TxtCodProv.Text <> "%" Then
            found = ver_proveedor("" & mytablex.Fields("producto"))

            If found = 0 Then GoTo seguy124

        End If

        If Combo1 = "Producto" Then
            tmp1 = "" & mytablex.Fields("Producto")

        End If

        If Combo1 = "Codigo" Then
            tmp1 = "" & mytablex.Fields("codigo")

        End If

        If Combo1 = "Vendedor" Then
            tmp1 = "" & mytablex.Fields("vendedor")

        End If

        If Combo1 = "Comisiones" Then
            tmp1 = "" & mytablex.Fields("vendedor")

        End If

        If Combo1 = "Zona" Then
            tmp1 = "" & mytablex.Fields("Zona")

        End If

        If Combo1 = "Cajero" Then
            tmp1 = "" & mytablex.Fields("Usuario")

        End If

        If Combo1 = "Caja" Then
            tmp1 = "" & mytablex.Fields("Caja")

        End If

        If Combo1 = "Turno" Then
            tmp1 = "" & mytablex.Fields("Turno")

        End If

        If Combo1 = "Familia" Then
            tmp1 = "" & mytablex.Fields("Familia")

        End If

        If Combo1 = "Subfamilia" Then
            tmp1 = "" & mytablex.Fields("Subfamilia")

        End If

        If Combo1 = "Marca" Then
            tmp1 = "" & mytablex.Fields("Marca")

        End If

        If Combo1 = "Ccosto" Then
            tmp1 = "" & mytablex.Fields("Ccosto")

        End If

        If Combo1 = "Almacen" Then
            tmp1 = "" & mytablex.Fields("bodega")

        End If

        If sw = 0 Then
            If Combo1 = "Ccosto" Then
                buf = "" & mytablex.Fields("Ccosto")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                found = formateaa("", 31, 0, 0)
                Tmp = "" & mytablex.Fields("Ccosto")

            End If

            If Combo1 = "Almacen" Then
                buf = "" & mytablex.Fields("bodega")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                found = formateaa("", 31, 0, 0)
  
                Tmp = "" & mytablex.Fields("bodega")

            End If

            If Combo1 = "Familia" Then
                buf = "" & mytablex.Fields("Familia")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                found = formateaa("", 31, 0, 0)
   
                Tmp = "" & mytablex.Fields("Familia")

            End If

            If Combo1 = "Subfamilia" Then
                buf = "" & mytablex.Fields("Subfamilia")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                found = formateaa("", 31, 0, 0)
  
                Tmp = "" & mytablex.Fields("Subfamilia")

            End If

            If Combo1 = "Marca" Then
                buf = "" & mytablex.Fields("Marca")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                found = formateaa("", 31, 0, 0)
  
                Tmp = "" & mytablex.Fields("Marca")

            End If

            If Combo1 = "Producto" Then
                buf = "" & mytablex.Fields("producto")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_producto(buf)
                found = formateaa(buf, 31, 0, 0)
  
                Tmp = "" & mytablex.Fields("producto")

            End If

            If Combo1 = "Codigo" Then
                buf = "" & mytablex.Fields("codigo")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_nombre(buf)
                found = formateaa(buf, 31, 0, 0)
   
                Tmp = "" & mytablex.Fields("codigo")

            End If

            If Combo1 = "Vendedor" Then
                buf = "" & mytablex.Fields("vendedor")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor(buf)
                found = formateaa(buf, 31, 0, 0)
   
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Comisiones" Then
                buf = "" & mytablex.Fields("vendedor")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor(buf)
                found = formateaa(buf, 31, 0, 0)
   
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                found = formateaa("", 31, 0, 0)
  
                Tmp = "" & mytablex.Fields("zona")

            End If

            If Combo1 = "Cajero" Then
                buf = "" & mytablex.Fields("usuario")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor(buf)
                found = formateaa(buf, 31, 0, 0)
  
                Tmp = "" & mytablex.Fields("usuario")

            End If

            If Combo1 = "Caja" Then
                buf = "" & mytablex.Fields("Caja")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                found = formateaa("", 31, 0, 0)
  
                Tmp = "" & mytablex.Fields("Caja")

            End If

            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("Turno")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                found = formateaa("", 31, 0, 0)
  
                Tmp = "" & mytablex.Fields("Turno")

            End If

            sw = 1
            suma1 = 0
            suma2 = 0
   
        End If

        If Tmp <> tmp1 Then
            sdx = 0

            For I = 1 To 31
                sdx = sdx + xmeses(I)
                found = formateaa("" & xmeses(I), 8, 0, 1)
                found = formateaa("", 1, 0, 0)
            Next I

            found = formateaa("" & sdx, 8, 0, 1)
            found = formateaa("", 1, 0, 0)
   
            found = formateaa("", 1, 2, 0)
            nlineas
   
            For I = 1 To 31
                xmeses(I) = 0
            Next I

            If Combo1 = "Ccosto" Then
                buf = "" & mytablex.Fields("Ccosto")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                found = formateaa("", 31, 0, 0)
                Tmp = "" & mytablex.Fields("Ccosto")

            End If

            If Combo1 = "Almacen" Then
                buf = "" & mytablex.Fields("bodega")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                found = formateaa("", 31, 0, 0)
  
                Tmp = "" & mytablex.Fields("bodega")

            End If

            If Combo1 = "Familia" Then
                buf = "" & mytablex.Fields("Familia")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                found = formateaa("", 31, 0, 0)
   
                Tmp = "" & mytablex.Fields("Familia")

            End If

            If Combo1 = "Subfamilia" Then
                buf = "" & mytablex.Fields("Subfamilia")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                found = formateaa("", 31, 0, 0)
  
                Tmp = "" & mytablex.Fields("Subfamilia")

            End If

            If Combo1 = "Marca" Then
                buf = "" & mytablex.Fields("Marca")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                found = formateaa("", 31, 0, 0)
  
                Tmp = "" & mytablex.Fields("Marca")

            End If

            If Combo1 = "Producto" Then
                buf = "" & mytablex.Fields("producto")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_producto(buf)
                found = formateaa(buf, 31, 0, 0)
  
                Tmp = "" & mytablex.Fields("producto")

            End If

            If Combo1 = "Codigo" Then
                buf = "" & mytablex.Fields("codigo")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_nombre(buf)
                found = formateaa(buf, 31, 0, 0)
   
                Tmp = "" & mytablex.Fields("codigo")

            End If

            If Combo1 = "Vendedor" Then
                buf = "" & mytablex.Fields("vendedor")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor(buf)
                found = formateaa(buf, 31, 0, 0)
   
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Comisiones" Then
                buf = "" & mytablex.Fields("vendedor")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor(buf)
                found = formateaa(buf, 31, 0, 0)
   
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                found = formateaa("", 31, 0, 0)
  
                Tmp = "" & mytablex.Fields("zona")

            End If

            If Combo1 = "Cajero" Then
                buf = "" & mytablex.Fields("usuario")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor(buf)
                found = formateaa(buf, 31, 0, 0)
  
                Tmp = "" & mytablex.Fields("usuario")

            End If

            If Combo1 = "Caja" Then
                buf = "" & mytablex.Fields("Caja")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                found = formateaa("", 31, 0, 0)
  
                Tmp = "" & mytablex.Fields("Caja")

            End If

            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("Turno")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                found = formateaa("", 31, 0, 0)
  
                Tmp = "" & mytablex.Fields("Turno")

            End If

        End If

        If orden = "MONTO" Then
            xmeses(Val("" & mytablex.Fields("xmes"))) = xmeses(Val("" & mytablex.Fields("xmes"))) + Val("" & mytablex.Fields("xtotal"))
            xxmeses(Val("" & mytablex.Fields("xmes"))) = xxmeses(Val("" & mytablex.Fields("xmes"))) + Val("" & mytablex.Fields("xtotal"))

        End If

        If orden = "CANT" Then
            xmeses(Val("" & mytablex.Fields("xmes"))) = xmeses(Val("" & mytablex.Fields("xmes"))) + Val("" & mytablex.Fields("xcanti"))
            xxmeses(Val("" & mytablex.Fields("xmes"))) = xxmeses(Val("" & mytablex.Fields("xmes"))) + Val("" & mytablex.Fields("xcanti"))

        End If

seguy124:
        mytablex.MoveNext
    Loop
    sdx = 0

    For I = 1 To 31
        sdx = sdx + xmeses(I)
        found = formateaa("" & xmeses(I), 8, 0, 1)
        found = formateaa("", 1, 0, 0)
    Next I

    found = formateaa("" & sdx, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
   
    found = formateaa("", 1, 2, 0)
    nlineas
   
    found = formateaa("Totales", 11, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("", 31, 0, 0)
  
    sdx = 0

    For I = 1 To 31
        sdx = sdx + xxmeses(I)
        found = formateaa("" & xxmeses(I), 8, 0, 1)
        found = formateaa("", 1, 0, 0)
    Next I

    found = formateaa("" & sdx, 8, 0, 1)
    found = formateaa("", 1, 2, 0)
   
    Exit Sub
cmd2666_err:
    MsgBox "Aviso en cuerpo programa producto n" + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub TxtCodProv_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_codigo

    End If

End Sub

