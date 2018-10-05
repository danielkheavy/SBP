VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tclieta 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operacion"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   13755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   13755
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Direcciones de Despacho"
      Height          =   3255
      Left            =   12120
      TabIndex        =   107
      Top             =   2760
      Visible         =   0   'False
      Width           =   8175
      Begin VB.TextBox direcciona 
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
         Left            =   1920
         MaxLength       =   60
         TabIndex        =   108
         Top             =   240
         Width           =   4935
      End
      Begin MSDataGridLib.DataGrid dbgrid6 
         Height          =   2415
         Left            =   120
         TabIndex        =   109
         Top             =   720
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   4260
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
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "+"
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
         Left            =   6960
         TabIndex        =   113
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
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
         Left            =   6960
         TabIndex        =   112
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label44 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cierra"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   111
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label43 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Direccion"
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
         TabIndex        =   110
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Telefonos"
      Height          =   3255
      Left            =   12120
      TabIndex        =   100
      Top             =   5280
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox telefonoa 
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
         Left            =   1920
         MaxLength       =   14
         TabIndex        =   105
         Top             =   240
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid dbgrid5 
         Height          =   2415
         Left            =   120
         TabIndex        =   101
         Top             =   720
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4260
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
      Begin VB.Label Label42 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Telefono"
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
         TabIndex        =   106
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cierra"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   104
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
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
         Left            =   5160
         TabIndex        =   103
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "+"
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
         Left            =   3600
         TabIndex        =   102
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Otros"
      Height          =   5655
      Left            =   12480
      TabIndex        =   81
      Top             =   1080
      Visible         =   0   'False
      Width           =   9015
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Sa&lir"
         Height          =   615
         Left            =   7800
         Picture         =   "tclieta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   4800
         Width           =   1095
      End
      Begin VB.TextBox tipovive 
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
         MaxLength       =   1
         TabIndex        =   96
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox hobbie 
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
         MaxLength       =   30
         TabIndex        =   94
         Top             =   2520
         Width           =   3615
      End
      Begin VB.TextBox cargo 
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
         MaxLength       =   30
         TabIndex        =   92
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox Trabajo 
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
         MaxLength       =   30
         TabIndex        =   90
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox nrodepe 
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
         MaxLength       =   2
         TabIndex        =   88
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox civil 
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
         MaxLength       =   1
         TabIndex        =   86
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox religion 
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
         MaxLength       =   30
         TabIndex        =   84
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox profesion 
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
         MaxLength       =   30
         TabIndex        =   82
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label38 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Vivienda (A.P.)"
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
         Left            =   240
         TabIndex        =   97
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label Label37 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hobbie"
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
         Left            =   240
         TabIndex        =   95
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label36 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cargo"
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
         Left            =   240
         TabIndex        =   93
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label xk44 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Centro Trabajo"
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
         Left            =   240
         TabIndex        =   91
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label34 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NroDependientes"
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
         Left            =   240
         TabIndex        =   89
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label33 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estado Civil (S/C)"
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
         Left            =   240
         TabIndex        =   87
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label32 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Religion"
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
         Left            =   240
         TabIndex        =   85
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Profesion"
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
         Left            =   240
         TabIndex        =   83
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consulta"
      Height          =   8895
      Left            =   11520
      TabIndex        =   76
      Top             =   2640
      Visible         =   0   'False
      Width           =   13695
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
         Left            =   8280
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   600
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   600
         Width           =   2295
      End
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
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   600
         Width           =   2895
      End
      Begin MSDataGridLib.DataGrid dbGrid3 
         Height          =   7575
         Left            =   240
         TabIndex        =   80
         Top             =   1200
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   13361
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
   Begin VB.TextBox barras 
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
      Left            =   6480
      MaxLength       =   15
      TabIndex        =   4
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox dni 
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
      Left            =   6480
      MaxLength       =   11
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox ruc 
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
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox especial 
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
      MaxLength       =   1
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   7320
      Width           =   495
   End
   Begin VB.TextBox estado 
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
      TabIndex        =   69
      Top             =   7680
      Width           =   495
   End
   Begin VB.TextBox clasifica 
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
      Left            =   5520
      MaxLength       =   6
      TabIndex        =   68
      Top             =   6240
      Width           =   1455
   End
   Begin VB.TextBox zona 
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
      TabIndex        =   67
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox tipoclie 
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
      TabIndex        =   2
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CheckBox viernes 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Viernes"
      Height          =   375
      Left            =   9000
      TabIndex        =   66
      Top             =   5400
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   13695
      TabIndex        =   65
      Top             =   0
      Width           =   13755
      Begin VB.CommandButton Command8 
         Caption         =   "Otros"
         Height          =   615
         Left            =   960
         Picture         =   "tclieta.frx":08DA
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H80000016&
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
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tclieta.frx":0A44
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Grabar registro"
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.TextBox codigo 
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   6
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox codigo1 
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
      Top             =   1560
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
      MaxLength       =   60
      TabIndex        =   3
      Top             =   2280
      Width           =   4695
   End
   Begin VB.TextBox nombrec 
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
      TabIndex        =   34
      Top             =   2640
      Width           =   4695
   End
   Begin VB.TextBox contacto 
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
      TabIndex        =   33
      Top             =   3000
      Width           =   4695
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
      MaxLength       =   60
      TabIndex        =   32
      Top             =   3480
      Width           =   4695
   End
   Begin VB.TextBox dpto 
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
      TabIndex        =   31
      Top             =   4200
      Width           =   3375
   End
   Begin VB.TextBox distrito 
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
      TabIndex        =   30
      Top             =   4560
      Width           =   3375
   End
   Begin VB.TextBox telefono 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   14
      TabIndex        =   29
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox telefono1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      MaxLength       =   14
      TabIndex        =   28
      Top             =   5760
      Width           =   1575
   End
   Begin VB.TextBox telefono2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      MaxLength       =   14
      TabIndex        =   27
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox correo 
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
      TabIndex        =   26
      Top             =   5280
      Width           =   4695
   End
   Begin VB.TextBox descuento 
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
      MaxLength       =   10
      TabIndex        =   25
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox diapago 
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
      MaxLength       =   2
      TabIndex        =   24
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox fpago 
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
      TabIndex        =   23
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox cuenta 
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
      MaxLength       =   20
      TabIndex        =   22
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox vendedor 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   11
      TabIndex        =   21
      Top             =   6240
      Width           =   1935
   End
   Begin VB.TextBox descuento1 
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
      MaxLength       =   8
      TabIndex        =   20
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox credito 
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
      MaxLength       =   8
      TabIndex        =   19
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CheckBox lunes 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Lunes"
      Height          =   375
      Left            =   7080
      TabIndex        =   18
      Top             =   5040
      Width           =   735
   End
   Begin VB.CheckBox martes 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Martes"
      Height          =   375
      Left            =   7920
      TabIndex        =   17
      Top             =   5040
      Width           =   855
   End
   Begin VB.CheckBox miercoles 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Miercoles"
      Height          =   375
      Left            =   8760
      TabIndex        =   16
      Top             =   5040
      Width           =   975
   End
   Begin VB.CheckBox jueves 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Jueves"
      Height          =   375
      Left            =   9840
      TabIndex        =   15
      Top             =   5040
      Width           =   855
   End
   Begin VB.CheckBox sabado 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sabado"
      Height          =   375
      Left            =   7080
      TabIndex        =   14
      Top             =   5400
      Width           =   855
   End
   Begin VB.CheckBox domingo 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Domingo"
      Height          =   375
      Left            =   7920
      TabIndex        =   13
      Top             =   5400
      Width           =   975
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
      Left            =   8760
      MaxLength       =   1
      TabIndex        =   12
      Top             =   4320
      Width           =   375
   End
   Begin VB.TextBox flete 
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
      MaxLength       =   8
      TabIndex        =   11
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox fechalta 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   10
      Top             =   7320
      Width           =   1935
   End
   Begin VB.TextBox referencia 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   11
      TabIndex        =   9
      Top             =   6600
      Width           =   1935
   End
   Begin VB.TextBox garantia 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   11
      TabIndex        =   8
      Top             =   6960
      Width           =   1935
   End
   Begin VB.TextBox referencias 
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
      TabIndex        =   7
      Top             =   3840
      Width           =   4695
   End
   Begin VB.Label Label31 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cod.Barras"
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
      Left            =   4320
      TabIndex        =   75
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label30 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dni"
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
      Left            =   4320
      TabIndex        =   73
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label28 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ruc"
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
      TabIndex        =   72
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label27 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cliente Especial 1.Si"
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
      Left            =   4200
      TabIndex        =   71
      Top             =   7320
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   64
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C.Extranjeria"
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
      TabIndex        =   63
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ApellidoNomb/RSocial "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   62
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre Comercial"
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
      TabIndex        =   61
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Contacto"
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
      TabIndex        =   60
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Direccion"
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
      TabIndex        =   59
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Departamento"
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
      TabIndex        =   58
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Distrito"
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
      TabIndex        =   57
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Telefono"
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
      TabIndex        =   56
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Correo Electronico"
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
      TabIndex        =   55
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Zona"
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
      TabIndex        =   54
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFC0&
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
      Left            =   120
      TabIndex        =   53
      Top             =   7680
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dto. por Defecto"
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
      Left            =   7080
      TabIndex        =   52
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nro Dias "
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
      Left            =   7080
      TabIndex        =   51
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CondicionVenta"
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
      Left            =   7080
      TabIndex        =   50
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cuenta"
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
      Left            =   7080
      TabIndex        =   49
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   48
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dto. pronto Pago"
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
      Left            =   7080
      TabIndex        =   47
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Linea de Credito"
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
      Left            =   7080
      TabIndex        =   46
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dias de Visita"
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
      Left            =   7080
      TabIndex        =   45
      Top             =   4680
      Width           =   2895
   End
   Begin VB.Label Label29 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Moneda"
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
      Left            =   7080
      TabIndex        =   44
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Flete"
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
      Left            =   7080
      TabIndex        =   43
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaNacimiento"
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
      TabIndex        =   42
      Top             =   7320
      Width           =   2175
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo Cliente"
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
      TabIndex        =   41
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label clasddd 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clasificacion"
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
      Left            =   4200
      TabIndex        =   40
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label24 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Referido Por"
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
      TabIndex        =   39
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Label Label25 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Garantia"
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
      TabIndex        =   38
      Top             =   6960
      Width           =   2175
   End
   Begin VB.Label ngarantia 
      BackColor       =   &H00FFFFC0&
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
      Left            =   4200
      TabIndex        =   37
      Top             =   6960
      Width           =   4695
   End
   Begin VB.Label nreferencia 
      BackColor       =   &H00FFFFC0&
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
      Left            =   4200
      TabIndex        =   36
      Top             =   6600
      Width           =   4695
   End
   Begin VB.Label Label35 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Referencia"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   35
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Menu dk23 
      Caption         =   "&Graba"
   End
   Begin VB.Menu flo223 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tclieta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cadena_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   If opcion1 = "1" Then
      Frame3.Visible = False
      Frame3.Enabled = False
      tipoclie.SetFocus
      Exit Sub
   End If
   If opcion1 = "2" Then
      Frame3.Visible = False
      Frame3.Enabled = False
      vendedor.SetFocus
      Exit Sub
   End If
   If opcion1 = "3" Then
      Frame3.Visible = False
      Frame3.Enabled = False
      clasifica.SetFocus
      Exit Sub
   End If
   If opcion1 = "300" Then
      Frame3.Visible = False
      Frame3.Enabled = False
      nombre.SetFocus
      Exit Sub
   End If

      If opcion1 = "4" Then
      Frame3.Visible = False
      Frame3.Enabled = False
      zona.SetFocus
      Exit Sub
   End If
      If opcion1 = "5" Then
      Frame3.Visible = False
      Frame3.Enabled = False
      fpago.SetFocus
      Exit Sub
   End If
   If opcion1 = "200" Then
      Frame3.Visible = False
      Frame3.Enabled = False
      referencia.SetFocus
      Exit Sub
   End If

   
End If
Command3_Click

End Sub

Private Sub buffer_Change()

End Sub

Private Sub cargo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
religion.SetFocus

End Sub

Private Sub civil_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
nrodepe.SetFocus

End Sub

Private Sub clasifica_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
referencia.SetFocus

End Sub

Private Sub clasifica_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   vendedor.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   Combo1.Clear
   Combo1.AddItem "Descripcio"
   Combo1.AddItem "Clasifica"
   Combo1.ListIndex = 0
   opcion1 = "3"
   Frame3.Visible = True
   Frame3.Enabled = True
   cadena = ""
   'cadena.SetFocus
   Command3_Click
   Exit Sub
End If


End Sub

Private Sub cmdSave_Click()
dk23_Click
End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
ruc.SetFocus
End Sub

Private Sub codigo1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
tipoclie.SetFocus

End Sub

Private Sub codigo1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   dni.SetFocus
   Exit Sub
End If
End Sub

Sub ejecuta(sw As Integer)
Dim buf As String
Dim mytablex As New ADODB.Recordset
If opcion1 = "1" Then
If Len(cadena) = 0 Then
buf = "select Descripcio,tipoclie from tipoclie "
Else
buf = "select Descripcio,tipoclie from tipoclie where " & Combo1 & " like '" & cadena & "%'"
End If
End If
If opcion1 = "2" Then
If Len(cadena) = 0 Then
buf = "select Nombre,Codigo from Vendedor "
Else
buf = "select Nombre,Codigo from Vendedor where " & Combo1 & " like '" & cadena & "%'"
End If
End If
If opcion1 = "200" Or opcion1 = "201" Then
If Len(cadena) = 0 Then
   buf = "select Nombre,Codigo from clientes "
Else
   buf = "select Nombre,Codigo from clientes where " & Combo1 & " like '" & cadena & "%'"
End If

End If

If opcion1 = "300" Then
   If Len(cadena) = 0 Then
      buf = "select Nombre,Codigo from clientes "
   Else
   buf = "select Nombre,Codigo from clientes where " & Combo1 & " like '" & cadena & "%'"
End If
End If



If opcion1 = "3" Then
If Len(cadena) = 0 Then
buf = "select Descripcio,Clasifica from clasifi "
Else
buf = "select Descripcio,Clasifica from Clasifi where " & Combo1 & " like '" & cadena & "%'"
End If
End If
If opcion1 = "4" Then
If Len(cadena) = 0 Then
buf = "select Descripcio,Zona from Zona "
Else
buf = "select Descripcio,Zona from Zona where " & Combo1 & " like '" & cadena & "%'"
End If
End If
If opcion1 = "5" Then
If Len(cadena) = 0 Then
buf = "select Descripcio,Fpago from Fpago "
Else
buf = "select Descripcio,Fpago from Fpago where " & Combo1 & " like '" & cadena & "%'"
End If
End If
'MsgBox buf


   If mytablex.State = 1 Then mytablex.Close
   mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
   Set dbGrid3.DataSource = mytablex
   dbGrid3.Columns(0).Width = 4000
   dbGrid3.Columns(1).Width = 2000
   If mytablex.RecordCount = 0 Then
      mytablex.Close
      cadena.SetFocus
      Exit Sub
   End If
   dbGrid3.SetFocus
End Sub


Private Sub Command1_Click()

End Sub

Private Sub Command3_Click()
ejecuta 1
End Sub


Private Sub Command8_Click()
Frame1.Visible = True
profesion.SetFocus

End Sub

Private Sub Command9_Click()
Frame1.Visible = False
End Sub

Private Sub contacto_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
direccion.SetFocus

End Sub

Private Sub contacto_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   nombrec.SetFocus
   Exit Sub
End If

End Sub

Private Sub correo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
vendedor.SetFocus

End Sub

Private Sub correo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   telefono2.SetFocus
   Exit Sub
End If

End Sub

Private Sub credito_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
moneda.SetFocus

End Sub

Private Sub credito_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   descuento1.SetFocus
   Exit Sub
End If

End Sub

Private Sub cuenta_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
flete.SetFocus

End Sub

Private Sub cuenta_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   fpago.SetFocus
   Exit Sub
End If

End Sub

Private Sub DBGrid3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   cadena.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
   If opcion1 = "1" Then
      tipoclie = dbGrid3.Columns(1)
      Frame3.Visible = False
      Frame3.Enabled = False
      tipoclie.SetFocus
   End If
      If opcion1 = "2" Then
      vendedor = dbGrid3.Columns(1)
      Frame3.Visible = False
      Frame3.Enabled = False
      vendedor.SetFocus
   End If
   If opcion1 = "3" Then
      clasifica = dbGrid3.Columns(1)
      Frame3.Visible = False
      Frame3.Enabled = False
      clasifica.SetFocus
   End If
   If opcion1 = "3" Then
      Exit Sub
      
   End If
   If opcion1 = "4" Then
      zona = dbGrid3.Columns(1)
      Frame3.Visible = False
      Frame3.Enabled = False
      zona.SetFocus
   End If
   If opcion1 = "5" Then
      fpago = dbGrid3.Columns(1)
      Frame3.Visible = False
      Frame3.Enabled = False
      fpago.SetFocus
   End If
   
   If opcion1 = "200" Then
      referencia = dbGrid3.Columns(1)
      nreferencia = dbGrid3.Columns(0)
      Frame3.Visible = False
      Frame3.Enabled = False
      referencia.SetFocus
   End If
   If opcion1 = "201" Then
      garantia = dbGrid3.Columns(1)
      nreferencia = dbGrid3.Columns(0)
      Frame3.Visible = False
      Frame3.Enabled = False
      garantia.SetFocus
   End If

End If

End Sub

Private Sub DBGrid1_Click()

End Sub

Private Sub descuento_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
diapago.SetFocus

End Sub

Private Sub descuento_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   estado.SetFocus
   Exit Sub
End If

End Sub

Private Sub descuento1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
credito.SetFocus
End Sub

Private Sub descuento1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   flete.SetFocus
   Exit Sub
End If

End Sub

Private Sub diapago_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
fpago.SetFocus

End Sub

Private Sub diapago_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   descuento.SetFocus
   Exit Sub
End If

End Sub

Private Sub direccion_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
referencias.SetFocus

End Sub

Private Sub direccion_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   contacto.SetFocus
   Exit Sub
End If

End Sub

Private Sub distrito_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
zona.SetFocus

End Sub

Private Sub distrito_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   dpto.SetFocus
   Exit Sub
End If

End Sub

Private Sub dk23_Click()
Dim found As Integer

found = grabando()
If found = 1 Then
   flo223_Click
   Exit Sub
End If
End Sub

Private Sub dni_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
codigo1.SetFocus

End Sub

Private Sub Dni_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   ruc.SetFocus
   Exit Sub
End If

End Sub

Private Sub dpto_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
distrito.SetFocus

End Sub

Private Sub dpto_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   referencias.SetFocus
   Exit Sub
End If

End Sub

Private Sub estado_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
descuento.SetFocus

End Sub

Private Sub estado_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   fechalta.SetFocus
   Exit Sub
End If
End Sub

Private Sub fechalta_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
estado.SetFocus
End Sub

Private Sub fechalta_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   garantia.SetFocus
   Exit Sub
End If
End Sub

Private Sub flete_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
descuento1.SetFocus

End Sub

Private Sub flete_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   cuenta.SetFocus
   Exit Sub
End If

End Sub

Private Sub flo223_Click()
If Frame1.Visible = True Then
   Frame1.Visible = False
   Exit Sub
End If

If Frame3.Visible = True Then
   cadena_KeyPress 27
   Exit Sub
End If
tclieta.Hide
Unload tclieta
End Sub

Private Sub Form_Activate()
nreferencia = busca_clientes("" & referencia)
ngarantia = busca_clientes("" & garantia)
End Sub

Private Sub fpago_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
cuenta.SetFocus
End Sub

Private Sub fpago_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   diapago.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   Combo1.Clear
   Combo1.AddItem "Descripcio"
   Combo1.AddItem "Fpago"
   Combo1.ListIndex = 0
   opcion1 = "5"
   Frame3.Visible = True
   Frame3.Enabled = True
   cadena = ""
   cadena.SetFocus
   Command3_Click
   Exit Sub
End If

End Sub

Private Sub garantia_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
fechalta.SetFocus

End Sub

Private Sub garantia_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   referencia.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   Combo1.Clear
   Combo1.AddItem "Nombre"
   Combo1.AddItem "Codigo"
   Combo1.ListIndex = 0
   opcion1 = "201"
   Frame3.Visible = True
   Frame3.Enabled = True
   cadena = ""
   cadena.SetFocus
   Command3_Click
   Exit Sub
End If


End Sub

Private Sub hobie_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
tipovive.SetFocus

End Sub

Private Sub Label39_Click()
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
   mytablex.Open "select * from telefono where codigo='" & codigo & "' and telefono='" & telefonoa & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount = 0 Then
      mytablex.AddNew
      mytablex.Fields("telefono") = telefonoa
      mytablex.Fields("codigo") = codigo
      mytablex.Update
      Else
      telefonoa.SetFocus
      Exit Sub
   End If
   mytablex.Close
   Label9_Click
   
End Sub

Private Sub Label40_Click()
On Error GoTo cmd567_err
cn.Execute ("delete from telefono where codigo='" & codigo & "' and telefono='" & "" & dbgrid5.Columns("telefono") & "'")
Label9_Click
Exit Sub
cmd567_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub
End Sub

Private Sub Label41_Click()
   Frame2.Visible = False
   Frame2.Enabled = False
End Sub

Private Sub Label44_Click()
Frame4.Visible = False
Frame4.Enabled = False
End Sub

Private Sub Label45_Click()
On Error GoTo cmd568_err
cn.Execute ("delete from despacho where codigo='" & codigo & "' and direccion='" & "" & dbgrid6.Columns("direccion") & "'")
Label6_Click
Exit Sub
cmd568_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub

End Sub

Private Sub Label46_Click()
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
   mytablex.Open "select * from despacho where codigo='" & codigo & "' and direccion='" & direcciona & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount = 0 Then
      mytablex.AddNew
      mytablex.Fields("direccion") = direcciona
      mytablex.Fields("codigo") = codigo
      mytablex.Update
      Else
      direcciona.SetFocus
      Exit Sub
   End If
   mytablex.Close
   Label6_Click

End Sub

Private Sub Label6_Click()
Dim mytablex As New ADODB.Recordset
If tclieta.Caption = "NUEVO" Then
   MsgBox "Solo en modificacion ", 48, "Aviso"
   Exit Sub
End If
   If mytablex.State = 1 Then mytablex.Close
   mytablex.Open "select * from despacho where codigo='" & codigo & "'", cn, adOpenStatic, adLockOptimistic
   Set dbgrid6.DataSource = mytablex
   dbgrid6.Columns(0).Width = 4000
   dbgrid6.Columns(1).Width = 1000
   Frame4.Visible = True
   Frame4.Enabled = True
   direcciona = ""
   direcciona.SetFocus

End Sub

Private Sub Label9_Click()
Dim mytablex As New ADODB.Recordset
If tclieta.Caption = "NUEVO" Then
   MsgBox "Solo en modificacion ", 48, "Aviso"
   Exit Sub
   Exit Sub
End If
   If mytablex.State = 1 Then mytablex.Close
   mytablex.Open "select * from telefono where codigo='" & codigo & "'", cn, adOpenStatic, adLockOptimistic
   Set dbgrid5.DataSource = mytablex
   dbgrid5.Columns(0).Width = 2000
   dbgrid5.Columns(1).Width = 2000
   Frame2.Visible = True
   Frame2.Enabled = True
   telefonoa = ""
   telefonoa.SetFocus
End Sub

Private Sub moneda_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
End Sub

Private Sub moneda_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   credito.SetFocus
   Exit Sub
End If

End Sub

Private Sub nombre_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
nombrec.SetFocus
End Sub

Private Sub nombre_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   tipoclie.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   Combo1.Clear
   Combo1.AddItem "Nombre"
   Combo1.ListIndex = 0
   opcion1 = "300"
   Frame3.Visible = True
   Frame3.Enabled = True
   cadena = ""
   If Len(nombre) > 0 Then
      cadena = "%" & nombre & "%"
   End If
   cadena.SetFocus
   Command3_Click
   Exit Sub
End If


End Sub

Private Sub nombrec_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
contacto.SetFocus

End Sub

Private Sub nombrec_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   nombre.SetFocus
   Exit Sub
End If

End Sub

Private Sub nrodepe_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
hobbie.SetFocus

End Sub

Private Sub profesion_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
Trabajo.SetFocus
End Sub

Private Sub referencia_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
garantia.SetFocus

End Sub

Private Sub referencia_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   clasifica.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   Combo1.Clear
   Combo1.AddItem "Nombre"
   Combo1.AddItem "Codigo"
   Combo1.ListIndex = 0
   opcion1 = "200"
   Frame3.Visible = True
   Frame3.Enabled = True
   cadena = ""
   cadena.SetFocus
   Command3_Click
   Exit Sub
End If



End Sub

Private Sub referencias_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
dpto.SetFocus

End Sub

Private Sub referencias_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   direccion.SetFocus
   Exit Sub
End If

End Sub

Private Sub religion_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
civil.SetFocus

End Sub

Private Sub ruc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
dni.SetFocus
End Sub

Private Sub ruc_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   If codigo.Enabled = True Then
      codigo.SetFocus
   End If
   Exit Sub
End If

End Sub

Private Sub telefono_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
telefono1.SetFocus

End Sub

Private Sub telefono_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   zona.SetFocus
   Exit Sub
End If

End Sub

Private Sub telefono1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
telefono2.SetFocus

End Sub

Private Sub telefono1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   telefono.SetFocus
   Exit Sub
End If

End Sub

Private Sub telefono2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
correo.SetFocus

End Sub

Private Sub telefono2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   telefono1.SetFocus
   Exit Sub
End If

End Sub

Private Sub tipoclie_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
nombre.SetFocus

End Sub

Private Sub tipoclie_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   codigo1.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   Combo1.Clear
   Combo1.AddItem "Descripcio"
   Combo1.AddItem "Familia"
   Combo1.ListIndex = 0
   opcion1 = "1"
   Frame3.Visible = True
   Frame3.Enabled = True
   cadena = ""
   cadena.SetFocus
   Command3_Click
   Exit Sub
End If


End Sub

Private Sub Trabajo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
cargo.SetFocus

End Sub

Private Sub vendedor_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
clasifica.SetFocus

End Sub

Private Sub vendedor_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   correo.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   Combo1.Clear
   Combo1.AddItem "Nombre"
   Combo1.AddItem "Codigo"
   Combo1.ListIndex = 0
   opcion1 = "2"
   Frame3.Visible = True
   Frame3.Enabled = True
   cadena = ""
   cadena.SetFocus
   Command3_Click
   Exit Sub
End If


End Sub

Private Sub zona_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
telefono.SetFocus

End Sub

Private Sub zona_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   distrito.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   Combo1.Clear
   Combo1.AddItem "Descripcio"
   Combo1.AddItem "Zona"
   Combo1.ListIndex = 0
   opcion1 = "4"
   Frame3.Visible = True
   Frame3.Enabled = True
   cadena = ""
   cadena.SetFocus
   Command3_Click
   Exit Sub
End If


End Sub
Function grabando()
Dim found As Integer
found = valida()
If found = 0 Then
   MsgBox "Campos invalidos No se puede grabar", 48, "Aviso"
   Exit Function
End If
If MsgBox("Desea Grabar???", 1, "Aviso") <> 1 Then Exit Function
If tclieta.Caption = "MODIFICA" Then
End If
If tclieta.Caption = "NUEVO" Then
   dbclie.AddNew
End If
 dbclie.Fields("lunes") = lunes.Value
 dbclie.Fields("martes") = martes.Value
 dbclie.Fields("miercoles") = miercoles.Value
 dbclie.Fields("jueves") = jueves.Value
 dbclie.Fields("viernes") = viernes.Value
 dbclie.Fields("sabado") = sabado.Value
 dbclie.Fields("domingo") = domingo.Value
 dbclie.Fields("flete") = Val(flete)
 dbclie.Fields("REFERENCIA") = referencia
 dbclie.Fields("GARANTIA") = garantia
 dbclie.Fields("observa") = referencias
 dbclie.Fields("tipoclie") = tipoclie
 dbclie.Fields("especial") = especial
 dbclie.Fields("clasifica") = clasifica
If Len(fechalta) = 0 Then
    dbclie.Fields("fechanac") = Format(Now, "dd/mm/yyyy")
   Else
   If IsDate(fechalta) Then
    dbclie.Fields("fechanac") = fechalta
   End If
End If
 dbclie.Fields("moneda") = moneda
 dbclie.Fields("vendedor") = vendedor
 dbclie.Fields("descuento1") = Val(descuento1)
 dbclie.Fields("credito") = Val(credito)
 dbclie.Fields("barras") = barras
 dbclie.Fields("dni") = dni
 dbclie.Fields("ruc") = ruc
 'dbclie.Fields("codigo") = codigo
 dbclie.Fields("extranjeria") = codigo1
 dbclie.Fields("nombre") = nombre
 dbclie.Fields("nombrec") = nombrec
 dbclie.Fields("contacto") = contacto
 dbclie.Fields("direccion") = direccion
 dbclie.Fields("dpto") = dpto
 dbclie.Fields("distrito") = distrito
 dbclie.Fields("zona") = zona
 dbclie.Fields("telefono") = telefono
 dbclie.Fields("telefono1") = telefono1
 dbclie.Fields("telefono2") = telefono2
 dbclie.Fields("correo") = correo
 dbclie.Fields("estado") = estado
 dbclie.Fields("descuento") = Val(descuento)
 dbclie.Fields("diapago") = diapago
 dbclie.Fields("fpago") = fpago
 dbclie.Fields("cuenta") = cuenta
 
 
  dbclie.Fields("profesion") = profesion
  dbclie.Fields("trabajo") = Trabajo
  dbclie.Fields("religion") = religion
  dbclie.Fields("nrodepe") = nrodepe
  dbclie.Fields("cargo") = cargo
  dbclie.Fields("hobbie") = hobbie
  dbclie.Fields("civil") = civil
  dbclie.Fields("tipovive") = tipovive
  
 
 dbclie.Update
 If Len(telefono) > 0 Then
    agregar_telefono "" & telefono
 End If
 If Len(telefono1) > 0 Then
    agregar_telefono "" & telefono1
 End If
 If Len(telefono2) > 0 Then
    agregar_telefono "" & telefono2
 End If
grabando = 1
End Function
Function valida()
Dim mytablex As New ADODB.Recordset
'If Len(codigo) = 0 Then
'   codigo.SetFocus
'   Exit Function
'End If
If Len(nombre) = 0 Then
   nombre.SetFocus
   Exit Function
End If
If moneda <> "S" And moneda <> "D" Then
   moneda.SetFocus
   Exit Function
End If
If Len(fechalta) > 0 Then
If Not IsDate(fechalta) Then
   fechalta = ""
   fechalta.SetFocus
   Exit Function
End If
End If
If Len(dni) > 0 Then
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "select codigo,dni from clientes where dni='" & dni & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   If "" & mytablex.Fields("codigo") <> codigo Then
      MsgBox "Dni ya usado en " + mytablex.Fields("codigo")
      mytablex.Close
      Exit Function
   End If
End If
mytablex.Close
End If
If Len(ruc) > 0 And Len(ruc) < 11 Then
   MsgBox "Ruc no Valido "
   Exit Function
End If
If Len(ruc) > 0 Then
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "select codigo,ruc from clientes where ruc='" & ruc & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   If "" & mytablex.Fields("codigo") <> codigo Then
      MsgBox "Ruc ya usado en " + mytablex.Fields("codigo")
      mytablex.Close
      Exit Function
   End If
End If
mytablex.Close
End If
If Len(dni) > 0 Then
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "select codigo,dni from clientes where Dni='" & dni & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   If "" & mytablex.Fields("codigo") <> codigo Then
      MsgBox "Dni barras ya usado en " + mytablex.Fields("codigo")
      mytablex.Close
      Exit Function
   End If
End If
mytablex.Close
End If
If Len(barras) > 0 Then
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "select codigo,barras from clientes where barras='" & barras & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   If "" & mytablex.Fields("codigo") <> codigo Then
      MsgBox "Codigo barras ya usado en " + mytablex.Fields("codigo")
      mytablex.Close
      Exit Function
   End If
End If
mytablex.Close
End If
valida = 1
End Function
Function busca_clientes(buf As String) As String
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "select nombre from clientes where codigo='" & buf & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   busca_clientes = "" & mytablex.Fields("nombre")
End If
mytablex.Close

End Function
Sub agregar_telefono(buf As String)
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "select * from telefono where codigo='" & codigo & "' and telefono='" & buf & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount = 0 Then
   mytablex.AddNew
   mytablex.Fields("codigo") = codigo
   mytablex.Fields("telefono") = buf
   mytablex.Update
End If
mytablex.Close

End Sub



