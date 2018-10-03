VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form trecaja 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comprobantes de Caja"
   ClientHeight    =   8580
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   9870
   StartUpPosition =   1  'CenterOwner
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
      Height          =   8085
      Left            =   360
      TabIndex        =   130
      Top             =   8400
      Visible         =   0   'False
      Width           =   10320
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "Aceptar"
         Height          =   870
         Left            =   8385
         TabIndex        =   135
         Top             =   840
         Width           =   795
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
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
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
         Left            =   75
         Style           =   2  'Dropdown List
         TabIndex        =   132
         TabStop         =   0   'False
         Top             =   270
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
         Left            =   6630
         TabIndex        =   131
         TabStop         =   0   'False
         Top             =   225
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   7125
         Left            =   75
         TabIndex        =   134
         Top             =   855
         Width           =   8250
         _ExtentX        =   14552
         _ExtentY        =   12568
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF80&
      Height          =   5715
      Left            =   6375
      TabIndex        =   113
      Top             =   9450
      Visible         =   0   'False
      Width           =   4800
      Begin VB.TextBox letipo 
         Enabled         =   0   'False
         Height          =   495
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   122
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox lenumero 
         Height          =   495
         Left            =   1920
         MaxLength       =   11
         TabIndex        =   121
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox lefechai 
         Height          =   495
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   120
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox lefechaf 
         Height          =   495
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   119
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox lemoneda 
         Height          =   495
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   118
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox levalor 
         Height          =   495
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   117
         Top             =   3120
         Width           =   2055
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF80&
         Caption         =   "Grabar"
         Height          =   615
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   116
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFFF80&
         Caption         =   "Close"
         Height          =   615
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox leserie 
         Enabled         =   0   'False
         Height          =   495
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   114
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label36 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
         Height          =   495
         Left            =   270
         TabIndex        =   129
         Top             =   165
         Width           =   1815
      End
      Begin VB.Label Label37 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero"
         Height          =   495
         Left            =   120
         TabIndex        =   128
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label38 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaInicio"
         Height          =   495
         Left            =   120
         TabIndex        =   127
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label39 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaFinal"
         Height          =   495
         Left            =   120
         TabIndex        =   126
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label40 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Moneda"
         Height          =   495
         Left            =   120
         TabIndex        =   125
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label41 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         Height          =   495
         Left            =   120
         TabIndex        =   124
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label42 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie"
         Height          =   495
         Left            =   120
         TabIndex        =   123
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Canje Letra"
      Height          =   4185
      Left            =   5070
      TabIndex        =   93
      Top             =   9960
      Visible         =   0   'False
      Width           =   3705
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FFFF80&
         Caption         =   "Finalizar y Grabar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   3480
         Width           =   1935
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFF80&
         Caption         =   "Close"
         Height          =   615
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   2760
         Width           =   1935
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFF80&
         Caption         =   "Borra"
         Height          =   615
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFF80&
         Caption         =   "Modifica"
         Height          =   615
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFF80&
         Caption         =   "Nuevo"
         Height          =   615
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   960
         Width           =   1935
      End
      Begin MSDataGridLib.DataGrid dbgrid3 
         Height          =   4335
         Left            =   120
         TabIndex        =   99
         Top             =   2640
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   7646
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
            DataField       =   "Tipo"
            Caption         =   "Tipo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3179
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Serie"
            Caption         =   "Serie"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3179
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Numero"
            Caption         =   "Numero"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3179
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Fechai"
            Caption         =   "FechaI"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3179
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Fechaf"
            Caption         =   "Fechaf"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3179
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Moneda"
            Caption         =   "M"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3179
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Valor"
            Caption         =   "Valor"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3179
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1635.024
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   2190.047
            EndProperty
         EndProperty
      End
      Begin VB.Label fpago 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   112
         Top             =   1200
         Width           =   105
      End
      Begin VB.Label lstotal 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   6480
         TabIndex        =   111
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label34 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         Height          =   495
         Left            =   4680
         TabIndex        =   110
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lsmoneda 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   6480
         TabIndex        =   109
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label32 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Moneda"
         Height          =   495
         Left            =   4680
         TabIndex        =   108
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lsnombre 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   1920
         TabIndex        =   107
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label30 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         Height          =   495
         Left            =   120
         TabIndex        =   106
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lscodigo 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   1920
         TabIndex        =   105
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         Height          =   495
         Left            =   90
         TabIndex        =   104
         Top             =   255
         Width           =   1815
      End
      Begin VB.Label lssaldo 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   6240
         TabIndex        =   103
         Top             =   7680
         Width           =   1815
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Saldo"
         Height          =   495
         Left            =   4440
         TabIndex        =   102
         Top             =   7680
         Width           =   1815
      End
      Begin VB.Label lssubtotal 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   6240
         TabIndex        =   101
         Top             =   7200
         Width           =   1815
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Acumulado"
         Height          =   495
         Left            =   4440
         TabIndex        =   100
         Top             =   7200
         Width           =   1815
      End
   End
   Begin VB.ComboBox subconcepto 
      BackColor       =   &H00C0FFFF&
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
      Left            =   13875
      Style           =   2  'Dropdown List
      TabIndex        =   92
      Top             =   4740
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.ComboBox concepto 
      BackColor       =   &H00C0FFFF&
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
      Left            =   13875
      Style           =   2  'Dropdown List
      TabIndex        =   91
      Top             =   4380
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.ComboBox Combo2 
      Enabled         =   0   'False
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
      Left            =   2100
      Style           =   2  'Dropdown List
      TabIndex        =   90
      Top             =   1230
      Width           =   1500
   End
   Begin VB.CheckBox pagocash 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pago Cash"
      Height          =   495
      Left            =   3345
      TabIndex        =   89
      Top             =   7830
      Width           =   1695
   End
   Begin VB.TextBox local1 
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
      Left            =   9975
      MaxLength       =   3
      TabIndex        =   12
      Top             =   2805
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox seccion10 
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
      Left            =   12315
      MaxLength       =   10
      TabIndex        =   67
      Top             =   8865
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox seccion5 
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
      Left            =   10155
      MaxLength       =   10
      TabIndex        =   65
      Top             =   8865
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox seccion9 
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
      Left            =   12315
      MaxLength       =   10
      TabIndex        =   63
      Top             =   8505
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox seccion4 
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
      Left            =   10155
      MaxLength       =   10
      TabIndex        =   61
      Top             =   8505
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox seccion8 
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
      Left            =   12315
      MaxLength       =   10
      TabIndex        =   59
      Top             =   8145
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox seccion3 
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
      Left            =   10155
      MaxLength       =   10
      TabIndex        =   57
      Top             =   8145
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox seccion7 
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
      Left            =   12315
      MaxLength       =   10
      TabIndex        =   55
      Top             =   7785
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox seccion2 
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
      Left            =   13185
      MaxLength       =   10
      TabIndex        =   53
      Top             =   5640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox seccion6 
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
      Left            =   12315
      MaxLength       =   10
      TabIndex        =   51
      Top             =   7425
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox seccion1 
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
      Left            =   13185
      MaxLength       =   10
      TabIndex        =   47
      Top             =   5280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox moneda 
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
      Left            =   7680
      MaxLength       =   1
      TabIndex        =   7
      Text            =   "S"
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox tipoclie 
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
      Left            =   1500
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1215
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00808080&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   9810
      TabIndex        =   37
      Top             =   0
      Width           =   9870
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
         Picture         =   "trecaja.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   40
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
         Left            =   720
         Picture         =   "trecaja.frx":1212
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
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
         Left            =   0
         MaskColor       =   &H00E0E0E0&
         Picture         =   "trecaja.frx":2424
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Grabar registro"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox c11 
      BackColor       =   &H00FFFF00&
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
      Left            =   6600
      MaxLength       =   10
      TabIndex        =   31
      Top             =   9720
      Width           =   1215
   End
   Begin VB.TextBox c12 
      BackColor       =   &H00FFFF00&
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
      Left            =   6600
      MaxLength       =   10
      TabIndex        =   30
      Top             =   10080
      Width           =   1215
   End
   Begin VB.TextBox c13 
      BackColor       =   &H00FFFF00&
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
      Left            =   6600
      MaxLength       =   10
      TabIndex        =   29
      Top             =   9960
      Width           =   1215
   End
   Begin VB.TextBox c14 
      BackColor       =   &H00FFFF00&
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
      Left            =   6600
      MaxLength       =   10
      TabIndex        =   28
      Top             =   10320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Cobrar"
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
      Left            =   5385
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7830
      Width           =   2055
   End
   Begin VB.TextBox total 
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
      Height          =   495
      Left            =   5385
      MaxLength       =   10
      TabIndex        =   9
      Top             =   7230
      Width           =   2055
   End
   Begin VB.TextBox paridad 
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
      Left            =   9000
      MaxLength       =   10
      TabIndex        =   8
      Text            =   "1"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox vendedor 
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
      Left            =   7680
      MaxLength       =   11
      TabIndex        =   6
      Top             =   1560
      Width           =   2055
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
      Height          =   375
      Left            =   1485
      MaxLength       =   30
      TabIndex        =   5
      Top             =   2295
      Width           =   3495
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
      Left            =   1485
      MaxLength       =   60
      TabIndex        =   4
      Top             =   1905
      Width           =   3495
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
      Left            =   1500
      MaxLength       =   11
      TabIndex        =   3
      Top             =   1590
      Width           =   2055
   End
   Begin VB.TextBox fecha 
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
      Left            =   7680
      MaxLength       =   10
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox tipo 
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
      Left            =   1485
      MaxLength       =   3
      TabIndex        =   0
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox numero 
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
      Left            =   7680
      MaxLength       =   11
      TabIndex        =   10
      Top             =   855
      Width           =   2055
   End
   Begin VB.CommandButton Command10 
      Caption         =   "...."
      Height          =   360
      Left            =   5160
      TabIndex        =   138
      ToolTipText     =   "Obtiene Datos de Comprobantes seleccionados"
      Top             =   2280
      Width           =   405
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "trecaja.frx":3636
      Height          =   2895
      Left            =   240
      OleObjectBlob   =   "trecaja.frx":364A
      TabIndex        =   143
      TabStop         =   0   'False
      Top             =   3600
      Width           =   8460
   End
   Begin VB.TextBox fechai 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   7680
      MaxLength       =   10
      TabIndex        =   140
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox fechaf 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   7680
      MaxLength       =   10
      TabIndex        =   139
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label numerod 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "numerod"
      Height          =   195
      Left            =   7200
      TabIndex        =   147
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label seried 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "seried"
      Height          =   195
      Left            =   6600
      TabIndex        =   146
      Top             =   3240
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label tipod 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "tipod"
      Height          =   195
      Left            =   5760
      TabIndex        =   145
      Top             =   3240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label cargar 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cargar"
      Height          =   495
      Left            =   8760
      TabIndex        =   144
      Top             =   3780
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SaldoActual"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   420
      TabIndex        =   137
      Top             =   7230
      Width           =   1515
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo Anterior"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   420
      TabIndex        =   136
      Top             =   6810
      Width           =   1485
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3555
      Picture         =   "trecaja.frx":4B95
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   375
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3660
      Picture         =   "trecaja.frx":4E9F
      Stretch         =   -1  'True
      Top             =   2190
      Width           =   375
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2085
      Picture         =   "trecaja.frx":51A9
      Stretch         =   -1  'True
      Top             =   840
      Width           =   375
   End
   Begin VB.Label canjeletra 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   14160
      TabIndex        =   88
      Top             =   7680
      Width           =   105
   End
   Begin VB.Label xcuentacol 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   14160
      TabIndex        =   87
      Top             =   6960
      Width           =   105
   End
   Begin VB.Label XCUENTACO1 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   14160
      TabIndex        =   86
      Top             =   7440
      Width           =   105
   End
   Begin VB.Label xcuentaco 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   14160
      TabIndex        =   85
      Top             =   7200
      Width           =   105
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BorraFila"
      Height          =   495
      Left            =   8760
      TabIndex        =   84
      Top             =   5220
      Width           =   975
   End
   Begin VB.Label YACARGUE 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Height          =   195
      Left            =   11805
      TabIndex        =   83
      Top             =   5475
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label21 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Num."
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
      TabIndex        =   82
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Buscar"
      Height          =   495
      Left            =   8160
      TabIndex        =   81
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BorraTodos"
      Height          =   495
      Left            =   8760
      TabIndex        =   80
      Top             =   4740
      Width           =   975
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CargaTodos"
      Height          =   495
      Left            =   8760
      TabIndex        =   79
      Top             =   4260
      Width           =   975
   End
   Begin VB.Label Label17 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Doc.Sel."
      Height          =   495
      Left            =   3345
      TabIndex        =   78
      Top             =   6750
      Width           =   2055
   End
   Begin VB.Label totaldoc 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   5385
      TabIndex        =   77
      Top             =   6750
      Width           =   2055
   End
   Begin VB.Label anticipo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   11805
      TabIndex        =   76
      Top             =   5115
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label serie 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   10200
      TabIndex        =   75
      Top             =   3720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label vienede 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Height          =   195
      Left            =   9435
      TabIndex        =   74
      Top             =   9345
      Width           =   45
   End
   Begin VB.Label Label14 
      BackColor       =   &H00E0E0E0&
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
      Left            =   330
      TabIndex        =   73
      Top             =   855
      Width           =   1185
   End
   Begin VB.Label dia 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5880
      TabIndex        =   72
      Top             =   9600
      Width           =   105
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sumar"
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
      Left            =   11955
      TabIndex        =   71
      Top             =   9585
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dueños"
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
      Left            =   12585
      TabIndex        =   70
      Top             =   4650
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label tseccion 
      BackColor       =   &H00E0E0E0&
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
      Left            =   11955
      TabIndex        =   69
      Top             =   9225
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label26 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
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
      Left            =   10995
      TabIndex        =   68
      Top             =   9225
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label cseccion10 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "---"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1905
      TabIndex        =   66
      Top             =   7215
      Width           =   1245
   End
   Begin VB.Label cseccion5 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9195
      TabIndex        =   64
      Top             =   8865
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label cseccion9 
      BackColor       =   &H00E0E0E0&
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
      Height          =   435
      Left            =   1890
      TabIndex        =   62
      Top             =   6810
      Width           =   1275
   End
   Begin VB.Label cseccion4 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9195
      TabIndex        =   60
      Top             =   8505
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label cseccion8 
      BackColor       =   &H00E0E0E0&
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
      Left            =   11355
      TabIndex        =   58
      Top             =   8145
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label cseccion3 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9735
      TabIndex        =   56
      Top             =   8310
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label cseccion7 
      BackColor       =   &H00E0E0E0&
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
      Left            =   11355
      TabIndex        =   54
      Top             =   7785
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label cseccion2 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12585
      TabIndex        =   52
      Top             =   5730
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label cseccion6 
      BackColor       =   &H00E0E0E0&
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
      Left            =   11355
      TabIndex        =   50
      Top             =   7425
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label16 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Monto"
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
      Left            =   12315
      TabIndex        =   49
      Top             =   7065
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label15 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Seccion"
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
      Left            =   11355
      TabIndex        =   48
      Top             =   7065
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label cseccion1 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12585
      TabIndex        =   46
      Top             =   5370
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Monto"
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
      Left            =   13185
      TabIndex        =   45
      Top             =   4920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Seccion"
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
      Left            =   12585
      TabIndex        =   44
      Top             =   5010
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label turno 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2280
      TabIndex        =   43
      Top             =   9600
      Width           =   495
   End
   Begin VB.Label caja 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1560
      TabIndex        =   42
      Top             =   9600
      Width           =   615
   End
   Begin VB.Label cajero 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   41
      Top             =   9600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label bandera 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9120
      TabIndex        =   36
      Top             =   3120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label nc4 
      BackColor       =   &H00FFFF00&
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
      Left            =   5160
      TabIndex        =   35
      Top             =   10320
      Width           =   1455
   End
   Begin VB.Label nc3 
      BackColor       =   &H00FFFF00&
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
      Left            =   5160
      TabIndex        =   34
      Top             =   9960
      Width           =   1455
   End
   Begin VB.Label nc2 
      BackColor       =   &H00FFFF00&
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
      Left            =   5160
      TabIndex        =   33
      Top             =   10080
      Width           =   1455
   End
   Begin VB.Label nc1 
      BackColor       =   &H00FFFF00&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5160
      TabIndex        =   32
      Top             =   9720
      Width           =   1455
   End
   Begin VB.Label afecta 
      BackColor       =   &H00FFFF00&
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
      Left            =   4560
      TabIndex        =   27
      Top             =   9960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label saldos 
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1740
      TabIndex        =   26
      Top             =   2940
      Width           =   1455
   End
   Begin VB.Label acu 
      BackColor       =   &H00FFFF00&
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
      Left            =   4320
      TabIndex        =   25
      Top             =   9960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3345
      TabIndex        =   24
      Top             =   7230
      Width           =   2055
   End
   Begin VB.Label saldod 
      BackColor       =   &H00E0E0E0&
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
      Left            =   3180
      TabIndex        =   23
      Top             =   2940
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo Credito"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   300
      TabIndex        =   22
      Top             =   2940
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "T/C"
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
      TabIndex        =   21
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Moneda(S/D)"
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
      TabIndex        =   20
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cobr/vend"
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
      TabIndex        =   19
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Observac."
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
      Left            =   315
      TabIndex        =   18
      Top             =   2295
      Width           =   1200
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre"
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
      Left            =   300
      TabIndex        =   17
      Top             =   1935
      Width           =   1185
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
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
      Left            =   315
      TabIndex        =   16
      Top             =   1575
      Width           =   1185
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TCliente(CPV)"
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
      Left            =   315
      TabIndex        =   15
      Top             =   1215
      Width           =   1185
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
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
      Left            =   6240
      TabIndex        =   14
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9480
      TabIndex        =   13
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblFechaInicio 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F.Inicio"
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
      TabIndex        =   142
      Top             =   2280
      Width           =   1545
   End
   Begin VB.Label lblFechaFinal 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F.Final"
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
      TabIndex        =   141
      Top             =   2640
      Width           =   1425
   End
   Begin VB.Menu ch89343 
      Caption         =   "&Copia"
      Visible         =   0   'False
   End
   Begin VB.Menu d7823 
      Caption         =   "&Anula"
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "trecaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tabletra    As New ADODB.Recordset

'10/06/2017 kenyo cobro credito por rango de fecha
Dim opcionfocus As String

'10/06/2017 kenyo cobro credito por rango de fecha
     
Private Sub ajdu1_Click()

    If Frame1.Visible = True Then Exit Sub
    inicializa
    tipo = ""
    Numero = ""

    If tipo.Enabled = True Then
        tipo.SetFocus

    End If

End Sub

Private Sub bo712_Click()

End Sub

Private Sub buffer_DblClick()
    Command1_Click

End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        dlo132_Click
        Exit Sub

    End If

    Command1_Click

End Sub

Private Sub buffer_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode <> 13 And KeyCode <> 27 Then
        ejecuta 0

    End If

End Sub

Private Sub cargar_Click()

    Dim found As Integer

    found = busca_tipo(5)

    '10/06/2017 kenyo cobro credito por rango de fecha
    If Len(fechai) <> 10 Then Exit Sub
    If Len(fechaf) <> 10 Then Exit Sub
    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub
    '10/06/2017 kenyo cobro credito por rango de fecha

    If found <> 5 Then

    End If

    If Len(codigo) > 0 Then
        found = busca_codigo()

        If found = 0 Then Exit Sub

    End If

    If tipoclie = "P" Then 'proveedor
        'If afecta <> "L" Then
        found = busca_saldoCobroCredito()

        'End If
        'If afecta = "L" Then
        '   found = busca_saldo_letra1()
        'End If
    End If

    If tipoclie = "C" Or tipoclie = "V" Then 'proveedor
        found = busca_saldoCobroCredito()

    End If

    found = sumar_creditos()

    cseccion9.Caption = Format(Val(saldos), "0.00") 'saldo anterior
    cseccion10.Caption = Format((Val(saldos) - Val(total)), "0.00") ' saldo actual

End Sub

Private Sub ch89343_Click()

    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    consulta_copia

End Sub

Private Sub CmdAceptar_Click()
    CargaDatos

End Sub

Private Sub cmdAddEntry_Click()
    ajdu1_Click

End Sub

Private Sub cmdDelete_Click()

End Sub

Private Sub cmdExit_Click()
    dlo132_Click

End Sub

Private Sub cmdHelp_Click()

End Sub

Private Sub cmdPrint_Click()

End Sub

Private Sub cmdSave_Click()

    Dim found As Integer

    If Frame1.Visible = True Then Exit Sub
    found = grabar()

    If found = 0 Then Exit Sub
    Frame2.Visible = False
    proceso_impresion1 "" & local1, "" & tipo, "" & serie, "" & Numero, "" & acu

    '16/03/2018 No sale error en Ingreso/Egreso desde Menu
    ''''21/09/2017 kenyo Abrir gaveta al realizar recibos
    If AbreGaveta = "S" Then
        found = abre_puerto(Trim("" & mytable11.Fields("capuerto")), Val("" & mytable11.Fields("catipo")), "" & mytable11.Fields("gavetacola"))
        AbreGaveta = "N"

    End If

    ''''21/09/2017 kenyo Abrir gaveta al realizar recibos
    '16/03/2018 No sale error en Ingreso/Egreso desde Menu

    dlo132_Click
    'cmdAddEntry_Click
    'tipo.SetFocus

End Sub

Private Sub cmdSort_Click()

End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If valida_esencial() = 0 Then
        MsgBox "Aviso en valida esencial ", 48, "Aviso"

        If tipoclie.Enabled = True Then
            tipoclie.SetFocus

        End If

        Exit Sub

    End If

    If Len(codigo) > 0 Then
        found = busca_codigo()

    End If

    If tipoclie = "P" Or tipoclie = "C" Or tipoclie = "V" Then
        'If afecta <> "L" Then
        found = busca_saldo()

        'End If
        'If afecta = "L" Then
        '   found = busca_saldo_letra()
        'End If
    End If

    saldos = Format(suma1, "0.00")
    saldod = Format(suma2, "0.00")
    nombre.SetFocus

End Sub

Private Sub Codigo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        If tipoclie <> "C" And tipoclie <> "P" And tipoclie <> "V" Then
            If tipoclie.Enabled = True Then
                tipoclie.SetFocus

            End If

            Exit Sub

        End If

        consulta_codigo

    End If

    If KeyCode = &H26 Then
        If tipoclie.Enabled = False Then
            If fecha.Enabled = True Then
                fecha.SetFocus

            End If

            Exit Sub

        End If

        tipoclie.SetFocus
        Exit Sub

    End If

End Sub

Private Sub Combo2_Click()

    If Combo2.Text = "C" Or Combo2.Text = "V" Then
        tipoclie = Combo2.Text
        xcuentaco = "cuentac"
        XCUENTACO1 = "cuentacd"

    End If

    If Combo2.Text = "P" Then
        tipoclie = Combo2.Text
        xcuentaco = "cuentap"
        XCUENTACO1 = "cuentapd"

    End If

End Sub

Private Sub Command1_Click()
    ejecuta 1

End Sub

Sub ejecuta(sw As Integer)

    Dim buf       As String

    Dim buf1      As String

    Dim buf2      As String

    Dim buf3      As String

    Dim buf4      As String

    Dim rconsulta As New ADODB.Recordset

    'MsgBox afecta
    buf4 = "cpedidov"

    If tipoclie = "C" Then
        buf1 = "clientes"
        buf2 = "cuentac"
        buf3 = "letrav"

    End If

    If tipoclie = "P" Then
        buf1 = "proveedo"
        buf2 = "cuentap"
        buf3 = "letrac"

    End If

    If tipoclie = "V" Then  '
        buf1 = "vendedor"
        buf2 = "cuentac"
        buf3 = "letrav"

    End If

    If opcion1 = "1280" Then   'ver producto
        'MsgBox Data2.Recordset.Fields("tipo") & " " & Data2.Recordset.Fields("serie")
        buf = "select Descripcio,Producto,Unidad as Und,Factor as Fx,Cantidad as Cant,Precio,Total,Observa1,Observa2,Observa3,Observa4 from detalle where   tipo='" & "" & Data2.Recordset.Fields("tipo1") & "' and serie='" & "" & Data2.Recordset.Fields("serie1") & "' and numero='" & "" & Data2.Recordset.Fields("numero1") & "'"

        'MsgBox buf
    End If

    If opcion1 = "12" Or opcion1 = "13" Or opcion1 = "14" Or opcion1 = "15" Then  'CUENTA CORRIENTE
        If Len(buffer) = 0 Then
            buf = "select Local,tipo,serie,numero,Cuota,Moneda as M,Total,Saldo,Fecha,fechav,c1,c2,c3,c4 from " & xcuentaco & " where fpago='C' and tipoclie='" & tipoclie & "' and codigo='" & codigo & "' and saldo>0"
        Else
            buf = "select Local,tipo,serie,numero,Cuota,Moneda as M,Total,Saldo,Fecha,fechav,c1,c2,c3,c4 from  " & xcuentaco & " where fpago='C' and tipoclie='" & tipoclie & "' and codigo='" & codigo & "' and " & Combo1 & " like '" & buffer & "%' and saldo>0"

        End If

    End If

    If opcion1 = "3" Or opcion1 = "4" Or opcion1 = "5" Or opcion1 = "6" Then
        If Len(buffer) = 0 Then
            buf = "select tipo,serie,numero,cuota,fecha,fechav,Moneda as M,Total,saldo,c1,c2,c3,c4 from " & xcuentaco & " where fpago='C' and tipoclie='" & tipoclie & "' and codigo='" & codigo & "' order by fechav "
        Else
            buf = "select tipo,serie,numero,cuota,fecha,fechav,Moneda as M,Total,saldo,c1,c2,c3,c4 from  " & xcuentaco & " where fpagov='C' and tipoclie='" & tipoclie & "' and codigo='" & codigo & "' and " & Combo1 & " like '" & buffer & "%' order by fechav"

        End If

    End If

    If opcion1 = "ANULA" Or opcion1 = "COPIA" Then
        If Len(buffer) = 0 Then
            buf = "select Local,Estado as E,tipo,Serie,Numero,fecha,Codigo,Nombre,Moneda as M,Total,Usuario,caja,Turno,tipoclie,acu from recibo where local='" & local1 & "' and acu='" & acu & "' and usuario='" & cajero & "' and caja='" & caja & "' and turno='" & turno & "'"
            buf = buf & " and fecha='" & Format(dia, "YYYYMMDD") & "'"
        Else
            buf = "select Local,Estado as E,tipo,Serie,Numero,fecha,Codigo,Nombre,Moneda as M,Total,Usuario,Caja,Turno,tipoclie,acu from recibo where local='" & local1 & "' and acu='" & acu & "'  and usuario='" & cajero & "' and caja='" & caja & "' and turno='" & turno & "' " & Combo1 & " like '" & buffer & "%'"
            buf = buf & " and  fecha='" & Format(dia, "YYYYMMDD") & "'"

        End If

    End If

    If opcion1 = "1" Then
        If Len(buffer) = 0 Then
            buf = "select Descripcio,tipo from tipo where tipodoc='" & acu & "' order by tipo "
        Else
            buf = "select Descripcio,tipo from tipo where  tipodoc='" & acu & "' and " & Combo1 & " like '" & buffer & "%' order by tipo"

        End If

    End If

    'If opcion1 = "1" Then
    '   If Len(buffer) = 0 Then
    '  buf = "select Descripcio,tipo from tipo where (tipodoc='V' or tipodoc='W') order by tipo "
    '  Else
    '  buf = "select Descripcio,tipo from tipo where  (tipodoc='V' or tipodoc='W') and " & Combo1 & " like '" & buffer & "%' order by tipo"
    '  End If
    'End If
   
    If opcion1 = "20" Then
        If Len(buffer) = 0 Then
            buf = "select Tipo,Numero,Fecha,Codigo,Nombre,Moneda as M,Total from recibo where tipo='" & tipo & "'"
        Else
            buf = "select Tipo,Numero,Fecha,Codigo,Nombre,Moneda as M,Total from recibo where tipo='" & tipo & "' and " & Combo1 & " like '" & buffer & "%'"

        End If

    End If

    If opcion1 = "2" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo from  " & buf1
        Else
            buf = "select Nombre,Codigo from " & buf1 & " where " & Combo1 & " like '" & buffer & "%'"

        End If
      
    End If

    If opcion1 = "22" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo from  vendedor "
        Else
            buf = "select Nombre,Codigo from vendedor where " & Combo1 & " like '" & buffer & "%'"

        End If
      
    End If

    If opcion1 = "231" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo from  tlocal"
        Else
            buf = "select Nombre,Codigo from tlocal where " & Combo1 & " like '" & buffer & "%'"

        End If
      
    End If

    'MsgBox buf
    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open buf, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = rconsulta
   
    '24/06/2017 kenyo CORRECCION acepta-busqueda
    If rconsulta.RecordCount = 0 Then
        CmdAceptar.Enabled = False
    Else
        CmdAceptar.Enabled = True

    End If

    '24/06/2017 kenyo CORRECCION acepta-busqueda
   
    If rconsulta.EOF = True And rconsulta.BOF = True Then
        rconsulta.Close
        buffer.SetFocus
        Exit Sub

    End If
            
    If opcion1 = "ANULA" Or opcion1 = "COPIA" Then
        dbGrid1.columns(0).Width = 500
        dbGrid1.columns(1).Width = 500
        dbGrid1.columns(2).Width = 800
        dbGrid1.columns(3).Width = 800
        dbGrid1.columns(4).Width = 1300
        dbGrid1.columns(5).Width = 1300
        dbGrid1.columns(6).Width = 1500
        dbGrid1.columns(7).Width = 3500
        dbGrid1.columns(8).Width = 600
        dbGrid1.columns(9).Width = 1300

    End If
               
    If opcion1 = "12" Or opcion1 = "13" Or opcion1 = "14" Or opcion1 = "15" Then
        dbGrid1.columns(0).Width = 700
        dbGrid1.columns(1).Width = 700
        dbGrid1.columns(2).Width = 700
        dbGrid1.columns(3).Width = 1300
        dbGrid1.columns(4).Width = 500
        dbGrid1.columns(5).Width = 500
        dbGrid1.columns(6).Width = 1500
        dbGrid1.columns(7).Width = 1500
        dbGrid1.columns(8).Width = 1500
        dbGrid1.columns(9).Width = 1500
        dbGrid1.columns(10).Width = 500
        dbGrid1.columns(11).Width = 500
        dbGrid1.columns(12).Width = 500
        dbGrid1.columns(13).Width = 500

    End If

    If opcion1 = "20" Then
        dbGrid1.columns(0).Width = 800
        dbGrid1.columns(1).Width = 1300
        dbGrid1.columns(2).Width = 1300
        dbGrid1.columns(3).Width = 1500
        dbGrid1.columns(4).Width = 3500
        dbGrid1.columns(5).Width = 800
        dbGrid1.columns(6).Width = 1000

    End If
               
    If opcion1 = "3" Then
        dbGrid1.columns(0).Width = 800
        dbGrid1.columns(1).Width = 800
        dbGrid1.columns(2).Width = 1500
        dbGrid1.columns(3).Width = 800
        dbGrid1.columns(4).Width = 1500
        dbGrid1.columns(5).Width = 1500
        dbGrid1.columns(6).Width = 500
        dbGrid1.columns(7).Width = 1200
        dbGrid1.columns(8).Width = 1200
        dbGrid1.columns(9).Width = 1200
        dbGrid1.columns(10).Width = 1200
        dbGrid1.columns(11).Width = 1200
        dbGrid1.columns(12).Width = 1200

    End If

    If opcion1 = "8" Then
        dbGrid1.columns(0).Width = 1300
        dbGrid1.columns(1).Width = 1500
        dbGrid1.columns(2).Width = 1500
        dbGrid1.columns(3).Width = 1500
        dbGrid1.columns(4).Width = 500
        dbGrid1.columns(5).Width = 1000
        dbGrid1.columns(6).Width = 1000
        dbGrid1.columns(7).Width = 500
        dbGrid1.columns(8).Width = 500
        dbGrid1.columns(9).Width = 500
        dbGrid1.columns(10).Width = 500
                  
    End If

    If opcion1 = "231" Or opcion1 = "1" Or opcion1 = "2" Or opcion1 = "4" Or opcion1 = "5" Or opcion2 = "6" Or opcion1 = "7" Or opcion2 = "8" Or opcion1 = "9" Or opcion1 = "10" Or opcion1 = "11" Or opcion1 = "22" Then
        dbGrid1.columns(0).Width = 4000
        dbGrid1.columns(1).Width = 2000

    End If

    If opcion1 = "1280" Then
        dbGrid1.columns(0).Width = 3000
        dbGrid1.columns(1).Width = 1500
        dbGrid1.columns(2).Width = 1000
        dbGrid1.columns(3).Width = 1000
        dbGrid1.columns(4).Width = 1000

    End If

    If sw = 1 Then
        dbGrid1.SetFocus

    End If

End Sub

Private Sub Command10_Click()

    Dim items As String

    items = DBGrid2.VisibleRows

    If items > 0 Then
        observa.Text = ""

        Dim valor As String

        Dim I     As Integer

        Dim Texto As String

        valor = DBGrid2.VisibleRows

        For I = 0 To valor - 1
            DBGrid2.Row = I
            DBGrid2.Col = DBGrid2.VisibleRows - 1
            DBGrid2.SetFocus
            Texto = Texto + DBGrid2.columns(2) + "-" + DBGrid2.columns(3) + " / "
            observa.Text = Texto
        Next I

        observa = Mid$(observa, 1, (Len(observa) - 3))

    End If

End Sub

Private Sub Command2_Click()

    ''' 25/08/2017 kenyo. Cobro Credito desde ventana de ventas
    Dim found As Integer

    found = sumar_creditos()
    ''' 25/08/2017 kenyo. Cobro Credito desde ventana de ventas

    total_KeyPress 13

End Sub

Private Sub Command3_Click()
    Frame3.Visible = True
    inicializa_letra
    Frame2.Enabled = False
    Frame3.Caption = "Nuevo"
    letipo.Enabled = True
    leserie.Enabled = True
    lenumero.Enabled = True
    lenumero.SetFocus

End Sub

Private Sub Command4_Click()

    On Error GoTo cmd99002_err

    letipo = "" & tabletra.Fields("tipo")
    leserie = "" & tabletra.Fields("serie")
    lenumero = "" & tabletra.Fields("numero")
    letipo.Enabled = False
    leserie.Enabled = False
    lenumero.Enabled = False

    lefechai = Format("" & tabletra.Fields("fechai"), "dd/mm/yyyy")
    lefechaf = Format("" & tabletra.Fields("fechaf"), "dd/mm/yyyy")
    lemoneda = "" & tabletra.Fields("moneda")
    levalor = "" & tabletra.Fields("valor")
    Frame3.Caption = "Modifica"
    Frame3.Visible = True
    Frame2.Enabled = False
    Exit Sub
cmd99002_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command5_Click()

    On Error GoTo cmd9900_err

    tabletra.Delete
    suma_tabletra
    Exit Sub
cmd9900_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command6_Click()

    If tabletra.State = 1 Then tabletra.Close
    Frame2.Visible = False

End Sub

Private Sub Command7_Click()

    Dim found As Integer

    If Len(letipo) = 0 Then
        letipo.SetFocus
        Exit Sub

    End If

    If Len(lenumero) = 0 Then
        lenumero.SetFocus
        Exit Sub

    End If

    If Len(lefechai) = 0 Then
        lefechai.SetFocus
        Exit Sub

    End If

    If Len(lefechaf) = 0 Then
        lefechaf.SetFocus
        Exit Sub

    End If

    If Len(lemoneda) = 0 Then
        lemoneda.SetFocus
        Exit Sub

    End If

    If Len(levalor) = 0 Then
        levalor.SetFocus
        Exit Sub

    End If

    If lemoneda <> "S" And lemoneda <> "D" Then
        moneda.SetFocus
        Exit Sub

    End If

    If Frame3.Caption = "Modifica" Then
        tabletra.Fields("tipo") = letipo
        tabletra.Fields("serie") = leserie
        tabletra.Fields("numero") = lenumero
        tabletra.Fields("fechai") = Format(lefechai, "")
        tabletra.Fields("fechaf") = Format(lefechaf, "")
        tabletra.Fields("moneda") = lemoneda
        tabletra.Fields("valor") = Val(levalor)
        tabletra.Update
        Frame3.Visible = False
        Frame2.Enabled = True
        suma_tabletra
        Exit Sub
   
    End If

    If Frame3.Caption = "Nuevo" Then
        found = existe_letra()

        If found = 1 Then
            MsgBox "Ya existe numero Seleccionado ", 48, "Aviso"
            Exit Sub

        End If

        found = existe_letrad()

        If found = 1 Then
            MsgBox "Ya existe numero Utilizado en otro cliente ", 48, "Aviso"
            Exit Sub

        End If

        tabletra.AddNew
        tabletra.Fields("tipo") = letipo
        tabletra.Fields("serie") = leserie
        tabletra.Fields("numero") = lenumero
        tabletra.Fields("fechai") = Format(lefechai, "")
        tabletra.Fields("fechaf") = Format(lefechaf, "")
        tabletra.Fields("moneda") = lemoneda
        tabletra.Fields("valor") = Val(levalor)
        tabletra.Update
        Frame2.Enabled = True
        Frame3.Visible = False
        suma_tabletra
        Exit Sub

    End If

    Exit Sub

End Sub

Private Sub Command8_Click()
    Frame2.Enabled = True
    Frame3.Visible = False

End Sub

Private Sub Command9_Click()

    If tabletra.RecordCount = 0 Then
        MsgBox "No hay Registros ", 48, "Aviso"
        Exit Sub

    End If

    cmdSave_Click
    Frame2.Visible = False

End Sub

Private Sub concepto_Click()

    Dim buf As String

    buf = extra_loquesea(concepto)

    If buf <> "%" Then
        carga_subconcepto buf

    End If

End Sub

Private Sub d7823_Click()

    If Frame2.Visible = True Then Exit Sub

    If Frame1.Visible = True Then Exit Sub
    consulta_anula

End Sub

Sub CargaDatos()

    Dim found As Integer

    '14/06/2017 kenyo No se cuelga el Sistema al aceptar
    'If dbgrid1.VisibleRows =  Then
    '  Exit Sub
    'End If
    '14/06/2017 kenyo No se cuelga el Sistema al aceptar

    If opcion1 = "1280" Then
        'tipo = Trim(dbGrid1.columns(1))
        Frame1.Enabled = False
        Frame1.Visible = False
        'tipo.SetFocus
        'tipo_KeyPress 13
        Exit Sub

    End If

    If opcion1 = "1" Then
        tipo = Trim(dbGrid1.columns(1))
        Frame1.Enabled = False
        Frame1.Visible = False
        tipo.SetFocus
        tipo_KeyPress 13

    End If

    If opcion1 = "ANULA" Then
        If "" & dbGrid1.columns("e") = "1" Then
            MsgBox "Ya se encuentra anulado ", 48, "Aviso"
            dbGrid1.SetFocus
            Exit Sub

        End If

        If MsgBox("Desea Anular ", 1, "Aviso") <> 1 Then Exit Sub
        tipoclie = Trim("" & dbGrid1.columns("tipoclie"))
        found = anular_recibo(Trim("" & dbGrid1.columns("local")), Trim("" & dbGrid1.columns("tipo")), Trim("" & dbGrid1.columns("serie")), Trim("" & dbGrid1.columns("numero")))

        If found = 1 Then
            MsgBox "Anulacion satisfactoria", 48, "Aviso"
            proceso_impresion1 Trim("" & dbGrid1.columns("local")), Trim("" & dbGrid1.columns("tipo")), Trim("" & dbGrid1.columns("serie")), Trim("" & dbGrid1.columns("numero")), Trim("" & dbGrid1.columns("acu"))

        End If

        Frame1.Visible = False
        Frame1.Enabled = False

        If tipo.Enabled = True Then
            tipo.SetFocus

        End If

        Exit Sub

    End If

    If opcion1 = "COPIA" Then
        tipoclie = Trim("" & dbGrid1.columns("tipoclie"))
        proceso_impresion1 Trim("" & dbGrid1.columns("local")), Trim("" & dbGrid1.columns("tipo")), Trim("" & dbGrid1.columns("serie")), Trim("" & dbGrid1.columns("numero")), Trim("" & dbGrid1.columns("acu"))
        Frame1.Visible = False
        Frame1.Enabled = False

        If tipo.Enabled = True Then
            tipo.SetFocus

        End If

        Exit Sub

    End If

    If opcion1 = "12" Then  'selecionar la cuenta corriente

        'pone los registros
        If existe_seleccionado(Trim("" & dbGrid1.columns("local")), Trim("" & dbGrid1.columns("tipo")), Trim("" & dbGrid1.columns("serie")), Trim("" & dbGrid1.columns("numero"))) = 1 Then
            MsgBox "Documento ya cargado ", 48, "Aviso"
            dbGrid1.SetFocus
            Exit Sub

        End If

        Data2.Recordset.AddNew
        Data2.Recordset.Fields("tipoclie") = "" & tipoclie
        Data2.Recordset.Fields("codigo") = "" & codigo
        Data2.Recordset.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
        Data2.Recordset.Fields("hora") = Format(Now, "HH:MM:SS")
        Data2.Recordset.Fields("tipo") = "" & tipo
        Data2.Recordset.Fields("local") = "" & local1
        Data2.Recordset.Fields("serie") = "" & serie
        Data2.Recordset.Fields("numero") = "" & Numero
        Data2.Recordset.Fields("acu") = "" & acu
        Data2.Recordset.Fields("usuario") = ""
        Data2.Recordset.Fields("paridad") = 0
        Data2.Recordset.Fields("local1") = Trim("" & dbGrid1.columns("local"))
        Data2.Recordset.Fields("tipo1") = Trim("" & dbGrid1.columns("tipo"))
        Data2.Recordset.Fields("serie1") = Trim("" & dbGrid1.columns("serie"))
        Data2.Recordset.Fields("numero1") = Trim("" & dbGrid1.columns("numero"))
        Data2.Recordset.Fields("cuota") = Trim("" & dbGrid1.columns("cuota"))
        Data2.Recordset.Fields("moneda") = Trim("" & dbGrid1.columns("M"))
        Data2.Recordset.Fields("total") = Val("" & dbGrid1.columns("saldo"))
        Data2.Recordset.Fields("paga") = Val("" & dbGrid1.columns("saldo"))
        Data2.Recordset.Fields("estado") = "2"
        Data2.Recordset.Update
        Data2.refresh
        found = sumar_creditos()
        Data2.refresh
        Frame1.Visible = False
        Frame1.Enabled = False
        DBGrid2.refresh
        DBGrid2.SetFocus

    End If

    If opcion1 = "3" Then

        'pone los registros
        If existe_seleccionado(Trim("" & dbGrid1.columns("local")), Trim("" & dbGrid1.columns("tipo")), Trim("" & dbGrid1.columns("serie")), Trim("" & dbGrid1.columns("numero"))) = 1 Then
            MsgBox "Documento ya cargado ", 48, "Aviso"
            dbGrid1.SetFocus
            Exit Sub

        End If

        Data2.Recordset.AddNew
        Data2.Recordset.Fields("codigo") = codigo
        Data2.Recordset.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
        Data2.Recordset.Fields("hora") = Format(Now, "HH:MM:SS")
        Data2.Recordset.Fields("tipo") = "" & tipo
        Data2.Recordset.Fields("serie") = "" & serie
        Data2.Recordset.Fields("numero") = "" & Numero
        Data2.Recordset.Fields("acu") = "" & acu
        Data2.Recordset.Fields("usuario") = ""
        Data2.Recordset.Fields("paridad") = 0
        Data2.Recordset.Fields("tipo1") = Trim(dbGrid1.columns(0))
        Data2.Recordset.Fields("serie1") = Trim(dbGrid1.columns(1))
        Data2.Recordset.Fields("numero1") = Trim(dbGrid1.columns(2))
        Data2.Recordset.Fields("cuota") = Trim(dbGrid1.columns(3))
        Data2.Recordset.Fields("moneda") = Trim(dbGrid1.columns(6))
        Data2.Recordset.Fields("total") = Trim(dbGrid1.columns(8))
        Data2.Recordset.Fields("paga") = Trim(dbGrid1.columns(8))
        'Data2.Recordset.Fields("L1") = Val("" & dbGrid1.Columns(9))
        'Data2.Recordset.Fields("L2") = Val("" & dbGrid1.Columns(10))
        'Data2.Recordset.Fields("L3") = Val("" & dbGrid1.Columns(11))
        'Data2.Recordset.Fields("L4") = Val("" & dbGrid1.Columns(12))
        Data2.Recordset.Fields("estado") = "2"
        Data2.Recordset.Update
        Frame1.Visible = False
        Frame1.Enabled = False
        DBGrid2.SetFocus

    End If

    If opcion1 = "8" Then 'si es letra

        'pone los registros
        If existe_seleccionado("", "LE", "LE", Trim(dbGrid1.columns(0))) = 1 Then
            MsgBox "Documento ya cargado ", 48, "Aviso"
            dbGrid1.SetFocus
            Exit Sub

        End If

        Data2.Recordset.AddNew
        Data2.Recordset.Fields("codigo") = codigo
        Data2.Recordset.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
        Data2.Recordset.Fields("hora") = Format(Now, "HH:MM:SS")
        Data2.Recordset.Fields("tipo") = "" & tipo
        Data2.Recordset.Fields("serie") = "" & serie
        Data2.Recordset.Fields("numero") = "" & Numero
        Data2.Recordset.Fields("acu") = "" & acu
        Data2.Recordset.Fields("usuario") = ""
        Data2.Recordset.Fields("paridad") = 0
        Data2.Recordset.Fields("tipo1") = "LE"
        Data2.Recordset.Fields("serie1") = "LE"
        Data2.Recordset.Fields("numero1") = Trim(dbGrid1.columns(0))
        Data2.Recordset.Fields("cuota") = "1"
        Data2.Recordset.Fields("moneda") = Trim(dbGrid1.columns(4))
        Data2.Recordset.Fields("total") = Trim(dbGrid1.columns(5))
        'Data2.Recordset.Fields("L1") = DBGrid1.Columns(7)
        'Data2.Recordset.Fields("L2") = DBGrid1.Columns(8)
        'Data2.Recordset.Fields("L3") = DBGrid1.Columns(9)
        'Data2.Recordset.Fields("L4") = DBGrid1.Columns(10)
        Data2.Recordset.Fields("estado") = "2"
        Data2.Recordset.Update
        Frame1.Visible = False
        Frame1.Enabled = False
        DBGrid2.SetFocus

    End If

    If opcion1 = "2" Then
        codigo = Trim(dbGrid1.columns(1))
        Frame1.Visible = False
        Frame1.Enabled = False
        codigo.SetFocus
        codigo_KeyPress 13

    End If

    If opcion1 = "22" Then
        vendedor = Trim(dbGrid1.columns(1))
        Frame1.Visible = False
        Frame1.Enabled = False
        vendedor.SetFocus
        vendedor_KeyPress 13

    End If

    If opcion1 = "231" Then
        local1 = Trim(dbGrid1.columns(1))
        Frame1.Visible = False
        Frame1.Enabled = False
        local1.SetFocus
        local1_KeyPress 13

    End If

    If opcion1 = "20" Then
        Numero = Trim(dbGrid1.columns(1))
        Frame1.Visible = False
        Frame1.Enabled = False
        Numero.SetFocus
        numero_KeyPress 13

    End If

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        CargaDatos

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

Private Sub dbgrid1_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim buf  As String

    Dim buf2 As String

    Dim sw   As Integer

    If KeyCode <> 13 And KeyCode <> 27 Then

        'MsgBox KeyCode
        If KeyCode >= 48 And KeyCode <= 57 Then
            GoTo sigue9

        End If

        If KeyCode >= 65 And KeyCode <= 90 Then
            GoTo sigue9

        End If

        If KeyCode >= 97 And KeyCode <= 122 Then
            GoTo sigue9

        End If

        If KeyCode = 8 And Chr(KeyCode) = "*" Then
            GoTo sigue9

        End If

        Exit Sub
sigue9:

        If KeyCode = 8 Then
            If Len(buffer) > 0 Then
                buf = Mid$(buffer, 1, Len(buffer) - 1)
                buffer = buf
                KeyCode = 0
            Else
                KeyCode = 0
                Exit Sub

            End If

        End If

        buf = Chr(KeyCode)

        If Chr(KeyCode) = "*" Then
            buf = ""
            buffer = buf

        End If

        If KeyCode <> 13 Then
            buffer = buffer + buf

        End If

        buf = buffer
        ejecuta 0

    End If

End Sub

Private Sub dbgrid2_AfterColUpdate(ByVal ColIndex As Integer)

    Select Case ColIndex

        Case 6

    End Select
            
End Sub

Private Sub dbgrid2_BeforeColEdit(ByVal ColIndex As Integer, _
                                  ByVal KeyAscii As Integer, _
                                  Cancel As Integer)

    Select Case ColIndex

        Case 0, 1, 2, 3, 4, 5, 8
            Cancel = True
            Exit Sub

        Case 6

            If Len("" & DBGrid2.columns(0)) = 0 Then
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

        Case 6

            If Not IsNumeric(DBGrid2.columns(6)) Then
                Cancel = True
                Exit Sub

            End If

            If Val("" & DBGrid2.columns(5)) < Val("" & DBGrid2.columns(6)) Then
                Cancel = True
                Exit Sub

            End If
            
            '10/06/2017 kenyo cobro credito por rango de fecha
        Case 7

            If Not IsNumeric(DBGrid2.columns(7)) Then
                Cancel = True
                Exit Sub

            End If

            If Val("" & DBGrid2.columns(6)) < Val("" & DBGrid2.columns(7)) Then
                Cancel = True
                MsgBox "Monto no puede ser mayor al total !", vbCritical, "Message"
                opcionfocus = "S"
                Exit Sub
            Else
                opcionfocus = "N"
             
            End If

            '10/06/2017 kenyo cobro credito por rango de fecha
            
    End Select

End Sub

Private Sub dbgrid2_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim found As Integer

    Dim sdx1  As Double

    If KeyCode = 13 Then
        found = sumar_creditos()
   
        If saldos <> "" Then
            cseccion9.Caption = Format(Val(saldos), "0.00") 'saldo anterior
            cseccion10.Caption = Format((Val(saldos) - Val(total)), "0.00") ' saldo actual

        End If
      
        '10/06/2017 kenyo cobro credito por rango de fecha
        'total.SetFocus
        If opcionfocus = "N" Then total.SetFocus
        '10/06/2017 kenyo cobro credito por rango de fecha
  
        Exit Sub

    End If

    'If KeyCode = &H70 Then  'f1
    '   Label10_Click
    'End If

End Sub

Private Sub DBGrid2_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim found As Integer

    On Error GoTo cmd135_err

    If KeyCode = &H26 Then
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        consulta_productos
        Exit Sub

    End If

    If KeyCode = &H2E Then  'borrar linea
        Data2.Recordset.Delete
        'Data2.refresh
        found = sumar_creditos()
        total.SetFocus

    End If

    Exit Sub
cmd135_err:
    Exit Sub

End Sub

Private Sub djuer1_Click()

End Sub

Private Sub dlo132_Click()

    If Frame3.Visible = True Then
        Frame3.Visible = False
        Frame2.Enabled = True
        Exit Sub

    End If

    If Frame2.Visible = True Then
        Frame2.Visible = False
        Exit Sub

    End If

    If Frame1.Visible = True Then
        If opcion1 = "ANULA" Or opcion1 = "COPIA" Then
            Frame1.Visible = False

            If Numero.Enabled = True Then
                Numero.SetFocus

            End If

            Exit Sub

        End If

        If opcion1 = "1280" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            DBGrid2.SetFocus
            Exit Sub

        End If
   
        If opcion1 = "20" Then
            Frame1.Visible = False
            Frame1.Enabled = False

            If Numero.Enabled = True Then
                Numero.SetFocus

            End If

            Exit Sub

        End If

        If opcion1 = "1" Then
            Frame1.Visible = False
            Frame1.Enabled = False

            If tipo.Enabled = True Then
                tipo.SetFocus

            End If

            Exit Sub

        End If

        If opcion1 = "231" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            local1.SetFocus
            Exit Sub

        End If

        If opcion1 = "2" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            codigo.SetFocus
            Exit Sub

        End If

        If opcion1 = "3" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            total.SetFocus
            Exit Sub

        End If

        If opcion1 = "12" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            total.SetFocus
            Exit Sub

        End If

        If opcion1 = "8" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            total.SetFocus
            Exit Sub

        End If

        Exit Sub
   
    End If

    trecaja.Hide
    Unload trecaja

End Sub

Private Sub fecha_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If valida_esencial() = 0 Then Exit Sub
    If Len(fecha) = 0 Then
        fecha = Format(Now, "dd/mm/yyyy")

    End If

    If Len(fecha) <> 10 Then Exit Sub
    If Not IsDate(fecha) Then
        fecha = ""
        Exit Sub

    End If

    If tipoclie.Enabled = False Then
        If codigo.Enabled = False Then
            observa.SetFocus
            Exit Sub

        End If

        Exit Sub

    End If

    If tipoclie.Enabled = True Then
        tipoclie.SetFocus
        Exit Sub

    End If

    If codigo.Enabled = True Then
        codigo.SetFocus
        Exit Sub

    End If

    observa.SetFocus

End Sub

Private Sub fecha_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        If tipo.Enabled = True Then
            tipo.SetFocus

        End If

        Exit Sub

    End If

End Sub

Private Sub Form_Activate()

    Dim found As Integer

    carga_seccion
    carga_tmpcta

    If Len(caja) = 0 Then
        caja = "00"

    End If

    If acu = "V" Then
        Combo2.Clear
        Combo2.AddItem "%"
        Combo2.AddItem "C"
        Combo2.AddItem "P"
        Combo2.AddItem "V"
        Combo2.ListIndex = 0

    End If

    If acu = "W" Then
        Combo2.Clear
        Combo2.AddItem "%"
        Combo2.AddItem "C"
        Combo2.AddItem "V"
        Combo2.ListIndex = 0

    End If

    '10/06/2017 kenyo cobro credito por rango de fecha
    fechaf = Format(Now, "dd/mm/yyyy")
    fechai = "01" & "/" & "01" & "/" & Format(Year(Now), "0000")
    '10/06/2017 kenyo cobro credito por rango de fecha

    If vienede = "CXC" Then
        'Data2.Database.Execute "DELETE FROM " & "_r" & gusuario
        'Data2.Refresh
        carga_viene
        found = sumar_creditos()

    End If

    If Len(xcuentaco) = 0 Then
        MsgBox "No existe xcuentaco", 48, "Aviso"

    End If

    carga_concepto
    subconcepto.AddItem "%"

    If cargar.Visible = True Then
        codigo_KeyPress 13
        cargar_Click
    
    End If

End Sub

Sub carga_tmpcta()
    Data2.Connect = "foxpro 2.5;"
    Data2.DatabaseName = globaldir
    Data2.RecordSource = "select * from _r" & gusuario & " where len(numero)>0  "
    Data2.refresh
             
End Sub

Sub inicializa_letra()
    letipo = "LE"
    leserie = "001"
    lenumero = ""
    lefechai = Format(Now, "dd/mm/yyyy")
    lefechaf = Format(Now, "dd/mm/yyyy")
    lemoneda = moneda
    levalor = lssaldo

End Sub

Sub carga_viene()

    Dim found As Integer

    If tipoclie = "P" Or tipoclie = "C" Or tipoclie = "V" Then
        found = busca_saldo()

    End If

    saldos = Format(suma1, "0.00")
    saldod = Format(suma2, "0.00")
    Data2.refresh
    Do

        If Data2.Recordset.EOF Then Exit Do
        Data2.Recordset.MoveNext
    Loop

End Sub

Private Sub Form_Load()

    Dim found As Integer

    afecta = ""
    'local1 = glocal
    found = busca_paridad()
    Combo1.Clear
    Combo1.AddItem "Descripcio"
    Combo1.AddItem "tipo"
    Combo1.ListIndex = 0
    fecha = Format(Now, "dd/mm/yyyy")

    Frame1.Top = 735
    Frame1.Left = 60

End Sub

Sub inicializa()

    Dim found As Integer

    afecta = ""
    anticipo = ""
    seccion1 = ""
    seccion2 = ""
    seccion3 = ""
    seccion4 = ""
    seccion5 = ""
    seccion6 = ""
    seccion7 = ""
    seccion8 = ""
    seccion9 = ""
    seccion10 = ""

    cseccion1 = ""
    cseccion2 = ""
    cseccion3 = ""
    cseccion4 = ""
    cseccion5 = ""
    cseccion6 = ""
    cseccion7 = ""
    cseccion8 = ""
    cseccion9 = ""
    cseccion10 = ""

    tipoclie = "C"
    moneda = "S"
    saldos = ""
    saldod = ""
    fecha = Format(Now, "dd/mm/yyyy")
    codigo = ""
    nombre = ""
    observa = ""
    vendedor = ""
    paridad = ""
    total = ""
    c11 = ""
    c12 = ""
    c13 = ""
    c14 = ""
    borrar_data2
    found = busca_paridad()

    If Val(paridad) = 0 Then
        paridad = "1"

    End If

End Sub

Function busca_registro()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM recibo where  tipo='" & tipo & "' and serie='" & serie & "' and numero='" & Numero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        pone_registro mytablex
        busca_registro = 1

    End If

    mytablex.Close

End Function

Function valida_recibo()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM recibo where  local='" & local1 & "' and tipo='" & tipo & "' and serie='" & serie & "' and numero='" & Numero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        valida_recibo = 1

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Sub pone_registro(mytablex As ADODB.Recordset)
    local1 = "" & mytablex.Fields("local")
    tipo = "" & mytablex.Fields("tipo")
    serie = "" & mytablex.Fields("serie")
    Numero = "" & mytablex.Fields("numero")
    fecha = "" & mytablex.Fields("fecha")
    tipoclie = "" & mytablex.Fields("tipoclie")
    moneda = "" & mytablex.Fields("moneda")
    codigo = "" & mytablex.Fields("codigo")
    nombre = "" & mytablex.Fields("nombre")
    observa = "" & mytablex.Fields("observa")
    vendedor = "" & mytablex.Fields("vendedor")
    paridad = "" & mytablex.Fields("paridad")
    total = "" & mytablex.Fields("total")
    c11 = "" & mytablex.Fields("c1")
    c12 = "" & mytablex.Fields("c2")
    c13 = "" & mytablex.Fields("c3")
    c14 = "" & mytablex.Fields("c4")
    seccion1 = "" & mytablex.Fields("seccion1")
    seccion2 = "" & mytablex.Fields("seccion2")
    seccion3 = "" & mytablex.Fields("seccion3")
    seccion4 = "" & mytablex.Fields("seccion4")
    seccion5 = "" & mytablex.Fields("seccion5")
    seccion6 = "" & mytablex.Fields("seccion6")
    seccion7 = "" & mytablex.Fields("seccion7")
    seccion8 = "" & mytablex.Fields("seccion8")
    seccion9 = "" & mytablex.Fields("seccion9")
    seccion10 = "" & mytablex.Fields("seccion10")

    cseccion1 = "" & mytablex.Fields("cseccion1")
    cseccion2 = "" & mytablex.Fields("cseccion2")
    cseccion3 = "" & mytablex.Fields("cseccion3")
    cseccion4 = "" & mytablex.Fields("cseccion4")
    cseccion5 = "" & mytablex.Fields("cseccion5")
    cseccion6 = "" & mytablex.Fields("cseccion6")
    cseccion7 = "" & mytablex.Fields("cseccion7")
    cseccion8 = "" & mytablex.Fields("cseccion8")
    cseccion9 = "" & mytablex.Fields("cseccion9")
    cseccion10 = "" & mytablex.Fields("cseccion10")

End Sub

Sub grabando(mytablex As ADODB.Recordset)
    mytablex.Fields("seccion1") = Val(seccion1)
    mytablex.Fields("seccion2") = Val(seccion2)
    mytablex.Fields("seccion3") = Val(seccion3)
    mytablex.Fields("seccion4") = Val(seccion4)
    mytablex.Fields("seccion5") = Val(seccion5)
    mytablex.Fields("seccion6") = Val(seccion6)
    mytablex.Fields("seccion7") = Val(seccion7)
    mytablex.Fields("seccion8") = Val(seccion8)
    mytablex.Fields("seccion9") = Val(seccion9)
    mytablex.Fields("seccion10") = Val(seccion10)

    mytablex.Fields("cseccion1") = cseccion1
    mytablex.Fields("cseccion2") = cseccion2
    mytablex.Fields("cseccion3") = cseccion3
    mytablex.Fields("cseccion4") = cseccion4
    mytablex.Fields("cseccion5") = cseccion5
    mytablex.Fields("cseccion6") = cseccion6
    mytablex.Fields("cseccion7") = cseccion7
    mytablex.Fields("cseccion8") = cseccion8
    mytablex.Fields("cseccion9") = cseccion9
    mytablex.Fields("cseccion10") = cseccion10

    mytablex.Fields("local") = local1
    mytablex.Fields("afecta") = afecta
    mytablex.Fields("tipo") = tipo
    mytablex.Fields("usuario") = cajero
    mytablex.Fields("caja") = caja

    If Len(caja) = 0 Then
        mytablex.Fields("caja") = "00"

    End If

    mytablex.Fields("turno") = turno
    mytablex.Fields("numero") = Numero
    mytablex.Fields("serie") = serie
    mytablex.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
    mytablex.Fields("hora") = Format(Now, "hh:mm")
    mytablex.Fields("tipoclie") = tipoclie
    mytablex.Fields("codigo") = codigo
    mytablex.Fields("nombre") = Mid$(nombre, 1, 60)
    mytablex.Fields("observa") = observa
    mytablex.Fields("vendedor") = vendedor
    mytablex.Fields("moneda") = moneda
    mytablex.Fields("paridad") = Val(paridad)
    mytablex.Fields("total") = Val(total)
    mytablex.Fields("estado") = "2"
    mytablex.Fields("acu") = acu
    mytablex.Fields("servicio") = acu
    mytablex.Fields("c1") = Val(c11)
    mytablex.Fields("c2") = Val(c12)
    mytablex.Fields("c3") = Val(c13)
    mytablex.Fields("c4") = Val(c14)

    mytablex.Fields("concepto") = Trim(extra_loquesea(concepto))
    mytablex.Fields("subconcepto") = Trim(extra_loquesea(subconcepto))

End Sub

Private Sub grba1_Click()

End Sub

Private Sub Image1_Click()

    If tipoclie <> "C" And tipoclie <> "P" And tipoclie <> "V" Then
        If tipoclie.Enabled = True Then
            tipoclie.SetFocus

        End If

        Exit Sub

    End If

    consulta_codigo

End Sub

Private Sub Image2_Click()
    consulta_tipo

End Sub

Private Sub Image3_Click()

    If tipoclie <> "C" And tipoclie <> "P" And tipoclie <> "V" Then
        If tipoclie.Enabled = True Then
            tipoclie.SetFocus

        End If

    End If

    consulta_codigo

    If tipoclie.Enabled = False Then
        If fecha.Enabled = True Then
            fecha.SetFocus

        End If

    End If

End Sub

Private Sub Label1_Click()
    cmdSort_Click

End Sub

Function grabar()

    Dim found    As Integer

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    found = valida()

    If found = 0 Then
        MsgBox "Campos invalidos", 48, "Aviso"
        Exit Function

    End If

    'MsgBox ""
    'If bandera = "NUEVO" Then
    If caja = "00" Then
        mytablex.Open "SELECT * FROM tipo where  tipo='" & tipo & "'", cn, adOpenKeyset, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            sdx = Val("" & mytablex.Fields("numero")) + 1
            Numero = "" & sdx

        End If

        mytablex.Close

    End If
   
    If caja <> "00" Then
        If acu = "W" Then 'ingreso
            sdx = Val("" & mytable11.Fields("numerori")) + 1
            Numero = "" & sdx

        End If

        If acu = "V" Then 'egreso
            sdx = Val("" & mytable11.Fields("numerore")) + 1
            Numero = "" & sdx

        End If

    End If

ainicio:
    mytablex.Open "SELECT * FROM recibo where  local='" & local1 & "' and tipo='" & tipo & "' and serie='" & serie & "' and numero='" & Numero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        sdx = Val(Numero) + 1
        Numero = "" & sdx
        mytablex.Close
        GoTo ainicio

    End If

    'If mytablex.RecordCount = 0 Then
    mytablex.AddNew
    grabando mytablex
    mytablex.Update
   
    If pagocash <> 1 Then
        If Frame2.Visible = False Then
            graba_fpagov

        End If

        If Frame2.Visible = True Then
            graba_fpagovle

        End If

    End If
   
    If pagocash = 1 Then
        graba_fpagocash

    End If
   
    'MsgBox "Hopla"
    If Val(totaldoc) > 0 And Val(totaldoc) = Val(total) Then
        If tipoclie = "C" Or tipoclie = "P" Or tipoclie = "V" Then

            'MsgBox afecta
            If Trim("" & afecta) <> "L" Then
                graba_tmpcta 0
                'MsgBox "abcd"
                found = descarga_cuentac(local1, tipo, serie, Numero, "+1")

            End If

            If Trim(afecta) = "L" Then
                graba_tmpcta 1

                'found = descarga_letra(Tipo, serie, numero, "+1")
            End If
      
        End If

    End If
   
    If caja = "00" Then
        found = busca_tipo(2)
    Else

        '------------------
        If acu = "W" Then 'ingreso
            'mytable11.Edit
            mytable11.Fields("numerori") = Numero
            mytable11.Update

        End If

        If acu = "V" Then 'egreso
            'mytable11.Edit
            mytable11.Fields("numerore") = Numero
            mytable11.Update

        End If

        '------------------
    End If

    'found = busca_anticipo()
    If busca_anticipo() = "S" Then
      
        found = graba_datos()

    End If

    If busca_anticipo() = "B" Then  'deposito adelantado bancos
        found = graba_datosb()

    End If

    'MsgBox "QQQ"
    grabar = 1
    mytablex.Close

End Function

Function graba_datos()

    Dim mytablex As New ADODB.Recordset

    Dim found    As Integer

    mytablex.Open "SELECT * FROM " & xcuentaco & " where  local='" & local1 & "' and tipo='" & tipo & "' and serie='" & serie & "' and numero='" & Numero & "' and cuota='1'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        'mytablex.Edit
        'pone_registroG mytablex
        'mytablex.Update
    Else
        mytablex.AddNew
        pone_registroG mytablex
        mytablex.Update

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Sub pone_registroG(mytablex As ADODB.Recordset)

    If Len(caja) = 0 Then
        mytablex.Fields("caja") = "00"

    End If

    mytablex.Fields("usuario") = cajero
    mytablex.Fields("caja") = caja
    mytablex.Fields("grupo") = "A"  'anticipo
    mytablex.Fields("acu") = acu
    mytablex.Fields("fpago") = "A"
    mytablex.Fields("observa") = "ADEL.EFECTIVO"

    mytablex.Fields("turno") = turno
    mytablex.Fields("local") = local1
    mytablex.Fields("tipo") = tipo
    mytablex.Fields("serie") = serie
    mytablex.Fields("numero") = Numero
    mytablex.Fields("cuota") = "1"
    mytablex.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
    mytablex.Fields("fechav") = Format(fecha, "dd/mm/yyyy")
    mytablex.Fields("tipoclie") = tipoclie
    mytablex.Fields("codigo") = codigo
    mytablex.Fields("nombre") = nombre
    mytablex.Fields("zona") = ""
    mytablex.Fields("vendedor") = vendedor
    mytablex.Fields("moneda") = moneda
    mytablex.Fields("total") = Val(total)
    mytablex.Fields("interes") = 0
    mytablex.Fields("abono") = 0
    mytablex.Fields("saldo") = Val(total)
    mytablex.Fields("estado") = "0"
    mytablex.Fields("anticipo") = "1"
    mytablex.Fields("observa") = Mid$("" & observa, 1, 20)

End Sub

Function graba_datosb()

    Dim mytablex As New ADODB.Recordset

    Dim found    As Integer

    mytablex.Open "SELECT * FROM " & xcuentaco & " where  local='" & local1 & "' and tipo='" & tipo & "' and serie='" & serie & "' and numero='" & Numero & "' and cuota='1'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        'mytablex.Edit
        'pone_registroG mytablex
        'mytablex.Update
    Else
        mytablex.AddNew
        pone_registroGb mytablex
        mytablex.Update

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Sub pone_registroGb(mytablex As ADODB.Recordset)

    If Len(caja) = 0 Then
        mytablex.Fields("caja") = "00"

    End If

    mytablex.Fields("usuario") = cajero
    mytablex.Fields("caja") = caja
    mytablex.Fields("grupo") = "D"  'anticipo
    mytablex.Fields("acu") = acu
    mytablex.Fields("fpago") = "H"
    mytablex.Fields("observa") = "DEPOSITOBCO"

    mytablex.Fields("turno") = turno
    mytablex.Fields("local") = local1
    mytablex.Fields("tipo") = tipo
    mytablex.Fields("serie") = serie
    mytablex.Fields("numero") = Numero
    mytablex.Fields("cuota") = "1"
    mytablex.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
    mytablex.Fields("fechav") = Format(fecha, "dd/mm/yyyy")
    mytablex.Fields("tipoclie") = tipoclie
    mytablex.Fields("codigo") = codigo
    mytablex.Fields("nombre") = nombre
    mytablex.Fields("zona") = ""
    mytablex.Fields("vendedor") = vendedor
    mytablex.Fields("moneda") = moneda
    mytablex.Fields("total") = Val(total)
    mytablex.Fields("interes") = 0
    mytablex.Fields("abono") = 0
    mytablex.Fields("saldo") = Val(total)
    mytablex.Fields("estado") = "0"
    mytablex.Fields("anticipo") = "1"
    mytablex.Fields("observa") = Mid$("" & observa, 1, 20)

End Sub

Function valida()

    Dim found As Integer

    If valida_esencial() = 0 Then
        Exit Function

    End If

    found = busca_tipo(0)

    If found = 0 Then
        MsgBox "No existe Tipo", 48, "Aviso"
        Exit Function

    End If

    If valida_fecha("" & fecha) = 0 Then
        fecha = ""
        fecha.SetFocus
        Exit Function

    End If

    If moneda <> "S" And moneda <> "D" Then
        moneda = ""
        moneda.SetFocus
        Exit Function

    End If

    If tipoclie <> "C" And tipoclie <> "V" And tipoclie <> "P" Then
        tipoclie = ""
        tipoclie.SetFocus
        Exit Function

    End If

    'If Len(codigo) = 0 Then
    '   codigo.SetFocus
    '   Exit Function
    'End If
    If Len(nombre) = 0 Then
        nombre.SetFocus
        Exit Function

    End If

    If Len(vendedor) > 0 Then
        found = busca_vendedor()

        If found = 0 Then
            vendedor = ""
            vendedor.SetFocus
            Exit Function

        End If

    End If

    If Val(total) = 0 Then
        total.SetFocus
        Exit Function

    End If

    found = busca_tipo(5)

    If found = 5 Then
        If Val(totaldoc) = 0 Then

            'MsgBox "Obligatorio el ingreso de documentos a pagar", 48, "Aviso"
            'Exit Function
        End If

        If Val(totaldoc) <> Val(total) Then

            'total = totaldoc
            'Exit Function
        End If

        If concepto = "%" Then
            MsgBox "Seleccione un concepto ", 48, "Aviso"
            Exit Function

        End If

        If subconcepto = "%" Then
            MsgBox "Seleccione un Subconcepto ", 48, "Aviso"
            Exit Function

        End If
  
    End If

    valida = 1

End Function

Private Sub Label11_Click()

    'serie1 = ""
    'serie2 = ""
    'serie3 = ""
    'serie4 = ""
    'numero1 = ""
    'numero2 = ""
    'numero3 = ""
    'numero4 = ""
    'tipo1 = ""
    'tipo2 = ""
    'tipo3 = ""
    'tipo4 = ""
    'moneda1 = ""
    'moneda2 = ""
    'moneda3 = ""
    'moneda4 = ""
    'cuota1 = ""
    'cuota2 = ""
    'cuota3 = ""
    'cuota4 = ""
    'total1 = ""
    'total2 = ""
    'total3 = ""
    'total4 = ""
    'paga1 = ""
    'paga2 = ""
    'paga3 = ""
    'paga4 = ""
End Sub

Private Sub Label14_Click()
    consulta_tipo

End Sub

Private Sub Label17_Click()
    Label10_Click

End Sub

Private Sub Label10_Click()

    Dim found As Integer

    Exit Sub
    found = busca_tipo(5)

    If found <> 5 Then

        'MsgBox "Tipo Documento debe permitir cruce documentos ", 48, "Aviso"
        'Exit Sub
    End If

    If tipoclie <> "C" And tipoclie <> "P" And tipoclie <> "V" Then
        tipoclie.SetFocus
        Exit Sub

    End If

    If Len(codigo) = 0 Then
        codigo.SetFocus
        Exit Sub

    End If

    If moneda <> "S" And moneda <> "D" Then
        moneda.SetFocus
        Exit Sub

    End If

    If valida_esencial() = 0 Then
        tipoclie.SetFocus
        Exit Sub

    End If

    consulta_tipo1

End Sub

Private Sub Label12_Click()

    Dim found As Integer

    found = sumar_creditos()
    total.SetFocus

End Sub

Private Sub Label18_Click()

    Dim found As Integer

    found = busca_tipo(5)

    '10/06/2017 kenyo cobro credito por rango de fecha
    If Len(fechai) <> 10 Then Exit Sub
    If Len(fechaf) <> 10 Then Exit Sub
    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub
    '10/06/2017 kenyo cobro credito por rango de fecha

    If found <> 5 Then

        'MsgBox "Tipo Documento Sin permiso de cruce documentos ", 48, "Aviso"
        'Exit Sub
    End If

    If Len(codigo) > 0 Then
        found = busca_codigo()

        If found = 0 Then Exit Sub

    End If

    If tipoclie = "P" Then 'proveedor
        'If afecta <> "L" Then
        found = busca_saldo1()

        'End If
        'If afecta = "L" Then
        '   found = busca_saldo_letra1()
        'End If
    End If

    If tipoclie = "C" Or tipoclie = "V" Then 'proveedor
        'If afecta <> "L" Then
        found = busca_saldo1()

        'End If
        'If afecta = "L" Then
        '   found = busca_saldo_letra1()
        'End If
    End If

    found = sumar_creditos()

    cseccion9.Caption = Format(Val(saldos), "0.00") 'saldo anterior
    cseccion10.Caption = Format((Val(saldos) - Val(total)), "0.00") ' saldo actual

End Sub

Private Sub Label19_Click()

    Dim found As Integer

    Data2.Database.Execute "DELETE FROM " & "_r" & gusuario
    Data2.refresh
    totaldoc = ""
    total = ""
    found = sumar_creditos()

    total.SetFocus

End Sub

Private Sub Label20_Click()

    Dim found As Integer

    found = busca_tipo(5)

    If found = 5 Then

        'MsgBox "Tipo Documento Sin cruce", 48, "Aviso"
        'Exit Sub
    End If

    Label10_Click

End Sub

Private Sub Label22_Click()

    Dim found As Integer

    On Error GoTo cmd6754_err

    Data2.Recordset.Delete
    'Data2.Recordset.Close
    'carga_tmpcta

    found = sumar_creditos()
 
    '10/06/2017 kenyo cobro credito por rango de fecha
    If saldos <> "" Then
        cseccion9.Caption = Format(Val(saldos), "0.00") 'saldo anterior
        cseccion10.Caption = Format((Val(saldos) - Val(total)), "0.00") ' saldo actual

    End If

    '10/06/2017 kenyo cobro credito por rango de fecha
    
    Exit Sub
cmd6754_err:
    Data2.refresh
    Exit Sub

End Sub

Private Sub Label29_Click()

    Dim sdx As Double

    sdx = Val(seccion1) + Val(seccion2) + Val(seccion3) + Val(seccion4) + Val(seccion5) + Val(seccion6) + Val(seccion7) + Val(seccion8) + Val(seccion9) + Val(seccion10)
    tseccion = Format(sdx, "0.00")
    total = Format(sdx, "0.00")

End Sub

Private Sub lefechaf_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    lemoneda.SetFocus

End Sub

Private Sub lefechai_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    lefechaf.SetFocus

End Sub

Private Sub lemoneda_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    levalor.SetFocus

End Sub

Private Sub lenumero_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    lefechai.SetFocus

End Sub

Private Sub lesaldo_Click()

End Sub

Private Sub letipo_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    lenumero.SetFocus

End Sub

Private Sub levalor_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub local1_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If Len(local1) = 0 Then
        consulta_local
        Exit Sub

    End If

    found = busca_local()

    If found = 0 Then
        MsgBox "No existe Local", 48, "Aviso"
        Exit Sub

    End If

    If tipo.Enabled = True Then
        tipo.SetFocus

    End If

End Sub

Private Sub local1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_local

    End If

End Sub

Private Sub moneda_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If valida_esencial() = 0 Then Exit Sub
    If moneda <> "S" And moneda <> "D" Then
        moneda = "S"

    End If

    paridad.SetFocus

End Sub

Private Sub moneda_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        vendedor.SetFocus
        Exit Sub

    End If

End Sub

Private Sub nombre_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If valida_esencial() = 0 Then Exit Sub
    If Len(nombre) = 0 Then Exit Sub
    observa.SetFocus

End Sub

Private Sub nombre_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        If codigo.Enabled = True Then
            codigo.SetFocus
        Else
            fecha.SetFocus

        End If

        Exit Sub

    End If

End Sub

Private Sub numero_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        dlo132_Click
        Exit Sub

    End If

    If Len(local1) = 0 Then
        local1.SetFocus
        Exit Sub

    End If

    found = busca_local()

    If found = 0 Then
        local1.SetFocus
        Exit Sub

    End If

    If Len(tipo) = 0 Then
        If tipo.Enabled = True Then
            tipo.SetFocus

        End If

        Exit Sub

    End If

    found = busca_tipo(0)

    If found = 0 Then
        tipo = ""

        If tipo.Enabled = True Then
            tipo.SetFocus

        End If

        Exit Sub

    End If

    If Len(Numero) = 0 Then
        found = busca_tipo(1)

        If found = 0 Then
            tipo = ""
            Numero = ""

            If tipo.Enabled = True Then
                tipo.SetFocus

            End If

            Exit Sub

        End If

    End If

    found = valida_recibo()

    If found = 1 Then
        If bandera = "NUEVO" Then
            MsgBox "ya existe Numero", 48, "Aviso"
            Numero = ""
            Numero.SetFocus
            Exit Sub

        End If

    End If

    'consulta_sqlx
    If fecha.Enabled = False Then
        tipoclie.SetFocus
        Exit Sub

    End If

    fecha.SetFocus

End Sub

Private Sub numero_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        If tipo.Enabled = True Then
            tipo.SetFocus

        End If

        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        Exit Sub

        'If Len(tipo) = 0 Then
        '   tipo.SetFocus
        '   Exit Sub
        'End If
        'consulta_documento
    End If

End Sub

Private Sub observa_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If valida_esencial() = 0 Then Exit Sub
    vendedor.SetFocus

End Sub

Private Sub observa_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        nombre.SetFocus
        Exit Sub

    End If

End Sub

Private Sub paridad_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If valida_esencial() = 0 Then Exit Sub
    If Not IsDate(fecha) Then
        fecha.SetFocus
        Exit Sub

    End If

    found = busca_paridad()

    If found = 0 Then Exit Sub
    If Val(paridad) = 0 Then
        paridad = "1"

    End If

    If vienede = "CXC" Then

        'Label10_Click
        'Exit Sub
    End If

    total.SetFocus

End Sub

Private Sub paridad_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        If moneda.Enabled = True Then
            moneda.SetFocus
        Else
            observa.SetFocus

        End If

        Exit Sub

    End If

End Sub

Private Sub tipo_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        Exit Sub

    End If

    If Len(tipo) = 0 Then
        consulta_tipo
        Exit Sub

    End If

    found = busca_tipo(0)

    If found = 0 Then
        MsgBox "No existe Tipo", 48, "Aviso"
        Exit Sub

    End If

    'If caja <> "00" Then
    '   numero = valida_ingreso()
    'End If
    If fecha.Enabled = True Then
        fecha.SetFocus
        Exit Sub

    End If

    If tipoclie.Enabled = True Then
        tipoclie.SetFocus
        Exit Sub

    End If

    codigo.SetFocus

End Sub

Private Sub tipo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        'local1.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        consulta_tipo

    End If

End Sub

Private Sub tipo1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub tipo1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        paridad.SetFocus
        Exit Sub

    End If

End Sub

Private Sub tipo2_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub tipo2_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
   
        Exit Sub

    End If

End Sub

Private Sub tipo3_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub tipo3_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
   
        Exit Sub

    End If

End Sub

Private Sub tipo4_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub tipo4_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
   
        Exit Sub

    End If

End Sub

Private Sub tipoclie_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(tipoclie) = 0 Then
        tipoclie = "C"

    End If

    If tipoclie <> "C" And tipoclie <> "P" And tipoclie <> "V" Then
        tipoclie = ""
        tipoclie.SetFocus
        Exit Sub

    End If

    If valida_esencial() = 0 Then
        tipoclie.SetFocus
        Exit Sub

    End If

    codigo.SetFocus

End Sub

Private Sub tipoclie_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        If fecha.Enabled = False Then
            'numero.SetFocus
            Exit Sub

        End If

        fecha.SetFocus
        Exit Sub

    End If

End Sub

Private Sub total_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If Val(total) = 0 Then Exit Sub

    If pagocash.Value = 1 Then
        If MsgBox("Desea Grabar ", 1, "Aviso") <> 1 Then Exit Sub
   
    End If

    found = valida()

    If found = 0 Then
        Exit Sub

    End If

    'found = valida_registro()
    'If found = 1 Then
    '    MsgBox "Numero Recibo ya existe " + numero, 48, "Aviso"
    '    numero.SetFocus
    '    Exit Sub
    'End If
    If vienede = "CXC" Then
        total = ""
        found = sumar_creditos()

        If Val(total) = 0 Then
            MsgBox "Debe de Existir Cuentas Corrientes ", 48, "Aviso"
            Exit Sub

        End If

    End If

    'borrar datos de temporalf

    opcion2 = 0

    If Val(paridad) = 0 Then
        paridad = "1"

    End If

    anticipoo = anticipo

    'anticipoo = ""
    If canjeletra = "S" Then
        fpago = ""
        found = verifica_fpagoletra()

        If found = 0 Then
            MsgBox "No existe forma pago letra", 48, "Aviso"
            Exit Sub

        End If

        borratmpletra
        cn.Execute ("select * into _k" & gusuario & " from tmletra")

        If tabletra.State = 1 Then tabletra.Close
        tabletra.Open "SELECT * from _k" & gusuario, cn, adOpenKeyset, adLockOptimistic
        Set dbgrid3.DataSource = tabletra
        dbgrid3.refresh
        lscodigo = codigo
        lsnombre = nombre
        lsmoneda = moneda
        lstotal = total
        lssubtotal = ""
        lssaldo = total
        suma_tabletra
        Frame2.Visible = True
        Exit Sub

    End If

    If pagocash.Value = 1 Then
        cmdSave_Click
        Exit Sub

    End If

    fpusuarior = "_l" & gusuario
    forpago.tipoclie = tipoclie
    forpago.paridad = paridad

    If moneda = "D" Then
        forpago.txtotals = Format(Val(total) * Val(paridad), "0.00")
        forpago.txtotald = Format(Val(total), "0.00")
        forpago.stxtotals = Format(Val(total) * Val(paridad), "0.00")
        forpago.stxtotald = Format(Val(total), "0.00")

    End If

    If moneda = "S" Then
        forpago.txtotals = Format(Val(total), "0.00")
        forpago.txtotald = Format(Val(total) / Val(paridad), "0.00")
        forpago.stxtotals = Format(Val(total), "0.00")
        forpago.stxtotald = Format(Val(total) / Val(paridad), "0.00")

    End If

    forpago.fechadia = Format(Now, "dd/mm/yyyy")
    forpago.CAMPO1 = codigo
    forpago.CAMPO2 = nombre
    forpago.Show 1

    If opcion2 = 10000 Then
        cmdSave_Click

    End If

End Sub

Sub borratmpletra()

    On Error GoTo cmdn78_err

    cn.Execute ("drop table _k" & gusuario)
    Exit Sub
cmdn78_err:
    Exit Sub

End Sub

Private Sub total_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        paridad.SetFocus
   
    End If

End Sub

Private Sub total4_Click()

End Sub

Private Sub vendedor_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If valida_esencial() = 0 Then Exit Sub
    If Len(vendedor) > 0 Then
        found = busca_vendedor()

        If found = 0 Then
            vendedor = ""
            Exit Sub

        End If

    End If

    If moneda.Enabled = True Then
        moneda.SetFocus
    Else
        paridad.SetFocus

    End If

End Sub

Private Sub vendedor_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        observa.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        consulta_vendedor

    End If

End Sub

Function busca_paridad()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM parame where  codigo='01'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        paridad = Format("" & mytablex.Fields("parivta"))
        busca_paridad = 1

    End If

    mytablex.Close

End Function

Function busca_vendedor()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM vendedor where  codigo='" & vendedor & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        busca_vendedor = 1

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_codigo()

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    If tipoclie = "C" Then
        buf = "clientes"

    End If

    If tipoclie = "P" Then
        buf = "proveedo"

    End If

    If tipoclie = "V" Then
        buf = "vendedor"

    End If

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM " & buf & " where  codigo='" & codigo & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        busca_codigo = 1
        nombre = "" & mytablex.Fields("nombre")

    End If

    mytablex.Close

End Function

Function busca_tipo(sw As Integer)

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Double

    afecta = ""
    mytablex.Open "SELECT * FROM tipo where tipo='" & tipo & "' and tipodoc='" & acu & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        anticipo = "" & mytablex.Fields("anticipo")
        serie = "" & mytablex.Fields("serie")
        'MsgBox anticipo
        busca_tipo = 1

        If sw = 1 Then
            If caja = "00" Then
                serie = "" & mytablex.Fields("serie")

                '   sdx = Val("" & mytablex.Fields("numero")) + 1
                '   numero = "" & sdx
            End If

        End If

        If sw = 2 Then
            'mytablex.Edit
            mytablex.Fields("numero") = "" & Numero
            mytablex.Update

        End If

        If sw = 5 Then
            If "" & mytablex.Fields("obliga") = "S" Then
                busca_tipo = 5

            End If

        End If

        If sw = 6 Then  'aque tipo afectta letra
            afecta = "" & mytablex.Fields("cajachica")

        End If

    End If

    mytablex.Close

End Function

Sub consulta_tipo()
    'cerrar_data1
    Combo1.Clear
    Combo1.AddItem "Descripcio"
    Combo1.AddItem "Tipo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "1"
    Command1_Click

End Sub

Sub consulta_productos()
    'cerrar_data1
    Combo1.Clear
    Combo1.AddItem "Descripcio"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "1280"
    Command1_Click

End Sub

Sub consulta_documento()
    'cerrar_data1
    Combo1.Clear
    Combo1.AddItem "Numero"
    Combo1.AddItem "Tipo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "20"
    Command1_Click

End Sub

Sub consulta_vendedor()
    'cerrar_data1
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "22"
    Command1_Click

End Sub

Sub consulta_codigo()
    'cerrar_data1
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "2"
    Command1_Click

End Sub

Sub consulta_local()
    'cerrar_data1
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "231"
    Command1_Click

End Sub

Sub consulta_copia()
    Combo1.Clear
    Combo1.AddItem "Tipo"
    Combo1.AddItem "Serie"
    Combo1.AddItem "Numero"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "COPIA"
    Command1_Click
    Exit Sub

End Sub

Sub consulta_anula()
    Combo1.Clear
    Combo1.AddItem "Tipo"
    Combo1.AddItem "Serie"
    Combo1.AddItem "Numero"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "ANULA"
    Command1_Click
    Exit Sub

End Sub

Sub consulta_tipo1()

    'cerrar_data1
    If tipoclie = "P" Or tipoclie = "C" Then  'cliente o proveedor,VENDEDOR
   
        Combo1.Clear
        Combo1.AddItem "Tipo"
        Combo1.AddItem "Serie"
        Combo1.AddItem "Numero"
        Combo1.AddItem "Codigo"
        Combo1.ListIndex = 0
        Frame1.Visible = True
        Frame1.Enabled = True
        buffer = ""
        buffer.SetFocus
        opcion1 = "12"
        Command1_Click
        Exit Sub

    End If

    If tipoclie = "V" Then
        Combo1.Clear
        Combo1.AddItem "Tipo"
        Combo1.AddItem "serie"
        Combo1.AddItem "numero"
        Combo1.ListIndex = 0
        Frame1.Visible = True
        Frame1.Enabled = True
        buffer = ""
        buffer.SetFocus
        opcion1 = "3"
        Command1_Click

    End If

End Sub

Sub consulta_tipo2()

    'cerrar_data1
    If tipoclie = "P" Then
        Combo1.Clear
        Combo1.AddItem "Tipo"
        Combo1.AddItem "Serie"
        Combo1.AddItem "Numero"
        Combo1.AddItem "Codigo"
        Combo1.ListIndex = 0
        Frame1.Visible = True
        Frame1.Enabled = True
        buffer = ""
        buffer.SetFocus
        opcion1 = "13"
        Command1_Click
        Exit Sub

    End If

    If afecta = "L" Then
        Combo1.Clear
        Combo1.AddItem "letra"
        Combo1.AddItem "aceptante"
        Combo1.ListIndex = 0
        Frame1.Visible = True
        Frame1.Enabled = True
        buffer = ""
        buffer.SetFocus
        opcion1 = "9"
        Command1_Click
        Exit Sub

    End If

    Combo1.Clear
    Combo1.AddItem "Tipo"
    Combo1.AddItem "serie"
    Combo1.AddItem "numero"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "4"
    Command1_Click

End Sub

Sub consulta_tipo3()

    'cerrar_data1
    If afecta = "P" Then
        Combo1.Clear
        Combo1.AddItem "Tipo"
        Combo1.AddItem "Serie"
        Combo1.AddItem "Numero"
        Combo1.AddItem "Codigo"
        Combo1.ListIndex = 0
        Frame1.Visible = True
        Frame1.Enabled = True
        buffer = ""
        buffer.SetFocus
        opcion1 = "14"
        Command1_Click
        Exit Sub

    End If

    If afecta = "L" Then
        Combo1.Clear
        Combo1.AddItem "letra"
        Combo1.AddItem "aceptante"
        Combo1.ListIndex = 0
        Frame1.Visible = True
        Frame1.Enabled = True
        buffer = ""
        buffer.SetFocus
        opcion1 = "10"
        Command1_Click
        Exit Sub

    End If

    Combo1.Clear
    Combo1.AddItem "Tipo"
    Combo1.AddItem "serie"
    Combo1.AddItem "numero"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "5"
    Command1_Click

End Sub

Sub consulta_tipo4()

    'cerrar_data1
    If afecta = "P" Then
        Combo1.Clear
        Combo1.AddItem "Tipo"
        Combo1.AddItem "Serie"
        Combo1.AddItem "Numero"
        Combo1.AddItem "Codigo"
        Combo1.ListIndex = 0
        Frame1.Visible = True
        Frame1.Enabled = True
        buffer = ""
        buffer.SetFocus
        opcion1 = "15"
        Command1_Click
        Exit Sub

    End If

    If afecta = "L" Then
        Combo1.Clear
        Combo1.AddItem "letra"
        Combo1.AddItem "aceptante"
        Combo1.ListIndex = 0
        Frame1.Visible = True
        Frame1.Enabled = True
        buffer = ""
        buffer.SetFocus
        opcion1 = "11"
        Command1_Click
        Exit Sub

    End If

    Combo1.Clear
    Combo1.AddItem "Tipo"
    Combo1.AddItem "serie"
    Combo1.AddItem "numero"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "6"
    Command1_Click

End Sub

Function busca_saldo_letra1()

    Dim mytablex As Table

    Dim found    As Integer

    Dim buf      As String

    If tipoclie = "C" Or tipoclie = "V" Then
        buf = "letrav"

    End If

    If tipoclie = "P" Then
        buf = "letrac"

    End If

    'consulta_sqlx
    Data2.Database.Execute "DELETE FROM " & fgusuario
    'consulta_sqlx
    totaldoc = ""
    total = ""
    mytablex.Open "SELECT * FROM " & xcuentaco & " where  tipoclie='" & tipoclie & "' and aceptante='" & codigo & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then  'si existe
        mytablex.Close
        Exit Function

    End If

    Do

        If mytablex.EOF Then Exit Do

        '-----------------------------------------
        If Val("" & mytablex.Fields("saldo")) > 0 Then
            Data2.Recordset.AddNew
            Data2.Recordset.Fields("tipoclie") = "" & tipoclie
            Data2.Recordset.Fields("codigo") = "" & codigo
            Data2.Recordset.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
            Data2.Recordset.Fields("hora") = Format(Now, "HH:MM:SS")
            Data2.Recordset.Fields("tipo") = "" & tipo
            Data2.Recordset.Fields("local") = "" & local1
            Data2.Recordset.Fields("serie") = "" & serie
            Data2.Recordset.Fields("numero") = "" & Numero
            Data2.Recordset.Fields("acu") = "" & acu
            Data2.Recordset.Fields("usuario") = ""
            Data2.Recordset.Fields("paridad") = 0
            Data2.Recordset.Fields("local1") = "" & mytablex.Fields("local")
            Data2.Recordset.Fields("tipo1") = "" & mytablex.Fields("tipo")
            Data2.Recordset.Fields("serie1") = "" & mytablex.Fields("serie")
            Data2.Recordset.Fields("numero1") = "" & mytablex.Fields("numero")
            Data2.Recordset.Fields("cuota") = "" & mytablex.Fields("cuota")
            Data2.Recordset.Fields("moneda") = "" & mytablex.Fields("Moneda")
            Data2.Recordset.Fields("total") = Val("" & mytablex.Fields("saldo"))
            Data2.Recordset.Fields("paga") = Val("" & mytablex.Fields("saldo"))
            Data2.Recordset.Fields("L1") = 0
            Data2.Recordset.Fields("L2") = 0
            Data2.Recordset.Fields("L3") = 0
            Data2.Recordset.Fields("L4") = 0
            Data2.Recordset.Fields("estado") = "2"
            Data2.Recordset.Update

        End If

        mytablex.MoveNext
    Loop
    mytablex.Close
    found = sumar_creditos()
    found = ir_inicio(1)

End Function

Function busca_saldo1()

    Dim mytablex As New ADODB.Recordset

    Dim found    As Integer

    Dim buf      As String

    'Data2.refresh
    Data2.Database.Execute "DELETE FROM " & "_r" & gusuario
    Data2.refresh
    totaldoc = ""
    total = ""

    '10/06/2017 kenyo cobro credito por rango de fecha
    'mytablex.Open "SELECT * FROM " & xcuentaco & " where fecha ='24/04/2017' and tipoclie='" & tipoclie & "' and codigo='" & codigo & "' order by numero", cn, adOpenKeyset, adLockOptimistic
    mytablex.Open "SELECT * FROM " & xcuentaco & " where  fecha>='" & Format(fechai, "YYYYMMDD") & "'  and fecha<='" & Format(fechaf, "YYYYMMDD") & "' and tipoclie='" & tipoclie & "' and codigo='" & codigo & "' order by fecha", cn, adOpenKeyset, adLockOptimistic
    '10/06/2017 kenyo cobro credito por rango de fecha

    If mytablex.RecordCount = 0 Then  'si existe
        mytablex.Close
        Exit Function

    End If

    Do

        If mytablex.EOF Then Exit Do

        '-----------------------------------------
        If Val("" & mytablex.Fields("saldo")) > 0 And "" & mytablex.Fields("grupo") = "C" Then
            Data2.Recordset.AddNew
            Data2.Recordset.Fields("tipoclie") = "" & tipoclie
            Data2.Recordset.Fields("codigo") = "" & codigo
   
            '10/06/2017 kenyo cobro credito por rango de fecha
            Data2.Recordset.Fields("fecha") = "" & mytablex.Fields("fecha")
            '10/06/2017 kenyo cobro credito por rango de fecha
   
            Data2.Recordset.Fields("hora") = Format(Now, "HH:MM:SS")
            Data2.Recordset.Fields("tipo") = "" & tipo
            Data2.Recordset.Fields("local") = "" & local1
            Data2.Recordset.Fields("serie") = "" & serie
            Data2.Recordset.Fields("numero") = "" & Numero
            Data2.Recordset.Fields("acu") = "" & acu
            Data2.Recordset.Fields("usuario") = ""
            Data2.Recordset.Fields("paridad") = 0
            Data2.Recordset.Fields("local1") = "" & mytablex.Fields("local")
            Data2.Recordset.Fields("tipo1") = "" & mytablex.Fields("tipo")
            Data2.Recordset.Fields("serie1") = "" & mytablex.Fields("serie")
            Data2.Recordset.Fields("numero1") = "" & mytablex.Fields("numero")
            Data2.Recordset.Fields("cuota") = "" & mytablex.Fields("cuota")
            Data2.Recordset.Fields("moneda") = "" & mytablex.Fields("Moneda")
            Data2.Recordset.Fields("total") = Val("" & mytablex.Fields("saldo"))
            Data2.Recordset.Fields("paga") = Val("" & mytablex.Fields("saldo"))
            Data2.Recordset.Fields("L1") = 0
            Data2.Recordset.Fields("L2") = 0
            Data2.Recordset.Fields("L3") = 0
            Data2.Recordset.Fields("L4") = 0
            Data2.Recordset.Fields("estado") = "2"
            Data2.Recordset.Update

            'Data2.refresh
        End If

        mytablex.MoveNext
    Loop
    mytablex.Close
    'Data2.Recordset.Close
    'carga_tmpcta

    found = sumar_creditos()
    found = ir_inicio(1)

End Function

''' 25/08/2017 kenyo. Cobro Credito desde ventana de ventas
Function busca_saldoCobroCredito()

    Dim mytablex As New ADODB.Recordset

    Dim found    As Integer

    Dim buf      As String

    'Data2.refresh
    Data2.Database.Execute "DELETE FROM " & "_r" & gusuario
    Data2.refresh
    totaldoc = ""
    total = ""

    mytablex.Open "SELECT * FROM " & xcuentaco & " where  tipo>='" & tipod & "'  and serie<='" & seried & "' and numero='" & numerod & "' and tipoclie='" & tipoclie & "' and codigo='" & codigo & "' order by fecha", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then  'si existe
        mytablex.Close
        Exit Function

    End If

    Do

        If mytablex.EOF Then Exit Do

        '-----------------------------------------
        If Val("" & mytablex.Fields("saldo")) > 0 And "" & mytablex.Fields("grupo") = "C" Then
            Data2.Recordset.AddNew
            Data2.Recordset.Fields("tipoclie") = "" & tipoclie
            Data2.Recordset.Fields("codigo") = "" & codigo
            Data2.Recordset.Fields("fecha") = "" & mytablex.Fields("fecha")
   
            Data2.Recordset.Fields("hora") = Format(Now, "HH:MM:SS")
            Data2.Recordset.Fields("tipo") = "" & tipo
            Data2.Recordset.Fields("local") = "" & local1
            Data2.Recordset.Fields("serie") = "" & serie
            Data2.Recordset.Fields("numero") = "" & Numero
            Data2.Recordset.Fields("acu") = "" & acu
            Data2.Recordset.Fields("usuario") = ""
            Data2.Recordset.Fields("paridad") = 0
            Data2.Recordset.Fields("local1") = "" & mytablex.Fields("local")
            Data2.Recordset.Fields("tipo1") = "" & mytablex.Fields("tipo")
            Data2.Recordset.Fields("serie1") = "" & mytablex.Fields("serie")
            Data2.Recordset.Fields("numero1") = "" & mytablex.Fields("numero")
            Data2.Recordset.Fields("cuota") = "" & mytablex.Fields("cuota")
            Data2.Recordset.Fields("moneda") = "" & mytablex.Fields("Moneda")
            Data2.Recordset.Fields("total") = Val("" & mytablex.Fields("saldo"))
            Data2.Recordset.Fields("paga") = Val("" & mytablex.Fields("saldo"))
            Data2.Recordset.Fields("L1") = 0
            Data2.Recordset.Fields("L2") = 0
            Data2.Recordset.Fields("L3") = 0
            Data2.Recordset.Fields("L4") = 0
            Data2.Recordset.Fields("estado") = "2"
            Data2.Recordset.Update

            'Data2.refresh
        End If

        mytablex.MoveNext
    Loop
    mytablex.Close
    'Data2.Recordset.Close
    'carga_tmpcta

    found = sumar_creditos()
    found = ir_inicio(1)

End Function

''' 25/08/2017 kenyo. Cobro Credito desde ventana de ventas

Function busca_saldo()

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    Dim buf1     As String

    suma1 = 0
    suma2 = 0
    'If tipoclie = "C" Or tipoclie = "V" Then
    '   buf = "cuentac"
    'End If
    'If tipoclie = "P" Then
    '   buf = "cuentap"
    'End If
    buf1 = "SELECT * FROM " & xcuentaco & " where  tipoclie='" & tipoclie & "' and codigo='" & codigo & "'"
    mytablex.Open buf1, cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then  'si existe
        mytablex.Close
        Exit Function

    End If

    busca_saldo = 1
    Do

        If mytablex.EOF Then Exit Do

        '-----------------------------------------
        If "" & mytablex.Fields("grupo") = "C" Then
            If "" & mytablex.Fields("moneda") = "S" Then
                suma1 = suma1 + Val("" & mytablex.Fields("saldo"))

            End If

            If "" & mytablex.Fields("moneda") = "D" Then
                suma2 = suma2 + Val("" & mytablex.Fields("saldo"))

            End If

        End If

        mytablex.MoveNext
    Loop
    '------------------------------------- ------------
    mytablex.Close

End Function

Function descarga_cuentac(xlocal1 As String, _
                          xtipo1 As String, _
                          xserie1 As String, _
                          xnumero1 As String, _
                          signo As String)

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset

    Dim buf      As String

    Dim buf1     As String

    '----primero debe habecerse grabdo el temporal y luego seleccionar para descargar
    If tipoclie = "C" Or tipoclie = "V" Then
        buf = "cuentac"
        buf1 = "cuentacd"

    End If

    If tipoclie = "P" Then
        buf = "cuentap"
        buf1 = "cuentapd"

    End If

    'MsgBox xnumero1
    'MsgBox XCUENTACO1 & " " & xlocal1 & "" & xtipo1 & " " & xserie1 & " " & xnumero1

    mytabley.Open "SELECT * FROM " & XCUENTACO1 & " where  local='" & xlocal1 & "' and tipo='" & xtipo1 & "' and serie='" & xserie1 & "' and numero='" & xnumero1 & "'", cn, adOpenKeyset, adLockOptimistic

    If mytabley.RecordCount = 0 Then  'si existe
        mytabley.Close
        Exit Function

    End If

    'MsgBox "Hola"
    Do

        If mytabley.EOF Then Exit Do
        mytablex.Open "SELECT * FROM " & xcuentaco & " where  local='" & mytabley.Fields("local1") & "' and tipo='" & mytabley.Fields("tipo1") & "' and serie='" & mytabley.Fields("serie1") & "' and numero='" & mytabley.Fields("numero1") & "' and cuota='" & mytabley.Fields("cuota") & "'", cn, adOpenKeyset, adLockOptimistic

        If mytablex.RecordCount > 0 Then  'si existe
            'mytablex.Edit
            sdx = Val("" & mytablex.Fields("abono")) + Val(signo) * Val("" & mytabley.Fields("paga"))
            mytablex.Fields("abono") = Val(Format(sdx, "0.00"))
            sdx = Val("" & mytablex.Fields("total")) - Val("" & mytablex.Fields("abono")) + Val("" & mytablex.Fields("interes"))
            mytablex.Fields("saldo") = Val(Format(sdx, "0.00"))
            sdx = Val("" & mytablex.Fields("c1")) - Val(signo) * Val("" & mytabley.Fields("l1"))
            mytablex.Fields("c1") = Format(sdx, "0.00")
            sdx = Val("" & mytablex.Fields("c2")) - Val(signo) * Val("" & mytabley.Fields("l2"))
            mytablex.Fields("c2") = Format(sdx, "0.00")
            sdx = Val("" & mytablex.Fields("c3")) - Val(signo) * Val("" & mytabley.Fields("l3"))
            mytablex.Fields("c3") = Format(sdx, "0.00")
            sdx = Val("" & mytablex.Fields("c4")) - Val(signo) * Val("" & mytabley.Fields("l4"))
            mytablex.Fields("c4") = Format(sdx, "0.00")
            mytablex.Update

            'ahora en pedido
            '--------------------------------------------
            If mytabley.Fields("tipo1") = "P" Then
                mytablez.Open "SELECT * FROM cpedidov where  local='" & mytabley.Fields("local1") & "' and tipo='" & mytabley.Fields("tipo1") & "' and serie='" & mytabley.Fields("serie1") & "' and numero='" & mytabley.Fields("numero1") & "'", cn, adOpenKeyset, adLockOptimistic

                If mytablez.RecordCount > 0 Then  'si existe
                    mytablez.Fields("acuenta") = Val("" & mytablez.Fields("acuenta")) + Val(signo) * Val("" & mytabley.Fields("paga"))
                    mytablez.Update

                End If

                mytablez.Close

            End If

            '--------------------------------------------
        
        End If

        mytablex.Close
        mytabley.MoveNext
    Loop
    mytabley.Close

    If Val(signo) = -1 Then
        cn.Execute ("delete from " & XCUENTACO1 & " where local='" & xlocal1 & "' and tipo='" & xtipo1 & "' and serie='" & xserie1 & "' and numero='" & xnumero1 & "'")

    End If

End Function

Function valida_registro()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM recibo where  local='" & local1 & "' and tipo='" & tipo & "' and serie='" & serie & "' and numero='" & Numero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        valida_registro = 1

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function graba_fpagov()

    Dim mytabley As New ADODB.Recordset

    Dim mytablex As Table

    'Dim mytablez As New adodb.Recordset
    Dim buf      As String

    '---------- validando si es cuenta corriente
    'MsgBox gofpago
    If tipoclie = "C" Or tipoclie = "V" Then

        'If afecta <> "L" Then
        '   buf = "cuentac"
        '
        'End If
        'If afecta = "L" Then
        '   buf = "cuentac"
        '
        'End If
    End If

    If tipoclie = "P" Then

        'If afecta <> "L" Then
        '   buf = "cuentap"
        '
        'End If
        'If afecta = "L" Then
        '   buf = "cuentap"
        '
        'End If
    End If

    Set mytablex = mydbxglo.OpenTable(fpusuarior)
    mytabley.Open "SELECT * FROM " & gofpago & " where local='" & local1 & "' and tipo='" & tipo & "' and serie='" & serie & "' and numero='" & Numero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytabley.RecordCount = 0 Then
        Do

            If mytablex.EOF Then Exit Do
            mytabley.AddNew
            grabar_registro_fpagov mytablex, mytabley, xcuentaco
            mytabley.Update
            'verifica si es credito}
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close
    mytabley.Close

    'mytablez.Close
End Function

Sub grabar_registro_fpagov(mytablex As Table, mytabley As ADODB.Recordset, buf1 As String)

    Dim buf As String

    mytabley.Fields("local") = "" & local1
    mytabley.Fields("tipo") = "" & tipo
    mytabley.Fields("serie") = "" & serie
    mytabley.Fields("numero") = "" & Numero
    mytabley.Fields("tipoclie") = "" & tipoclie
    mytabley.Fields("codigo") = "" & codigo
    mytabley.Fields("nombre") = "" & nombre
    mytabley.Fields("vendedor") = "" & vendedor
    mytabley.Fields("usuario") = "" & cajero
    mytabley.Fields("caja") = "" & caja

    If Len(caja) = 0 Then
        mytabley.Fields("caja") = "00"

    End If

    mytabley.Fields("turno") = "" & turno
    mytabley.Fields("fecha") = Format("" & fecha, "dd/mm/yyyy")
    mytabley.Fields("moneda") = "" & mytablex.Fields("moneda")
    mytabley.Fields("total") = Val(total)
    mytabley.Fields("recibe") = Val("" & mytablex.Fields("recibe"))
    mytabley.Fields("recibes") = Val("" & mytablex.Fields("recibes"))
    mytabley.Fields("recibed") = Val("" & mytablex.Fields("recibed"))
    mytabley.Fields("saldos") = Val("" & mytablex.Fields("saldos"))
    mytabley.Fields("saldod") = Val("" & mytablex.Fields("saldod"))
    mytabley.Fields("nombre") = "" & mytablex.Fields("nombre")

    If Len(Trim("" & mytablex.Fields("nombre"))) = 0 Then
        mytabley.Fields("nombre") = nombre

    End If

    mytabley.Fields("orden") = "" & mytablex.Fields("orden")
    mytabley.Fields("observa") = Trim("" & observa)
    mytabley.Fields("dias") = "" & mytablex.Fields("dias")
    mytabley.Fields("fpago") = "" & mytablex.Fields("fpago")
    buf = busca_fpago("" & mytablex.Fields("fpago"))
    'MsgBox buf
    mytabley.Fields("acufp") = buf
    mytabley.Fields("descripcio") = "" & mytablex.Fields("descripcio")
    mytabley.Fields("acu") = "" & acu
    mytabley.Fields("servicio") = "" & acu
    mytabley.Fields("estado") = "2"

    If buf = "" Then  '

    End If

    If buf = "C" Then   'credito
        graba_credito mytablex, buf, "C", buf1

    End If

    If buf = "G" Then   'letras

        'graba_letras mytablex
        'MsgBox xcuentacol
        If Len(xcuentacol) > 0 Then
            graba_credito_letra mytablex, buf, "C", xcuentacol

        End If

    End If
   
    If buf = "H" Then   'bancos
        graba_tarjetas mytabley

        'graba_credito mytablex, mytablez, buf, "H"
    End If
   
End Sub

Function graba_fpagocash()

    Dim mytabley As New ADODB.Recordset

    Dim buf      As String

    Dim sdx      As Double

    sdx = Val(total)
    'If tabletra.RecordCount = 0 Then Exit Function
    mytabley.Open "SELECT * FROM " & gofpago & " where local='" & local1 & "' and tipo='" & tipo & "' and serie='" & serie & "' and numero='" & Numero & "'", cn, adOpenKeyset, adLockOptimistic
    mytabley.AddNew
    mytabley.Fields("local") = "" & local1
    mytabley.Fields("tipo") = "" & tipo
    mytabley.Fields("serie") = "" & serie
    mytabley.Fields("numero") = "" & Numero
    mytabley.Fields("tipoclie") = "" & tipoclie
    mytabley.Fields("codigo") = "" & codigo
    mytabley.Fields("nombre") = "" & nombre
    mytabley.Fields("vendedor") = "" & vendedor
    mytabley.Fields("usuario") = "" & cajero
    mytabley.Fields("caja") = "" & caja

    If Len(caja) = 0 Then
        mytabley.Fields("caja") = "00"

    End If

    mytabley.Fields("turno") = "" & turno
    mytabley.Fields("fecha") = Format("" & fecha, "dd/mm/yyyy")
    mytabley.Fields("moneda") = moneda
    mytabley.Fields("total") = Val(total)
    mytabley.Fields("recibe") = Val(total)
    mytabley.Fields("recibes") = Val(total)
    mytabley.Fields("recibed") = 0
    mytabley.Fields("saldos") = 0
    mytabley.Fields("saldod") = 0
    mytabley.Fields("nombre") = "" & nombre
   
    mytabley.Fields("orden") = ""
    mytabley.Fields("observa") = "Cash"
    mytabley.Fields("dias") = ""
    mytabley.Fields("fpago") = "1"
    buf = busca_fpago("1")
    'MsgBox buf
    mytabley.Fields("acufp") = buf
    mytabley.Fields("descripcio") = ""
    mytabley.Fields("acu") = "" & acu
    mytabley.Fields("servicio") = "" & acu
    mytabley.Fields("estado") = "2"
    mytabley.Update

End Function

Function graba_fpagovle()

    Dim mytabley As New ADODB.Recordset

    Dim buf      As String

    Dim sdx      As Double

    sdx = Val(total)

    If tabletra.RecordCount = 0 Then Exit Function
    mytabley.Open "SELECT * FROM " & gofpago & " where local='" & local1 & "' and tipo='" & tipo & "' and serie='" & serie & "' and numero='" & Numero & "'", cn, adOpenKeyset, adLockOptimistic
    'If mytabley.RecordCount > 0 Then
    '   mytabley.Close
    '   Exit Function
    'End If
    tabletra.MoveFirst

    Do

        If tabletra.EOF Then Exit Do
        mytabley.AddNew
        mytabley.Fields("local") = "" & local1
        mytabley.Fields("tipo") = "" & tipo
        mytabley.Fields("serie") = "" & serie
        mytabley.Fields("numero") = "" & Numero
        mytabley.Fields("tipoclie") = "" & tipoclie
        mytabley.Fields("codigo") = "" & codigo
        mytabley.Fields("nombre") = "" & nombre
        mytabley.Fields("vendedor") = "" & vendedor
        mytabley.Fields("usuario") = "" & cajero
        mytabley.Fields("caja") = "" & caja

        If Len(caja) = 0 Then
            mytabley.Fields("caja") = "00"

        End If

        mytabley.Fields("turno") = "" & turno
        mytabley.Fields("fecha") = Format("" & fecha, "dd/mm/yyyy")
        mytabley.Fields("moneda") = "" & tabletra.Fields("moneda")
        mytabley.Fields("total") = sdx
        mytabley.Fields("recibe") = Val("" & tabletra.Fields("valor"))
        mytabley.Fields("recibes") = Val("" & tabletra.Fields("valor"))
        mytabley.Fields("recibed") = 0
        mytabley.Fields("saldos") = sdx - Val("" & tabletra.Fields("valor"))
        mytabley.Fields("saldod") = 0
        mytabley.Fields("nombre") = "" & nombre
   
        mytabley.Fields("orden") = ""
        mytabley.Fields("observa") = "Canje Letras"
        mytabley.Fields("dias") = ""
        mytabley.Fields("fpago") = "" & fpago
        buf = busca_fpago("" & fpago)
        'MsgBox buf
        mytabley.Fields("acufp") = buf
        mytabley.Fields("descripcio") = "Pago Letras"
        mytabley.Fields("acu") = "" & acu
        mytabley.Fields("servicio") = "" & acu
        mytabley.Fields("estado") = "2"
        mytabley.Update

        If buf = "" Then  '

        End If

        If buf = "G" Then   'letras
            If Len(xcuentacol) > 0 Then
                graba_credito_letrae buf, "C", xcuentacol

            End If

        End If

        If buf = "H" Then   'bancos
            graba_tarjetas mytabley

            'graba_credito mytablex, mytablez, buf, "H"
        End If

        tabletra.MoveNext
    Loop

End Function

Function borra_fpagov(xtipo As String, xserie As String, xnumero As String)
    cn.Execute ("delete from " & gofpago & " where local='" & local1 & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'")

End Function

Sub carga_empresa()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM grupos where  grupos='01'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        nc1 = "" & mytablex.Fields("l1")
        nc2 = "" & mytablex.Fields("l2")
        nc3 = "" & mytablex.Fields("l3")
        nc4 = "" & mytablex.Fields("l4")

    End If

    '------------------------------------- ------------
    mytablex.Close
 
End Sub

Function existe_seleccionado(buf0 As String, _
                             buf1 As String, _
                             buf2 As String, _
                             buf3 As String)

    Dim mytablex As Table

    Set mytablex = mydbxglo.OpenTable(fgusuario)
    mytablex.Index = "tmpcta1"
    mytablex.Seek "=", buf0, buf1, buf2, buf3

    If Not mytablex.NoMatch Then
        existe_seleccionado = 1

    End If

    mytablex.Close

End Function

Function ir_inicio(sw As Integer)

    On Error GoTo cmd1000_err

    If sw = 0 Then
        Data2.Recordset.MoveFirst

    End If

    If sw = 1 Then
        Data2.Recordset.MoveLast

    End If

    ir_inicio = 1
    Exit Function
cmd1000_err:
    Exit Function

End Function

Function sumar_creditos()

    Dim sdx   As Double

    Dim sdx1  As Double

    Dim sdx2  As Double

    Dim sdx3  As Double

    Dim sdx4  As Double

    Dim found As Integer

    found = ir_inicio(0)
    'If found = 0 Then Exit function
    sdx = 0
    sdx1 = 0
    sdx2 = 0
    sdx3 = 0
    sdx4 = 0

    If Val(paridad) <= 0 Then
        paridad = "1"

    End If

    Do

        If Data2.Recordset.EOF Then Exit Do
        If moneda = "S" Then
            If "" & Data2.Recordset.Fields("moneda") = "S" Then
                sdx = sdx + Val("" & Data2.Recordset.Fields("paga"))
                sdx1 = sdx1 + Val("" & Data2.Recordset.Fields("l1"))
                sdx2 = sdx2 + Val("" & Data2.Recordset.Fields("l2"))
                sdx3 = sdx3 + Val("" & Data2.Recordset.Fields("l3"))
                sdx4 = sdx4 + Val("" & Data2.Recordset.Fields("l4"))

            End If

            If "" & Data2.Recordset.Fields("moneda") = "D" Then
                sdx = sdx + Val("" & Data2.Recordset.Fields("paga")) * Val(paridad)
                sdx1 = sdx1 + Val("" & Data2.Recordset.Fields("l1")) * Val(paridad)
                sdx2 = sdx2 + Val("" & Data2.Recordset.Fields("l2")) * Val(paridad)
                sdx3 = sdx3 + Val("" & Data2.Recordset.Fields("l3")) * Val(paridad)
                sdx4 = sdx4 + Val("" & Data2.Recordset.Fields("l4")) * Val(paridad)

            End If

        End If

        If moneda = "D" Then
            If "" & Data2.Recordset.Fields("moneda") = "S" Then
                sdx = sdx + Val("" & Data2.Recordset.Fields("paga")) / Val(paridad)
                sdx1 = sdx1 + Val("" & Data2.Recordset.Fields("l1")) / Val(paridad)
                sdx2 = sdx2 + Val("" & Data2.Recordset.Fields("l2")) / Val(paridad)
                sdx3 = sdx3 + Val("" & Data2.Recordset.Fields("l3")) / Val(paridad)
                sdx4 = sdx4 + Val("" & Data2.Recordset.Fields("l4")) / Val(paridad)

            End If

            If "" & Data2.Recordset.Fields("moneda") = "D" Then
                sdx = sdx + Val("" & Data2.Recordset.Fields("paga"))
                sdx1 = sdx1 + Val("" & Data2.Recordset.Fields("l1"))
                sdx2 = sdx2 + Val("" & Data2.Recordset.Fields("l2"))
                sdx3 = sdx3 + Val("" & Data2.Recordset.Fields("l3"))
                sdx4 = sdx4 + Val("" & Data2.Recordset.Fields("l4"))

            End If

        End If

        Data2.Recordset.MoveNext
    Loop
    totaldoc = ""

    If sdx > 0 Then
        total = Format(sdx, "0.00")
        totaldoc = total
        sumar_creditos = 1
        c11 = Format(sdx1, "0.00")
        c12 = Format(sdx2, "0.00")
        c13 = Format(sdx3, "0.00")
        c14 = Format(sdx4, "0.00")

    End If

End Function

Sub graba_tmpcta(sw As Integer)

    Dim mytablex As New ADODB.Recordset

    Dim found    As Integer

    Dim I        As Integer

    Dim buf      As String

    found = ir_inicio(0)

    'Data2.refresh
    'If sw = 0 Then
    If tipoclie = "C" Or tipoclie = "V" Then

        'If afecta = "L" Then
        '   buf = "letracd"
        'End If
        'If afecta <> "L" Then
        '   buf = "cuentacd"
        'End If
    End If

    If tipoclie = "P" Then

        'If afecta = "L" Then
        '   buf = "letrapd"
        'End If
        'If afecta <> "L" Then
        '   buf = "cuentapd"
        'End If
    End If
   
    Select Case xcuentaco

        Case "CUENTAC"
            buf = "cuentacd"

        Case "CUENTAP"
            buf = "cuentapd"

    End Select

    'MsgBox "descarga"
    ''MsgBox XCUENTACO1
   
    mytablex.Open "SELECT * FROM " & XCUENTACO1, cn, adOpenKeyset, adLockOptimistic
    Do

        If Data2.Recordset.EOF Then Exit Do

        'MsgBox Val("" & Data2.Recordset.Fields("paga"))
        If Val("" & Data2.Recordset.Fields("paga")) > 0 Then
            mytablex.AddNew

            For I = 0 To Data2.Recordset.Fields.count - 1
                mytablex.Fields(I) = Data2.Recordset.Fields(I)
            Next I

            mytablex.Fields("local") = "" & local1
            mytablex.Fields("tipo") = "" & tipo
            mytablex.Fields("serie") = "" & serie
            mytablex.Fields("numero") = "" & Numero
            mytablex.Fields("usuario") = cajero
            mytablex.Fields("caja") = caja
      
            ''15/07/2017 kenyo fecha ingreso en finanzas cuentas por cobrar
            mytablex.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
            ''15/07/2017 kenyo fecha ingreso en finanzas cuentas por cobrar
      
            If Len(caja) = 0 Then
                mytablex.Fields("caja") = "00"

            End If

            mytablex.Fields("turno") = turno
            mytablex.Update

        End If

        Data2.Recordset.MoveNext
    Loop
    mytablex.Close
    'End If

End Sub

Function busca_fpago(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM fpago where  fpago='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_fpago = "" & mytablex.Fields("tipo")

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Sub borrar_data2()

    Dim found As Integer

    On Error GoTo cmd345_err

denuevo1:
    found = ir_inicio(0)
    Data2.Recordset.Delete
    Data2.refresh
    GoTo denuevo1
    Exit Sub
cmd345_err:
    Data2.refresh
    Exit Sub

End Sub

Function valida_esencial()

    Dim found As Integer

    If Len(local1) = 0 Then
        If local1.Enabled = True Then
            local1.SetFocus

        End If

        Exit Function

    End If

    found = busca_local()

    If found = 0 Then
        If local1.Enabled = True Then
            local1.SetFocus

        End If

        Exit Function

    End If

    If Len(tipo) = 0 Then
        If tipo.Enabled = True Then
            tipo.SetFocus

        End If

        Exit Function

    End If

    valida_esencial = 1

End Function

Sub carga_seccion()

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM VENDEDOR where dueno='S' ", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        Exit Sub

    End If

    '----------------------------------------
    I = 1
    Do

        If mytablex.EOF Then Exit Do

        Select Case I

            Case 1
                cseccion1 = "" & mytablex.Fields("codigo")

            Case 2
                cseccion2 = "" & mytablex.Fields("codigo")

            Case 3
                cseccion3 = "" & mytablex.Fields("codigo")

            Case 4
                cseccion4 = "" & mytablex.Fields("codigo")

            Case 5
                cseccion5 = "" & mytablex.Fields("codigo")

            Case 6
                cseccion6 = "" & mytablex.Fields("codigo")

            Case 7
                cseccion7 = "" & mytablex.Fields("codigo")

            Case 8
                cseccion8 = "" & mytablex.Fields("codigo")

            Case 9
                cseccion9 = "" & mytablex.Fields("codigo")

            Case 10
                cseccion10 = "" & mytablex.Fields("codigo")

        End Select

        I = I + 1
        mytablex.MoveNext
    Loop
    '------------------------------------- ------------
    mytablex.Close
 
End Sub

Sub proceso_impresion1(bxlocal As String, _
                       bxtipo As String, _
                       bxserie As String, _
                       bxnumero As String, _
                       bxacu As String)

    Dim found    As Integer

    Dim archivot As String

    On Error GoTo cmd6_err:

    cerrar_archivo
    factura_formatox bxlocal, "" & bxtipo, bxserie, "" & bxnumero, "", bxacu
    cerrar_archivo
    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    found = control_impresionxx(bxtipo, bxacu)

    If found = 1 Then
        genver.file = globaldir & "\temporal\" & gusuario & ".txt"
        genver.Show 1

    End If

    Exit Sub
cmd6_err:
    MsgBox "Mensaje, Error al iniciar Impresion " & error$
    Exit Sub

End Sub

Function control_impresionxx(bxtipo As String, bxacu As String)

    Dim found  As Integer

    Dim xcolax As String

    Dim oldprinter

    Dim mytablex As New ADODB.Recordset

    Dim xxpuerto As String

    Dim sFile    As String

    Dim sw       As String

    On Error GoTo cmd67111_err

    xxpuerto = "X_"
    xcolax = ""
    sw = ""

    mytablex.Open "SELECT * FROM tipo where  tipo='" & bxtipo & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        xxpuerto = "" & mytablex.Fields("puerto")

        Select Case bxacu

            Case "V"
            
                If Len(cajero) > 0 Then
                    If caja <> "00" Then
                        xcolax = "" & mytable11.Fields("crin")
                        xxpuerto = "" & mytable11.Fields("puertore")
                        sw = "" & mytable11.Fields("iri")

                    End If

                End If
            
            Case "W"
            
                If Len(cajero) > 0 Then
                    If caja <> "00" Then
                        xcolax = "" & mytable11.Fields("creg")
                        xxpuerto = "" & mytable11.Fields("puertori")
                        sw = "" & mytable11.Fields("ire")

                    End If

                End If

        End Select

    End If

    mytablex.Close
    control_impresionxx = 1

    'MsgBox ""
    'MsgBox xxpuerto
    If Len(cajero) > 0 Then  'imprime en caja registradora
        If caja <> "00" Then
            If sw = "S" Then
                If MsgBox("Desea Imprimir", 1 + 256, "Aviso") <> 1 Then
                    control_impresionxx = 1
                    Exit Function

                End If

            End If

            If xcolax = "S" Then
                oldprinter = Printer.DeviceName
                selecciona_impresoras (xxpuerto)
                sFile = globaldir & "\temporal\" & gusuario & ".txt"
                found = Imprime_archivojj(sFile, 0, "8", "", "S", "")
                control_impresionxx = 2
                selecciona_impresoras (oldprinter)

            End If
   
            If xcolax <> "S" Then
                found = abre_puerto("" & mytable11.Fields("capuerto"), Val("" & mytable11.Fields("catipo")), "" & mytable11.Fields("gavetacola"))
                found = star_sp342(xxpuerto, 0)
                found = corte_papel(xxpuerto, Val("" & mytable11.Fields("catipo")))
                control_impresionxx = 0

            End If

        End If 'fin de 00

    End If

    Exit Function
cmd67111_err:
    MsgBox "Error en control inpresionxx " + error$, 48, "Aviso"
    Exit Function

End Function

Function busca_archivo_formato(bxtipo As String, bxacu) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM tipo where  tipo='" & bxtipo & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then

        Select Case bxacu

            Case "V"  'egreso
                busca_archivo_formato = "" & mytablex.Fields("archivoe")

                If caja <> "00" Then
                    busca_archivo_formato = "" & mytable11.Fields("archivorE")

                End If

            Case "W"  'ingreso
                busca_archivo_formato = "" & mytablex.Fields("archivo")

                If caja <> "00" Then
                    busca_archivo_formato = "" & mytable11.Fields("archivori")

                End If

        End Select

    End If

    mytablex.Close
 
End Function

Sub factura_formatox(bxlocal As String, _
                     bxtipo As String, _
                     bxserie As String, _
                     bxnumero As String, _
                     ascopia As String, _
                     axacu As String)

    Dim vacu            As String

    Dim mytablex        As New ADODB.Recordset

    Dim found           As Integer

    Dim nro_lineas      As Integer

    Dim contando        As Integer

    Dim faltante        As Integer

    Dim I               As Integer

    Dim archivo_formato As String

    Dim xtipoarchivo    As String

    Dim mytabley        As New ADODB.Recordset

    On Error GoTo cmd450009_err

    'If tipoclie = "C" Then
    'xtipoarchivo = "CUENTACD"
    'End If
    'If tipoclie = "P" Then
    'xtipoarchivo = "CUENTAPD"
    'End If
    'If tipoclie = "V" Then
    'xtipoarchivo = "CUENTACD"
    'End If
    xtipoarchivo = xcuentaco

    vacu = ""
    nro_lineas = 13
    contando = 0
    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    found = borra_nombre("" & FileName)
    archivo_formato = busca_archivo_formato(bxtipo, axacu)

    If Len(archivo_formato) = 0 Then
        MsgBox "No existe archivo formato ", 48, "Aviso"
        Exit Sub

    End If

    'recibo
       
    mytabley.Open "SELECT * FROM recibo where  local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytabley.RecordCount = 0 Then
        Exit Sub

    End If

    'proceso_formatos(archivo_formato , mydbx , mytablex , ubicacioni , ubicacionf , basedatos , indice , tipo , numero , ascopia , contando )
    found = proceso_formatos(archivo_formato, mytabley, "{", "}", "RECIBO", "RECIBO", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    vacu = "" & mytabley.Fields("acu")
    'End If
    'mytabley.Close
    '
    'detalle
    flag_contando = 0
    'Set mydbx = OpenDatabase(globaldat, False, False, "foxpro 2.5;")
    'MsgBox xtipoarchivo
    mytablex.Open "SELECT * FROM " & xtipoarchivo & "  where  local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        Do

            If mytablex.EOF Then Exit Do
            flag_contando = contando + 1
            'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
            'found = proceso_formatos(archivo_formato, mytablex, "/", "\", xtipoarchivo, "tmpcta", bxlocal, bxtipo, bxserie, bxnumero, ascopia, contando)
            found = proceso_formatos(archivo_formato, mytablex, "/", "\", xtipoarchivo, "tmpcta", bxlocal, bxtipo, bxserie, bxnumero, ascopia, contando)
            'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
            contando = contando + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close
    '
    'If bxtipo <> "1" And bxtipo <> "2" And bxtipo <> "5" Then
    '   If contando < nro_lineas Then
    '      For i = contando To nro_lineas
    '          Open filename For Append As #1
    '          found = formateaa("", 1, 2, 0)
    '          Close #1
    '      Next i
    '   End If
    'End If
    '----- SUBTOTAL
       
    'Set mydbx = OpenDatabase(globaldat, False, False, "foxpro 2.5;")
    'Set mytablex = mydbxglo.OpenTable("RECIBO")
    'mytablex.Index = "RECIBO"
       
    'mytabley.Seek "=", bxlocal, bxtipo, bxserie, bxnumero
    'If Not mytabley.NoMatch Then
    'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
    'found = proceso_formatos(archivo_formato, mytabley, "$", "?", "recibo", "recibo", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    found = proceso_formatos(archivo_formato, mytabley, "$", "?", "recibo", "recibo", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
          
    'End If
    'mytablex.Close
    '
    'forma de pago
    'Set mydbx = OpenDatabase(globaldat, False, False, "foxpro 2.5;")
        
    mytablex.Open "SELECT * FROM " & gofpago & " where  local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
       
        Do

            If mytablex.EOF Then Exit Do
            'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
            'found = proceso_formatos(archivo_formato, mytablex, "<", ">", gofpago, "fpagov", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
            found = proceso_formatos(archivo_formato, mytablex, "<", ">", gofpago, "fpagov", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
            'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
                
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close
    '
    '----------pie de paginatotal  xxxx
    'Set mydbx = OpenDatabase(globaldat, False, False, "foxpro 2.5;")
    'Set mytablex = mydbxglo.OpenTable("recibo")
    'mytablex.Index = "recibo"
    'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
    'found = proceso_formatos(archivo_formato, mytabley, "^", "&", "RECIBO", "RECIBO", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    found = proceso_formatos(archivo_formato, mytabley, "^", "&", "RECIBO", "RECIBO", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
           
    mytabley.Close
    Exit Sub
cmd450009_err:
    MsgBox "Aviso Proceso en Compilando " & error$, 24, "Aviso"
    Exit Sub

End Sub

Function anular_recibo(xlocal As String, _
                       xtipo As String, _
                       xserie As String, _
                       xnumero As String)

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim sw       As Integer

    Dim found    As Integer

    On Error GoTo cmd4312_err

    mytablex.Open "SELECT * FROM recibo where  local='" & xlocal & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        'mytablex.Edit
        mytablex.Fields("estado") = "1"
        mytablex.Update
        sw = 1

    End If

    mytablex.Close
    'ahora forma de pago
    mytablex.Open "SELECT * FROM " & gofpago & " where  local='" & xlocal & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        'mytablex.Edit
        mytablex.Fields("estado") = "1"
        mytablex.Update

    End If

    mytablex.Close
    'ahora los temporales----------------------------------------------------------------
    found = descarga_cuentac(xlocal, xtipo, xserie, xnumero, "-1")
    '------------------------------------------------------------------------------------
    'si existe en bancos borrarlo
as1k:
    mytablex.Open "SELECT * FROM chequemo where  transaccio='" & caja & xnumero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        mytablex.Delete
        mytablex.Close
        GoTo as1k

    End If

    mytablex.Close
    anular_recibo = 1
    Exit Function
cmd4312_err:
    mytablex.Close
 
    Exit Function

End Function

Function ver_credito(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM fpago where  fpago='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        ver_credito = "" & mytablex.Fields("tipo")

    End If

    mytablex.Close

End Function

Sub anula_credito(xlocal As String, xtipo As String, xserie As String, xnumero As String)

    Dim buf1 As String

    Select Case xcuentaco

        Case "CUENTAC"
            buf1 = "cuentacd"

        Case "CUENTAP"
            buf1 = "cuentapd"

    End Select

    cn.Execute ("delete from " & xcuentaco & " where local='" & xlocal & "' and tipo='" & xtipo & "' and serie='" & xserie & " ' and numero='" & xnumero & "")
    cn.Execute ("delete from " & XCUENTACO1 & " where local='" & xlocal & "' and tipo='" & xtipo & "' and serie='" & xserie & " ' and numero='" & xnumero & "")

End Sub

Function busca_local()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM tlocal where  codigo='" & local1 & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_local = 1

    End If

    mytablex.Close

End Function

Function valida_ingreso() As String

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    Dim sdx      As Double

    sdx = 0

    If acu = "W" Then  'ingreso
        sdx = Val("" & mytable11.Fields("numerori")) + 1

    End If

    If acu = "V" Then
        sdx = Val("" & mytable11.Fields("numerore")) + 1

    End If

    MsgBox serie
vienea:
    buf = "" & sdx
    mytablex.Open "SELECT * FROM recibo where  local='" & local1 & "' and tipo='" & tipo & "' and serie='" & serie & "' and numero='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        mytablex.Close
        sdx = sdx + 1
        GoTo vienea

    End If

    valida_ingreso = "" & sdx
    mytablex.Close

End Function

Function graba_credito(mytabley As Table, buf As String, buf2 As String, buf3 As String)

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd6712121_err

    mytablex.Open "SELECT * FROM " & buf3 & " where  local='" & local1 & "' and tipo='" & tipo & "' and serie='" & serie & "' and numero='" & Numero & "' and cuota='1'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.AddNew
        mytablex.Fields("grupo") = buf2
        mytablex.Fields("acu") = acu
        mytablex.Fields("observa") = Mid$("" & mytabley.Fields("descripcio"), 1, 20)
        mytablex.Fields("fpago") = buf
        mytablex.Fields("tipo") = "" & tipo
        mytablex.Fields("serie") = "" & serie
        mytablex.Fields("numero") = "" & Numero
        mytablex.Fields("dias") = Val("" & mytabley.Fields("dias"))
        mytablex.Fields("cuota") = "1"
        mytablex.Fields("tipoclie") = tipoclie
        mytablex.Fields("codigo") = "" & mytabley.Fields("codigo")
        mytablex.Fields("nombre") = "" & mytabley.Fields("nombre")
        mytablex.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
        mytablex.Fields("fechav") = Format(CVDate(fecha) + Val("" & mytabley.Fields("dias")), "dd/mm/yyyy")
        mytablex.Fields("moneda") = "" & mytabley.Fields("moneda")
        mytablex.Fields("total") = Val("" & mytabley.Fields("recibe"))
        mytablex.Fields("abono") = 0
        mytablex.Fields("interes") = 0
        mytablex.Fields("saldo") = Val("" & mytabley.Fields("recibe"))
        mytablex.Fields("estado") = "0"
        mytablex.Fields("vendedor") = vendedor
        mytablex.Fields("usuario") = cajero
        mytablex.Fields("caja") = caja
        mytablex.Fields("turno") = turno
        mytablex.Fields("zona") = ""
        mytablex.Fields("local") = "" & local1
        mytablex.Fields("observa") = Mid$("" & observa, 1, 20)
        mytablex.Update

    End If

    mytablex.Close
    Exit Function
cmd6712121_err:
    MsgBox "Aviso en Graba Credito " + error$, 48, "Aviso"
    Exit Function

End Function

Function graba_credito_letra(mytabley As ADODB.Recordset, _
                             buf As String, _
                             buf2 As String, _
                             buf3 As String)

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd67121211_err

    'MsgBox buf3
    mytablex.Open "SELECT * FROM " & buf3 & " where  local='" & local1 & "' and tipo='LE' and serie='LE' and numero='" & "" & mytabley.Fields("dias") & "' ", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.AddNew
        mytablex.Fields("grupo") = buf2
        mytablex.Fields("acu") = acu
        mytablex.Fields("observa") = "" 'Mid$("" & mytabley.Fields("descripcio"), 1, 20)
        mytablex.Fields("fpago") = buf
        mytablex.Fields("tipo") = "LE"
        mytablex.Fields("serie") = "LE" '& serie
        mytablex.Fields("numero") = "" & mytabley.Fields("numero")
        'mytablex.Fields("dias") = Val("" & mytabley.Fields("dias"))
        mytablex.Fields("cuota") = "1"
        mytablex.Fields("tipoclie") = tipoclie
        mytablex.Fields("codigo") = "" & codigo
        mytablex.Fields("nombre") = "" & nombre
        mytablex.Fields("fecha") = Format("" & mytabley.Fields("fechai"), "dd/mm/yyyy")
        mytablex.Fields("fechav") = Format("" & mytabley.Fields("fechaf"), "dd/mm/yyyy")
        mytablex.Fields("moneda") = "" & mytabley.Fields("moneda")
        mytablex.Fields("total") = Val("" & mytabley.Fields("valor"))
        mytablex.Fields("abono") = 0
        mytablex.Fields("interes") = 0
        mytablex.Fields("saldo") = Val("" & mytabley.Fields("valor"))
        mytablex.Fields("estado") = "0"
        mytablex.Fields("vendedor") = vendedor
        mytablex.Fields("usuario") = cajero
        mytablex.Fields("caja") = caja
        mytablex.Fields("turno") = turno
        mytablex.Fields("zona") = ""
        mytablex.Fields("local") = "" & local1
        mytablex.Fields("observa") = "" ' Mid$("" & observa, 1, 20)
        mytablex.Update

    End If

    mytablex.Close
    Exit Function
cmd67121211_err:
    MsgBox "Aviso en Graba Credito Letras " + error$, 48, "Aviso"
    Exit Function

End Function

Function graba_credito_letrae(buf As String, buf2 As String, buf3 As String)

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd6712121112_err

    mytablex.Open "SELECT * FROM " & buf3 & " where  local='" & local1 & "' and tipo='" & "" & tabletra.Fields("tipo") & "' and serie='" & "" & tabletra.Fields("tipo") & "' and numero='" & "" & tabletra.Fields("NUMERO") & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.AddNew
        mytablex.Fields("grupo") = buf2
        mytablex.Fields("acu") = acu
        mytablex.Fields("fpago") = buf
        mytablex.Fields("tipo") = "" & tabletra.Fields("tipo")
        mytablex.Fields("serie") = "" & tabletra.Fields("serie") '& serie
        mytablex.Fields("numero") = "" & tabletra.Fields("numero")
        mytablex.Fields("dias") = 1
        mytablex.Fields("cuota") = "1"
        mytablex.Fields("tipoclie") = tipoclie
        mytablex.Fields("codigo") = "" & codigo
        mytablex.Fields("nombre") = "" & nombre
        mytablex.Fields("fecha") = Format("" & tabletra.Fields("fechai"), "dd/mm/yyyy")
        mytablex.Fields("fechav") = Format("" & tabletra.Fields("fechaf"), "dd/mm/yyyy")
        mytablex.Fields("moneda") = "" & tabletra.Fields("moneda")
        mytablex.Fields("total") = Val("" & tabletra.Fields("valor"))
        mytablex.Fields("abono") = 0
        mytablex.Fields("interes") = 0
        mytablex.Fields("saldo") = Val("" & tabletra.Fields("valor"))
        mytablex.Fields("estado") = "0"
        mytablex.Fields("vendedor") = vendedor
        mytablex.Fields("usuario") = cajero
        mytablex.Fields("caja") = caja
        mytablex.Fields("turno") = turno
        mytablex.Fields("zona") = ""
        mytablex.Fields("local") = "" & local1
        mytablex.Fields("observa") = "" ' Mid$("" & observa, 1, 20)
        mytablex.Update

    End If

    mytablex.Close
    Exit Function
cmd6712121112_err:
    MsgBox "Aviso en Graba Credito Letras e " + error$, 48, "Aviso"
    Exit Function

End Function

Sub graba_tarjetas(mytabley As ADODB.Recordset)

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Double

    On Error GoTo cmd7811_err

    'sdx = busca_banco_numero()

    mytablex.Open "SELECT * FROM chequemo where  transaccio='" & caja & Numero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.AddNew
        mytablex.Fields("transaccio") = caja & Numero
        mytablex.Fields("moneda") = "" & mytabley.Fields("moneda")
        mytablex.Fields("tipoclie") = "C"
        mytablex.Fields("codigo") = "" & mytabley.Fields("codigo")
        mytablex.Fields("banco") = ""
        mytablex.Fields("cuenta") = ""
        mytablex.Fields("tipo") = "72"
        mytablex.Fields("numero") = ""
        mytablex.Fields("fechan") = Format("" & mytabley.Fields("fecha"), "dd/mm/yyyy")
        mytablex.Fields("fechae") = Format("" & mytabley.Fields("fecha"), "dd/mm/yyyy")
        mytablex.Fields("nombre") = "" & mytabley.Fields("nombre")
        mytablex.Fields("conciliado") = "N"
        mytablex.Fields("concepto") = "" & mytabley.Fields("descripcio")
        mytablex.Fields("acu") = "X"
        mytablex.Fields("comenta") = ""
        mytablex.Fields("total") = Val("" & mytabley.Fields("recibe"))
        mytablex.Fields("descuento") = 0
        mytablex.Fields("recargo") = 0
        mytablex.Fields("abono") = 0
        mytablex.Fields("neto") = Val("" & mytabley.Fields("recibe"))
        mytablex.Fields("saldo") = Val("" & mytabley.Fields("recibe"))
        mytablex.Fields("cajero") = "" & cajero
        mytablex.Fields("caja") = "" & caja
        mytablex.Fields("turno") = "" & turno
        mytablex.Fields("xtipo") = "" & mytabley.Fields("tipo")
        mytablex.Fields("xserie") = "" & mytabley.Fields("serie")
        mytablex.Fields("xnumero") = "" & mytabley.Fields("numero")
        mytablex.Update

    End If

    mytablex.Close
    Exit Sub
cmd7811_err:
    MsgBox "Aviso en graba tarjetas " + error$, 48, "Aviso"
    Exit Sub

End Sub

Function busca_banco_numero() As Double

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM parame where  codigo='01'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_banco_numero = Val("" & mytablex.Fields("banco"))

    End If

    mytablex.Close

End Function

Function busca_anticipo() As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM tipo where  tipo='" & tipo & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_anticipo = "" & mytablex.Fields("anticipo")

    End If

    mytablex.Close

End Function

Function verifica_fpagoletra()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM fpago where  tipo='G'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        fpago = "" & mytablex.Fields("fpago")
        verifica_fpagoletra = 1

    End If

    mytablex.Close

End Function

Function existe_letra()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM _k" & gusuario & " where  tipo='" & letipo & "' and serie='" & leserie & "' and numero='" & lenumero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        existe_letra = 1

    End If

    mytablex.Close

End Function

Sub suma_tabletra()

    Dim sdx As Double

    sdx = 0
    lssubtotal = ""
    lssaldo = lstotal

    If tabletra.RecordCount > 0 Then
        tabletra.MoveFirst

    End If

    Do

        If tabletra.EOF Then Exit Do
        sdx = sdx + Val("" & tabletra.Fields("valor"))
        tabletra.MoveNext
    Loop
    lssubtotal = Format(sdx, "0.00")
    sdx = Val(lstotal) - Val(lssubtotal)
    lssaldo = Format(sdx, "0.00")

End Sub

Function existe_letrad()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM " & xcuentacol & " where  tipo='" & letipo & "' and serie='" & leserie & "' and numero='" & lenumero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        existe_letrad = 1

    End If

    mytablex.Close

End Function

Sub carga_concepto()

    Dim mytablex As New ADODB.Recordset

    concepto.Clear
    concepto.AddItem "%"
    mytablex.Open "SELECT * FROM concepto where grupo='" & acu & "'", cn, adOpenKeyset, adLockOptimistic
    'mytablex.Open "SELECT * FROM concepto ", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        concepto.AddItem Trim("" & mytablex.Fields("concepto")) & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    mytablex.Close
    concepto.ListIndex = 0

End Sub

Sub carga_subconcepto(buf As String)

    Dim mytablex As New ADODB.Recordset

    subconcepto.Clear
    subconcepto.AddItem "%"
    mytablex.Open "SELECT * FROM subconcepto where concepto='" & buf & "'", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        subconcepto.AddItem Trim("" & mytablex.Fields("subconcepto")) & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    mytablex.Close
    subconcepto.ListIndex = 0

End Sub

