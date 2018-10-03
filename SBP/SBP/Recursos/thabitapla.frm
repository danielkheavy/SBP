VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form thabitapla 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Planning de Habitaciones"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   12240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Consulta"
      Height          =   4215
      Left            =   0
      TabIndex        =   110
      Top             =   0
      Width           =   8175
      Begin VB.CommandButton Command3 
         Caption         =   "Filtrar"
         Height          =   495
         Left            =   3240
         TabIndex        =   113
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox fecha 
         Height          =   495
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   112
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label17 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Desde"
         Height          =   495
         Left            =   120
         TabIndex        =   111
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Leyenda"
      Height          =   4215
      Left            =   0
      TabIndex        =   92
      Top             =   0
      Visible         =   0   'False
      Width           =   8175
      Begin VB.CommandButton Command2 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   3360
         TabIndex        =   109
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label16 
         BackColor       =   &H00800080&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5280
         TabIndex        =   108
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "Reserva Congresos"
         Height          =   495
         Left            =   3720
         TabIndex        =   107
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5280
         TabIndex        =   106
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Reserva Empresa"
         Height          =   495
         Left            =   3720
         TabIndex        =   105
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5280
         TabIndex        =   104
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Reserva Agencia"
         Height          =   495
         Left            =   3720
         TabIndex        =   103
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5280
         TabIndex        =   102
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Reserva Directa"
         Height          =   495
         Left            =   3720
         TabIndex        =   101
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1800
         TabIndex        =   100
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Posible Reservar"
         Height          =   495
         Left            =   240
         TabIndex        =   99
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1800
         TabIndex        =   98
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Fuera Servicio"
         Height          =   495
         Left            =   240
         TabIndex        =   97
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1800
         TabIndex        =   96
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Ocupado"
         Height          =   495
         Left            =   240
         TabIndex        =   95
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1800
         TabIndex        =   94
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Libre"
         Height          =   495
         Left            =   240
         TabIndex        =   93
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Leyenda"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   91
      Top             =   8280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid dbgrid1 
      Height          =   7815
      Left            =   0
      TabIndex        =   90
      Top             =   1800
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   13785
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   0
      RowHeight       =   18
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   31
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
      BeginProperty Column05 
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
      BeginProperty Column06 
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
      BeginProperty Column07 
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
      BeginProperty Column08 
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
      BeginProperty Column09 
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
      BeginProperty Column10 
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
      BeginProperty Column11 
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
      BeginProperty Column12 
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
      BeginProperty Column13 
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
      BeginProperty Column14 
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
      BeginProperty Column15 
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
      BeginProperty Column16 
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
      BeginProperty Column17 
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
      BeginProperty Column18 
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
      BeginProperty Column19 
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
      BeginProperty Column20 
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
      BeginProperty Column21 
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
      BeginProperty Column22 
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
      BeginProperty Column23 
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
      BeginProperty Column24 
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
      BeginProperty Column25 
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
      BeginProperty Column26 
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
      BeginProperty Column27 
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
      BeginProperty Column28 
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
      BeginProperty Column29 
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
      BeginProperty Column30 
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
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   360
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   360
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   360
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   329.953
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   390.047
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   299.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   390.047
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   329.953
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   390.047
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   329.953
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   329.953
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   345.26
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   390.047
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   345.26
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   390.047
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   315.213
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   360
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   329.953
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   404.787
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   360
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   345.26
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   404.787
         EndProperty
         BeginProperty Column24 
            ColumnWidth     =   345.26
         EndProperty
         BeginProperty Column25 
            ColumnWidth     =   360
         EndProperty
         BeginProperty Column26 
            ColumnWidth     =   360
         EndProperty
         BeginProperty Column27 
            ColumnWidth     =   360
         EndProperty
         BeginProperty Column28 
            ColumnWidth     =   360
         EndProperty
         BeginProperty Column29 
            ColumnWidth     =   360
         EndProperty
         BeginProperty Column30 
            ColumnWidth     =   315.213
         EndProperty
      EndProperty
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   29
      Left            =   11760
      TabIndex        =   89
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   28
      Left            =   11400
      TabIndex        =   88
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   27
      Left            =   11040
      TabIndex        =   87
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   26
      Left            =   10680
      TabIndex        =   86
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   25
      Left            =   10320
      TabIndex        =   85
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   24
      Left            =   9960
      TabIndex        =   84
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   23
      Left            =   9600
      TabIndex        =   83
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   22
      Left            =   9240
      TabIndex        =   82
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   21
      Left            =   8880
      TabIndex        =   81
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   20
      Left            =   8520
      TabIndex        =   80
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   19
      Left            =   8160
      TabIndex        =   79
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   18
      Left            =   7800
      TabIndex        =   78
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   17
      Left            =   7440
      TabIndex        =   77
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   16
      Left            =   7080
      TabIndex        =   76
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   6720
      TabIndex        =   75
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   6360
      TabIndex        =   74
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   6000
      TabIndex        =   73
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   5640
      TabIndex        =   72
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   5280
      TabIndex        =   71
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   4920
      TabIndex        =   70
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   4560
      TabIndex        =   69
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   4200
      TabIndex        =   68
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   3840
      TabIndex        =   67
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   3480
      TabIndex        =   66
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   3120
      TabIndex        =   65
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   2760
      TabIndex        =   64
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   2400
      TabIndex        =   63
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   2040
      TabIndex        =   62
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1680
      TabIndex        =   61
      Top             =   720
      Width           =   375
   End
   Begin VB.Label numero 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1320
      TabIndex        =   60
      Top             =   720
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   29
      Left            =   11760
      TabIndex        =   59
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   28
      Left            =   11400
      TabIndex        =   58
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   27
      Left            =   11040
      TabIndex        =   57
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   26
      Left            =   10680
      TabIndex        =   56
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   25
      Left            =   10320
      TabIndex        =   55
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   24
      Left            =   9960
      TabIndex        =   54
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   23
      Left            =   9600
      TabIndex        =   53
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   22
      Left            =   9240
      TabIndex        =   52
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   21
      Left            =   8880
      TabIndex        =   51
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   20
      Left            =   8520
      TabIndex        =   50
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   19
      Left            =   8160
      TabIndex        =   49
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   18
      Left            =   7800
      TabIndex        =   48
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   17
      Left            =   7440
      TabIndex        =   47
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   16
      Left            =   7080
      TabIndex        =   46
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   15
      Left            =   6720
      TabIndex        =   45
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   14
      Left            =   6360
      TabIndex        =   44
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   13
      Left            =   6000
      TabIndex        =   43
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   12
      Left            =   5640
      TabIndex        =   42
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   11
      Left            =   5280
      TabIndex        =   41
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   10
      Left            =   4920
      TabIndex        =   40
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   9
      Left            =   4560
      TabIndex        =   39
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   8
      Left            =   4200
      TabIndex        =   38
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   7
      Left            =   3840
      TabIndex        =   37
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   6
      Left            =   3480
      TabIndex        =   36
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   5
      Left            =   3120
      TabIndex        =   35
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   4
      Left            =   2760
      TabIndex        =   34
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   3
      Left            =   2400
      TabIndex        =   33
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   2
      Left            =   2040
      TabIndex        =   32
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   1
      Left            =   1680
      TabIndex        =   31
      Top             =   360
      Width           =   375
   End
   Begin VB.Label dias 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   0
      Left            =   1320
      TabIndex        =   30
      Top             =   360
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   29
      Left            =   11760
      TabIndex        =   29
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   28
      Left            =   11400
      TabIndex        =   28
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   27
      Left            =   11040
      TabIndex        =   27
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   26
      Left            =   10680
      TabIndex        =   26
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   25
      Left            =   10320
      TabIndex        =   25
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   24
      Left            =   9960
      TabIndex        =   24
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   23
      Left            =   9600
      TabIndex        =   23
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   22
      Left            =   9240
      TabIndex        =   22
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   21
      Left            =   8880
      TabIndex        =   21
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   20
      Left            =   8520
      TabIndex        =   20
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   19
      Left            =   8160
      TabIndex        =   19
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   18
      Left            =   7800
      TabIndex        =   18
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   17
      Left            =   7440
      TabIndex        =   17
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   16
      Left            =   7080
      TabIndex        =   16
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   15
      Left            =   6720
      TabIndex        =   15
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   14
      Left            =   6360
      TabIndex        =   14
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   13
      Left            =   6000
      TabIndex        =   13
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   12
      Left            =   5640
      TabIndex        =   12
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   11
      Left            =   5280
      TabIndex        =   11
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   10
      Left            =   4920
      TabIndex        =   10
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   9
      Left            =   4560
      TabIndex        =   9
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   8
      Left            =   4200
      TabIndex        =   8
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   7
      Left            =   3840
      TabIndex        =   7
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   6
      Left            =   3480
      TabIndex        =   6
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   5
      Left            =   3120
      TabIndex        =   5
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   4
      Left            =   2760
      TabIndex        =   4
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   3
      Left            =   2400
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   2
      Left            =   2040
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
   Begin VB.Label meses 
      BorderStyle     =   1  'Fixed Single
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
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
   Begin VB.Menu dl883 
      Caption         =   "&Consulta"
   End
   Begin VB.Menu fk444 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "thabitapla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub sumar_dias()

    Dim I    As Integer

    Dim D    As Integer

    Dim hoyi As String

    Dim hoyf As String

    Dim sw   As Integer

    Dim dd   As String

    Dim mm   As String

    hoyi = Format(fecha, "dd/mm/yyyy")
    hoyf = hoyi

    For I = 0 To 29
        Numero(I).Caption = Day(CVDate(hoyf))
        D = Weekday(CVDate(hoyf))

        Select Case D

            Case 1: dias(I).Caption = "Dom"

            Case 2: dias(I).Caption = "Lun"

            Case 3: dias(I).Caption = "Mar"

            Case 4: dias(I).Caption = "Mie"

            Case 5: dias(I).Caption = "Jue"

            Case 6: dias(I).Caption = "Vie"

            Case 7: dias(I).Caption = "Sab"

        End Select

        selecciona_mes Month(hoyf), I
        hoyi = DateAdd("D", 1, hoyi)
        hoyf = Format(hoyi, "dd/mm/yyyy")

    Next I

End Sub

Sub selecciona_mes(buf As Integer, I As Integer)

    Select Case buf

        Case 1:   meses(I).Caption = "Ene"

        Case 2:   meses(I).Caption = "Feb"

        Case 3:   meses(I).Caption = "Mar"

        Case 4:   meses(I).Caption = "Abr"

        Case 5:   meses(I).Caption = "May"

        Case 6:   meses(I).Caption = "Jun"

        Case 7:   meses(I).Caption = "Jul"

        Case 8:   meses(I).Caption = "Ago"

        Case 9:   meses(I).Caption = "Set"

        Case 10:  meses(I).Caption = "Oct"

        Case 11:  meses(I).Caption = "Nov"

        Case 12:  meses(I).Caption = "Dic"

    End Select

End Sub

Private Sub Command1_Click()
    Frame1.Visible = True

End Sub

Private Sub Command2_Click()
    Frame1.Visible = False

End Sub

Private Sub Command3_Click()

    If Not IsDate(fecha) Then
        fecha = Format(Now, "dd/mm/yyyy")
        Exit Sub

    End If

    Frame2.Visible = False
    sumar_dias
    carga_estado

End Sub

Private Sub dl883_Click()

    If Frame1.Visible = True Then Exit Sub
    Frame2.Visible = True
    fecha = Format(Now, "dd/mm/yyyy")
    fecha.SetFocus

End Sub

Private Sub fecha_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    Command3_Click

End Sub

Private Sub fk444_Click()

    If Frame1.Visible = True Then
        Frame1.Visible = False
        Exit Sub

    End If

    If Frame2.Visible = True Then
        Frame2.Visible = False
        Exit Sub

    End If

    thabitapla.Hide
    Unload thabitapla

End Sub

Private Sub Form_Load()
    fecha = Format(Now, "dd/mm/yyyy")

End Sub

Sub carga_estado()

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    cn.Execute ("delete from habitacionpla")
    mytabley.Open "select * from habitacionpla ", cn, adOpenStatic, adLockOptimistic
    mytablex.Open "select * from habitacion", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        mytabley.AddNew
        mytabley.Fields("habitacion") = Trim("" & mytablex.Fields("habitacion"))
        adiciona_dias mytabley, Trim("" & mytablex.Fields("habitacion"))
        mytabley.Update
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close
    visualiza_plan

End Sub

Sub adiciona_dias(mytablex As ADODB.Recordset, habitacion As String)

    Dim I        As Integer

    Dim buf      As String

    Dim hoyi     As String

    Dim hoyf     As String

    Dim X        As Integer

    Dim mytabley As New ADODB.Recordset

    hoyi = Format(fecha, "dd/mm/yyyy")
    hoyf = hoyi
    X = 1

    If Not IsDate(hoyf) Then Exit Sub

    For I = 0 To 29
        buf = "SELECT     dbo.hotelcheckin.ESTADO,dbo.hotelcheckin.checkin, dbo.hotelcheckin.habitacion"
        buf = buf & " FROM         dbo.hotelcheckin "
        buf = buf & " where   dbo.hotelcheckin.arribofecha<='" & Format(hoyf, "YYYYMMDD") & "'"
        buf = buf & " and   dbo.hotelcheckin.arribofechaf>='" & Format(hoyf, "YYYYMMDD") & "'"
        buf = buf & " and dbo.hotelcheckin.habitacion='" & Trim(habitacion) & "'"
        mytabley.Open buf, cn, adOpenStatic, adLockOptimistic

        If mytabley.RecordCount > 0 Then
            If "" & mytabley.Fields("estado") = "RESERVA" Then
                mytablex.Fields("l" & X) = "R"

            End If

            If "" & mytabley.Fields("estado") = "ENTRADA" Then
                mytablex.Fields("l" & X) = "O"

            End If

        Else
            mytablex.Fields("l" & X) = ""

        End If

        mytabley.Close
        hoyi = DateAdd("D", 1, hoyi)
        hoyf = Format(hoyi, "dd/mm/yyyy")
        X = X + 1
    Next I

End Sub

Sub visualiza_plan()

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from habitacionpla order by habitacion ", cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = mytablex
    dbGrid1.columns(0).Width = 1000

    For I = 1 To 29
        dbGrid1.columns(I).Width = 360
    Next I

    dbGrid1.refresh

End Sub
