VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tvendedo 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tabla de Personal"
   ClientHeight    =   8385
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   13470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   13470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Consulta"
      Height          =   8295
      Left            =   4560
      TabIndex        =   162
      Top             =   360
      Visible         =   0   'False
      Width           =   12615
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
         Left            =   8280
         TabIndex        =   165
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
         TabIndex        =   164
         TabStop         =   0   'False
         Top             =   600
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
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   163
         TabStop         =   0   'False
         Top             =   600
         Width           =   2895
      End
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   6735
         Left            =   240
         TabIndex        =   166
         Top             =   1200
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   11880
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Parametros Adicionales"
      Height          =   6615
      Left            =   2760
      TabIndex        =   92
      Top             =   840
      Visible         =   0   'False
      Width           =   8775
      Begin VB.TextBox clave 
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
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   131
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox rw1 
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
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   1
         TabIndex        =   130
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox v1 
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
         IMEMode         =   3  'DISABLE
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   129
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox rw2 
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
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   1
         TabIndex        =   128
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox v2 
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
         IMEMode         =   3  'DISABLE
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   127
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox rw3 
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
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   1
         TabIndex        =   126
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox v3 
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
         IMEMode         =   3  'DISABLE
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   125
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox rw4 
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
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   1
         TabIndex        =   124
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox v4 
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
         IMEMode         =   3  'DISABLE
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   123
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox rw5 
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
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   1
         TabIndex        =   122
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox v5 
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
         IMEMode         =   3  'DISABLE
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   121
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox rw6 
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
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   1
         TabIndex        =   120
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox v6 
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
         IMEMode         =   3  'DISABLE
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   119
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox rw7 
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
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   1
         TabIndex        =   118
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox v7 
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
         IMEMode         =   3  'DISABLE
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   117
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox rw8 
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
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   1
         TabIndex        =   116
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox v8 
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
         IMEMode         =   3  'DISABLE
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   115
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox rw9 
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
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   1
         TabIndex        =   114
         Top             =   3960
         Width           =   615
      End
      Begin VB.TextBox v9 
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
         IMEMode         =   3  'DISABLE
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   113
         Top             =   3960
         Width           =   615
      End
      Begin VB.TextBox rw10 
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
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   1
         TabIndex        =   112
         Top             =   4320
         Width           =   615
      End
      Begin VB.TextBox v10 
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
         IMEMode         =   3  'DISABLE
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   111
         Top             =   4320
         Width           =   615
      End
      Begin VB.TextBox vevend 
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
         IMEMode         =   3  'DISABLE
         Left            =   4200
         MaxLength       =   1
         TabIndex        =   110
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox veclave 
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
         IMEMode         =   3  'DISABLE
         Left            =   5400
         MaxLength       =   1
         TabIndex        =   109
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox caja 
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
         IMEMode         =   3  'DISABLE
         Left            =   6480
         MaxLength       =   1
         TabIndex        =   108
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox parame 
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
         IMEMode         =   3  'DISABLE
         Left            =   6480
         MaxLength       =   1
         TabIndex        =   107
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox parame1 
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
         IMEMode         =   3  'DISABLE
         Left            =   6480
         MaxLength       =   3
         TabIndex        =   106
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox parame3 
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
         IMEMode         =   3  'DISABLE
         Left            =   6480
         MaxLength       =   3
         TabIndex        =   105
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox parame2 
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
         IMEMode         =   3  'DISABLE
         Left            =   6480
         MaxLength       =   3
         TabIndex        =   104
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox parame4 
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
         IMEMode         =   3  'DISABLE
         Left            =   6480
         MaxLength       =   3
         TabIndex        =   103
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox cuadre 
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
         IMEMode         =   3  'DISABLE
         Left            =   5760
         MaxLength       =   1
         TabIndex        =   102
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox anula 
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
         IMEMode         =   3  'DISABLE
         Left            =   5760
         MaxLength       =   1
         TabIndex        =   101
         Top             =   4440
         Width           =   615
      End
      Begin VB.TextBox copia 
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
         IMEMode         =   3  'DISABLE
         Left            =   5760
         MaxLength       =   1
         TabIndex        =   100
         Top             =   4800
         Width           =   615
      End
      Begin VB.TextBox congela 
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
         IMEMode         =   3  'DISABLE
         Left            =   5760
         MaxLength       =   1
         TabIndex        =   99
         Top             =   5160
         Width           =   615
      End
      Begin VB.TextBox apertura 
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
         IMEMode         =   3  'DISABLE
         Left            =   5760
         MaxLength       =   1
         TabIndex        =   98
         Top             =   5520
         Width           =   615
      End
      Begin VB.TextBox v11 
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
         IMEMode         =   3  'DISABLE
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   97
         Top             =   4680
         Width           =   615
      End
      Begin VB.TextBox rw11 
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
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   1
         TabIndex        =   96
         Top             =   4680
         Width           =   615
      End
      Begin VB.TextBox rw12 
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
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   1
         TabIndex        =   95
         Top             =   5040
         Width           =   615
      End
      Begin VB.TextBox v12 
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
         IMEMode         =   3  'DISABLE
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   94
         Top             =   5040
         Width           =   615
      End
      Begin VB.TextBox vecostoimp 
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
         IMEMode         =   3  'DISABLE
         Left            =   5760
         MaxLength       =   1
         TabIndex        =   93
         Top             =   5880
         Width           =   615
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clave Acceso"
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
         TabIndex        =   159
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Menu"
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
         TabIndex        =   158
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R/W"
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
         TabIndex        =   157
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Visible"
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
         TabIndex        =   156
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label28 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tablas"
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
         Left            =   120
         TabIndex        =   155
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label29 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tienda"
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
         Left            =   120
         TabIndex        =   154
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label30 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ventas"
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
         Left            =   120
         TabIndex        =   153
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label31 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Compras"
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
         Left            =   120
         TabIndex        =   152
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label32 
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   151
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label33 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tesoreria"
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
         Left            =   120
         TabIndex        =   150
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label Label34 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Produccion"
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
         Left            =   120
         TabIndex        =   149
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label Label35 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Personal"
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
         Left            =   120
         TabIndex        =   148
         Top             =   3600
         Width           =   2175
      End
      Begin VB.Label Label36 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Contable"
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
         Left            =   120
         TabIndex        =   147
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Label Label37 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reporte"
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
         Left            =   120
         TabIndex        =   146
         Top             =   4320
         Width           =   2175
      End
      Begin VB.Label Label52 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vend"
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
         TabIndex        =   145
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label53 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clave"
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
         Left            =   4800
         TabIndex        =   144
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label54 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1.Ger 2.Adm 3.Sup 4.Cajero"
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
         TabIndex        =   143
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label55 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Solamente estas cajas?"
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
         Left            =   3600
         TabIndex        =   142
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label56 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cajas 01,02,03,Terminales T1,T2,T3"
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
         Height          =   1455
         Left            =   3600
         TabIndex        =   141
         Top             =   2520
         Width           =   2895
      End
      Begin VB.Label Label58 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuadre Caja"
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
         Left            =   3600
         TabIndex        =   140
         Top             =   4080
         Width           =   2175
      End
      Begin VB.Label Label59 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Anula en Caja"
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
         Left            =   3600
         TabIndex        =   139
         Top             =   4440
         Width           =   2175
      End
      Begin VB.Label Label60 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Saca Copia en Caja"
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
         Left            =   3600
         TabIndex        =   138
         Top             =   4800
         Width           =   2175
      End
      Begin VB.Label Label61 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descongela"
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
         Left            =   3600
         TabIndex        =   137
         Top             =   5160
         Width           =   2175
      End
      Begin VB.Label Label62 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Abre Gaveta"
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
         Left            =   3600
         TabIndex        =   136
         Top             =   5520
         Width           =   2175
      End
      Begin VB.Label clavex 
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
         Left            =   3600
         TabIndex        =   135
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label63 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Historia Clinica"
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
         Left            =   120
         TabIndex        =   134
         Top             =   4680
         Width           =   2175
      End
      Begin VB.Label Label64 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Importaciones"
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
         Left            =   120
         TabIndex        =   133
         Top             =   5040
         Width           =   2175
      End
      Begin VB.Label Label65 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ve Costo Import."
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
         Left            =   3600
         TabIndex        =   132
         Top             =   5880
         Width           =   2175
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   7695
      Left            =   8520
      TabIndex        =   51
      Top             =   600
      Width           =   4095
      Begin VB.TextBox ipss 
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
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   160
         Top             =   6600
         Width           =   1335
      End
      Begin VB.ComboBox regimen1 
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
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox fechaing 
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   80
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox fechacese 
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   79
         Top             =   3600
         Width           =   1335
      End
      Begin VB.ComboBox moneda 
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
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   4680
         Width           =   1335
      End
      Begin VB.TextBox basico 
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   77
         Top             =   5040
         Width           =   1335
      End
      Begin VB.TextBox jornal 
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   76
         Top             =   5400
         Width           =   1335
      End
      Begin VB.TextBox mtardanza 
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   75
         Top             =   5880
         Width           =   1335
      End
      Begin VB.TextBox dtardanza 
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   74
         Top             =   6240
         Width           =   1335
      End
      Begin VB.TextBox fechavaca 
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   73
         Top             =   4320
         Width           =   1335
      End
      Begin VB.TextBox ini1 
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
         MaxLength       =   10
         TabIndex        =   67
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox fin1 
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
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   66
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox por1 
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
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   65
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox valor1 
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
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   64
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox ini2 
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
         MaxLength       =   10
         TabIndex        =   63
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox fin2 
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
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   62
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox por2 
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
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   61
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox valor2 
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
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   60
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox ini3 
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
         MaxLength       =   10
         TabIndex        =   59
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox fin3 
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
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   58
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox por3 
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
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   57
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox valor3 
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
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   56
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox ini4 
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
         MaxLength       =   10
         TabIndex        =   55
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox fin4 
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
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   54
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox por4 
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
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   53
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox valor4 
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
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   52
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label51 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ipss"
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
         TabIndex        =   161
         Top             =   6600
         Width           =   1335
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Datos Generales"
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
         TabIndex        =   91
         Top             =   2880
         Width           =   3855
      End
      Begin VB.Label regimen 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Regimen"
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
         TabIndex        =   90
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F.Ingreso"
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
         TabIndex        =   89
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F.Cese"
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
         TabIndex        =   88
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label19 
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
         TabIndex        =   87
         Top             =   4680
         Width           =   1335
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Basico"
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
         TabIndex        =   86
         Top             =   5040
         Width           =   1335
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Jornal"
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
         TabIndex        =   85
         Top             =   5400
         Width           =   1335
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tolerancia"
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
         TabIndex        =   84
         Top             =   5880
         Width           =   1335
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DsctoTarda"
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
         TabIndex        =   83
         Top             =   6240
         Width           =   1335
      End
      Begin VB.Label Label66 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaVacac"
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
         TabIndex        =   82
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tabla Comisiones"
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
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ini"
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
         TabIndex        =   71
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Final"
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
         Left            =   1080
         TabIndex        =   70
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "%"
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
         Left            =   2040
         TabIndex        =   69
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor"
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
         Left            =   3000
         TabIndex        =   68
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.TextBox pocket 
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
      Left            =   7440
      MaxLength       =   3
      TabIndex        =   49
      Top             =   6840
      Width           =   735
   End
   Begin VB.TextBox clavere 
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
      Left            =   7440
      MaxLength       =   10
      TabIndex        =   47
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox reloj 
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
      TabIndex        =   45
      Top             =   6480
      Width           =   1575
   End
   Begin VB.TextBox seccion 
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
      Left            =   7440
      MaxLength       =   6
      TabIndex        =   44
      Top             =   5760
      Width           =   975
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
      TabIndex        =   42
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancelar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      MaskColor       =   &H00FFFFFF&
      Picture         =   "tvendedo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   120
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   495
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
      Height          =   375
      Left            =   6360
      MaskColor       =   &H00FFFFFF&
      Picture         =   "tvendedo.frx":07AE
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   120
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   495
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
      Left            =   2280
      MaxLength       =   15
      TabIndex        =   39
      Top             =   6120
      Width           =   3375
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
      Left            =   2280
      MaxLength       =   15
      TabIndex        =   37
      Top             =   5760
      Width           =   3375
   End
   Begin VB.ComboBox civil 
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
      Left            =   5040
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   5280
      Width           =   1335
   End
   Begin VB.ComboBox sexo 
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
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   5280
      Width           =   1335
   End
   Begin VB.ComboBox estado 
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
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   6960
      Width           =   1935
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
      TabIndex        =   11
      Top             =   4800
      Width           =   6135
   End
   Begin VB.TextBox telefono2 
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
      MaxLength       =   15
      TabIndex        =   10
      Top             =   4440
      Width           =   2295
   End
   Begin VB.TextBox telefono1 
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
      Left            =   4200
      MaxLength       =   15
      TabIndex        =   9
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox telefono 
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
      TabIndex        =   8
      Top             =   4440
      Width           =   1695
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
      TabIndex        =   7
      Top             =   3600
      Width           =   3375
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
      TabIndex        =   6
      Top             =   3240
      Width           =   3375
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
      TabIndex        =   5
      Top             =   2880
      Width           =   6135
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
      TabIndex        =   4
      Top             =   2400
      Width           =   6135
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
      TabIndex        =   3
      Top             =   2040
      Width           =   6135
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
      TabIndex        =   2
      Top             =   1680
      Width           =   6135
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
      Top             =   1200
      Width           =   1815
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
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   0
      Top             =   840
      Width           =   1455
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
      Left            =   1440
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tvendedo.frx":0F5C
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Grabar registro"
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
      Picture         =   "tvendedo.frx":216E
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Nuevo registro"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Claves"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tvendedo.frx":3380
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Ayuda"
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
      Picture         =   "tvendedo.frx":4592
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Imprimir"
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
      Left            =   4320
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tvendedo.frx":57A4
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   735
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
      Height          =   615
      Left            =   720
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tvendedo.frx":69B6
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Borrar registro"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tvendedo.frx":7BC8
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Consulta"
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label57 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Terminal Pocket"
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
      Left            =   5760
      TabIndex        =   50
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label Label45 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cod.AccesoReloj"
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
      Left            =   5760
      TabIndex        =   48
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Label Label44 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cod.ProgramacionHoraria"
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
      TabIndex        =   46
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Label Label43 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SeccionProdcc."
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
      Left            =   5760
      TabIndex        =   43
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFF00&
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
      Left            =   120
      TabIndex        =   38
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFF00&
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
      Left            =   120
      TabIndex        =   36
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estado Civil"
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
      TabIndex        =   35
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sexo"
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
      TabIndex        =   33
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFF00&
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
      TabIndex        =   31
      Top             =   6960
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFF00&
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
      TabIndex        =   29
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFF00&
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
      TabIndex        =   28
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Telefonos"
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
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFF00&
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
      TabIndex        =   26
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF00&
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
      TabIndex        =   25
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF00&
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
      TabIndex        =   24
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
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
      TabIndex        =   23
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
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
      TabIndex        =   22
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Apellidos Nombres"
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
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo Alterno"
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
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
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
      TabIndex        =   19
      Top             =   840
      Width           =   2175
   End
   Begin VB.Menu ajdu1 
      Caption         =   "&Add"
   End
   Begin VB.Menu bo712 
      Caption         =   "&Borrar"
   End
   Begin VB.Menu grba1 
      Caption         =   "&Grabar"
   End
   Begin VB.Menu djuer1 
      Caption         =   "&Reporte"
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tvendedo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub ajdu1_Click()
'If Frame3.Visible = True Then Exit Sub
If Frame2.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
inicializa
codigo = ""
codigo.SetFocus

End Sub

Private Sub bo712_Click()
Dim found As Integer
'If Frame3.Visible = True Then Exit Sub
If Frame2.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
found = busca_registro()
If found = 0 Then
   MsgBox "No existe registro", 48, "Aviso"
   Exit Sub
End If
found = borra_registro()
If found = 0 Then Exit Sub
MsgBox "Ok,Registro Borrado", 48, "Aviso"
codigo = ""
inicializa
codigo.SetFocus
End Sub



Private Sub buffer_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   Frame1.Visible = False
   Frame1.Enabled = False
   codigo.SetFocus
   Exit Sub
End If
Command1_Click

End Sub

Private Sub buffer_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> 13 And KeyCode <> 27 Then
   ejecuta 0
End If
End Sub

Private Sub cargo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
profesion.SetFocus

End Sub

Private Sub cargo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   regimen1.SetFocus
   Exit Sub
End If

End Sub

Private Sub civil_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
regimen1.SetFocus

End Sub

Private Sub civil_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   sexo.SetFocus
   Exit Sub
End If

End Sub

Private Sub clave_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
clave = UCase(clave)
End Sub

Private Sub cmdAddEntry_Click()
ajdu1_Click
End Sub

Private Sub cmdDelete_Click()
bo712_Click
End Sub

Private Sub cmdExit_Click()
dlo132_Click

End Sub

Private Sub cmdHelp_Click()
Dim found As Integer
found = busca_registro()
If found = 0 Then
   MsgBox "No existe Codigo", 48, "Aviso"
   codigo.SetFocus
   Exit Sub
End If
If busca_clave1("" & gusuario) <> "S" Then
   MsgBox "No tiene permiso", 48, "Aviso"
   Exit Sub
End If
Frame2.Visible = True
clave.SetFocus
End Sub

Private Sub cmdPrint_Click()
djuer1_Click
End Sub

Private Sub cmdSave_Click()
grba1_Click
End Sub

Private Sub cmdSort_Click()
Dim cad As String
   Dim rconsulta As New ADODB.Recordset
   cad = "SELECT * FROM vendedor  "
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True Or rconsulta.BOF = True Then
      rconsulta.Close
      Exit Sub
   End If
   

Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.ListIndex = 0

consulta_vendedor
End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   dlo132_Click
   Exit Sub
End If
If Len(codigo) = 0 Then Exit Sub
found = busca_registro()
If found = 0 Then
   inicializa
End If
codigo1.SetFocus
End Sub

Private Sub codigo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
cmdSort_Click
End If
End Sub

Private Sub codigo1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
nombre.SetFocus
End Sub

Private Sub codigo1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   codigo.SetFocus
   Exit Sub
End If

End Sub

Private Sub Command1_Click()
ejecuta 1
End Sub
Function ejecuta(sw As Integer)
Dim rconsulta As New ADODB.Recordset
Dim cad As String
If opcion1 = "1" Then  'bodega
   If Len(buffer) = 0 Then
      cad = "select Nombre,Codigo,Direccion,Telefono from vendedor "
   End If
   If Len(buffer) > 0 Then
      cad = "SELECT Nombre,Codigo,Direccion,Telefono from vendedor   where  " & Combo1 & " like '" & buffer & "%'"
   End If
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True And rconsulta.BOF = True Then
      rconsulta.Close
      buffer.SetFocus
      Exit Function
   End If
   Set dbGrid1.DataSource = rconsulta
   dbGrid1.Columns(0).Width = 4000
   dbGrid1.Columns(1).Width = 2000
   dbGrid1.Columns(2).Width = 4000
   dbGrid1.Columns(3).Width = 2000
   If sw = 1 Then
      dbGrid1.SetFocus
   End If
   Exit Function
End If
If opcion1 = "2" Then  'bodega
   If Len(buffer) = 0 Then
      cad = "select Descripcio,Seccion from pseccion "
   End If
   If Len(buffer) > 0 Then
      cad = "SELECT Descripcio,Seccion from pseccion   where  " & Combo1 & " like '" & buffer & "%'"
   End If
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True And rconsulta.BOF = True Then
      rconsulta.Close
      buffer.SetFocus
      Exit Function
   End If
   Set dbGrid1.DataSource = rconsulta
   dbGrid1.Columns(0).Width = 4000
   dbGrid1.Columns(1).Width = 2000
   If sw = 1 Then
      dbGrid1.SetFocus
   End If
   Exit Function
End If
If opcion1 = "22" Then  'bodega
   If Len(buffer) = 0 Then
      cad = "select Descripcio,tprohora from tprohora "
   End If
   If Len(buffer) > 0 Then
      cad = "SELECT Descripcio,tprohora from tprohora   where  " & Combo1 & " like '" & buffer & "%'"
   End If
   
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True And rconsulta.BOF = True Then
      rconsulta.Close
      buffer.SetFocus
      Exit Function
   End If
   Set dbGrid1.DataSource = rconsulta
   dbGrid1.Columns(0).Width = 4000
   dbGrid1.Columns(1).Width = 2000
   If sw = 1 Then
      dbGrid1.SetFocus
   End If
   Exit Function
End If




If opcion1 = "4" Then  'bodega
   If Len(buffer) = 0 Then
      cad = "select Descripcio,tipopla from tipopla "
   End If
   If Len(buffer) > 0 Then
      cad = "SELECT Descripcio,tipopla from tipopla   where  " & Combo1 & " like '" & buffer & "%'"
   End If
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True And rconsulta.BOF = True Then
      rconsulta.Close
      buffer.SetFocus
      Exit Function
   End If
   Set dbGrid1.DataSource = rconsulta
   dbGrid1.Columns(0).Width = 4000
   dbGrid1.Columns(1).Width = 2000
   
   If sw = 1 Then
      dbGrid1.SetFocus
   End If
   Exit Function
End If



End Function

Private Sub Command2_Click()
Dim found As Integer
'found = valida_sisper()
'If found = 0 Then
'   MsgBox "Parametros Invalidos", 48, "Aviso"
'   Exit Sub
'End If
'found = grabar_sisper()
'dlo132_Click
End Sub

Private Sub Command3_Click()
dlo132_Click
End Sub


Private Sub contacto_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
direccion.SetFocus

End Sub

Private Sub contacto_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   nombrec.SetFocus
   Exit Sub
End If

End Sub

Private Sub correo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
sexo.SetFocus

End Sub

Private Sub correo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
  telefono2.SetFocus
   Exit Sub
End If

End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   buffer.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
   If opcion1 = "1" Then
      codigo = dbGrid1.Columns(1)
      Frame1.Visible = False
      Frame1.Enabled = False
      codigo.SetFocus
      codigo_KeyPress 13
   End If
      If opcion1 = "2" Then
      seccion = dbGrid1.Columns(1)
      Frame1.Visible = False
      Frame1.Enabled = False
      seccion.SetFocus
      End If
      If opcion1 = "4" Then
      'tipopla = DBGrid1.Columns(1)
      'Frame1.Visible = False
      'tipopla.SetFocus
      End If
      If opcion1 = "22" Then
      reloj = dbGrid1.Columns(1)
      Frame1.Visible = False
      Frame1.Enabled = False
      reloj.SetFocus
      End If
End If
End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
Dim buf As String
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

Private Sub Dia_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Len(dia) <> 10 Then Exit Sub
If Not IsDate(dia) Then
   dia = ""
   Exit Sub
End If
'limpia_sisper
'found = busca_sisper()
If found = 0 Then
   MsgBox "No hay ninguna Transaccion", 48, "Aviso"
   Exit Sub
End If
'eh1.SetFocus
End Sub

Private Sub direccion_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
dpto.SetFocus

End Sub

Private Sub direccion_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   contacto.SetFocus
   Exit Sub
End If

End Sub

Private Sub distrito_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
zona.SetFocus

End Sub

Private Sub distrito_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   dpto.SetFocus
   Exit Sub
End If

End Sub

Private Sub djuer1_Click()
'If Frame3.Visible = True Then Exit Sub
If Frame2.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
reporgen.NAMETABLA = "vendedor"
reporgen.Show 1

End Sub

Private Sub dlo132_Click()
If Frame2.Visible = True Then
   Frame2.Visible = False
   nombre.SetFocus
   Exit Sub
End If


If opcion1 = "1" Then
If Frame1.Visible = True Then
   Frame1.Visible = False
   codigo.SetFocus
   Exit Sub
End If
End If
If opcion1 = "2" Then
If Frame1.Visible = True Then
   Frame1.Visible = False
   seccion.SetFocus
   Exit Sub
End If
End If
If opcion1 = "22" Then
If Frame1.Visible = True Then
   Frame1.Visible = False
   reloj.SetFocus
   Exit Sub
End If
End If

If opcion1 = "4" Then
'If Frame1.Visible = True Then
'   Frame1.Visible = False
'   tipopla.SetFocus
'   Exit Sub
'End If
End If

tvendedo.Hide
Unload tvendedo
End Sub

Private Sub dpto_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
distrito.SetFocus

End Sub

Private Sub dpto_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   direccion.SetFocus
   Exit Sub
End If

End Sub

Private Sub estado_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   fechacese.SetFocus
   Exit Sub
End If

End Sub

Private Sub fechacese_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
estado.SetFocus

End Sub

Private Sub fechacese_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   fechaing.SetFocus
   Exit Sub
End If

End Sub

Private Sub fechaing_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
fechacese.SetFocus

End Sub

Private Sub fechaing_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   profesion.SetFocus
   Exit Sub
End If

End Sub

Private Sub Form_Load()


regimen1.Clear
regimen1.AddItem "E"
regimen1.AddItem "O"
regimen1.AddItem "I"
regimen1.ListIndex = 0

sexo.Clear
sexo.AddItem "M"
sexo.AddItem "F"
sexo.ListIndex = 0

moneda.Clear
moneda.AddItem "S"
moneda.AddItem "D"
moneda.ListIndex = 0


civil.Clear
civil.AddItem "S"
civil.AddItem "C"
civil.ListIndex = 0


'Combo1.Clear
'Combo1.AddItem "NOMBRE"
'Combo1.AddItem "CODIGO"
'Combo1.ListIndex = 0


estado.Clear
estado.AddItem "ACTIVO"
estado.AddItem "NOACTIVO"
estado.ListIndex = 0
End Sub
Sub inicializa()
'tipopla = ""
ipss = ""
fechavaca = ""
vecostoimp = ""
apertura = ""
anula = ""
copia = ""
congela = ""
cuadre = ""
pocket = ""
'tipoca = ""
'puertoca = ""
reloj = ""
caja = ""
'tipopla = ""
parame = ""
parame1 = ""
parame2 = ""
parame3 = ""
parame4 = ""

seccion = ""
valor1 = ""
valor2 = ""
valor3 = ""
valor4 = ""
por1 = ""
por2 = ""
por3 = ""
por4 = ""
ini1 = ""
ini2 = ""
ini3 = ""
ini4 = ""
fin1 = ""
fin2 = ""
fin3 = ""
fin4 = ""



rw1 = "W"
rw2 = "W"
rw3 = "W"
rw4 = "W"
rw5 = "W"
rw6 = "W"
rw7 = "W"
rw8 = "W"
rw9 = "W"
rw10 = "W"
rw11 = ""
v1 = "S"
v2 = "S"
v3 = "N"
v4 = "N"
v5 = "N"
v6 = "N"
v7 = "N"
v8 = "N"
v9 = "N"
v10 = "S"
v11 = "N"
v12 = "N"
vevend = ""
veclave = ""

clavere = ""
clave = ""
codigo1 = ""
nombre = ""
nombrec = ""
contacto = ""
direccion = ""
dpto = ""
distrito = ""
zona = ""
telefono = ""
telefono1 = ""
telefono2 = ""
correo = ""

regimen1.ListIndex = 0
sexo.ListIndex = 0
civil.ListIndex = 0
cargo = ""
profesion = ""
fechaing = ""
fechacese = ""
moneda.ListIndex = 0
basico = ""
jornal = ""
mtardanza = ""
dtardanza = ""
End Sub
Function borra_registro()
On Error GoTo cmd56_err

If MsgBox("Desea Borrar ", 1, "Aviso") <> 1 Then Exit Function
cn.Execute ("DELETE   FROM vendedor WHERE codigo='" & Trim(codigo) & "'")
borra_registro = 1
Exit Function
cmd56_err:
MsgBox "Aviso en borra " + error$, 48, "Aviso"
Exit Function

End Function
Function busca_registro()
Dim rsexiste As New ADODB.Recordset
   rsexiste.Open "SELECT * FROM vendedor where  codigo='" & Trim(codigo) & "'", cn, adOpenKeyset, adLockOptimistic
   If rsexiste.RecordCount > 0 Then  'si existe
      pone_registro rsexiste
      busca_registro = 1
   End If


End Function
Sub pone_registro(mytablex As ADODB.Recordset)
vecostoimp = "" & mytablex.Fields("vecostoimp")
rw1 = "" & mytablex.Fields("rw1")
rw2 = "" & mytablex.Fields("rw2")
rw3 = "" & mytablex.Fields("rw3")
rw4 = "" & mytablex.Fields("rw4")
rw5 = "" & mytablex.Fields("rw5")
rw6 = "" & mytablex.Fields("rw6")
rw7 = "" & mytablex.Fields("rw7")
rw8 = "" & mytablex.Fields("rw8")
rw9 = "" & mytablex.Fields("rw9")
rw10 = "" & mytablex.Fields("rw10")
rw11 = "" & mytablex.Fields("rw11")
rw12 = "" & mytablex.Fields("rw12")

v1 = "" & mytablex.Fields("v1")
v2 = "" & mytablex.Fields("v2")
v3 = "" & mytablex.Fields("v3")
v4 = "" & mytablex.Fields("v4")
v5 = "" & mytablex.Fields("v5")
v6 = "" & mytablex.Fields("v6")
v7 = "" & mytablex.Fields("v7")
v8 = "" & mytablex.Fields("v8")
v9 = "" & mytablex.Fields("v9")
v10 = "" & mytablex.Fields("v10")
v11 = "" & mytablex.Fields("v11")
v12 = "" & mytablex.Fields("v12")

pocket = "" & mytablex.Fields("pocket")
vevend = "" & mytablex.Fields("vevend")
veclave = "" & mytablex.Fields("veclave")
clavere = "" & mytablex.Fields("claverE")
reloj = "" & mytablex.Fields("reloj")
caja = "" & mytablex.Fields("caja")
cuadre = "" & mytablex.Fields("cuadre")
anula = "" & mytablex.Fields("anula")
congela = "" & mytablex.Fields("congela")
copia = "" & mytablex.Fields("copia")
apertura = "" & mytablex.Fields("apertura")

seccion = "" & mytablex.Fields("seccion")

parame = "" & mytablex.Fields("parame")
parame1 = "" & mytablex.Fields("parame1")
parame2 = "" & mytablex.Fields("parame2")
parame3 = "" & mytablex.Fields("parame3")
parame4 = "" & mytablex.Fields("parame4")

ini1 = "" & mytablex.Fields("ini1")
ini2 = "" & mytablex.Fields("ini2")
ini3 = "" & mytablex.Fields("ini3")
ini4 = "" & mytablex.Fields("ini4")
fin1 = "" & mytablex.Fields("fin1")
fin2 = "" & mytablex.Fields("fin2")
fin3 = "" & mytablex.Fields("fin3")
fin4 = "" & mytablex.Fields("fin4")
por1 = "" & mytablex.Fields("por1")
por2 = "" & mytablex.Fields("por2")
por3 = "" & mytablex.Fields("por3")
por4 = "" & mytablex.Fields("por4")
valor1 = "" & mytablex.Fields("valor1")
valor2 = "" & mytablex.Fields("valor2")
valor3 = "" & mytablex.Fields("valor3")
valor4 = "" & mytablex.Fields("valor4")

clave = "" & mytablex.Fields("clave")
codigo = "" & mytablex.Fields("codigo")
codigo1 = "" & mytablex.Fields("codigo1")
nombre = "" & mytablex.Fields("nombre")
nombrec = "" & mytablex.Fields("nombrec")
contacto = "" & mytablex.Fields("contacto")
direccion = "" & mytablex.Fields("direccion")
dpto = "" & mytablex.Fields("dpto")
distrito = "" & mytablex.Fields("distrito")
zona = "" & mytablex.Fields("zona")
telefono = "" & mytablex.Fields("telefono")
telefono1 = "" & mytablex.Fields("telefono1")
telefono2 = "" & mytablex.Fields("telefono2")
correo = "" & mytablex.Fields("correo")
estado.ListIndex = 0
If "" & mytablex.Fields("estado") = "NOACTIVO" Then
   estado.ListIndex = 1
End If
regimen1.ListIndex = 0
If "" & mytablex.Fields("regimen") = "O" Then
   regimen1.ListIndex = 1
End If
If "" & mytablex.Fields("regimen") = "X" Then
   regimen1.ListIndex = 2
End If

cargo = "" & mytablex.Fields("cargo")
profesion = "" & mytablex.Fields("profesion")

civil.ListIndex = 0
If "" & mytablex.Fields("civil") = "C" Then
   civil.ListIndex = 1
End If
sexo.ListIndex = 0
If "" & mytablex.Fields("sexo") = "F" Then
   sexo.ListIndex = 1
End If
fechaing = "" & mytablex.Fields("fechaingr")
fechacese = "" & mytablex.Fields("fechacese")
moneda.ListIndex = 0
If "" & mytablex.Fields("moneda") = "D" Then
   moneda.ListIndex = 1
End If
basico = "" & mytablex.Fields("basico")
jornal = "" & mytablex.Fields("jornal")
mtardanza = "" & mytablex.Fields("mtardanza")
dtardanza = "" & mytablex.Fields("dtardanza")


End Sub
Sub grabando(sw As Integer)
Dim cad As String
Dim buf As String
If Len(rw1) = 0 Then
   rw1 = "N"
End If
If Len(rw2) = 0 Then
   rw2 = "N"
End If
If Len(rw3) = 0 Then
   rw3 = "N"
End If
If Len(rw4) = 0 Then
   rw4 = "N"
End If
If Len(rw5) = 0 Then
   rw5 = "N"
End If
If Len(rw6) = 0 Then
   rw6 = "N"
End If
If Len(rw7) = 0 Then
   rw7 = "N"
End If
If Len(rw8) = 0 Then
   rw8 = "N"
End If
If Len(rw9) = 0 Then
   rw9 = "N"
End If
If Len(rw10) = 0 Then
   rw10 = "N"
End If
If Len(rw11) = 0 Then
   rw11 = "N"
End If
If Len(rw12) = 0 Then
   rw12 = "N"
End If

If Len(v1) = 0 Then
   v1 = "N"
End If
If Len(v2) = 0 Then
   v2 = "N"
End If
If Len(v3) = 0 Then
   v3 = "N"
End If
If Len(v4) = 0 Then
   v4 = "N"
End If
If Len(v5) = 0 Then
   v5 = "N"
End If
If Len(v6) = 0 Then
   v6 = "N"
End If
If Len(v7) = 0 Then
   v7 = "N"
End If
If Len(v8) = 0 Then
   v8 = "N"
End If
If Len(v9) = 0 Then
   v9 = "N"
End If
If Len(v10) = 0 Then
   v10 = "N"
End If
If Len(v11) = 0 Then
   v11 = "N"
End If
If Len(v12) = 0 Then
   v12 = "N"
End If




If sw = 0 Then
   cad = "INSERT INTO vendedor VALUES('" & Trim(codigo) & "','"
   cad = cad & Trim(codigo1) & "','"
   cad = cad & Trim(nombre) & "','"
   cad = cad & Trim(nombrec) & "','"
   cad = cad & Trim(contacto) & "','"
   cad = cad & Trim(direccion) & "','"
   cad = cad & Trim(dpto) & "','"
   cad = cad & Trim(distrito) & "','"
   cad = cad & Trim(zona) & "','"
   cad = cad & Trim(telefono) & "','"
   cad = cad & Trim(telefono1) & "','"
   cad = cad & Trim(telefono2) & "','"
   cad = cad & Trim(correo) & "','"
   cad = cad & Trim(estado) & "','"
   cad = cad & Trim(regimen) & "','"
   cad = cad & Trim(cargo) & "','"
   cad = cad & Trim(profesion) & "','"
   cad = cad & Trim(civil) & "','"
   cad = cad & Trim(sexo) & "','"
   cad = cad & Trim(fechaing) & "','"
   cad = cad & Trim(fechavaca) & "','"
   cad = cad & Trim(fechacese) & "','"
   cad = cad & Trim(moneda) & "',"
   cad = cad & Val(basico) & ","
   cad = cad & Val(jornal) & ",'"
   cad = cad & Trim(ipss) & "','"
   cad = cad & Trim(mtardanza) & "',"
   cad = cad & Val(dtardanza) & ",'"
   cad = cad & Trim(clave) & "','"
   cad = cad & Trim(seccion) & "',"
   cad = cad & Val(ini1) & ","
   cad = cad & Val(ini2) & ","
   cad = cad & Val(ini3) & ","
   cad = cad & Val(ini4) & ","
   cad = cad & Val(fin1) & ","
   cad = cad & Val(fin2) & ","
   cad = cad & Val(fin3) & ","
   cad = cad & Val(fin4) & ","
   cad = cad & Val(por1) & ","
   cad = cad & Val(por2) & ","
   cad = cad & Val(por3) & ","
   cad = cad & Val(por4) & ","
   cad = cad & Val(valor1) & ","
   cad = cad & Val(valor2) & ","
   cad = cad & Val(valor3) & ","
   cad = cad & Val(valor4) & ",'"
   'cad = cad & Trim(tipopla) & "','"
   cad = cad & Trim(reloj) & "','"
   cad = cad & Trim(clavere) & "','"
   cad = cad & Trim(clave) & "','"
   cad = cad & Trim(vevend) & "','"
   cad = cad & Trim(veclave) & "','"
   cad = cad & Trim(caja) & "','"
   cad = cad & Trim(parame) & "','"
   cad = cad & Trim(parame1) & "','"
   cad = cad & Trim(parame2) & "','"
   cad = cad & Trim(parame3) & "','"
   cad = cad & Trim(parame4) & "','"
   cad = cad & Trim(pocket) & "','"
   cad = cad & Trim(cuadre) & "','"
   cad = cad & Trim(anula) & "','"
   cad = cad & Trim(copia) & "','"
   cad = cad & Trim(congela) & "','"
   cad = cad & Trim(apertura) & "','"
   cad = cad & Trim(v1) & "','"
   cad = cad & Trim(v2) & "','"
   cad = cad & Trim(v3) & "','"
   cad = cad & Trim(v4) & "','"
   cad = cad & Trim(v5) & "','"
   cad = cad & Trim(v6) & "','"
   cad = cad & Trim(v7) & "','"
   cad = cad & Trim(v8) & "','"
   cad = cad & Trim(v9) & "','"
   cad = cad & Trim(v10) & "','"
   cad = cad & Trim(v11) & "','"
   cad = cad & Trim(v12) & "','"
   cad = cad & Trim(rw1) & "','"
   cad = cad & Trim(rw2) & "','"
   cad = cad & Trim(rw3) & "','"
   cad = cad & Trim(rw4) & "','"
   cad = cad & Trim(rw5) & "','"
   cad = cad & Trim(rw6) & "','"
   cad = cad & Trim(rw7) & "','"
   cad = cad & Trim(rw8) & "','"
   cad = cad & Trim(rw9) & "','"
   cad = cad & Trim(rw10) & "','"
   cad = cad & Trim(rw11) & "','"
   cad = cad & Trim(rw12) & "','"
   cad = cad & Trim(vecostoimp) & "')"
   cn.Execute (cad)
   MsgBox "Adicion exitosa", 48, "Aviso"
   codigo.SetFocus
   Exit Sub
End If

If sw = 1 Then
   cad = "UPDATE vendedor SET "
   cad = cad & "codigo1 = '" & Trim(codigo1) & "'"
   cad = cad & ",nombre = '" & Trim(nombre) & "'"
   cad = cad & ",nombrec = '" & Trim(nombrec) & "'"
   cad = cad & ",contacto = '" & Trim(contacto) & "'"
   cad = cad & ",direccion = '" & Trim(direccion) & "'"
   cad = cad & ",dpto = '" & Trim(dpto) & "'"
   cad = cad & ",distrito = '" & Trim(distrito) & "'"
   cad = cad & ",zona = '" & Trim(zona) & "'"
   cad = cad & ",telefono = '" & Trim(telefono) & "'"
   cad = cad & ",telefono1 = '" & Trim(telefono1) & "'"
   cad = cad & ",telefono2 = '" & Trim(telefono2) & "'"
   cad = cad & ",correo = '" & Trim(correo) & "'"
   cad = cad & ",estado = '" & Trim(estado) & "'"
   cad = cad & ",regimen = '" & Trim(regimen) & "'"
   cad = cad & ",cargo = '" & Trim(cargo) & "'"
   cad = cad & ",profesion = '" & Trim(profesion) & "'"
   cad = cad & ",civil = '" & Trim(civil) & "'"
   cad = cad & ",sexo = '" & Trim(sexo) & "'"
   cad = cad & ",fechaingr = '" & Trim(fechaing) & "'"
   cad = cad & ",fechavaca = '" & Trim(fechavaca) & "'"
   cad = cad & ",fechacese = '" & Trim(fechacese) & "'"
   cad = cad & ",moneda = '" & Trim(moneda) & "'"
   cad = cad & ",basico = " & Val(basico) & ""
   cad = cad & ",jornal = " & Val(jornal) & ""
   cad = cad & ",ipss = '" & Trim(ipss) & "'"
   cad = cad & ",mtardanza = '" & Trim(mtardanza) & "'"
   cad = cad & ",dtardanza = " & Val(dtardanza) & ""
   cad = cad & ",clave = '" & Trim(clave) & "'"
   cad = cad & ",seccion = '" & Trim(seccion) & "'"
   cad = cad & ", ini1= " & Val(ini1) & ""
   cad = cad & ", ini2= " & Val(ini2) & ""
   cad = cad & ", ini3= " & Val(ini3) & ""
   cad = cad & ", ini4= " & Val(ini4) & ""
   cad = cad & ", fin1= " & Val(fin1) & ""
   cad = cad & ", fin2= " & Val(fin2) & ""
   cad = cad & ", fin3= " & Val(fin3) & ""
   cad = cad & ", fin4= " & Val(fin4) & ""
   cad = cad & ", por1= " & Val(por1) & ""
   cad = cad & ", por2= " & Val(por2) & ""
   cad = cad & ", por3= " & Val(por3) & ""
   cad = cad & ", por4= " & Val(por4) & ""
   cad = cad & ", valor1= " & Val(valor1) & ""
   cad = cad & ", valor2= " & Val(valor2) & ""
   cad = cad & ", valor3= " & Val(valor3) & ""
   cad = cad & " ,valor4= " & Val(valor4) & ""
   'cad = cad & ",tipopla = '" & Trim(tipopla) & "'"
   cad = cad & ",reloj = '" & Trim(reloj) & "'"
   cad = cad & ",clavere = '" & Trim(clavere) & "'"
   cad = cad & ",vevend = '" & Trim(vevend) & "'"
   cad = cad & ",veclave = '" & Trim(veclave) & "'"
   cad = cad & ",caja = '" & Trim(caja) & "'"
   cad = cad & ",parame = '" & Trim(parame) & "'"
   cad = cad & ",parame1 = '" & Trim(parame1) & "'"
   cad = cad & ",parame2 = '" & Trim(parame2) & "'"
   cad = cad & ",parame3 = '" & Trim(parame3) & "'"
   cad = cad & ",parame4 = '" & Trim(parame4) & "'"
   cad = cad & ",pocket = '" & Trim(pocket) & "'"
   cad = cad & ",cuadre = '" & Trim(cuadre) & "'"
   cad = cad & ",anula = '" & Trim(anula) & "'"
   cad = cad & ",copia = '" & Trim(copia) & "'"
   cad = cad & ",congela = '" & Trim(congela) & "'"
   cad = cad & ",apertura = '" & Trim(apertura) & "'"
   cad = cad & ",v1 = '" & Trim(v1) & "'"
   cad = cad & ",v2 = '" & Trim(v2) & "'"
   cad = cad & ",v3 = '" & Trim(v3) & "'"
   cad = cad & ",v4 = '" & Trim(v4) & "'"
   cad = cad & ",v5 = '" & Trim(v5) & "'"
   cad = cad & ",v6 = '" & Trim(v6) & "'"
   cad = cad & ",v7 = '" & Trim(v7) & "'"
   cad = cad & ",v8 = '" & Trim(v8) & "'"
   cad = cad & ",v9 = '" & Trim(v9) & "'"
   cad = cad & ",v10 = '" & Trim(v10) & "'"
   cad = cad & ",v11 = '" & Trim(v11) & "'"
   cad = cad & ",v12 = '" & Trim(v12) & "'"
   cad = cad & ",rw1 = '" & Trim(rw1) & "'"
   cad = cad & ",rw2 = '" & Trim(rw2) & "'"
   cad = cad & ",rw3 = '" & Trim(rw3) & "'"
   cad = cad & ",rw4 = '" & Trim(rw4) & "'"
   cad = cad & ",rw5 = '" & Trim(rw5) & "'"
   cad = cad & ",rw6 = '" & Trim(rw6) & "'"
   cad = cad & ",rw7 = '" & Trim(rw7) & "'"
   cad = cad & ",rw8 = '" & Trim(rw8) & "'"
   cad = cad & ",rw9 = '" & Trim(rw9) & "'"
   cad = cad & ",rw10 = '" & Trim(rw10) & "'"
   cad = cad & ",rw11 = '" & Trim(rw11) & "'"
   cad = cad & ",rw12 = '" & Trim(rw12) & "'"
   cad = cad & ",vecostoimp = '" & Trim(vecostoimp) & "'"
   cn.Execute (cad)
   MsgBox "Rescripcion exitosa", 48, "Aviso"
   codigo.SetFocus
   Exit Sub
End If


End Sub

Private Sub grba1_Click()
Dim found As Integer
'If Frame3.Visible = True Then Exit Sub
'If Frame2.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
found = grabar()
If found = 0 Then Exit Sub
codigo.SetFocus

End Sub

Private Sub Label1_Click()
cmdSort_Click
End Sub

Private Sub Label24_Click()
clavex = clave
End Sub

Private Sub nombre_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
nombrec.SetFocus

End Sub

Private Sub nombre_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   codigo1.SetFocus
   Exit Sub
End If

End Sub

Private Sub nombrec_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
contacto.SetFocus

End Sub

Private Sub nombrec_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   nombre.SetFocus
   Exit Sub
End If

End Sub

Private Sub profesion_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
fechaing.SetFocus

End Sub

Private Sub profesion_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   cargo.SetFocus
   Exit Sub
End If

End Sub

Private Sub regimen1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
cargo.SetFocus

End Sub

Private Sub regimen1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   civil.SetFocus
   Exit Sub
End If

End Sub

Private Sub reloj_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_reloj
End If

End Sub

Private Sub seccion_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_seccion
End If

End Sub

Private Sub sexo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
civil.SetFocus

End Sub

Private Sub sexo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   correo.SetFocus
   Exit Sub
End If

End Sub

Private Sub telefono_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
telefono1.SetFocus

End Sub

Private Sub telefono_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   zona.SetFocus
   Exit Sub
End If

End Sub

Private Sub telefono1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
telefono2.SetFocus

End Sub

Private Sub telefono1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   telefono.SetFocus
   Exit Sub
End If

End Sub

Private Sub telefono2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
correo.SetFocus

End Sub

Private Sub telefono2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   telefono1.SetFocus
   Exit Sub
End If

End Sub


Private Sub zona_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
telefono.SetFocus

End Sub

Private Sub zona_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   distrito.SetFocus
   Exit Sub
End If

End Sub
Function grabar()
Dim found As Integer
Dim rsexiste As New ADODB.Recordset
found = valida()
If found = 0 Then
   MsgBox "Campos invalidos", 48, "Aviso"
   Exit Function
End If

rsexiste.Open "SELECT * FROM vendedor where  codigo='" & Trim(codigo) & "'", cn, adOpenKeyset, adLockOptimistic
   If rsexiste.RecordCount > 0 Then  'si existe
      If MsgBox("Desea Reescribir? ", 1, "Aviso") <> 1 Then
         codigo.SetFocus
         Exit Function
      End If
      grabando 1
      Exit Function
   End If
   If MsgBox("Desea Adicionar ? ", 1, "Aviso") <> 1 Then
      codigo.SetFocus
      Exit Function
   End If
   grabando 0


End Function

Function valida()
Dim found As Integer
If Len(codigo) = 0 Then
   codigo.SetFocus
   Exit Function
End If
If Len(nombre) = 0 Then
   nombre.SetFocus
   Exit Function
End If
found = validar_clave()
If found = 1 Then
   MsgBox "Clave asignada a otra persona,cambielo...", 48, "Aviso"
   nombre.SetFocus
   Exit Function
End If
valida = 1
End Function
Function validar_clave()
Dim rconsulta As New ADODB.Recordset
Dim cad As String
cad = "select * from vendedor WHERE clave='" & clave & "'"
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True And rconsulta.BOF = True Then
      Exit Function
   End If
   If "" & rconsulta.Fields("codigo") <> codigo Then
      validar_clave = 1
   End If


End Function
Sub consulta_vendedor()
Dim rconsulta As New ADODB.Recordset
Dim cad As String
cad = "select * from vendedor "
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True And rconsulta.BOF = True Then
      Exit Sub
   End If
Frame1.Visible = True
Frame1.Enabled = True
buffer = ""
opcion1 = "1"
ejecuta 1
End Sub
Sub consulta_reloj()
Dim rconsulta As New ADODB.Recordset
Dim cad As String
   cad = "select * from tprohora "
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True And rconsulta.BOF = True Then
      Exit Sub
   End If

Frame1.Visible = True
Frame1.Enabled = True
buffer = ""
opcion1 = "22"
ejecuta 1


End Sub
Sub consulta_seccion()
Dim rconsulta As New ADODB.Recordset
Dim cad As String
cad = "select * from pseccion "
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True And rconsulta.BOF = True Then
      Exit Sub
   End If

Frame1.Enabled = True
Frame1.Visible = True
buffer = ""
opcion1 = "2"
ejecuta 1


End Sub
Sub consulta_planilla()
Dim rconsulta As New ADODB.Recordset
Dim cad As String
cad = "select * from tipopla "
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True And rconsulta.BOF = True Then
      Exit Sub
   End If

Frame1.Visible = True
Frame1.Enabled = True
buffer = ""
opcion1 = "4"
ejecuta 1

End Sub
Function busca_clave1(buf As String) As String
Dim rconsulta As New ADODB.Recordset
Dim cad As String
cad = "select * from vendedor where codigo='" & buf & "'"
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True And rconsulta.BOF = True Then
      Exit Function
   End If
   busca_clave1 = "" & rconsulta.Fields("veclave")
End Function
