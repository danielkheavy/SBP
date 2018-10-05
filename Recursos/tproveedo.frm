VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tproveedo 
   BackColor       =   &H00FFFF00&
   Caption         =   "Tabla de Proveedores"
   ClientHeight    =   9270
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   12315
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Consulta"
      Height          =   8295
      Left            =   8040
      TabIndex        =   151
      Top             =   5640
      Visible         =   0   'False
      Width           =   11895
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
         TabIndex        =   154
         TabStop         =   0   'False
         Top             =   600
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   153
         TabStop         =   0   'False
         Top             =   600
         Width           =   2295
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
         Left            =   8280
         TabIndex        =   152
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   6735
         Left            =   240
         TabIndex        =   155
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
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ranking Productos"
      Height          =   6015
      Left            =   6960
      TabIndex        =   144
      Top             =   4080
      Visible         =   0   'False
      Width           =   11415
      Begin VB.TextBox ranno 
         Height          =   375
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   149
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Compras"
         Height          =   375
         Left            =   3000
         TabIndex        =   147
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Deliveri"
         Height          =   375
         Left            =   3000
         TabIndex        =   146
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
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
         Left            =   4320
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tproveedo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   145
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid dbgrid6 
         Height          =   4455
         Left            =   120
         TabIndex        =   150
         Top             =   1320
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   7858
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
      Begin VB.Label Label43 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Año"
         Height          =   375
         Left            =   120
         TabIndex        =   148
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFF80&
      Caption         =   "Cuentas Corrientes"
      Height          =   7695
      Left            =   8280
      TabIndex        =   132
      Top             =   2880
      Visible         =   0   'False
      Width           =   12255
      Begin MSDataGridLib.DataGrid dbgrid4 
         Height          =   2895
         Left            =   120
         TabIndex        =   133
         Top             =   600
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   5106
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
      Begin MSDataGridLib.DataGrid dbgrid5 
         Height          =   2895
         Left            =   120
         TabIndex        =   134
         Top             =   3960
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   5106
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
      Begin VB.Label Label37 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuentas por Cobrar"
         Height          =   375
         Left            =   120
         TabIndex        =   143
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         Height          =   375
         Left            =   6960
         TabIndex        =   142
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label totalsc 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   8520
         TabIndex        =   141
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label40 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5040
         TabIndex        =   140
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label totaldc 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   10080
         TabIndex        =   139
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label41 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Letras por Cobrar"
         Height          =   375
         Left            =   120
         TabIndex        =   138
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label totaldc1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   9960
         TabIndex        =   137
         Top             =   7080
         Width           =   1575
      End
      Begin VB.Label totalsc1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   8400
         TabIndex        =   136
         Top             =   7080
         Width           =   1575
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         Height          =   375
         Left            =   6840
         TabIndex        =   135
         Top             =   7080
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Datos Economicos"
      Height          =   6975
      Left            =   9840
      TabIndex        =   87
      Top             =   1920
      Visible         =   0   'False
      Width           =   10935
      Begin VB.CommandButton Command6 
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
         Picture         =   "tproveedo.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Deliveri"
         Height          =   375
         Left            =   4200
         TabIndex        =   90
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFF00&
         Caption         =   "Compras"
         Height          =   375
         Left            =   4200
         TabIndex        =   89
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox anno 
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
         MaxLength       =   11
         TabIndex        =   88
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Año(AAAA)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   131
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Enero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   130
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label eneros 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   129
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label enerod 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   128
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Febrero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   127
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label febreros 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   126
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label febrerod 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   125
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Marzo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   124
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label marzos 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   123
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label marzod 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   122
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Abril"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   121
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label abrils 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   120
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label abrild 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   119
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mayo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   118
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label mayos 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   117
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label mayod 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   116
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Junio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   115
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label junios 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   114
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label juniod 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   113
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Julio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   112
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label julios 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   111
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label juliod 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   110
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label51 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Agosto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   109
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label agostos 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   108
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label agostod 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   107
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label54 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Setiembre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   106
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label setiembres 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   105
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label setiembred 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   104
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Octubre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   103
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label octubres 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   102
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label octubred 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   101
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label Label60 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Noviembre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   100
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label noviembres 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   99
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label noviembred 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   98
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label Label63 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Diciembre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Top             =   5160
         Width           =   1695
      End
      Begin VB.Label diciembres 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   96
         Top             =   5160
         Width           =   1695
      End
      Begin VB.Label diciembred 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   95
         Top             =   5160
         Width           =   1695
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   94
         Top             =   5760
         Width           =   1695
      End
      Begin VB.Label totals 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   93
         Top             =   5760
         Width           =   1695
      End
      Begin VB.Label totald 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   92
         Top             =   5760
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Comprobantes Compra"
      Height          =   6975
      Left            =   5760
      TabIndex        =   81
      Top             =   4800
      Visible         =   0   'False
      Width           =   11535
      Begin VB.CommandButton Command8 
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
         Left            =   9960
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tproveedo.frx":0F5C
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "Salir"
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFF80&
         Caption         =   "Compras"
         Height          =   375
         Left            =   6960
         TabIndex        =   84
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFF80&
         Caption         =   "Pedidos"
         Height          =   375
         Left            =   6960
         TabIndex        =   83
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
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
         Height          =   855
         Left            =   8400
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tproveedo.frx":216E
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid dbgrid2 
         Height          =   5415
         Left            =   120
         TabIndex        =   85
         Top             =   1320
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   9551
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
   Begin VB.TextBox observa 
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
      MaxLength       =   30
      TabIndex        =   79
      Top             =   7680
      Width           =   1935
   End
   Begin VB.TextBox fechanac 
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
      TabIndex        =   77
      Top             =   6960
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ranking"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   76
      ToolTipText     =   "Ayuda"
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Grafico"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   75
      ToolTipText     =   "Ayuda"
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Deudas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   74
      ToolTipText     =   "Ayuda"
      Top             =   0
      Width           =   855
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
      TabIndex        =   72
      Top             =   3360
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mensual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   71
      ToolTipText     =   "Ayuda"
      Top             =   0
      Width           =   855
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
      TabIndex        =   67
      Top             =   6600
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
      TabIndex        =   65
      Top             =   6240
      Width           =   1935
   End
   Begin VB.ComboBox clasifica 
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
      TabIndex        =   63
      Top             =   5880
      Width           =   2655
   End
   Begin VB.ComboBox tipoclie 
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
      TabIndex        =   61
      Top             =   1440
      Width           =   2895
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
      TabIndex        =   59
      Top             =   8040
      Width           =   1935
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
      MaxLength       =   10
      TabIndex        =   57
      Top             =   3000
      Width           =   1335
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
      TabIndex        =   55
      Top             =   4200
      Width           =   375
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7080
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   9480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CheckBox domingo 
      BackColor       =   &H00FFFF00&
      Caption         =   "Domingo"
      Height          =   375
      Left            =   7920
      TabIndex        =   53
      Top             =   5280
      Width           =   975
   End
   Begin VB.CheckBox sabado 
      BackColor       =   &H00FFFF00&
      Caption         =   "Sabado"
      Height          =   375
      Left            =   7080
      TabIndex        =   52
      Top             =   5280
      Width           =   855
   End
   Begin VB.CheckBox viernes 
      BackColor       =   &H00FFFF00&
      Caption         =   "Viernes"
      Height          =   375
      Left            =   10680
      TabIndex        =   51
      Top             =   4920
      Width           =   855
   End
   Begin VB.CheckBox jueves 
      BackColor       =   &H00FFFF00&
      Caption         =   "Jueves"
      Height          =   375
      Left            =   9840
      TabIndex        =   50
      Top             =   4920
      Width           =   855
   End
   Begin VB.CheckBox miercoles 
      BackColor       =   &H00FFFF00&
      Caption         =   "Miercoles"
      Height          =   375
      Left            =   8760
      TabIndex        =   49
      Top             =   4920
      Width           =   975
   End
   Begin VB.CheckBox martes 
      BackColor       =   &H00FFFF00&
      Caption         =   "Martes"
      Height          =   375
      Left            =   7920
      TabIndex        =   48
      Top             =   4920
      Width           =   855
   End
   Begin VB.CheckBox lunes 
      BackColor       =   &H00FFFF00&
      Caption         =   "Lunes"
      Height          =   375
      Left            =   7080
      TabIndex        =   47
      Top             =   4920
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
      MaxLength       =   10
      TabIndex        =   45
      Top             =   3840
      Width           =   1215
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
      MaxLength       =   10
      TabIndex        =   43
      Top             =   3480
      Width           =   735
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
      TabIndex        =   41
      Top             =   5520
      Width           =   1935
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
      MaxLength       =   11
      TabIndex        =   40
      Top             =   2640
      Width           =   2055
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
      MaxLength       =   11
      TabIndex        =   37
      Top             =   2280
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
      MaxLength       =   11
      TabIndex        =   35
      Top             =   1920
      Width           =   735
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
      TabIndex        =   33
      Top             =   1560
      Width           =   735
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
      TabIndex        =   31
      Top             =   7320
      Width           =   1935
   End
   Begin VB.ComboBox zona 
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
      TabIndex        =   8
      Top             =   4440
      Width           =   1215
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
      TabIndex        =   12
      Top             =   5160
      Width           =   4695
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
      MaxLength       =   15
      TabIndex        =   11
      Top             =   4800
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
      MaxLength       =   15
      TabIndex        =   10
      Top             =   4800
      Width           =   1575
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
      MaxLength       =   15
      TabIndex        =   9
      Top             =   4800
      Width           =   1455
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
      Top             =   4080
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
      Top             =   3720
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
      Top             =   3000
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
      TabIndex        =   4
      Top             =   2520
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
      TabIndex        =   3
      Top             =   2160
      Width           =   4695
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
      Top             =   1800
      Width           =   4695
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
      Top             =   1080
      Width           =   1935
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
      MaxLength       =   11
      TabIndex        =   0
      Top             =   720
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
      Left            =   1440
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tproveedo.frx":291C
      Style           =   1  'Graphical
      TabIndex        =   19
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
      Picture         =   "tproveedo.frx":3B2E
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Nuevo registro"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Compras"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Ayuda"
      Top             =   0
      Width           =   855
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
      Picture         =   "tproveedo.frx":4D40
      Style           =   1  'Graphical
      TabIndex        =   16
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
      Left            =   7800
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tproveedo.frx":5F52
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   855
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
      Picture         =   "tproveedo.frx":7164
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Borrar registro"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tproveedo.frx":8376
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Consulta"
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label47 
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
      TabIndex        =   80
      Top             =   7680
      Width           =   2175
   End
   Begin VB.Label Label44 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Alta"
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
      TabIndex        =   78
      Top             =   8040
      Width           =   2175
   End
   Begin VB.Label Label35 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Referencias"
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
      TabIndex        =   73
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label nreferencia 
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
      Left            =   4200
      TabIndex        =   70
      Top             =   6240
      Width           =   4695
   End
   Begin VB.Label ngarantia 
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
      Left            =   4200
      TabIndex        =   69
      Top             =   6600
      Width           =   4695
   End
   Begin VB.Label Label25 
      BackColor       =   &H00FFFF00&
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
      TabIndex        =   68
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Label Label24 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Referencias"
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
      TabIndex        =   66
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label clasddd 
      BackColor       =   &H00FFFF00&
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
      Left            =   120
      TabIndex        =   64
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFF00&
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
      TabIndex        =   62
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFFF00&
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
      TabIndex        =   60
      Top             =   6960
      Width           =   2175
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFFF00&
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
      TabIndex        =   58
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label29 
      BackColor       =   &H00FFFF00&
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
      TabIndex        =   56
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFF00&
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
      TabIndex        =   54
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFF00&
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
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFF00&
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
      TabIndex        =   44
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFF00&
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
      TabIndex        =   42
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFF00&
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
      TabIndex        =   39
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFF00&
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
      TabIndex        =   38
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFF00&
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
      TabIndex        =   36
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFF00&
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
      TabIndex        =   34
      Top             =   1560
      Width           =   1695
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
      TabIndex        =   32
      Top             =   7320
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
      TabIndex        =   30
      Top             =   4440
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
      TabIndex        =   29
      Top             =   5160
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
      TabIndex        =   28
      Top             =   4800
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
      TabIndex        =   27
      Top             =   4080
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
      TabIndex        =   26
      Top             =   3720
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
      TabIndex        =   25
      Top             =   3000
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
      TabIndex        =   24
      Top             =   2520
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
      TabIndex        =   23
      Top             =   2160
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
      TabIndex        =   22
      Top             =   1800
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
      TabIndex        =   21
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
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
      TabIndex        =   20
      Top             =   720
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
      Begin VB.Menu dk23ge 
         Caption         =   "&1.Generico"
      End
      Begin VB.Menu dk8922 
         Caption         =   "&2.generador"
      End
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tproveedo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub ajdu1_Click()
Dim found As Integer
'If Frame6.Visible = True Then Exit Sub

If Frame5.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
If Frame4.Visible = True Then Exit Sub
If Frame2.Visible = True Then Exit Sub
If Frame3.Visible = True Then Exit Sub
inicializa
codigo = ""
'found = busca_parame(0)
codigo.SetFocus
End Sub

Private Sub anno_Click()
sumar_mensual
End Sub

Private Sub bo712_Click()
Dim found As Integer
'If Frame6.Visible = True Then Exit Sub

If Frame5.Visible = True Then Exit Sub
If Frame4.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
If Frame2.Visible = True Then Exit Sub
If Frame3.Visible = True Then Exit Sub
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

Private Sub cmdAddEntry_Click()
ajdu1_Click
End Sub

Private Sub cmdDelete_Click()
bo712_Click
End Sub

Private Sub cmdExit_Click()
dlo132_Click

End Sub

Private Sub cmdGrabar_Click()
Label27_Click
End Sub

Private Sub cmdHelp_Click()
Dim found As Integer
Dim buf As String
Dim rconsulta As New ADODB.Recordset
found = busca_registro()
If found = 0 Then
   MsgBox "No ha seleccionado un cliente ", 48, "Aviso"
   Exit Sub
End If
buf = "select Local,Tipo,Serie,Numero,Fecha,Codigo,Nombre,Moneda as M,Total,Estado as E from factura where (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F' )   "
buf = buf & " and codigo='" & codigo & "'"
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open buf, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True And rconsulta.BOF = True Then
      Exit Sub
   End If
   Set dbgrid2.DataSource = rconsulta
   Frame2.Visible = True
   dbgrid2.SetFocus
End Sub

Private Sub cmdPrint_Click()
'djuer1_Click
End Sub

Private Sub cmdSave_Click()
grba1_Click
End Sub

Private Sub cmdSort_Click()
Dim cad As String
   Dim rconsulta As New ADODB.Recordset
   cad = "SELECT * FROM proveedo  "
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True Or rconsulta.BOF = True Then
      rconsulta.Close
      Exit Sub
   End If

Combo1.Clear
Combo1.AddItem "NOMBRE"
Combo1.AddItem "CODIGO"
Combo1.ListIndex = 0
opcion1 = "1"
Frame1.Visible = True
Frame1.Enabled = True
buffer = ""
buffer.SetFocus
Command1_Click
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
Sub ejecuta(sw As Integer)
Dim rconsulta As New ADODB.Recordset
Dim buf As String
If opcion1 = "4" Then
      If Len(buffer) = 0 Then
      buf = "select Descripcio,Tipo from Tipo where tipodoc='R'"
      Else
      buf = "select Descripcio,Tipo from Tipo where tipodoc='R' and " & Combo1 & " like '" & buffer & "%'"
      End If
   End If
   

If opcion1 = "1" Or opcion1 = "5" Or opcion1 = "6" Then
   If Len(buffer) = 0 Then
      buf = "select Nombre,Codigo,Direccion,Telefono from proveedo "
      Else
      buf = "select Nombre,Codigo,Direccion,Telefono from proveedo where " & Combo1 & " like '" & buffer & "%'"
   End If
   
End If
If opcion1 = "2" Then
   If Len(buffer) = 0 Then
      buf = "select Descripcio,fpago from fpago "
      Else
      buf = "select Descripcio,fpago from fpago where " & Combo1 & " like '" & buffer & "%'"
   End If
   
End If
If opcion1 = "3" Then
   If Len(buffer) = 0 Then
      buf = "select Nombre,Codigo from vendedor "
      Else
      buf = "select Nombre,Codigo from Vendedor where " & Combo1 & " like '" & buffer & "%'"
   End If
   
End If
If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open buf, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True And rconsulta.BOF = True Then
      rconsulta.Close
      buffer.SetFocus
      Exit Sub
   End If
   Set dbGrid1.DataSource = rconsulta
   
   
               If opcion1 = "1" Then
               dbGrid1.Columns(0).Width = 4000
               dbGrid1.Columns(1).Width = 2000
               dbGrid1.Columns(2).Width = 4000
               dbGrid1.Columns(3).Width = 2000
               End If
               If opcion1 = "2" Or opcion1 = "3" Or opcion1 = "4" Then
               dbGrid1.Columns(0).Width = 4000
               dbGrid1.Columns(1).Width = 2000
               End If
               If sw = 1 Then
               dbGrid1.SetFocus
               End If

End Sub


Private Sub Command2_Click()
Dim found As Integer
Dim buf As String
found = busca_registro()
If found = 0 Then
   MsgBox "No ha seleccionado un cliente ", 48, "Aviso"
   Exit Sub
End If
anno = Format(Year(Now), "0000")
Frame3.Visible = True
sumar_mensual
anno.SetFocus


End Sub

Private Sub Command3_Click()
Dim found As Integer
Dim buf As String
Dim mytablex As New ADODB.Recordset
totalsc = ""
totaldc = ""
found = busca_registro()
If found = 0 Then
   MsgBox "No ha seleccionado un cliente ", 48, "Aviso"
   Exit Sub
End If
      mytablex.Open "SELECT * FROM cuentac where  codigo='" & Trim(codigo) & "'", cn, adOpenKeyset, adLockOptimistic
      If mytablex.RecordCount > 0 Then  'si existe
         Set dbgrid4.DataSource = mytablex
         suma_cuentac mytablex
      End If
      mytablex.Close
      mytablex.Open "SELECT * FROM letrav where  aceptante='" & Trim(codigo) & "'", cn, adOpenKeyset, adLockOptimistic
      If mytablex.RecordCount > 0 Then  'si existe
         Set dbgrid5.DataSource = mytablex
         suma_letras mytablex
      End If
      mytablex.Close
      Frame4.Visible = True
End Sub

Private Sub Command4_Click()
Dim found As Integer
found = busca_registro()
If found = 0 Then
   MsgBox "No existe registro", 48, "Aviso"
   Exit Sub
End If
FrmChart.codigo = codigo
'FrmChart.deliveri.Visible = True
'FrmChart.compras.Visible = False
'FrmChart.ventas.Visible = True
'FrmChart.ventas.Value = True
FrmChart.acu = "V"
FrmChart.docu = "2"
FrmChart.Show 1
End Sub

Private Sub Command5_Click()
Dim found As Integer
Dim i As Integer
found = busca_registro()
If found = 0 Then
   MsgBox "No existe registro", 48, "Aviso"
   Exit Sub
End If
ranno = Format(Year(Now), "0000")
Frame5.Visible = True
sql_ranking
ranno.SetFocus

End Sub

Private Sub Command6_Click()
sumar_mensual
End Sub

Private Sub Command7_Click()
sql_ranking
End Sub

Private Sub Command8_Click()
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
estado.SetFocus

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
   fpago = dbGrid1.Columns(1)
   Frame1.Visible = False
    Frame1.Enabled = False
   fpago.SetFocus
   End If
   If opcion1 = "3" Then
   vendedor = dbGrid1.Columns(1)
   Frame1.Visible = False
    Frame1.Enabled = False
   vendedor.SetFocus
   End If
   If opcion1 = "5" Then
   referencia = dbGrid1.Columns(1)
   Frame1.Visible = False
    Frame1.Enabled = False
   referencia.SetFocus
   End If
   If opcion1 = "6" Then
   garantia = dbGrid1.Columns(1)
   Frame1.Visible = False
    Frame1.Enabled = False
   garantia.SetFocus
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

Private Sub dj88232_Click()
End Sub

Private Sub dlo132_Click()
'If DBGrid3.Visible = True Then
'   DBGrid3.Visible = False
'   Exit Sub
'End If

If Frame5.Visible = True Then
   Frame5.Visible = False
   codigo.SetFocus
   Exit Sub
End If

If Frame4.Visible = True Then
   Frame4.Visible = False
   codigo.SetFocus
   Exit Sub
End If


If Frame2.Visible = True Then
   Frame2.Visible = False
   codigo.SetFocus
   Exit Sub
End If
If Frame3.Visible = True Then
   Frame3.Visible = False
   codigo.SetFocus
   Exit Sub
End If


If opcion1 = "1" Then
If Frame1.Visible = True Then
   Frame1.Visible = False
   Frame1.Enabled = False
   codigo.SetFocus
   Exit Sub
End If
End If
If opcion1 = "2" Then
If Frame1.Visible = True Then
   Frame1.Visible = False
   Frame1.Enabled = False
   fpago.SetFocus
   Exit Sub
End If
End If
If opcion1 = "3" Then
If Frame1.Visible = True Then
   Frame1.Visible = False
   Frame1.Enabled = False
   vendedor.SetFocus
   Exit Sub
End If
End If
If amsw = 1 Then
   tdeliver.dcodigo = codigo
End If

tproveedo.Hide
Unload tproveedo
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
   correo.SetFocus
   Exit Sub
End If

End Sub

Private Sub Form_Activate()
Dim found As Integer
Dim mytablex As New ADODB.Recordset
Dim atelefono As String
Dim sdx As Double
tipoclie.Clear
clasifica.Clear
zona.Clear
tipoclie.AddItem ""
clasifica.AddItem ""
zona.AddItem ""
   If mytablex.State = 1 Then mytablex.Close
   mytablex.Open "select * from tipoclie", cn, adOpenStatic, adLockOptimistic
Do
If mytablex.EOF Then Exit Do
tipoclie.AddItem "" & mytablex.Fields("tipoclie")
mytablex.MoveNext
Loop
mytablex.Close
tipoclie.ListIndex = 0
If mytablex.State = 1 Then mytablex.Close
   mytablex.Open "select * from clasifi", cn, adOpenStatic, adLockOptimistic
Do
If mytablex.EOF Then Exit Do
clasifica.AddItem "" & mytablex.Fields("clasifica")
mytablex.MoveNext
Loop
clasifica.ListIndex = 0

If mytablex.State = 1 Then mytablex.Close
   mytablex.Open "select * from zona", cn, adOpenStatic, adLockOptimistic
Do
If mytablex.EOF Then Exit Do
zona.AddItem "" & mytablex.Fields("zona")
mytablex.MoveNext
Loop
zona.ListIndex = 0

If amsw = 1 Then
amj1:
   found = busca_registro()
   If found = 1 Then
      'si ya existe registro buscar el sigueinte
      sdx = Val(codigo) + 1
      codigo = "" & sdx
      GoTo amj1
   End If
   atelefono = telefono
   inicializa
   telefono = atelefono
   nombre.SetFocus
End If

End Sub

Private Sub Form_Load()
Dim sdx As Double
Dim i As Integer

Combo1.Clear
Combo1.AddItem "NOMBRE"
Combo1.AddItem "CODIGO"
Combo1.ListIndex = 0



estado.Clear
estado.AddItem "ACTIVO"
estado.AddItem "NOACTIVO"
estado.ListIndex = 0


End Sub
Sub inicializa()

observa = ""
referencias = ""
flete = ""
ngarantia = ""
nreferencia = ""
garantia = ""
referencia = ""

fechalta = Format(Now, "dd/mm/yyyy")
'fechalta = ""
lunes.Value = False
martes.Value = False
miercoles.Value = False
jueves.Value = False
viernes.Value = False
sabado.Value = False
domingo.Value = False
moneda = "S"
credito = ""
descuento1 = ""
vendedor = ""
descuento = ""
diapago = ""
fpago = ""
cuenta = ""
codigo1 = ""
nombre = ""
nombrec = ""
contacto = ""
direccion = ""
dpto = ""
distrito = ""
clasifica.ListIndex = 0
tipoclie.ListIndex = 0
zona.ListIndex = 0
telefono = ""
telefono1 = ""
telefono2 = ""
correo = ""

End Sub
Function borra_registro()

On Error GoTo cmd56_err

cn.Execute ("DELETE   FROM proveedo WHERE codigo='" & Trim(codigo) & "'")
borra_registro = 1
Exit Function
cmd56_err:
MsgBox "Aviso en borra " + error$, 48, "Aviso"
Exit Function

End Function
Function busca_registro()
Dim rsexiste As New ADODB.Recordset
   rsexiste.Open "SELECT * FROM proveedo where  codigo='" & Trim(codigo) & "'", cn, adOpenKeyset, adLockOptimistic
   If rsexiste.RecordCount > 0 Then  'si existe
      pone_registro rsexiste
      busca_registro = 1
   End If


 
End Function
Function busca_referencia(buf As String, sw As Integer)
Dim rsexiste As New ADODB.Recordset
   rsexiste.Open "SELECT * FROM proveedo where  codigo='" & Trim(buf) & "'", cn, adOpenKeyset, adLockOptimistic
   If rsexiste.RecordCount > 0 Then  'si existe
   If sw = 0 Then
      nreferencia = "" & rsexiste.Fields("nombre")
   End If
   If sw = 1 Then
      ngarantia = "" & rsexiste.Fields("nombre")
   End If
   End If

 
End Function
Sub pone_registro(mytablex As ADODB.Recordset)
Dim found As Integer
found = busca_combox(clasifica, "" & mytablex.Fields("clasifica"))
clasifica.ListIndex = found
found = busca_combox(tipoclie, "" & mytablex.Fields("tipoclie"))
tipoclie.ListIndex = found
found = busca_combox(zona, "" & mytablex.Fields("zona"))
zona.ListIndex = found
lunes.Value = Val("" & mytablex.Fields("lunes"))
martes.Value = Val("" & mytablex.Fields("martes"))
miercoles.Value = Val("" & mytablex.Fields("miercoles"))
jueves.Value = Val("" & mytablex.Fields("jueves"))
viernes.Value = Val("" & mytablex.Fields("viernes"))
sabado.Value = Val("" & mytablex.Fields("sabado"))
domingo.Value = Val("" & mytablex.Fields("domingo"))
found = busca_combox(tipoclie, "" & mytablex.Fields("tipoclie"))
fechalta = "" & mytablex.Fields("fechanac")
referencias = "" & mytablex.Fields("observa")
referencia = "" & mytablex.Fields("referencia")
garantia = "" & mytablex.Fields("garantia")
flete = "" & mytablex.Fields("flete")
moneda = "" & mytablex.Fields("moneda")
descuento1 = "" & mytablex.Fields("descuento1")
credito = "" & mytablex.Fields("credito")
vendedor = "" & mytablex.Fields("vendedor")

descuento = "" & mytablex.Fields("descuento")
diapago = "" & mytablex.Fields("diapago")
fpago = "" & mytablex.Fields("fpago")
cuenta = "" & mytablex.Fields("cuenta")

codigo = "" & mytablex.Fields("codigo")
codigo1 = "" & mytablex.Fields("codigo1")
nombre = "" & mytablex.Fields("nombre")
nombrec = "" & mytablex.Fields("nombrec")
contacto = "" & mytablex.Fields("contacto")
direccion = "" & mytablex.Fields("direccion")
dpto = "" & mytablex.Fields("dpto")
distrito = "" & mytablex.Fields("distrito")
telefono = "" & mytablex.Fields("telefono")
telefono1 = "" & mytablex.Fields("telefono1")
telefono2 = "" & mytablex.Fields("telefono2")
correo = "" & mytablex.Fields("correo")
estado.ListIndex = 0
If "" & mytablex.Fields("estado") = "NOACTIVO" Then
   estado.ListIndex = 1
End If
found = busca_referencia("" & referencia, 0)
found = busca_referencia("" & garantia, 0)
End Sub
Sub grabando(sw As Integer)

Dim cad As String

If sw = 0 Then
   cad = "INSERT INTO proveedo VALUES('" & Trim(codigo) & "','"
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
   cad = cad & Trim(estado) & "',"
   cad = cad & Val(descuento) & ",'"
   cad = cad & Trim(diapago) & "','"
   cad = cad & Trim(fpago) & "','"
   cad = cad & Trim(cuenta) & "','"
   cad = cad & Trim(descuento1) & "','"
   cad = cad & Trim(vendedor) & "',"
   cad = cad & Val(credito) & ",'"
   cad = cad & Trim(lunes) & "','"
   cad = cad & Trim(martes) & "','"
   cad = cad & Trim(miercoles) & "','"
   cad = cad & Trim(jueves) & "','"
   cad = cad & Trim(viernes) & "','"
   cad = cad & Trim(sabado) & "','"
   cad = cad & Trim(domingo) & "','"
   cad = cad & Trim(moneda) & "',"
   cad = cad & Val(flete) & ",'"
   cad = cad & Trim(fechalta) & "','"
   cad = cad & Trim(clasifica) & "','"
   cad = cad & Trim(tipoclie) & "','"
   cad = cad & Trim(referencia) & "','"
   cad = cad & Trim(garantia) & "','"
   cad = cad & Trim(fechanac) & "','"
   cad = cad & Trim(observa) & "')"
   cn.Execute (cad)
   MsgBox "Adicion exitosa", 48, "Aviso"
   codigo.SetFocus
   Exit Sub
End If

If sw = 1 Then
   cad = "UPDATE proveedo SET "
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
   cad = cad & ",descuento = " & Val(descuento) & ""
   cad = cad & ",diapago = '" & Trim(diapago) & "'"
   cad = cad & ",fpago = '" & Trim(fpago) & "'"
   cad = cad & ",cuenta = '" & Trim(cuenta) & "'"
   cad = cad & ",descuento1 = '" & Trim(descuento1) & "'"
   cad = cad & ",vendedor = '" & Trim(vendedor) & "'"
   cad = cad & ",credito = " & Val(credito) & ""
   cad = cad & ",lunes = '" & Trim(lunes) & "'"
   cad = cad & ",martes = '" & Trim(martes) & "'"
   cad = cad & ",miercoles = '" & Trim(miercoles) & "'"
   cad = cad & ",jueves = '" & Trim(jueves) & "'"
   cad = cad & ",viernes = '" & Trim(viernes) & "'"
   cad = cad & ",sabado = '" & Trim(sabado) & "'"
   cad = cad & ",domingo = '" & Trim(domingo) & "'"
   cad = cad & ",moneda = '" & Trim(moneda) & "'"
   cad = cad & ",flete = " & Val(codigo1) & ""
   cad = cad & ",fechalta = '" & Trim(fechalta) & "'"
   cad = cad & ",clasifica = '" & Trim(clasifica) & "'"
   cad = cad & ",tipoclie = '" & Trim(tipoclie) & "'"
   cad = cad & ",referencia = '" & Trim(referencia) & "'"
   cad = cad & ",garantia = '" & Trim(garantia) & "'"
   cad = cad & ",fechanac = '" & Trim(fechanac) & "'"
   cad = cad & ",observa = '" & Trim(observa) & "'"
   cad = cad & " WHERE  codigo='" & Trim(codigo) & "'"
   cn.Execute (cad)
   MsgBox "Rescripcion exitosa", 48, "Aviso"
   codigo.SetFocus
   Exit Sub
End If



End Sub

Private Sub Form_Unload(Cancel As Integer)
If amsw = 1 Then
   tdeliver.codigo = codigo
End If

End Sub

Private Sub fpago_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Exit Sub
End If
KeyAscii = 0
End Sub

Private Sub fpago_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
consulta_fpago
End If

End Sub

Private Sub garantia_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
consulta_garantia
End If

End Sub

Private Sub grba1_Click()
Dim found As Integer


If Frame5.Visible = True Then Exit Sub
If Frame4.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
If Frame2.Visible = True Then Exit Sub
If Frame3.Visible = True Then Exit Sub
found = grabar()
If found = 0 Then Exit Sub
codigo.SetFocus

End Sub

Private Sub Label1_Click()
cmdSort_Click
End Sub

Private Sub Label27_Click()
Dim buf As String
Dim buf1 As String
Dim rconsulta As New ADODB.Recordset

buf1 = "factura"
If Option2.Value = -1 Then
   buf1 = "cpedidov"
End If

buf = "select * from " & buf1 & "  where "
buf = buf & " codigo='" & codigo & "'"
If buf1 = "factura" Then
   buf = buf & " and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F')"
End If
If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open buf, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True And rconsulta.BOF = True Then
      Exit Sub
   End If
   Set dbgrid2.DataSource = rconsulta
   Frame2.Visible = True
   dbgrid2.SetFocus
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

Private Sub referencia_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
consulta_Referencia
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

Private Sub vendedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Exit Sub
End If
KeyAscii = 0

End Sub

Private Sub vendedor_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
consulta_vendedor
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
Dim cad As String
Dim rsexiste As New ADODB.Recordset
found = valida()
If found = 0 Then
   MsgBox "Campos invalidos", 48, "Aviso"
   Exit Function
End If

   rsexiste.Open "SELECT * FROM proveedo where  codigo='" & Trim(codigo) & "'", cn, adOpenKeyset, adLockOptimistic
   If rsexiste.RecordCount > 0 Then  'si existe
      If MsgBox("Desea Reescribir? ", 1, "Aviso") <> 1 Then
         codigo.SetFocus
         Exit Function
      End If
      grabando 1
       grabar = 1
      Exit Function
   End If
   If MsgBox("Desea Adicionar ? ", 1, "Aviso") <> 1 Then
      codigo.SetFocus
      Exit Function
   End If
   grabando 0
   found = busca_parame(1)
   grabar = 1
   
End Function

Function valida()
If Len(codigo) = 0 Then
   codigo.SetFocus
   Exit Function
End If
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
valida = 1
End Function
Sub consulta_fpago()
Combo1.Clear
Combo1.AddItem "DESCRIPCIO"
Combo1.AddItem "FPAGO"
Combo1.ListIndex = 0
opcion1 = "2"
Frame1.Visible = True
Frame1.Enabled = True
buffer = ""
buffer.SetFocus
Command1_Click

End Sub
Sub consulta_vendedor()
Combo1.Clear
Combo1.AddItem "NOMBRE"
Combo1.AddItem "CODIGO"
Combo1.ListIndex = 0
opcion1 = "3"
Frame1.Visible = True
Frame1.Enabled = True
buffer = ""
buffer.SetFocus
Command1_Click

End Sub
Sub consulta_Referencia()
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.AddItem "Codigo"
Combo1.ListIndex = 0
opcion1 = "5"
Frame1.Visible = True
Frame1.Enabled = True
buffer = ""
buffer.SetFocus
Command1_Click

End Sub
Sub consulta_garantia()
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.AddItem "Codigo"
Combo1.ListIndex = 0
opcion1 = "6"
Frame1.Visible = True
Frame1.Enabled = True
buffer = ""
buffer.SetFocus
Command1_Click

End Sub
Sub suma_total()
End Sub
Sub ir_inicio1()
End Sub
Sub suma_cuentac(mytablex As ADODB.Recordset)
Dim sdx As Double
Dim sdx1 As Double
sdx = 0
sdx1 = 0

Do
If mytablex.EOF Then Exit Do
If "" & mytablex.Fields("moneda") = "S" Then
sdx = sdx + Val("" & mytablex.Fields("saldo"))
End If
If "" & mytablex.Fields("moneda") = "D" Then
sdx1 = sdx1 + Val("" & mytablex.Fields("saldo"))
End If
mytablex.MoveNext
Loop
totalsc = Format(sdx, "0.00")
totaldc = Format(sdx1, "0.00")
End Sub
Sub suma_letras(mytablex As ADODB.Recordset)
Dim sdx As Double
Dim sdx1 As Double
sdx = 0
sdx1 = 0

Do
If mytablex.EOF Then Exit Do
If "" & mytablex.Fields("moneda") = "S" Then
sdx = sdx + Val("" & mytablex.Fields("saldo"))
End If
If "" & mytablex.Fields("moneda") = "D" Then
sdx1 = sdx1 + Val("" & mytablex.Fields("saldo"))
End If
mytablex.MoveNext
Loop
totalsc1 = Format(sdx, "0.00")
totaldc1 = Format(sdx1, "0.00")
End Sub
Sub sumar_mensual()
Dim buf As String
Dim buf1 As String
Dim mytablex As New ADODB.Recordset
Dim sdx As Double

buf1 = "factura"
If Option3.Value = -1 Then
   buf1 = "cpedidov"
End If
inicializa_x
buf = "select month(fecha) as mes,moneda,sum(total) as xtotal  from " & buf1 & " where "
buf = buf & " codigo='" & codigo & "'"
If buf1 = "factura" Then
   buf = buf & " and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F')"
End If
buf = buf & " and year(fecha)=" & anno
buf = buf & "  group by month(fecha),moneda "

   If mytablex.State = 1 Then mytablex.Close
   mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
   If mytablex.EOF = True And mytablex.BOF = True Then
      Exit Sub
   End If
   
sdx = 0
Do
If mytablex.EOF Then Exit Do
If Val("" & mytablex.Fields("mes")) = 1 Then
   If "" & mytablex.Fields("moneda") = "S" Then
      sdx = Val(eneros) + Val("" & mytablex.Fields("xtotal"))
      eneros = "" & sdx
   End If
   If "" & mytablex.Fields("moneda") = "D" Then
      sdx = Val(enerod) + Val("" & mytablex.Fields("xtotal"))
      enerod = "" & sdx
   End If
End If
If Val("" & mytablex.Fields("mes")) = 2 Then
   If "" & mytablex.Fields("moneda") = "S" Then
      sdx = Val(febreros) + Val("" & mytablex.Fields("xtotal"))
      febreros = "" & sdx
   End If
   If "" & mytablex.Fields("moneda") = "D" Then
      sdx = Val(febrerod) + Val("" & mytablex.Fields("xtotal"))
      febrerod = "" & sdx
   End If
End If
If Val("" & mytablex.Fields("mes")) = 3 Then
   If "" & mytablex.Fields("moneda") = "S" Then
      sdx = Val(marzos) + Val("" & mytablex.Fields("xtotal"))
      marzos = "" & sdx
   End If
   If "" & mytablex.Fields("moneda") = "D" Then
      sdx = Val(marzod) + Val("" & mytablex.Fields("xtotal"))
      marzod = "" & sdx
   End If
End If
If Val("" & mytablex.Fields("mes")) = 4 Then
   If "" & mytablex.Fields("moneda") = "S" Then
      sdx = Val(abrils) + Val("" & mytablex.Fields("xtotal"))
      abrils = "" & sdx
   End If
   If "" & mytablex.Fields("moneda") = "D" Then
      sdx = Val(abrild) + Val("" & mytablex.Fields("xtotal"))
      abrild = "" & sdx
   End If
End If
If Val("" & mytablex.Fields("mes")) = 5 Then
   If "" & mytablex.Fields("moneda") = "S" Then
      sdx = Val(mayos) + Val("" & mytablex.Fields("xtotal"))
      mayos = "" & sdx
   End If
   If "" & mytablex.Fields("moneda") = "D" Then
      sdx = Val(mayod) + Val("" & mytablex.Fields("xtotal"))
      mayod = "" & sdx
   End If
End If

If Val("" & mytablex.Fields("mes")) = 6 Then
   If "" & mytablex.Fields("moneda") = "S" Then
      sdx = Val(junios) + Val("" & mytablex.Fields("xtotal"))
      junios = "" & sdx
   End If
   If "" & mytablex.Fields("moneda") = "D" Then
      sdx = Val(juniod) + Val("" & mytablex.Fields("xtotal"))
      juniod = "" & sdx
   End If
End If
If Val("" & mytablex.Fields("mes")) = 7 Then
   If "" & mytablex.Fields("moneda") = "S" Then
      sdx = Val(julios) + Val("" & mytablex.Fields("xtotal"))
      julios = "" & sdx
   End If
   If "" & mytablex.Fields("moneda") = "D" Then
      sdx = Val(juliod) + Val("" & mytablex.Fields("xtotal"))
      juliod = "" & sdx
   End If
End If

If Val("" & mytablex.Fields("mes")) = 8 Then
   If "" & mytablex.Fields("moneda") = "S" Then
      sdx = Val(agostos) + Val("" & mytablex.Fields("xtotal"))
      agostos = "" & sdx
   End If
   If "" & mytablex.Fields("moneda") = "D" Then
      sdx = Val(agostod) + Val("" & mytablex.Fields("xtotal"))
      agostod = "" & sdx
   End If
End If
If Val("" & mytablex.Fields("mes")) = 9 Then
   If "" & mytablex.Fields("moneda") = "S" Then
      sdx = Val(setiembres) + Val("" & mytablex.Fields("xtotal"))
      setiembres = "" & sdx
   End If
   If "" & mytablex.Fields("moneda") = "D" Then
      sdx = Val(setiembred) + Val("" & mytablex.Fields("xtotal"))
      setiembred = "" & sdx
   End If
End If
If Val("" & mytablex.Fields("mes")) = 10 Then
   If "" & mytablex.Fields("moneda") = "S" Then
      sdx = Val(octubres) + Val("" & mytablex.Fields("xtotal"))
      octubres = "" & sdx
   End If
   If "" & mytablex.Fields("moneda") = "D" Then
      sdx = Val(octubred) + Val("" & mytablex.Fields("xtotal"))
      octubred = "" & sdx
   End If
End If
If Val("" & mytablex.Fields("mes")) = 11 Then
   If "" & mytablex.Fields("moneda") = "S" Then
      sdx = Val(noviembres) + Val("" & mytablex.Fields("xtotal"))
      noviembres = "" & sdx
   End If
   If "" & mytablex.Fields("moneda") = "D" Then
      sdx = Val(noviembred) + Val("" & mytablex.Fields("xtotal"))
      noviembred = "" & sdx
   End If
End If
If Val("" & mytablex.Fields("mes")) = 12 Then
   If "" & mytablex.Fields("moneda") = "S" Then
      sdx = Val(diciembres) + Val("" & mytablex.Fields("xtotal"))
      diciembres = "" & sdx
   End If
   If "" & mytablex.Fields("moneda") = "D" Then
      sdx = Val(diciembred) + Val("" & mytablex.Fields("xtotal"))
      diciembred = "" & sdx
   End If
End If
mytablex.MoveNext
Loop
mytablex.Close
suma_meses
End Sub
Sub inicializa_x()
fechanac = ""
eneros = ""
enerod = ""
febreros = ""
febrerod = ""
marzos = ""
marzod = ""
abrils = ""
abrild = ""
mayos = ""
mayod = ""
junios = ""
juniod = ""
agostos = ""
agostod = ""
julios = ""
juliod = ""
setiembres = ""
setiembred = ""
agostos = ""
agostod = ""
octubres = ""
octubred = ""
noviembres = ""
noviembred = ""
diciembres = ""
diciembred = ""
suma_meses
End Sub
Sub suma_meses()
Dim sdx As Double
sdx = Val(eneros) + Val(febreros) + Val(marzos) + Val(abrils) + Val(mayos) + Val(junios) + Val(agostos) + Val(julios) + Val(setiembres) + Val(octubres) + Val(noviembres) + Val(diciembres)
totals = Format(sdx, "0.00")
sdx = Val(enerod) + Val(febrerod) + Val(marzod) + Val(abrild) + Val(mayod) + Val(juniod) + Val(agostod) + Val(juliod) + Val(setiembred) + Val(octubred) + Val(noviembred) + Val(diciembred)
totald = Format(sdx, "0.00")

End Sub
Function busca_parame(sw As Integer)
Dim sdx As Double
Dim mytablex As Table
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic
If mytablex.EOF = True And mytablex.BOF = True Then
   Exit Function
End If
   If sw = 0 Then
      sdx = Val("" & mytablex.Fields("proveedo")) + 1
      codigo = "" & sdx
   End If
   If sw = 1 Then
      mytablex.Edit
      mytablex.Fields("proveedo") = "" & codigo
      mytablex.Update
   End If
End Function
Sub sql_ranking()
Dim buf As String
Dim buf1 As String
Dim found As Integer
Dim rconsulta As New ADODB.Recordset

On Error GoTo cmd454_err
buf1 = "detalle"
If Option6.Value = -1 Then
   buf1 = "dpedidov"
End If
cerrar_archivo
   'MsgBox globaldat & "\_" & gusuario & ".dbf"
   found = borra_nombre(globaldat & "\_" & gusuario & ".dbf")
   buf = "select Producto,Descripcio,Unidad,Factor,moneda as m,sum(cantidad) as xcant,sum(total) as xtotal from " & buf1 & " where "
   buf = buf & " codigo='" & codigo & "'"
If buf1 = "detalle" Then
   buf = buf & " and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F')"
End If
   buf = buf & " and year(fecha)=" & ranno
   buf = buf & "  group by producto,Descripcio,Unidad,Factor,moneda  order  by SUM(total) DESC "
   
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open buf, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True And rconsulta.BOF = True Then
      Exit Sub
   End If
   Set dbgrid6.DataSource = rconsulta
               dbgrid6.Columns(0).Width = 1300
               dbgrid6.Columns(1).Width = 4500
               dbgrid6.Columns(2).Width = 800
               dbgrid6.Columns(3).Width = 800
               dbgrid6.Columns(4).Width = 500
               dbgrid6.Columns(5).Width = 1300
               dbgrid6.Columns(6).Width = 1300
               dbgrid6.SetFocus
               Exit Sub
cmd454_err:
MsgBox error$
Exit Sub

End Sub



