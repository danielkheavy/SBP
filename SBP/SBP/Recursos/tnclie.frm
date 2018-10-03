VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tnclie 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tabla de Clientes"
   ClientHeight    =   9165
   ClientLeft      =   150
   ClientTop       =   0
   ClientWidth     =   15540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   15540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   360
      Left            =   14565
      TabIndex        =   139
      Top             =   5820
      Width           =   990
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8895
      Left            =   4920
      TabIndex        =   127
      Top             =   8880
      Visible         =   0   'False
      Width           =   15015
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
         Left            =   3645
         MaxLength       =   10
         TabIndex        =   130
         TabStop         =   0   'False
         Top             =   600
         Width           =   3420
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
         TabIndex        =   129
         TabStop         =   0   'False
         Top             =   600
         Width           =   3375
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
         Left            =   8280
         TabIndex        =   128
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid dbGrid3 
         Height          =   7575
         Left            =   240
         TabIndex        =   131
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
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Otros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   6855
      Left            =   9960
      TabIndex        =   103
      Top             =   1200
      Visible         =   0   'False
      Width           =   12240
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
         TabIndex        =   118
         Top             =   405
         Width           =   3615
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
         TabIndex        =   117
         Top             =   1605
         Width           =   3615
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
         TabIndex        =   116
         Top             =   1965
         Width           =   375
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
         TabIndex        =   115
         Top             =   2325
         Width           =   1095
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
         TabIndex        =   114
         Top             =   765
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
         TabIndex        =   113
         Top             =   1125
         Width           =   3615
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
         TabIndex        =   112
         Top             =   2685
         Width           =   3615
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
         TabIndex        =   111
         Top             =   3045
         Width           =   375
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Sa&lir"
         Height          =   855
         Left            =   10650
         Picture         =   "tnclie.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   330
         Width           =   1095
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Direcciones de Despacho"
         Height          =   3255
         Left            =   240
         TabIndex        =   104
         Top             =   3450
         Width           =   10455
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
            TabIndex        =   105
            Top             =   240
            Width           =   7575
         End
         Begin MSDataGridLib.DataGrid dbgrid6 
            Height          =   2415
            Left            =   120
            TabIndex        =   106
            Top             =   720
            Width           =   9375
            _ExtentX        =   16536
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
         Begin VB.Label Label43 
            BackColor       =   &H00808080&
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
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   120
            TabIndex        =   109
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label45 
            Alignment       =   2  'Center
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
            Left            =   9600
            TabIndex        =   108
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label46 
            Alignment       =   2  'Center
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
            Left            =   9600
            TabIndex        =   107
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808080&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   126
         Top             =   405
         Width           =   2175
      End
      Begin VB.Label Label32 
         BackColor       =   &H00808080&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   125
         Top             =   1605
         Width           =   2175
      End
      Begin VB.Label Label33 
         BackColor       =   &H00808080&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   124
         Top             =   1965
         Width           =   2175
      End
      Begin VB.Label Label34 
         BackColor       =   &H00808080&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   123
         Top             =   2325
         Width           =   2175
      End
      Begin VB.Label xk44 
         BackColor       =   &H00808080&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   122
         Top             =   765
         Width           =   2175
      End
      Begin VB.Label Label36 
         BackColor       =   &H00808080&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   121
         Top             =   1125
         Width           =   2175
      End
      Begin VB.Label Label37 
         BackColor       =   &H00808080&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   120
         Top             =   2685
         Width           =   2175
      End
      Begin VB.Label Label38 
         BackColor       =   &H00808080&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   119
         Top             =   3045
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Height          =   8655
      Left            =   0
      TabIndex        =   56
      Top             =   720
      Visible         =   0   'False
      Width           =   15135
      Begin VB.TextBox provincia 
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
         TabIndex        =   141
         Top             =   4680
         Width           =   3375
      End
      Begin VB.TextBox clasesunat 
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
         Left            =   9570
         MaxLength       =   6
         TabIndex        =   33
         Top             =   6960
         Width           =   975
      End
      Begin VB.ComboBox clasesunat1 
         Height          =   315
         Left            =   10680
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   6720
         Width           =   2055
      End
      Begin VB.TextBox estadocredito 
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
         TabIndex        =   30
         Top             =   4080
         Width           =   735
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
         Left            =   2280
         MaxLength       =   1
         TabIndex        =   1
         Text            =   "N"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox cliente 
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
         TabIndex        =   0
         Top             =   240
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
         MaxLength       =   60
         TabIndex        =   3
         Top             =   1320
         Width           =   6015
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   11280
         Picture         =   "tnclie.frx":08DA
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Imprimir todo"
         Top             =   2040
         Width           =   1470
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   11280
         Picture         =   "tnclie.frx":11A4
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   960
         Width           =   1470
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
         Left            =   2280
         MaxLength       =   15
         TabIndex        =   6
         Top             =   2880
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
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   7680
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
         TabIndex        =   22
         Top             =   8040
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
         TabIndex        =   8
         Top             =   3240
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
         TabIndex        =   13
         Top             =   5400
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
         TabIndex        =   7
         Top             =   3240
         Width           =   1935
      End
      Begin VB.CheckBox viernes 
         BackColor       =   &H00808080&
         Caption         =   "Viernes"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9000
         TabIndex        =   63
         Top             =   5550
         Width           =   855
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
         TabIndex        =   2
         Top             =   960
         Width           =   1935
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
         TabIndex        =   4
         Top             =   1680
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
         TabIndex        =   5
         Top             =   2040
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
         MaxLength       =   100
         TabIndex        =   9
         Top             =   3600
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
         TabIndex        =   11
         Top             =   4320
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
         TabIndex        =   12
         Top             =   5040
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
         TabIndex        =   15
         Top             =   6120
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
         TabIndex        =   16
         Top             =   6120
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
         TabIndex        =   17
         Top             =   6120
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
         TabIndex        =   14
         Top             =   5760
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
         TabIndex        =   24
         Top             =   1800
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
         TabIndex        =   25
         Top             =   2160
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
         TabIndex        =   26
         Top             =   2520
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
         TabIndex        =   27
         Top             =   2880
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
         MaxLength       =   13
         TabIndex        =   18
         Top             =   6600
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
         TabIndex        =   29
         Top             =   3720
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
         TabIndex        =   31
         Top             =   4440
         Width           =   1215
      End
      Begin VB.CheckBox lunes 
         BackColor       =   &H00808080&
         Caption         =   "Lunes"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7080
         TabIndex        =   62
         Top             =   5190
         Width           =   735
      End
      Begin VB.CheckBox martes 
         BackColor       =   &H00808080&
         Caption         =   "Martes"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7920
         TabIndex        =   61
         Top             =   5190
         Width           =   855
      End
      Begin VB.CheckBox miercoles 
         BackColor       =   &H00808080&
         Caption         =   "Miercoles"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9000
         TabIndex        =   60
         Top             =   5190
         Width           =   975
      End
      Begin VB.CheckBox jueves 
         BackColor       =   &H00808080&
         Caption         =   "Jueves"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   10080
         TabIndex        =   59
         Top             =   5190
         Width           =   855
      End
      Begin VB.CheckBox sabado 
         BackColor       =   &H00808080&
         Caption         =   "Sabado"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7080
         TabIndex        =   58
         Top             =   5550
         Width           =   855
      End
      Begin VB.CheckBox domingo 
         BackColor       =   &H00808080&
         Caption         =   "Domingo"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7920
         TabIndex        =   57
         Top             =   5550
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
         TabIndex        =   32
         Top             =   4800
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
         TabIndex        =   28
         Top             =   3240
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
         TabIndex        =   21
         Top             =   7680
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
         MaxLength       =   13
         TabIndex        =   19
         Top             =   6960
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
         TabIndex        =   20
         Top             =   7320
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
         TabIndex        =   10
         Top             =   3960
         Width           =   4695
      End
      Begin VB.Label Label49 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Provincia"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   140
         Top             =   4680
         Width           =   2175
      End
      Begin VB.Label Label44 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Percep."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   8400
         TabIndex        =   102
         Top             =   6960
         Width           =   1215
      End
      Begin VB.Label Label39 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Credito Habilitado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   7080
         TabIndex        =   101
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "J (Juridica) ; D (DNI) ; N(Natural) P.Pasaporte X.Extranjeria O.Otros"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2670
         TabIndex        =   100
         Top             =   645
         Width           =   5865
      End
      Begin VB.Label Label28 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   99
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mas Datos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11250
         TabIndex        =   98
         Top             =   240
         Width           =   1470
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   97
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre/Razon Social"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   96
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label31 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cod.Barras"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   95
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label Label27 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cliente Especial 1.Si"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4200
         TabIndex        =   94
         Top             =   7680
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   93
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   92
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   91
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Direccion"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   90
         Top             =   3600
         Width           =   2175
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Departamento"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   89
         Top             =   4320
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distrito"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   88
         Top             =   5040
         Width           =   2175
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Telefono"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   87
         Top             =   6120
         Width           =   2175
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Correo Electronico"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   86
         Top             =   5760
         Width           =   2175
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Zona"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   85
         Top             =   5400
         Width           =   2175
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   84
         Top             =   8040
         Width           =   2175
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   7080
         TabIndex        =   83
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   7080
         TabIndex        =   82
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   7080
         TabIndex        =   81
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuenta"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   7080
         TabIndex        =   80
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vendedor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   79
         Top             =   6600
         Width           =   2175
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   7080
         TabIndex        =   78
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   7080
         TabIndex        =   77
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label Label29 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   7080
         TabIndex        =   76
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   7080
         TabIndex        =   75
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaNacimiento"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   74
         Top             =   7680
         Width           =   2175
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   73
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label clasddd 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clasificacion"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4200
         TabIndex        =   72
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label25 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Referido Por"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   71
         Top             =   6960
         Width           =   2175
      End
      Begin VB.Label Label26 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Garantia"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   70
         Top             =   7320
         Width           =   2175
      End
      Begin VB.Label ngarantia 
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
         Left            =   4200
         TabIndex        =   69
         Top             =   7320
         Width           =   4095
      End
      Begin VB.Label nreferencia 
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
         Left            =   4200
         TabIndex        =   68
         Top             =   6960
         Width           =   4095
      End
      Begin VB.Label Label35 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   67
         Top             =   3960
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "TestHuella"
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
      Left            =   12780
      TabIndex        =   55
      Top             =   6075
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Huella"
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
      Left            =   12780
      TabIndex        =   54
      Top             =   5595
      Width           =   1455
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   13620
      Style           =   2  'Dropdown List
      TabIndex        =   52
      Top             =   4515
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Enviar Correo"
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
      Left            =   12780
      TabIndex        =   50
      Top             =   5115
      Width           =   1455
   End
   Begin VB.ComboBox perfil 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   13620
      Style           =   2  'Dropdown List
      TabIndex        =   49
      Top             =   4155
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   15480
      TabIndex        =   36
      Top             =   0
      Width           =   15540
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
         Picture         =   "tnclie.frx":1A6E
         Style           =   1  'Graphical
         TabIndex        =   44
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
         Left            =   7560
         MaxLength       =   10
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   120
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "tnclie.frx":2C80
         Left            =   4560
         List            =   "tnclie.frx":2C8D
         Style           =   2  'Dropdown List
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   120
         Width           =   3015
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
         Left            =   10575
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   60
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
         Picture         =   "tnclie.frx":2CB3
         Style           =   1  'Graphical
         TabIndex        =   40
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
         Left            =   2760
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tnclie.frx":3EC5
         Style           =   1  'Graphical
         TabIndex        =   39
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
         Left            =   2160
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tnclie.frx":50D7
         Style           =   1  'Graphical
         TabIndex        =   38
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
         Picture         =   "tnclie.frx":62E9
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label flag 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   12300
         TabIndex        =   132
         Top             =   135
         Width           =   45
      End
      Begin VB.Label nreg 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3645
         TabIndex        =   47
         Top             =   150
         Width           =   825
      End
      Begin VB.Label DBPROV 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   720
         Visible         =   0   'False
         Width           =   45
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8295
      Left            =   0
      TabIndex        =   34
      Top             =   960
      Width           =   12495
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   7935
         Left            =   195
         TabIndex        =   35
         Top             =   300
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   13996
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
         ColumnCount     =   8
         BeginProperty Column00 
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
         BeginProperty Column01 
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
         BeginProperty Column02 
            DataField       =   "Tipoclie"
            Caption         =   "TipoClie"
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
            DataField       =   "Clasifica"
            Caption         =   "Clasifica"
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
            DataField       =   "Correo"
            Caption         =   "Correo"
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
         BeginProperty Column06 
            DataField       =   "Percepcion"
            Caption         =   "Percepcion"
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
            DataField       =   "ClaseSunat"
            Caption         =   "ClaseSunat"
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
               ColumnWidth     =   5940.284
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
         EndProperty
      End
   End
   Begin VB.CheckBox cumple 
      Caption         =   "Cumple años este mes"
      Height          =   435
      Left            =   12690
      TabIndex        =   133
      Top             =   975
      Width           =   1305
   End
   Begin VB.TextBox Ruta 
      Enabled         =   0   'False
      Height          =   315
      Left            =   12660
      TabIndex        =   135
      Top             =   2685
      Width           =   2955
   End
   Begin VB.CommandButton AbrirRuta 
      Caption         =   "..."
      Height          =   360
      Left            =   13170
      TabIndex        =   136
      Top             =   2250
      Width           =   510
   End
   Begin VB.TextBox Imagen 
      Enabled         =   0   'False
      Height          =   315
      Left            =   12615
      TabIndex        =   138
      Text            =   "archivo"
      Top             =   3450
      Width           =   2925
   End
   Begin VB.Label Label41 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Seleccion"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12660
      TabIndex        =   53
      Top             =   4515
      Width           =   975
   End
   Begin VB.Label Label42 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Perfil"
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
      Left            =   12660
      TabIndex        =   51
      Top             =   4155
      Width           =   975
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EnviarCorreo"
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
      Left            =   12660
      TabIndex        =   48
      Top             =   3795
      Width           =   2415
   End
   Begin VB.Label activado 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   46
      Top             =   9120
      Width           =   45
   End
   Begin VB.Label Label48 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre Archivo:"
      Height          =   195
      Left            =   12675
      TabIndex        =   137
      Top             =   3120
      Width           =   1185
   End
   Begin VB.Label Label47 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ruta:"
      Height          =   315
      Left            =   12660
      TabIndex        =   134
      Top             =   2340
      Width           =   390
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
   Begin VB.Menu fdlo893 
      Caption         =   "&ListaPrecios"
   End
   Begin VB.Menu djuer1 
      Caption         =   "&Reporte"
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tnclie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim txcliem As New ADODB.Recordset

Private Sub AbrirRuta_Click()

    Shell ("explorer.exe " & Ruta), vbMaximizedFocus

End Sub

Private Sub ajdu1_Click()

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    inicializa
    Frame2.Visible = True
    Frame2.Caption = "Nuevo"
    cmdGuardar.Enabled = True
    habilita 1

    '10/09/2017 kenyo habilita siempre campo codigo
    cliente.Enabled = True
    cliente.SetFocus
    '10/09/2017 kenyo habilita siempre campo codigo

    cliente.Text = ""

    'On Error Resume Next
    'cliente.SetFocus
End Sub

Private Sub bo712_Click()

    Dim buf As String

    On Error GoTo cmd656_err

    buf = "" & txcliem.Fields("CODIGO")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    If MsgBox("Desea Borra " + "" & txcliem.Fields("CODIGO"), 1, "Aviso") <> 1 Then
        Exit Sub

    End If

    'cn.Execute ("delete from precio1 where codigo='" & buf & "'")
    txcliem.Delete
    Command1_Click
    Exit Sub
cmd656_err:
    MsgBox "Seleccione un dato " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub buffer_DblClick()
    Command1_Click

End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    Command1_Click

End Sub

Private Sub Command6_Click()
    txcliem.Update

End Sub

Private Sub cumple_Click()

    If cumple.Value = 1 Then
        ejecuta2
    Else
        Frame1.Visible = True
        Frame1.Enabled = True
        opcion1 = "1"
        ejecuta 1

    End If

End Sub

Private Sub clasesunat1_Click()

    If clasesunat1 = "%" Then Exit Sub
    clasesunat = extra_loquesea(clasesunat1)

End Sub

Private Sub cmdAddEntry_Click()
    ajdu1_Click

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

Private Sub cmdGuardar_Click()

    Dim found As Integer

    found = grabar()

End Sub

Private Sub cmdPrint_Click()
    djuer1_Click

End Sub

Private Sub cmdSave_Click()
    f8443_Click

End Sub

Private Sub cliente_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        dlo132_Click
        Exit Sub

    End If

    If Len(cliente) = 0 Then
        cliente.SetFocus
        Exit Sub

    End If

    tipo.SetFocus

End Sub

Private Sub Command1_Click()
    Frame1.Visible = True
    Frame1.Enabled = True
    opcion1 = "1"
    ejecuta 1

End Sub

Sub ejecuta(sw As Integer)

    Dim cad As String

    If opcion1 = "1" Then
        If Len(buffer) = 0 Then
      
            cad = "SELECT * from " & DBPROV & " order by nombre " 'DBPROV = clientes

            'cad = "SELECT * from clientes order by nombre " 'DBPROV = clientes
        End If

        If Len(buffer) > 0 Then

            Select Case Combo1.Text

                Case "Nombre": cad = "SELECT * from " & DBPROV & " where nombre like '%" & buffer.Text & "%' order by nombre"

                Case "Nombre Comercial": cad = "SELECT * from " & DBPROV & "  where nombrec like '%" & buffer.Text & "%' order by nombrec"

                Case "Codigo": cad = "SELECT * from " & DBPROV & "  where codigo like '%" & buffer.Text & "%'"

                Case "Codigo1": cad = "SELECT * from " & DBPROV & "  where codigo1 like '%" & buffer.Text & "%'"

                Case "Tipo": cad = "SELECT * from " & DBPROV & "  where tipo like '%" & buffer.Text & "%'"

                Case "Contacto": cad = "SELECT * from " & DBPROV & "  where contacto like '%" & buffer.Text & "%'"

                Case "Direccion": cad = "SELECT * from " & DBPROV & "  where direccion like '%" & buffer.Text & "%'"

                Case "Referencia": cad = "SELECT * from " & DBPROV & "  where observa like '%" & buffer.Text & "%'"

                Case "Departamento": cad = "SELECT * from " & DBPROV & "  where DPTO like '%" & buffer.Text & "%'"

                Case "Distrito": cad = "SELECT * from " & DBPROV & "  where distrito like '%" & buffer.Text & "%'"

                Case "Zona": cad = "SELECT * from " & DBPROV & "  where zona like '%" & buffer.Text & "%'"

                Case "Correo": cad = "SELECT * from " & DBPROV & "  where correo like '%" & buffer.Text & "%'"

                Case "Telefono": cad = "SELECT * from " & DBPROV & "  where telefono like '%" & buffer.Text & "%'"

                Case "Telefono1": cad = "SELECT * from " & DBPROV & "  where telefono1 like '%" & buffer.Text & "%'"

                Case "Telefono2": cad = "SELECT * from " & DBPROV & "  where telefono2 like '%" & buffer.Text & "%'"

                Case "Clasifica": cad = "SELECT * from " & DBPROV & "  where CLASIFICA like '%" & buffer.Text & "%'"

            End Select

        End If

        If txcliem.State = 1 Then txcliem.Close
        txcliem.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbGrid1.DataSource = txcliem
        dbGrid1.columns(0).Width = 4000
        dbGrid1.columns(1).Width = 2000

        If txcliem.RecordCount > 0 Then
            dbGrid1.Enabled = True
            dbGrid1.SetFocus

        End If

        nreg = "" & txcliem.RecordCount

    End If

End Sub

'Sub ejecuta(sw As Integer)
'Dim cad As String
'If opcion1 = "1" Then
'    If Len(buffer) = 0 Then
'      'cad = "SELECT * from " & DBPROV & " order by nombre " 'DBPROV = clientes
'      cad = "SELECT * from clientes order by nombre " 'DBPROV = clientes
'    End If
'    If Len(buffer) > 0 Then
'        Select Case Combo1.Text
'            Case "Nombre": cad = "SELECT * from clientes where nombre like '%" & buffer.Text & "%' order by nombre"
'            Case "Nombre Comercial": cad = "SELECT * from clientes where nombrec like '%" & buffer.Text & "%' order by nombrec"
'            Case "Codigo": cad = "SELECT * from clientes where codigo like '%" & buffer.Text & "%'"
'            Case "Codigo1": cad = "SELECT * from clientes where codigo1 like '%" & buffer.Text & "%'"
'            Case "Tipo": cad = "SELECT * from clientes where tipo like '%" & buffer.Text & "%'"
'            Case "Contacto": cad = "SELECT * from clientes where contacto like '%" & buffer.Text & "%'"
'            Case "Direccion": cad = "SELECT * from clientes where direccion like '%" & buffer.Text & "%'"
'            Case "Referencia": cad = "SELECT * from clientes where observa like '%" & buffer.Text & "%'"
'            Case "Departamento": cad = "SELECT * from clientes where DPTO like '%" & buffer.Text & "%'"
'            Case "Distrito": cad = "SELECT * from clientes where distrito like '%" & buffer.Text & "%'"
'            Case "Zona": cad = "SELECT * from clientes where zona like '%" & buffer.Text & "%'"
'            Case "Correo": cad = "SELECT * from clientes where correo like '%" & buffer.Text & "%'"
'            Case "Telefono": cad = "SELECT * from clientes where telefono like '%" & buffer.Text & "%'"
'            Case "Telefono1": cad = "SELECT * from clientes where telefono1 like '%" & buffer.Text & "%'"
'            Case "Telefono2": cad = "SELECT * from clientes where telefono2 like '%" & buffer.Text & "%'"
'            Case "Clasifica": cad = "SELECT * from clientes where CLASIFICA like '%" & buffer.Text & "%'"
'        End Select
'    End If
'    If txcliem.State = 1 Then txcliem.Close
'    txcliem.Open cad, cn, adOpenStatic, adLockOptimistic
'    Set dbGrid1.DataSource = txcliem
'    dbGrid1.columns(0).Width = 4000
'    dbGrid1.columns(1).Width = 2000
'    If txcliem.RecordCount > 0 Then
'       dbGrid1.Enabled = True
'       dbGrid1.SetFocus
'    End If
'    nreg = "" & txcliem.RecordCount
'End If
'
'
''Combo1.AddItem "Nombre" 'nombre
''Combo1.AddItem "Nombre Comercial" ' nombrec
''Combo1.AddItem "Codigo" ' codigo
''Combo1.AddItem "Codigo1" ' codigo1
''Combo1.AddItem "Tipo" ' tipo
''Combo1.AddItem "Contacto" ' contacto
''Combo1.AddItem "Direccion" ' direccion
''Combo1.AddItem "Referencia" ' 'observa
''Combo1.AddItem "Departamento" ' dpto
''Combo1.AddItem "Distrito" ' distrito
''Combo1.AddItem "Zona" ' zona
''Combo1.AddItem "Correo" ' correo
''Combo1.AddItem "Telefono" 'telefono
''Combo1.AddItem "Telefono1" 'telefono1
''Combo1.AddItem "Telefono2" ' telefono2
'
'''If opcion1 = "1" Then  'bodega
'''   If Len(buffer) = 0 Then
'''      'cad = "SELECT * from " & DBPROV & " order by nombre " 'DBPROV = clientes
'''      cad = "SELECT * from clientes order by nombre " 'DBPROV = clientes
'''   End If
'''   If Len(buffer) > 0 Then
'''      cad = "SELECT * from clientes where  " & Combo1 & " like '" & buffer & "%' order by nombre"
'''   End If
'''   'MsgBox cad
'''   If txcliem.State = 1 Then txcliem.Close
'''   txcliem.Open cad, cn, adOpenStatic, adLockReadOnly 'adLockOptimistic
'''   Set dbGrid1.DataSource = txcliem
'''   dbGrid1.columns(0).Width = 4000
'''   dbGrid1.columns(1).Width = 2000
'''   If txcliem.RecordCount > 0 Then
'''     dbGrid1.Enabled = True
'''     dbGrid1.SetFocus
'''   End If
'''  nreg = "" & txcliem.RecordCount
'''End If
'
'End Sub
Sub ejecuta2()

    Dim cad As String

    If Len(buffer) = 0 Then
        'cad = "SELECT * from " & DBPROV & " order by nombre " 'DBPROV = clientes
        cad = "SELECT * from " & DBPROV & " where month(fechanac)=(select month(getdate())) and correo<>''   order by nombre " 'DBPROV = clientes

    End If
   
    If txcliem.State = 1 Then txcliem.Close
    txcliem.Open cad, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = txcliem
    dbGrid1.columns(0).Width = 4000
    dbGrid1.columns(1).Width = 2000

    If txcliem.RecordCount > 0 Then
        dbGrid1.Enabled = True
        dbGrid1.SetFocus

    End If

    nreg = "" & txcliem.RecordCount

End Sub

Private Sub Command2_Click()
    envio_correos

End Sub

Private Sub Command3_Click()
    ejecuta1 1

End Sub

Sub ejecuta1(sw As Integer)

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    If opcion1 = "1" Then
        If Len(cadena) = 0 Then
            buf = "select Descripcio,tipoclie from tipoclie "
        Else
            buf = "select Descripcio,tipoclie from tipoclie where " & Combo2 & " like '%" & cadena & "%'"

        End If

    End If

    If opcion1 = "2" Then
        If Len(cadena) = 0 Then
            buf = "select Nombre,Codigo from Vendedor "
        Else
            buf = "select Nombre,Codigo from Vendedor where " & Combo2 & " like '%" & cadena & "%'"

        End If

    End If

    If opcion1 = "200" Or opcion1 = "201" Then
        If Len(cadena) = 0 Then
            buf = "select Nombre,Codigo from  " & DBPROV
        Else
            buf = "select Nombre,Codigo from " & DBPROV & " where " & Combo2 & " like '%" & cadena & "%'"

        End If

    End If

    If opcion1 = "300" Then
        If Len(cadena) = 0 Then
            buf = "select Nombre,Codigo from  " & DBPROV
        Else
            buf = "select Nombre,Codigo from " & DBPROV & " where " & Combo2 & " like '%" & cadena & "%'"

        End If

    End If

    If opcion1 = "3" Then
        If Len(cadena) = 0 Then
            buf = "select Descripcio,Clasifica from clasifi "
        Else
            buf = "select Descripcio,Clasifica from Clasifi where " & Combo2 & " like '%" & cadena & "%'"

        End If

    End If

    If opcion1 = "4" Then
        If Len(cadena) = 0 Then
            buf = "select Descripcio,Zona from Zona "
        Else
            buf = "select Descripcio,Zona from Zona where " & Combo2 & " like '%" & cadena & "%'"

        End If

    End If

    If opcion1 = "5" Then
        If Len(cadena) = 0 Then
            buf = "select Descripcio,Fpago from Fpago "
        Else
            buf = "select Descripcio,Fpago from Fpago where " & Combo2 & " like '%" & cadena & "%'"

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

Private Sub Command4_Click()

    Dim buf As String

    On Error GoTo cmd8000_err

    buf = "" & txcliem.Fields("CODIGO")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    If Len(Trim(buf)) = 0 Then
        Exit Sub

    End If

    thuellad.tipo = "Clientes"
    thuellad.codigo = buf
    thuellad.nombre = "" & txcliem.Fields("nombre")
    thuellad.Show 1
    Exit Sub
cmd8000_err:
    MsgBox "Seleccione Dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command5_Click()

    Dim buf As String

    On Error GoTo cmd88000_err

    buf = "" & txcliem.Fields("CODIGO")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    If Len(Trim(buf)) = 0 Then
        Exit Sub

    End If

    thuellat.tipo = "Clientes"
    thuellat.codigo = buf
    thuellat.nombre = "" & txcliem.Fields("nombre")
    thuellat.Show 1
    Exit Sub
cmd88000_err:
    MsgBox "Seleccione Dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command9_Click()
    Frame4.Visible = False

End Sub

Private Sub dbgrid1_DblClick()
    f8443_Click

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then

        'cliente = dbGrid1.Columns(1)
        'Frame1.Visible = False
        'Frame1.Enabled = False
        'cliente.SetFocus
        'cliente_KeyPress 13
    End If

End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)

    Dim buf  As String

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

Private Sub DBGrid3_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        cadena.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = "1" Then
            tipoclie = dbgrid3.columns(1)
            Frame3.Visible = False
            Frame3.Enabled = False
            tipoclie.SetFocus

        End If

        If opcion1 = "2" Then
            vendedor = dbgrid3.columns(1)
            Frame3.Visible = False
            Frame3.Enabled = False
            vendedor.SetFocus

        End If

        If opcion1 = "3" Then
            clasifica = dbgrid3.columns(1)
            Frame3.Visible = False
            Frame3.Enabled = False
            clasifica.SetFocus

        End If

        If opcion1 = "3" Then
            Exit Sub
      
        End If

        If opcion1 = "4" Then
            zona = dbgrid3.columns(1)
            Frame3.Visible = False
            Frame3.Enabled = False
            zona.SetFocus

        End If

        If opcion1 = "5" Then
            fpago = dbgrid3.columns(1)
            Frame3.Visible = False
            Frame3.Enabled = False
            fpago.SetFocus

        End If
   
        If opcion1 = "200" Then
            referencia = dbgrid3.columns(1)
            nreferencia = dbgrid3.columns(0)
            Frame3.Visible = False
            Frame3.Enabled = False
            referencia.SetFocus

        End If

        If opcion1 = "201" Then
            garantia = dbgrid3.columns(1)
            nreferencia = dbgrid3.columns(0)
            Frame3.Visible = False
            Frame3.Enabled = False
            garantia.SetFocus

        End If

    End If

End Sub

Private Sub djuer1_Click()
    'If Frame2.Visible = True Then Exit Sub
    'Call Excel_a_Access("C:\ORION.V5\tclientes.xlsx", 10, 3)   'FILA COLUMNAS
    'Exit Sub

    reporgen.NAMETABLA = DBPROV
    reporgen.Show 1

End Sub

Sub envio_correos()

    Dim txtserver     As String

    Dim txtusername   As String

    Dim txtpassword   As String

    Dim txtport       As String

    Dim txtto         As String

    Dim chkssl        As String

    Dim txtfromname   As String

    Dim txtfromemail  As String

    Dim txtattach     As String

    Dim txtsubject    As String

    Dim txtmsg        As String

    Dim retval        As String

    Dim txthtml       As String

    Dim txtselecciona As String

    'Dim txtselecciona As String
    Dim mytablex      As New ADODB.Recordset

    Dim buf           As String

    On Error GoTo cmd0905677_err

    buf = extra_loquesea1(perfil)

    If Trim(buf) = 0 Then Exit Sub
    mytablex.Open "select * from correos where cosms='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        txtserver = Trim("" & mytablex.Fields("txtserver"))
        txtusername = Trim("" & mytablex.Fields("txtusername"))
        txtpassword = Trim("" & mytablex.Fields("txtpassword"))
        txtfromname = Trim("" & mytablex.Fields("txtfromname"))
        txtfromemail = Trim("" & mytablex.Fields("txtfromemail"))
        txtport = Trim("" & mytablex.Fields("txtport"))
        txtselecciona = Trim("" & mytablex.Fields("txtselecciona"))
        'txtto = Trim("" & mytablex.Fields("txtto"))
        chkssl = Trim("" & mytablex.Fields("chkssl"))
        'txtfromname = Trim("" & nombre) 'Trim("" & mytablex.Fields("txtfromname"))
        txtto = Trim("" & correo) 'Trim("" & mytablex.Fields("txtfromemail"))
        txtattach = App.path & "\ico\archivo.jpg"
        txtsubject = Trim("" & mytablex.Fields("txtsubject"))
        txtmsg = Trim("" & mytablex.Fields("txtmsg"))
        txtmsg = txtmsg & Chr$(10) & Chr$(13) & ""
        txtmsg = txtmsg & Format(Now, "dd/mm/yyyy") + " " + Format(Now, "hh:mm:ss")

        If Combo3 = "Seleccionado" Then
            If Len(Trim("" & txcliem.Fields("correo"))) > 0 Then
                txtto = Trim("" & txcliem.Fields("correo"))
                retval = SendMail(Trim$(txtto), Trim$(txtsubject), Trim$(txtfromname) & "<" & Trim$(txtfromemail) & ">", Trim$(txtmsg), Trim$(txtserver), CInt(Trim$(txtport)), Trim$(txtusername), Trim$(txtpassword), Trim$(txtattach), True, txtselecciona, txthtml)

            End If

            MsgBox "Proceso Realizado ", 48, "Aviso"
            mytablex.Close

        End If

        If Combo3 = "Todos" Then
            Do

                If txcliem.EOF Then Exit Do
                If Len(Trim("" & txcliem.Fields("correo"))) > 0 Then
                    txtto = Trim("" & txcliem.Fields("correo"))
                    retval = SendMail(Trim$(txtto), Trim$(txtsubject), Trim$(txtfromname) & "<" & Trim$(txtfromemail) & ">", Trim$(txtmsg), Trim$(txtserver), CInt(Trim$(txtport)), Trim$(txtusername), Trim$(txtpassword), Trim$(txtattach), True, txtselecciona, txthtml)

                End If

                txcliem.MoveNext
            Loop
            MsgBox "Proceso Realizado ", 48, "Aviso"
            mytablex.Close

        End If

    End If

    Exit Sub
cmd0905677_err:
    MsgBox "No se Pudo enviar Correo... " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub dki8923_Click()
    envio_correos

End Sub

Private Sub dlo132_Click()

    If FLAG = "NUEVO" Then
        tnclie.Hide
        Unload tnclie
        Exit Sub

    End If

    If Frame3.Visible = True Then
        cadena_KeyPress 27
        Exit Sub

    End If

    If Frame2.Visible = True Then
        habilita 0
        Frame2.Visible = False
        dbGrid1.Enabled = True
        Exit Sub

    End If

    tnclie.Hide
    Unload tnclie

End Sub

Private Sub f8443_Click()

    Dim buf As String

    On Error GoTo cmd456_err

    buf = txcliem.Fields("CODIGO")

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
    cliente.Enabled = False
    nombre.SetFocus
    Exit Sub
cmd456_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fdlo893_Click()
    tpactop.codigo = "" & txcliem.Fields("codigo")
    tpactop.nombre = "" & txcliem.Fields("nombre")
    tpactop.Show 1

End Sub

Private Sub fjh433_Click()

    Dim buf As String

    On Error GoTo cmd556_err

    buf = txcliem.Fields("CODIGO")

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
    cliente.Enabled = False
    nombre.SetFocus
    Exit Sub
cmd556_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Form_Activate()

    'MsgBox DBPROV
    If activado <> "S" Then
        Command1_Click
        activado = "S"

    End If

    If FLAG = "NUEVO" Then
        ajdu1_Click

    End If

    Frame2.Top = 0: Frame2.Left = 0
    Frame3.Top = 0: Frame3.Left = 0
    Frame4.Top = 1125

End Sub

Private Sub Form_Load()
    Combo3.Clear
    Combo3.AddItem "Seleccionado"
    Combo3.AddItem "Todos"
    Combo3.ListIndex = 0

    Combo1.Clear
    Combo1.AddItem "Nombre" 'nombre
    Combo1.AddItem "Nombre Comercial" ' nombrec
    Combo1.AddItem "Codigo" ' codigo
    Combo1.AddItem "Codigo1" ' codigo1
    Combo1.AddItem "Tipo" ' tipo
    Combo1.AddItem "Contacto" ' contacto
    Combo1.AddItem "Direccion" ' direccion
    Combo1.AddItem "Referencia" ' 'observa
    Combo1.AddItem "Departamento" ' dpto
    Combo1.AddItem "Distrito" ' distrito
    Combo1.AddItem "Zona" ' zona
    Combo1.AddItem "Correo" ' correo
    Combo1.AddItem "Telefono" 'telefono
    Combo1.AddItem "Telefono1" 'telefono1
    Combo1.AddItem "Telefono2" ' telefono2
    Combo1.AddItem "Clasifica" '

    Combo1.ListIndex = 0
    carga_clasificacion
    carga_config
    'Command1_Click

    ''kenyo
    Ruta = App.path & "\ico"

End Sub

Sub inicializa()
    'clasesunat.Clear
    'clasesunat.AddItem ""
    'clasesunat.AddItem "EXONERADOS"
    'clasesunat.AddItem "AGENTE RETENCION"
    'clasesunat.AddItem "AGENTE PERCEPCION"
    'clasesunat.AddItem "CLIENTE NORMAL"
    'clasesunat.ListIndex = 0

    clasesunat = ""
    'percepcion = ""
    estadocredito = ""
    'dueno = ""
    nombre = ""
    profesion = ""
    religion = ""
    nrodepe = ""
    Trabajo = ""
    cargo = ""
    hobbie = ""
    civil = ""
    tipovive = ""

    Barras = ""
    'ruc = ""
    'dni = ""
    especial = ""
    clasifica = ""
    tipoclie = ""
    zona = ""
    fechalta = ""
    referencias = ""
    referencia = ""
    garantia = ""
    flete = ""
    moneda = "S"
    descuento1 = ""
    credito = ""
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

    '''' 19/07/2018 Campo Provincia en Cliente
    provincia = ""
    '''' 19/07/2018 Campo Provincia en Cliente

    telefono = ""
    telefono1 = ""
    telefono2 = ""
    correo = ""
    estado = "A"

End Sub

Sub pone_registro()
    'clasesunat.Clear
    'clasesunat.AddItem Trim("" & txcliem.Fields("clasesunat"))
    'clasesunat.AddItem "EXONERADOS"
    'clasesunat.AddItem "AGENTE RETENCION"
    'clasesunat.AddItem "AGENTE PERCEPCION"
    'clasesunat.AddItem "CLIENTE NORMAL"
    'clasesunat.ListIndex = 0

    clasesunat = Trim("" & txcliem.Fields("clasesunat"))

    cliente = Trim("" & txcliem.Fields("CODIGO"))
    nombre = Trim("" & txcliem.Fields("NOMBRE"))
    tipo = Trim("" & txcliem.Fields("tipo"))

    profesion = Trim("" & txcliem.Fields("profesion"))
    religion = Trim("" & txcliem.Fields("religion"))
    nrodepe = Trim("" & txcliem.Fields("nrodepe"))
    Trabajo = Trim("" & txcliem.Fields("trabajo"))
    cargo = Trim("" & txcliem.Fields("cargo"))
    hobbie = Trim("" & txcliem.Fields("hobbie"))
    civil = Trim("" & txcliem.Fields("civil"))
    tipovive = Trim("" & txcliem.Fields("tipovive"))

    Barras = Trim("" & txcliem.Fields("barras"))
    'ruc = Trim("" & txcliem.Fields("ruc"))
    'dni = Trim("" & txcliem.Fields("dni"))
    especial = Trim("" & txcliem.Fields("especial"))
    clasifica = Trim("" & txcliem.Fields("clasifica"))
    tipoclie = Trim("" & txcliem.Fields("tipoclie"))

    zona = Trim("" & txcliem.Fields("zona"))
    lunes.Value = Val("" & txcliem.Fields("lunes"))
    martes.Value = Val("" & txcliem.Fields("martes"))
    miercoles.Value = Val("" & txcliem.Fields("miercoles"))
    jueves.Value = Val("" & txcliem.Fields("jueves"))
    viernes.Value = Val("" & txcliem.Fields("viernes"))
    sabado.Value = Val("" & txcliem.Fields("sabado"))
    domingo.Value = Val("" & txcliem.Fields("domingo"))
    fechalta = Trim("" & txcliem.Fields("fechanac"))
    referencias = Trim("" & txcliem.Fields("observa"))
    referencia = Trim("" & txcliem.Fields("referencia"))
    garantia = Trim("" & txcliem.Fields("garantia"))
    flete = Trim("" & txcliem.Fields("flete"))
    moneda = Trim("" & txcliem.Fields("moneda"))
    descuento1 = Trim("" & txcliem.Fields("descuento1"))
    credito = Trim("" & txcliem.Fields("credito"))
    vendedor = Trim("" & txcliem.Fields("vendedor"))
    descuento = Trim("" & txcliem.Fields("descuento"))
    diapago = Trim("" & txcliem.Fields("diapago"))
    fpago = Trim("" & txcliem.Fields("fpago"))
    cuenta = Trim("" & txcliem.Fields("cuenta"))

    cliente = Trim("" & txcliem.Fields("codigo"))
    codigo1 = Trim("" & txcliem.Fields("codigo1"))
    nombre = Trim("" & txcliem.Fields("nombre"))
    nombrec = Trim("" & txcliem.Fields("nombrec"))
    contacto = Trim("" & txcliem.Fields("contacto"))
    direccion = Trim("" & txcliem.Fields("direccion"))
    dpto = Trim("" & txcliem.Fields("dpto"))
    distrito = Trim("" & txcliem.Fields("distrito"))

    '''' 19/07/2018 Campo Provincia en Cliente
    provincia = Trim("" & txcliem.Fields("provincia"))
    '''' 19/07/2018 Campo Provincia en Cliente

    telefono = Trim("" & txcliem.Fields("telefono"))
    telefono1 = Trim("" & txcliem.Fields("telefono1"))
    telefono2 = Trim("" & txcliem.Fields("telefono2"))
    correo = Trim("" & txcliem.Fields("correo"))
    estado = Trim("" & txcliem.Fields("estado"))
    estadocredito = Trim("" & txcliem.Fields("estadocredito"))

End Sub

Sub grabando()
    'txcliem.Fields("CODIGO") = Trim(cliente)
    txcliem.Fields("clasesunat") = Trim(clasesunat)
    'txcliem.Fields("percepcion") = Val(percepcion)

    txcliem.Fields("estadocredito") = Trim(estadocredito)
    txcliem.Fields("tipo") = Trim(tipo)
    txcliem.Fields("NOMBRE") = Trim(nombre)
    txcliem.Fields("lunes") = lunes.Value
    txcliem.Fields("martes") = martes.Value
    txcliem.Fields("miercoles") = miercoles.Value
    txcliem.Fields("jueves") = jueves.Value
    txcliem.Fields("viernes") = viernes.Value
    txcliem.Fields("sabado") = sabado.Value
    txcliem.Fields("domingo") = domingo.Value
    txcliem.Fields("flete") = Val(flete)
    txcliem.Fields("REFERENCIA") = referencia
    txcliem.Fields("GARANTIA") = garantia
    txcliem.Fields("observa") = referencias
    txcliem.Fields("tipoclie") = tipoclie
    txcliem.Fields("especial") = especial
    txcliem.Fields("clasifica") = clasifica

    If Len(fechalta) = 0 Then
        txcliem.Fields("fechanac") = Format(Now, "dd/mm/yyyy")
    Else

        If IsDate(fechalta) Then
            txcliem.Fields("fechanac") = fechalta

        End If

    End If

    txcliem.Fields("moneda") = moneda
    txcliem.Fields("vendedor") = vendedor
    txcliem.Fields("descuento1") = Val(descuento1)
    txcliem.Fields("credito") = Val(credito)
    txcliem.Fields("barras") = Barras
    'txcliem.Fields("dni") = dni
    'txcliem.Fields("ruc") = ruc
    txcliem.Fields("codigo") = cliente
    txcliem.Fields("codigo1") = codigo1
    txcliem.Fields("nombre") = nombre
    txcliem.Fields("nombrec") = nombrec
    txcliem.Fields("contacto") = contacto
    txcliem.Fields("direccion") = direccion
    txcliem.Fields("dpto") = dpto
    txcliem.Fields("distrito") = distrito
 
    '''' 19/07/2018 Campo Provincia en Cliente
    txcliem.Fields("provincia") = provincia
    '''' 19/07/2018 Campo Provincia en Cliente
 
    txcliem.Fields("zona") = zona
    txcliem.Fields("telefono") = telefono
    txcliem.Fields("telefono1") = telefono1
    txcliem.Fields("telefono2") = telefono2
    txcliem.Fields("correo") = correo
    txcliem.Fields("estado") = estado
    txcliem.Fields("descuento") = Val(descuento)
    txcliem.Fields("diapago") = diapago
    txcliem.Fields("fpago") = fpago
    txcliem.Fields("cuenta") = cuenta
 
    txcliem.Fields("profesion") = profesion
    txcliem.Fields("trabajo") = Trabajo
    txcliem.Fields("religion") = religion
    txcliem.Fields("nrodepe") = nrodepe
    txcliem.Fields("cargo") = cargo
    txcliem.Fields("hobbie") = hobbie
    txcliem.Fields("civil") = civil
    txcliem.Fields("tipovive") = tipovive
 
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
        If Len(cliente) = 0 Then
            cliente.SetFocus
            Exit Function

        End If

        rbusca.Open "select codigo from " & DBPROV & " where codigo='" & cliente & "'", cn, adOpenStatic, adLockOptimistic

        If rbusca.RecordCount > 0 Then
            rbusca.Close
            MsgBox "Ya existe cliente ", 48, "Aviso"
            Exit Function

        End If

        txcliem.AddNew
        txcliem.Fields("CODIGO") = cliente
        grabando
        txcliem.Update
        dlo132_Click
        Exit Function

    End If

    If Frame2.Caption = "Modifica" Then
        txcliem.Fields("CODIGO") = cliente
        grabando
        txcliem.Update
        dlo132_Click
        Exit Function

    End If

End Function

Function valida()

    Dim mytablex As New ADODB.Recordset

    If Len(cliente) = 0 Then
        cliente.SetFocus
        Exit Function

    End If

    If tipo <> "J" And tipo <> "X" And tipo <> "N" And tipo <> "D" And tipo <> "O" And tipo <> "P" Then
        tipo.SetFocus
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

    'MsgBox cliente
    If Len(codigo1) > 0 Then
        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open "select codigo1,Codigo from " & DBPROV & " where codigo1='" & codigo1 & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then

            'If "" & mytablex.Fields("codigo1") <> cliente Then
            '   MsgBox "Codigo1 ya usado en " + "" & mytablex.Fields("codigo")
            '   mytablex.Close
            '   Exit Function
            'End If
        End If

        mytablex.Close

    End If

    'If Len(ruc) > 0 And Len(ruc) < 11 Then
    '   MsgBox "Ruc no Valido "
    '   Exit Function
    'End If
    'If Len(ruc) > 0 Then
    'If mytablex.State = 1 Then mytablex.Close
    'mytablex.Open "select codigo,ruc from clientes where ruc='" & ruc & "'", cn, adOpenStatic, adLockOptimistic
    'If mytablex.RecordCount > 0 Then
    '   If "" & mytablex.Fields("codigo") <> cliente Then
    '      MsgBox "Ruc ya usado en " + mytablex.Fields("codigo")
    '      mytablex.Close
    '      Exit Function
    '   End If
    'End If
    'mytablex.Close
    'End If

    'If Len(dni) > 0 Then
    'If mytablex.State = 1 Then mytablex.Close
    'mytablex.Open "select codigo,dni from clientes where Dni='" & dni & "'", cn, adOpenStatic, adLockOptimistic
    'If mytablex.RecordCount > 0 Then
    '   If "" & mytablex.Fields("codigo") <> cliente Then
    '      MsgBox "Dni barras ya usado en " + mytablex.Fields("codigo")
    '      mytablex.Close
    '      Exit Function
    '   End If
    'End If
    'mytablex.Close
    'End If

    If Len(Barras) > 0 Then
        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open "select codigo,barras from " & DBPROV & "  where barras='" & Barras & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            If "" & mytablex.Fields("codigo") <> cliente Then
                MsgBox "Codigo barras ya usado en " + mytablex.Fields("codigo")
                mytablex.Close
                Exit Function

            End If

        End If

        mytablex.Close

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
        fdlo893.Enabled = True
            
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
        fdlo893.Enabled = False
            
    End If
      
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
        Combo2.Clear
        Combo2.AddItem "Descripcio"
        Combo2.AddItem "Clasifica"
        Combo2.ListIndex = 0
        opcion1 = "3"
        Frame3.Visible = True
        Frame3.Enabled = True
        cadena = ""
        'cadena.SetFocus
        Command3_Click
        Exit Sub

    End If

End Sub

Private Sub codigo1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    nombre.SetFocus

End Sub

Private Sub codigo1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        tipo.SetFocus
        Exit Sub

    End If

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
        Combo2.Clear
        Combo2.AddItem "Descripcio"
        Combo2.AddItem "Fpago"
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
        Combo2.Clear
        Combo2.AddItem "Nombre"
        Combo2.AddItem "Codigo"
        Combo2.ListIndex = 0
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

Private Sub Label18_Click()
    tctaprov.codigo = cliente
    tctaprov.Show 1

End Sub

Private Sub Label45_Click()

    On Error GoTo cmd568_err

    cn.Execute ("delete from despacho where codigo='" & cliente & "' and direccion='" & "" & dbgrid6.columns("direccion") & "'")
    consulta_direcciones
    Exit Sub
cmd568_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Label46_Click()

    Dim mytablex As New ADODB.Recordset

    If Len(direcciona) = 0 Then Exit Sub
    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from despacho where codigo='" & cliente & "' and direccion='" & direcciona & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.AddNew
        mytablex.Fields("direccion") = direcciona
        mytablex.Fields("codigo") = cliente
        mytablex.Update
    Else
        direcciona.SetFocus
      
    End If

    mytablex.Close
    consulta_direcciones

End Sub

Private Sub Label5_Click()

    Label46.Enabled = True
    Label45.Enabled = True

    If Frame2.Caption = "Nuevo" Then
        Label46.Enabled = False
        Label45.Enabled = False

    End If

    consulta_direcciones
    Frame4.Visible = True
    Frame4.Enabled = True
    direcciona = ""
    Frame4.Visible = True
    profesion.SetFocus

End Sub

Sub consulta_direcciones()

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select Direccion,Codigo from despacho where codigo='" & cliente & "'", cn, adOpenStatic, adLockOptimistic
    Set dbgrid6.DataSource = mytablex
    dbgrid6.columns(0).Width = 4000
    dbgrid6.columns(1).Width = 1000

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
        codigo1.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        Combo2.Clear
        Combo2.AddItem "Nombre"
        Combo2.ListIndex = 0
        opcion1 = "300"
        Frame3.Visible = True
        Frame3.Enabled = True
        cadena = ""

        If Len(nombre) > 0 Then
            cadena = "%" & nombre & "%"

        End If

        cadena.SetFocus
        'Command3_Click
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
        Combo2.Clear
        Combo2.AddItem "Nombre"
        Combo2.AddItem "Codigo"
        Combo2.ListIndex = 0
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

Private Sub tipo_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    codigo1.SetFocus

End Sub

Private Sub tipo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        If cliente.Enabled = True Then
            cliente.SetFocus

        End If

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
        Combo2.Clear
        Combo2.AddItem "Descripcio"
        Combo2.AddItem "Familia"
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
        Combo2.Clear
        Combo2.AddItem "Nombre"
        Combo2.AddItem "Codigo"
        Combo2.ListIndex = 0
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
        Combo2.Clear
        Combo2.AddItem "Descripcio"
        Combo2.AddItem "Zona"
        Combo2.ListIndex = 0
        opcion1 = "4"
        Frame3.Visible = True
        Frame3.Enabled = True
        cadena = ""
        cadena.SetFocus
        Command3_Click
        Exit Sub

    End If

End Sub

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

End Sub

Sub carga_clasificacion()

    Dim mytablex As New ADODB.Recordset

    clasesunat1.Clear
    clasesunat1.AddItem "%"
    mytablex.Open "select * from clasesunat ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        clasesunat1.AddItem "" & mytablex.Fields("clasesunat") & "|" & "" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    mytablex.Close
    clasesunat1.ListIndex = 0

End Sub

Sub carga_config()

    Dim mytablex As New ADODB.Recordset

    perfil.Clear
    perfil.AddItem ""
    mytablex.Open "select * from correos ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        perfil.AddItem Trim("" & mytablex.Fields("Descripcio")) & "|" & Trim("" & mytablex.Fields("cosms"))
        mytablex.MoveNext
    Loop
    mytablex.Close
    perfil.ListIndex = 0

End Sub
