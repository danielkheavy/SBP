VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tnprov 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Tabla de Proveedores"
   ClientHeight    =   10830
   ClientLeft      =   165
   ClientTop       =   135
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10830
   ScaleWidth      =   15240
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Otros"
      Height          =   6855
      Left            =   1920
      TabIndex        =   92
      Top             =   120
      Visible         =   0   'False
      Width           =   10935
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Direcciones de Despacho"
         Height          =   3255
         Left            =   240
         TabIndex        =   111
         Top             =   3360
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
            TabIndex        =   112
            Top             =   240
            Width           =   7575
         End
         Begin MSDataGridLib.DataGrid dbgrid6 
            Height          =   2415
            Left            =   120
            TabIndex        =   113
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
            Left            =   9600
            TabIndex        =   116
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
            Left            =   9600
            TabIndex        =   115
            Top             =   600
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
            TabIndex        =   114
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Sa&lir"
         Height          =   615
         Left            =   9720
         Picture         =   "TNPROV.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   360
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
         TabIndex        =   100
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
         TabIndex        =   99
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
         TabIndex        =   98
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
         TabIndex        =   97
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
         TabIndex        =   96
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
         TabIndex        =   95
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
         TabIndex        =   94
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
         TabIndex        =   93
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
         TabIndex        =   109
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
         TabIndex        =   108
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
         TabIndex        =   107
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
         TabIndex        =   106
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
         TabIndex        =   105
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
         TabIndex        =   104
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
         TabIndex        =   103
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label3 
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
         TabIndex        =   102
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consulta"
      Height          =   8895
      Left            =   0
      TabIndex        =   87
      Top             =   0
      Visible         =   0   'False
      Width           =   13695
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
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   600
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   600
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
         Left            =   8280
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid dbGrid3 
         Height          =   7575
         Left            =   240
         TabIndex        =   91
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Frame2"
      Height          =   8655
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   12855
      Begin VB.TextBox cliente 
         BeginProperty Font 
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
         TabIndex        =   53
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
         TabIndex        =   52
         Top             =   600
         Width           =   6015
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   11280
         Picture         =   "TNPROV.frx":08DA
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Imprimir todo"
         Top             =   2040
         Width           =   1470
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   11280
         Picture         =   "TNPROV.frx":11A4
         Style           =   1  'Graphical
         TabIndex        =   50
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
         TabIndex        =   49
         Top             =   2880
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
         Left            =   2280
         MaxLength       =   11
         TabIndex        =   48
         Top             =   2160
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
         TabIndex        =   47
         Top             =   1800
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
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   7440
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
         TabIndex        =   45
         Top             =   7800
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
         TabIndex        =   44
         Top             =   6360
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
         TabIndex        =   43
         Top             =   5040
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
         TabIndex        =   42
         Top             =   3240
         Width           =   1935
      End
      Begin VB.CheckBox viernes 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Viernes"
         Height          =   375
         Left            =   9000
         TabIndex        =   41
         Top             =   5520
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
         TabIndex        =   40
         Top             =   2520
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
         TabIndex        =   39
         Top             =   960
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
         TabIndex        =   38
         Top             =   1320
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
         Top             =   4680
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
         TabIndex        =   34
         Top             =   5880
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
         TabIndex        =   33
         Top             =   5880
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
         TabIndex        =   32
         Top             =   5880
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
         TabIndex        =   31
         Top             =   5400
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         MaxLength       =   11
         TabIndex        =   26
         Top             =   6360
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
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CheckBox lunes 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Lunes"
         Height          =   375
         Left            =   7080
         TabIndex        =   23
         Top             =   5160
         Width           =   735
      End
      Begin VB.CheckBox martes 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Martes"
         Height          =   375
         Left            =   7920
         TabIndex        =   22
         Top             =   5160
         Width           =   855
      End
      Begin VB.CheckBox miercoles 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Miercoles"
         Height          =   375
         Left            =   8760
         TabIndex        =   21
         Top             =   5160
         Width           =   975
      End
      Begin VB.CheckBox jueves 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Jueves"
         Height          =   375
         Left            =   9840
         TabIndex        =   20
         Top             =   5160
         Width           =   855
      End
      Begin VB.CheckBox sabado 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Sabado"
         Height          =   375
         Left            =   7080
         TabIndex        =   19
         Top             =   5520
         Width           =   855
      End
      Begin VB.CheckBox domingo 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Domingo"
         Height          =   375
         Left            =   7920
         TabIndex        =   18
         Top             =   5520
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
         TabIndex        =   17
         Top             =   4440
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
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   7440
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
         TabIndex        =   14
         Top             =   6720
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
         TabIndex        =   13
         Top             =   7080
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
         TabIndex        =   12
         Top             =   3960
         Width           =   4695
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
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
         Left            =   11280
         TabIndex        =   110
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
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
         TabIndex        =   86
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF00&
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
         Height          =   375
         Left            =   120
         TabIndex        =   85
         Top             =   600
         Width           =   2175
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
         Left            =   120
         TabIndex        =   84
         Top             =   2880
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
         Left            =   120
         TabIndex        =   83
         Top             =   2160
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
         TabIndex        =   82
         Top             =   1800
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
         TabIndex        =   81
         Top             =   7440
         Width           =   2175
      End
      Begin VB.Label Label4 
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
         TabIndex        =   80
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label6 
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
         TabIndex        =   79
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label7 
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
         TabIndex        =   78
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label8 
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
         TabIndex        =   77
         Top             =   3600
         Width           =   2175
      End
      Begin VB.Label Label9 
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
         TabIndex        =   76
         Top             =   4320
         Width           =   2175
      End
      Begin VB.Label Label10 
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
         TabIndex        =   75
         Top             =   4680
         Width           =   2175
      End
      Begin VB.Label Label11 
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
         TabIndex        =   74
         Top             =   5880
         Width           =   2175
      End
      Begin VB.Label Label12 
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
         TabIndex        =   73
         Top             =   5400
         Width           =   2175
      End
      Begin VB.Label Label13 
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
         TabIndex        =   72
         Top             =   5040
         Width           =   2175
      End
      Begin VB.Label Label14 
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
         TabIndex        =   71
         Top             =   7800
         Width           =   2175
      End
      Begin VB.Label Label15 
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
         TabIndex        =   70
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label16 
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
         TabIndex        =   69
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label17 
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
         TabIndex        =   68
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label18 
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
         TabIndex        =   67
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label19 
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
         TabIndex        =   66
         Top             =   6360
         Width           =   2175
      End
      Begin VB.Label Label20 
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
         TabIndex        =   65
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label21 
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
         TabIndex        =   64
         Top             =   4080
         Width           =   1695
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
         TabIndex        =   63
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label Label22 
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
         TabIndex        =   62
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label23 
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
         TabIndex        =   61
         Top             =   7440
         Width           =   2175
      End
      Begin VB.Label Label24 
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
         TabIndex        =   60
         Top             =   3240
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
         TabIndex        =   59
         Top             =   6360
         Width           =   1335
      End
      Begin VB.Label Label25 
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
         TabIndex        =   58
         Top             =   6720
         Width           =   2175
      End
      Begin VB.Label Label26 
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
         TabIndex        =   57
         Top             =   7080
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
         TabIndex        =   56
         Top             =   7080
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
         TabIndex        =   55
         Top             =   6720
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
         TabIndex        =   54
         Top             =   3960
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   12435
      TabIndex        =   2
      Top             =   0
      Width           =   12495
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
         Picture         =   "TNPROV.frx":1A6E
         Style           =   1  'Graphical
         TabIndex        =   10
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
         TabIndex        =   9
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
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   120
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
         Left            =   10560
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
         Left            =   1440
         MaskColor       =   &H00E0E0E0&
         Picture         =   "TNPROV.frx":2C80
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
         Left            =   2760
         MaskColor       =   &H00E0E0E0&
         Picture         =   "TNPROV.frx":3E92
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
         Left            =   2160
         MaskColor       =   &H00E0E0E0&
         Picture         =   "TNPROV.frx":50A4
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
         Picture         =   "TNPROV.frx":62B6
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consulta"
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   12495
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   7935
         Left            =   120
         TabIndex        =   1
         Top             =   240
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
         ColumnCount     =   2
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   5940.284
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
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
   Begin VB.Menu djuer1 
      Caption         =   "&Reporte"
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tnprov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim txempre As New ADODB.Recordset
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
cliente.Enabled = False
cliente = ""
nombre.SetFocus
End Sub

Private Sub bo712_Click()
Dim buf As String
On Error GoTo cmd656_err
buf = "" & txempre.Fields("CODIGO")
If Frame2.Visible = True Then
   dbGrid1.SetFocus
   Exit Sub
End If
If MsgBox("Desea Borra " + "" & txempre.Fields("CODIGO"), 1, "Aviso") <> 1 Then
   Exit Sub
End If
txempre.Delete
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

Private Sub cmdHelp_Click()

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
If Len(cliente) = 0 Then Exit Sub
nombre.SetFocus
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
If opcion1 = "1" Then  'bodega
   If Len(buffer) = 0 Then
      cad = "SELECT * from PROVEEDO    "
   End If
   If Len(buffer) > 0 Then
      cad = "SELECT *  from PROVEEDO   where  " & Combo1 & " like '" & buffer & "%'"
   End If
   If txempre.State = 1 Then txempre.Close
   txempre.Open cad, cn, adOpenStatic, adLockOptimistic
   Set dbGrid1.DataSource = txempre
   dbGrid1.Columns(0).Width = 4000
   dbGrid1.Columns(1).Width = 2000
   If txempre.RecordCount > 0 Then
     dbGrid1.SetFocus
  End If
End If
End Sub

Private Sub Command2_Click()

End Sub


Private Sub Command3_Click()
ejecuta1 1
End Sub
Sub ejecuta1(sw As Integer)
Dim buf As String
Dim mytablex As New ADODB.Recordset
If opcion1 = "1" Then
If Len(cadena) = 0 Then
buf = "select Descripcio,tipoclie from tipoclie "
Else
buf = "select Descripcio,tipoclie from tipoclie where " & Combo2 & " like '" & cadena & "%'"
End If
End If
If opcion1 = "2" Then
If Len(cadena) = 0 Then
buf = "select Nombre,Codigo from Vendedor "
Else
buf = "select Nombre,Codigo from Vendedor where " & Combo2 & " like '" & cadena & "%'"
End If
End If
If opcion1 = "200" Or opcion1 = "201" Then
If Len(cadena) = 0 Then
   buf = "select Nombre,Codigo from PROVEEDO "
Else
   buf = "select Nombre,Codigo from PROVEEDO where " & Combo2 & " like '" & cadena & "%'"
End If

End If

If opcion1 = "300" Then
   If Len(cadena) = 0 Then
      buf = "select Nombre,Codigo from PROVEEDO "
   Else
   buf = "select Nombre,Codigo from PROVEEDO where " & Combo2 & " like '" & cadena & "%'"
End If
End If



If opcion1 = "3" Then
If Len(cadena) = 0 Then
buf = "select Descripcio,Clasifica from clasifi "
Else
buf = "select Descripcio,Clasifica from Clasifi where " & Combo2 & " like '" & cadena & "%'"
End If
End If
If opcion1 = "4" Then
If Len(cadena) = 0 Then
buf = "select Descripcio,Zona from Zona "
Else
buf = "select Descripcio,Zona from Zona where " & Combo2 & " like '" & cadena & "%'"
End If
End If
If opcion1 = "5" Then
If Len(cadena) = 0 Then
buf = "select Descripcio,Fpago from Fpago "
Else
buf = "select Descripcio,Fpago from Fpago where " & Combo2 & " like '" & cadena & "%'"
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

Private Sub Command9_Click()
Frame4.Visible = False
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub djuer1_Click()
If Frame2.Visible = True Then Exit Sub
reporgen.NAMETABLA = "proveedo"
reporgen.Show 1

End Sub

Private Sub dlo132_Click()
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

tnprov.Hide
Unload tnprov
End Sub


Private Sub f8443_Click()
Dim buf As String
On Error GoTo cmd456_err
buf = txempre.Fields("CODIGO")
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

Private Sub fjh433_Click()
Dim buf As String
On Error GoTo cmd556_err
buf = txempre.Fields("CODIGO")
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
Command1_Click
End Sub

Private Sub Form_Load()
Combo1.Clear
Combo1.AddItem "NOMBRE"
Combo1.AddItem "CODIGO"
Combo1.ListIndex = 0

End Sub
Sub inicializa()


nombre = ""

profesion = ""
religion = ""
nrodepe = ""
Trabajo = ""
cargo = ""
hobbie = ""
civil = ""
tipovive = ""


barras = ""
ruc = ""
dni = ""
especial = ""
clasifica = ""
tipoclie = ""

zona = ""
fechalta = ""
referencias = ""
referencia = ""
garantia = ""
flete = ""
moneda = ""
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
telefono = ""
telefono1 = ""
telefono2 = ""
correo = ""
estado = ""

End Sub
Sub pone_registro()
cliente = Trim("" & txempre.Fields("CODIGO"))
nombre = Trim("" & txempre.Fields("NOMBRE"))

profesion = Trim("" & txempre.Fields("profesion"))
religion = Trim("" & txempre.Fields("religion"))
nrodepe = Trim("" & txempre.Fields("nrodepe"))
Trabajo = Trim("" & txempre.Fields("trabajo"))
cargo = Trim("" & txempre.Fields("cargo"))
hobbie = Trim("" & txempre.Fields("hobbie"))
civil = Trim("" & txempre.Fields("civil"))
tipovive = Trim("" & txempre.Fields("tipovive"))


barras = Trim("" & txempre.Fields("barras"))
ruc = Trim("" & txempre.Fields("ruc"))
dni = Trim("" & txempre.Fields("dni"))
especial = Trim("" & txempre.Fields("especial"))
clasifica = Trim("" & txempre.Fields("clasifica"))
tipoclie = Trim("" & txempre.Fields("tipoclie"))

zona = Trim("" & txempre.Fields("zona"))
lunes.Value = Val("" & txempre.Fields("lunes"))
martes.Value = Val("" & txempre.Fields("martes"))
miercoles.Value = Val("" & txempre.Fields("miercoles"))
jueves.Value = Val("" & txempre.Fields("jueves"))
viernes.Value = Val("" & txempre.Fields("viernes"))
sabado.Value = Val("" & txempre.Fields("sabado"))
domingo.Value = Val("" & txempre.Fields("domingo"))
fechalta = Trim("" & txempre.Fields("fechanac"))
referencias = Trim("" & txempre.Fields("observa"))
referencia = Trim("" & txempre.Fields("referencia"))
garantia = Trim("" & txempre.Fields("garantia"))
flete = Trim("" & txempre.Fields("flete"))
moneda = Trim("" & txempre.Fields("moneda"))
descuento1 = Trim("" & txempre.Fields("descuento1"))
credito = Trim("" & txempre.Fields("credito"))
vendedor = Trim("" & txempre.Fields("vendedor"))
descuento = Trim("" & txempre.Fields("descuento"))
diapago = Trim("" & txempre.Fields("diapago"))
fpago = Trim("" & txempre.Fields("fpago"))
cuenta = Trim("" & txempre.Fields("cuenta"))

cliente = Trim("" & txempre.Fields("codigo"))
codigo1 = Trim("" & txempre.Fields("extranjeria"))
nombre = Trim("" & txempre.Fields("nombre"))
nombrec = Trim("" & txempre.Fields("nombrec"))
contacto = Trim("" & txempre.Fields("contacto"))
direccion = Trim("" & txempre.Fields("direccion"))
dpto = Trim("" & txempre.Fields("dpto"))
distrito = Trim("" & txempre.Fields("distrito"))
telefono = Trim("" & txempre.Fields("telefono"))
telefono1 = Trim("" & txempre.Fields("telefono1"))
telefono2 = Trim("" & txempre.Fields("telefono2"))
correo = Trim("" & txempre.Fields("correo"))
estado = Trim("" & txempre.Fields("estado"))

End Sub
Sub grabando()
'txempre.Fields("CODIGO") = Trim(cliente)
txempre.Fields("NOMBRE") = Trim(nombre)
txempre.Fields("lunes") = lunes.Value
 txempre.Fields("martes") = martes.Value
 txempre.Fields("miercoles") = miercoles.Value
 txempre.Fields("jueves") = jueves.Value
 txempre.Fields("viernes") = viernes.Value
 txempre.Fields("sabado") = sabado.Value
 txempre.Fields("domingo") = domingo.Value
 txempre.Fields("flete") = Val(flete)
 txempre.Fields("REFERENCIA") = referencia
 txempre.Fields("GARANTIA") = garantia
 txempre.Fields("observa") = referencias
 txempre.Fields("tipoclie") = tipoclie
 txempre.Fields("especial") = especial
 txempre.Fields("clasifica") = clasifica
If Len(fechalta) = 0 Then
    txempre.Fields("fechanac") = Format(Now, "dd/mm/yyyy")
   Else
   If IsDate(fechalta) Then
    txempre.Fields("fechanac") = fechalta
   End If
End If
 txempre.Fields("moneda") = moneda
 txempre.Fields("vendedor") = vendedor
 txempre.Fields("descuento1") = Val(descuento1)
 txempre.Fields("credito") = Val(credito)
 txempre.Fields("barras") = barras
 txempre.Fields("dni") = dni
 txempre.Fields("ruc") = ruc
 'txempre.Fields("codigo") = codigo
 txempre.Fields("extranjeria") = codigo1
 txempre.Fields("nombre") = nombre
 txempre.Fields("nombrec") = nombrec
 txempre.Fields("contacto") = contacto
 txempre.Fields("direccion") = direccion
 txempre.Fields("dpto") = dpto
 txempre.Fields("distrito") = distrito
 txempre.Fields("zona") = zona
 txempre.Fields("telefono") = telefono
 txempre.Fields("telefono1") = telefono1
 txempre.Fields("telefono2") = telefono2
 txempre.Fields("correo") = correo
 txempre.Fields("estado") = estado
 txempre.Fields("descuento") = Val(descuento)
 txempre.Fields("diapago") = diapago
 txempre.Fields("fpago") = fpago
 txempre.Fields("cuenta") = cuenta
 
 
  txempre.Fields("profesion") = profesion
  txempre.Fields("trabajo") = Trabajo
  txempre.Fields("religion") = religion
  txempre.Fields("nrodepe") = nrodepe
  txempre.Fields("cargo") = cargo
  txempre.Fields("hobbie") = hobbie
  txempre.Fields("civil") = civil
  txempre.Fields("tipovive") = tipovive
 
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
   'If Len(cliente) = 0 Then
   '   .SetFocus
   '   Exit Function
   'End If
   'rbusca.Open "select cliente from PROVEEDO where codigo='" & cliente & "'", cn, adOpenStatic, adLockOptimistic
   'If rbusca.RecordCount > 0 Then
   '   rbusca.Close
   '   MsgBox "Ya existe cliente ", 48, "Aviso"
   '   Exit Function
   'End If
   txempre.AddNew
   'txempre.Fields("CODIGO") = cliente
   grabando
   txempre.Update
   dlo132_Click
   Exit Function
End If
If Frame2.Caption = "Modifica" Then
   'txempre.Fields("CODIGO") = cliente
   grabando
   txempre.Update
   dlo132_Click
   Exit Function
End If

End Function

Function valida()
Dim mytablex As New ADODB.Recordset
'If Len(cliente) = 0 Then
'   cliente.SetFocus
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
mytablex.Open "select codigo,dni from PROVEEDO where dni='" & dni & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   If "" & mytablex.Fields("codigo") <> cliente Then
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
mytablex.Open "select codigo,ruc from PROVEEDO where ruc='" & ruc & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   If "" & mytablex.Fields("codigo") <> cliente Then
      MsgBox "Ruc ya usado en " + mytablex.Fields("codigo")
      mytablex.Close
      Exit Function
   End If
End If
mytablex.Close
End If
If Len(dni) > 0 Then
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "select codigo,dni from PROVEEDO where Dni='" & dni & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   If "" & mytablex.Fields("codigo") <> cliente Then
      MsgBox "Dni barras ya usado en " + mytablex.Fields("codigo")
      mytablex.Close
      Exit Function
   End If
End If
mytablex.Close
End If
If Len(barras) > 0 Then
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "select codigo,barras from PROVEEDO where barras='" & barras & "'", cn, adOpenStatic, adLockOptimistic
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
tipoclie.SetFocus

End Sub

Private Sub codigo1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   dni.SetFocus
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

Private Sub Label44_Click()

End Sub

Private Sub Label45_Click()
On Error GoTo cmd568_err
cn.Execute ("delete from despachop where codigo='" & cliente & "' and direccion='" & "" & dbgrid6.Columns("direccion") & "'")
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
   mytablex.Open "select * from despachop where codigo='" & cliente & "' and direccion='" & direcciona & "'", cn, adOpenStatic, adLockOptimistic
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
If tnclie.Caption = "NUEVO" Then
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
   mytablex.Open "select Direccion,Codigo from despachop where codigo='" & cliente & "'", cn, adOpenStatic, adLockOptimistic
   Set dbgrid6.DataSource = mytablex
   dbgrid6.Columns(0).Width = 4000
   dbgrid6.Columns(1).Width = 1000

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

Private Sub ruc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
dni.SetFocus
End Sub

Private Sub ruc_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   'If codigo.Enabled = True Then
   '   codigo.SetFocus
   'End If
   'Exit Sub
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

