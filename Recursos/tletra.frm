VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tletra 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tabla de Letras"
   ClientHeight    =   8880
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   14745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   14745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
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
      Height          =   7455
      Left            =   0
      TabIndex        =   93
      Top             =   600
      Visible         =   0   'False
      Width           =   11415
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   480
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   6960
         Visible         =   0   'False
         Width           =   1140
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
         Left            =   8160
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   240
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   96
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
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin MSDataGridLib.DataGrid dbgrid1 
         Height          =   5175
         Left            =   240
         TabIndex        =   94
         Top             =   1080
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   9128
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
      TabIndex        =   16
      Top             =   5280
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Caption         =   "Renovacion de Letra"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   6255
      Left            =   8640
      TabIndex        =   54
      Top             =   6240
      Visible         =   0   'False
      Width           =   10575
      Begin VB.CommandButton Command4 
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
         Left            =   8040
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tletra.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Borrar registro"
         Top             =   3840
         Width           =   735
      End
      Begin VB.CommandButton Command3 
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
         Left            =   8040
         Picture         =   "tletra.frx":1212
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Nuevo registro"
         Top             =   3120
         Width           =   735
      End
      Begin VB.CommandButton Command2 
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
         Left            =   8040
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tletra.frx":2424
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Grabar registro"
         Top             =   4560
         Width           =   735
      End
      Begin VB.TextBox otrosf 
         BeginProperty Font 
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
         MaxLength       =   10
         TabIndex        =   63
         Top             =   4560
         Width           =   1455
      End
      Begin VB.TextBox protestof 
         BeginProperty Font 
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
         MaxLength       =   10
         TabIndex        =   62
         Top             =   4200
         Width           =   1455
      End
      Begin VB.TextBox interes2f 
         BeginProperty Font 
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
         MaxLength       =   10
         TabIndex        =   61
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox interes1f 
         BeginProperty Font 
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
         MaxLength       =   10
         TabIndex        =   60
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox importef 
         BeginProperty Font 
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
         MaxLength       =   10
         TabIndex        =   59
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox fechaff 
         BeginProperty Font 
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
         TabIndex        =   58
         Top             =   3840
         Width           =   2175
      End
      Begin VB.TextBox fechaif 
         BeginProperty Font 
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
         TabIndex        =   57
         Top             =   3480
         Width           =   2175
      End
      Begin VB.TextBox letraf 
         BeginProperty Font 
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
         MaxLength       =   11
         TabIndex        =   56
         Top             =   3120
         Width           =   2175
      End
      Begin VB.TextBox amortizaf 
         BeginProperty Font 
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
         TabIndex        =   55
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label giradora 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
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
         TabIndex        =   91
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label40 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Girador"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   90
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label importea 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
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
         TabIndex        =   89
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label38 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   88
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label ntotal 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
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
         TabIndex        =   87
         Top             =   5040
         Width           =   1455
      End
      Begin VB.Label Label36 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NuevoTotal"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   86
         Top             =   5040
         Width           =   1575
      End
      Begin VB.Label Label35 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Otros"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   85
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label34 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Protesto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   84
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label Label33 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Interes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   83
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label32 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Interes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   82
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label31 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   81
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label30 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha vencimiento"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   80
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label Label29 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Emision"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   79
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label Label28 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Letra Renovada"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   78
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amortiza"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   77
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label paridada 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
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
         TabIndex        =   76
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Paridad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   75
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label monedaa 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
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
         TabIndex        =   74
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label23 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   73
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label fechafa 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
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
         TabIndex        =   72
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label19 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F.Vencimiento"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   71
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label fechaia 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
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
         TabIndex        =   70
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label17 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F.Emision"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   69
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label letraa 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
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
         TabIndex        =   68
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label16 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Letra Anterior"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   67
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   675
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   14685
      TabIndex        =   50
      Top             =   0
      Width           =   14745
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
         Left            =   720
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tletra.frx":3636
         Style           =   1  'Graphical
         TabIndex        =   53
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
         Left            =   3720
         Picture         =   "tletra.frx":4848
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Visible         =   0   'False
         Width           =   735
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
         Left            =   0
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tletra.frx":5A5A
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Grabar registro"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.TextBox negociado 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      MaxLength       =   11
      TabIndex        =   13
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox ochodia 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      MaxLength       =   10
      TabIndex        =   12
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox nrounico 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      MaxLength       =   11
      TabIndex        =   11
      Top             =   2880
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
      Left            =   5760
      MaxLength       =   6
      TabIndex        =   9
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox abono 
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
      Left            =   5760
      MaxLength       =   10
      TabIndex        =   41
      Top             =   7560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox saldo 
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
      Left            =   5760
      MaxLength       =   10
      TabIndex        =   40
      Top             =   7920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox fechavp 
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
      Left            =   8400
      MaxLength       =   10
      TabIndex        =   38
      Top             =   7080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox estadop 
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
      Left            =   5760
      MaxLength       =   1
      TabIndex        =   36
      Top             =   7080
      Visible         =   0   'False
      Width           =   375
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
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   15
      Top             =   4920
      Width           =   5055
   End
   Begin VB.TextBox refactura 
      BeginProperty Font 
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
      TabIndex        =   14
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox agencia 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      MaxLength       =   11
      TabIndex        =   10
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox banco 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      MaxLength       =   6
      TabIndex        =   8
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox girador 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      MaxLength       =   11
      TabIndex        =   7
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox aceptante 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      MaxLength       =   11
      TabIndex        =   6
      Top             =   960
      Width           =   1575
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
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   5
      Top             =   3000
      Width           =   1575
   End
   Begin VB.ComboBox moneda 
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
      TabIndex        =   4
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox importe 
      BeginProperty Font 
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
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox fechaf 
      BeginProperty Font 
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
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox fechai 
      BeginProperty Font 
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
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox estador 
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
      Left            =   5760
      MaxLength       =   1
      TabIndex        =   17
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
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
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label negociadon 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7440
      TabIndex        =   98
      Top             =   3600
      Width           =   6615
   End
   Begin VB.Label Label41 
      BackColor       =   &H00C0C0C0&
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
      TabIndex        =   92
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label bandera 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
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
      Height          =   195
      Left            =   6000
      TabIndex        =   49
      Top             =   120
      Width           =   75
   End
   Begin VB.Label Label37 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Negociado"
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
      TabIndex        =   48
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Octavo Dia"
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
      TabIndex        =   47
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nro.Unico"
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
      TabIndex        =   46
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label nseccion 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7440
      TabIndex        =   45
      Top             =   2160
      Width           =   4095
   End
   Begin VB.Label acu 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7920
      TabIndex        =   44
      Top             =   4800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Abono"
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
      TabIndex        =   43
      Top             =   7560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo"
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
      TabIndex        =   42
      Top             =   7920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Vencimiento"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   39
      Top             =   7080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Protesto"
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
      TabIndex        =   37
      Top             =   7080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label nombreg 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7440
      TabIndex        =   35
      Top             =   1320
      Width           =   6615
   End
   Begin VB.Label nombrea 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7440
      TabIndex        =   34
      Top             =   960
      Width           =   6615
   End
   Begin VB.Label tipox 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5760
      TabIndex        =   33
      Top             =   960
      Width           =   1605
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ReferenciasFacturas"
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
      TabIndex        =   32
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Observacion"
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
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Letra,Factura "
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
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Agencia"
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
      TabIndex        =   29
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Banco"
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
      TabIndex        =   28
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Girador"
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
      TabIndex        =   27
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
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
      Left            =   3960
      TabIndex        =   26
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aceptante"
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
      TabIndex        =   25
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "T.Cambio"
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
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
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
      TabIndex        =   23
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Importe"
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
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Vencimiento"
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
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Emision"
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
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Renovado"
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
      TabIndex        =   19
      Top             =   6720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nro.Letra"
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
      Top             =   960
      Width           =   2175
   End
   Begin VB.Menu ajdu1 
      Caption         =   "&Add"
      Visible         =   0   'False
   End
   Begin VB.Menu grba1 
      Caption         =   "&Grabar"
   End
   Begin VB.Menu reloreno 
      Caption         =   "&Renovar"
      Visible         =   0   'False
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tletra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xnameclie As String

Dim xcuentaco As String

Private Sub aceptante_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If Len(aceptante) = 0 Then Exit Sub
    found = busca_codigo("" & aceptante, 0)

    If found = 0 Then
        MsgBox "No existe aceptante", 48, "Aviso"
        Exit Sub

    End If

    girador.SetFocus

End Sub

Private Sub aceptante_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        paridad.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        consulta_aceptante

    End If

    If KeyCode = &H76 Then  'f7

        'tnclie.show 1
    End If

End Sub

Private Sub agencia_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    nrounico.SetFocus

End Sub

Private Sub agencia_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        seccion.SetFocus
        Exit Sub

    End If

End Sub

Private Sub ajdu1_Click()

    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub

    'If tipox = "RENOVACION" Then
    '   MsgBox "Se encuentra modo renovacion", 48, "Aviso"
    '   Exit Sub
    'End If
    If bandera = "MODIFICA" Then Exit Sub
    inicializa
    codigo = ""
    codigo.SetFocus

End Sub

Private Sub amortizaf_KeyPress(KeyAscii As Integer)

    Dim sdx As Double

    If KeyAscii <> 13 Then Exit Sub
    sdx = Val(importea) - Val(amortizaf)
    importef = Format(sdx, "0.00")
    letraf.SetFocus

End Sub

Private Sub banco_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If Len(banco) > 0 Then
        found = busca_banco()

        If found = 0 Then
            MsgBox "No existe Banco", 48, "Aviso"
            Exit Sub

        End If

    End If

    seccion.SetFocus

End Sub

Private Sub banco_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        girador.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        consulta_banco

    End If

    If KeyCode = &H76 Then  'f7
        tbanco.Show 1

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

Sub borra_renovacion()
    letraf = ""
    fechaif = ""
    fechaff = ""
    importef = ""
    amortizaf = ""
    interes1f = ""
    interes2f = ""
    protestof = ""
    otrosf = ""

End Sub

Private Sub cmdPrint_Click()

End Sub

Private Sub cmdSave_Click()
    grba1_Click

End Sub

Private Sub cmdSort_Click()

    Combo1.Clear
    Combo1.AddItem "Aceptante"
    Combo1.AddItem "Nombrea"
    Combo1.AddItem "Letra"
    Combo1.AddItem "estado"
    Combo1.AddItem "estadop"
    Combo1.ListIndex = 3
    Frame1.Enabled = True
    Frame1.Visible = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "1"
    Command1_Click

End Sub

Private Sub cmdSort2_Click()

End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        dlo132_Click
        Exit Sub

    End If

    If Len(codigo) = 0 Then
        codigo.SetFocus
        Exit Sub

    End If

    found = valida_registro()

    If found = 1 Then
        If bandera = "NUEVO" Then
            MsgBox "Ya existe Numero de Letra,cambie por otro", 48, "Aviso"
            codigo.SetFocus
            Exit Sub

        End If

    End If

    'found = busca_registro()
    'If found = 0 Then
    '   inicializa
    '   If tipox = "PROTESTO" Or tipox = "RENOVACION" Then
    '      Exit Sub
    '   End If
    'End If
    'If tipox = "RENOVACION" Then
    '   codigo.SetFocus
    '   Exit Sub
    'End If
    'If tipox = "PROTESTO" Then
    '   estadop.SetFocus
    '   Exit Sub
    'End If
    fechai.SetFocus

End Sub

Private Sub Codigo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1

        'cmdSort_Click
    End If

End Sub

Private Sub Command1_Click()

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    If opcion1 = "2" Or opcion1 = "3" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo from " & xnameclie
        Else
            buf = "select Nombre,Codigo from " & xnameclie & " where " & Combo1 & " like '" & buffer & "%'"

        End If
      
    End If

    If opcion1 = "22" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo from " & xnameclie
        Else
            buf = "select Nombre,Codigo from " & xnameclie & " where " & Combo1 & " like '" & buffer & "%'"

        End If
      
    End If

    If opcion1 = "1" Then
        If Len(buffer) = 0 Then
            buf = "select Estado as R,Estadop as P,Letra,Aceptante,Nombrea,Moneda as M,Saldo,Fechai,Fechaf,Seccion from " & xcuentaco
        Else
            buf = "select Estado as R,Estadop as P,Letra,Aceptante,Nombrea,Moneda as M,Saldo,Fechai,Fechaf,Seccion from " & xcuentaco & " where " & Combo1 & " like '" & buffer & "%'"

        End If
      
    End If

    If opcion1 = "4" Then
        If Len(buffer) = 0 Then
            buf = "select Descripcio,Banco from Banco "
        Else
            buf = "select Descripcio,Banco from Banco where " & Combo1 & " like '" & buffer & "%'"

        End If
      
    End If

    If opcion1 = "5" Then
        If Len(buffer) = 0 Then
            buf = "select Descripcio,Carsec from carsec "
        Else
            buf = "select Descripcio,Carsec from carsec where " & Combo1 & " like '" & buffer & "%'"

        End If
      
    End If

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = mytablex

    If opcion1 = "1" Then
        dbGrid1.columns(0).Width = 400
        dbGrid1.columns(1).Width = 400
        dbGrid1.columns(2).Width = 2000
        dbGrid1.columns(3).Width = 2000
        dbGrid1.columns(4).Width = 3000
        dbGrid1.columns(5).Width = 400
        dbGrid1.columns(6).Width = 2000
        dbGrid1.columns(7).Width = 2000
        dbGrid1.columns(8).Width = 2000

    End If

    If opcion2 = "2" Or opcion2 = "22" Or opcion1 = "3" Or opcion1 = "4" Or opcion1 = "5" Then
        dbGrid1.columns(0).Width = 4000
        dbGrid1.columns(1).Width = 2000

    End If

    If mytablex.RecordCount = 0 Then
        buffer.SetFocus
        Exit Sub

    End If

    dbGrid1.SetFocus

End Sub

Private Sub Command2_Click()

    Dim found As Integer

    suma_letra

    If Len(letraf) = 0 Then
        letraf.SetFocus
        Exit Sub

    End If

    found = busca_letraf()

    If found = 1 Then
        MsgBox "Letra Ya existe ", 48, "Aviso"
        letraf = ""
        letraf.SetFocus
        Exit Sub

    End If

    If Not IsDate(fechaif) Or Len(fechaif) <> 10 Then
        fechaif = ""
        fechaif.SetFocus
        Exit Sub

    End If

    If Not IsDate(fechaff) Or Len(fechaff) <> 10 Then
        fechaff = ""
        fechaff.SetFocus
        Exit Sub

    End If

    If Val(importef) = 0 Then
        importef.SetFocus
        Exit Sub

    End If

    found = graba_letraf()

    If found = 0 Then
        MsgBox "No se pudo grabar", 48, "Aviso"
        Exit Sub

    End If

    dlo132_Click
    codigo_KeyPress 13

End Sub

Private Sub Command3_Click()
    borra_renovacion
    amortizaf.SetFocus

End Sub

Private Sub Command4_Click()
    dlo132_Click

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = "1" Then
            codigo = Trim(dbGrid1.columns(2))
            Frame1.Visible = False
            Frame1.Enabled = False
            codigo.SetFocus
            codigo_KeyPress 13

        End If

        If opcion1 = "22" Then
            negociado = Trim(dbGrid1.columns(1))
            Frame1.Visible = False
            Frame1.Enabled = False
            negociado.SetFocus
            negociado_KeyPress 13

        End If

        If opcion1 = "2" Then
            aceptante = Trim(dbGrid1.columns(1))
            Frame1.Visible = False
            Frame1.Enabled = False
            aceptante.SetFocus
            aceptante_KeyPress 13

        End If

        If opcion1 = "3" Then
            girador = Trim(dbGrid1.columns(1))
            Frame1.Visible = False
            Frame1.Enabled = False
            girador.SetFocus
            girador_KeyPress 13

        End If

        If opcion1 = "4" Then
            banco = Trim(dbGrid1.columns(1))
            Frame1.Visible = False
            Frame1.Enabled = False
            banco.SetFocus
            banco_KeyPress 13

        End If

        If opcion1 = "5" Then
            seccion = Trim(dbGrid1.columns(1))
            Frame1.Visible = False
            Frame1.Enabled = False
            seccion.SetFocus
            seccion_KeyPress 13

        End If

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

Private Sub djuer1_Click()

End Sub

Private Sub DataGrid1_Click()

End Sub

Private Sub dlo132_Click()

    If Frame2.Visible = True Then
        Frame2.Visible = False
        codigo.SetFocus
        Exit Sub

    End If

    If Frame1.Visible = True Then
        If opcion1 = "22" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            negociado.SetFocus
            Exit Sub

        End If

        If opcion1 = "1" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            codigo.SetFocus
            Exit Sub

        End If

        If opcion1 = "2" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            aceptante.SetFocus
            Exit Sub

        End If

        If opcion1 = "3" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            girador.SetFocus
            Exit Sub

        End If

        If opcion1 = "4" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            banco.SetFocus
            Exit Sub

        End If

        If opcion1 = "5" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            seccion.SetFocus
            Exit Sub

        End If

        Exit Sub

    End If

    tletra.Hide
    Unload tletra

End Sub

Private Sub estado_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        observa.SetFocus
        Exit Sub

    End If

End Sub

Private Sub estadop_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    fechavp.SetFocus

End Sub

Private Sub fechaf_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(fechaf) = 0 Then
        fechaf = Format(Now, "dd/mm/yyyy")

    End If

    If valida_fecha("" & fechaf) = 0 Then
        fechaf = ""
        fechaf.SetFocus
        Exit Sub

    End If

    importe.SetFocus

End Sub

Private Sub fechaf_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        fechai.SetFocus
        Exit Sub

    End If

End Sub

Private Sub fechaff_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(fechaff) = 0 Then
        fechaff = Format(Now, "dd/mm/yyyy")

    End If

    If Not IsDate(fechaff) Then Exit Sub
    If Len(fechaff) <> 10 Then Exit Sub
    importef.SetFocus

End Sub

Private Sub fechaff_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        fechaif.SetFocus
        Exit Sub

    End If

End Sub

Private Sub fechai_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(fechai) = 0 Then
        fechai = Format(Now, "dd/mm/yyyy")

    End If

    If valida_fecha("" & fechai) = 0 Then
        fechai = ""
        fechai.SetFocus
        Exit Sub

    End If

    fechaf.SetFocus

End Sub

Private Sub fechai_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        If bandera = "NUEVO" Then
            codigo.SetFocus

        End If

        Exit Sub

    End If

End Sub

Private Sub fechaif_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(fechaif) = 0 Then
        fechaif = Format(Now, "dd/mm/yyyy")

    End If

    If Not IsDate(fechaif) Then Exit Sub
    If Len(fechaif) <> 10 Then Exit Sub
    fechaff.SetFocus

End Sub

Private Sub fechaif_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        letraf.SetFocus
        Exit Sub

    End If

End Sub

Private Sub fechavp_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(fechavp) = 0 Then
        fechavp = Format(Now, "dd/mm/yyyy")

    End If

End Sub

Private Sub Form_Activate()

    Dim found As Integer

    If acu = "V" Then
        xnameclie = "clientes"
        xcuentaco = "letrav"

    End If

    If acu = "C" Then
        xnameclie = "proveedo"
        xcuentaco = "letrac"

    End If

    abono.Enabled = True
    saldo.Enabled = False
    inicializa

    If bandera = "NUEVO" Then

    End If

    If bandera = "MODIFICA" Then
        found = busca_registro()

    End If

End Sub

Private Sub Form_Load()

    Dim mytablex As New ADODB.Recordset

    moneda.Clear
    moneda.AddItem "S"
    moneda.AddItem "D"
    moneda.ListIndex = 0

    Combo1.Clear
    Combo1.AddItem "Aceptante"
    Combo1.AddItem "letra"
    Combo1.ListIndex = 0

End Sub

Sub inicializa()
    negociadon = ""
    negociado = ""
    nrounico = ""
    ochodia = ""
    nseccion = ""
    abono = ""
    saldo = ""

    estador = "0"
    estadop = "0"
    fechavp = ""
    fechai = ""
    fechaf = ""
    importe = ""
    moneda.ListIndex = 0
    paridad = ""
    aceptante = ""
    girador = ""
    banco = ""
    seccion = ""
    agencia = ""
    refactura = ""
    observa = ""
    nombrea = ""
    nombreg = ""
    estado = "0"

End Sub

Function borra_registro()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from " & xcuentaco & " where letra='" & codigo & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If MsgBox("Desea Borra el registro", 1, "Aviso") = "1" Then
            mytablex.Delete
            borra_registro = 1

        End If

    End If

    mytablex.Close

End Function

Function busca_registro()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from " & xcuentaco & " where letra='" & codigo & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        pone_registro mytablex
        busca_registro = 1

    End If

    mytablex.Close
 
End Function

Function valida_registro()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from " & xcuentaco & " where letra='" & codigo & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        valida_registro = 1

    End If

    mytablex.Close

End Function

Sub pone_registro(mytablex As ADODB.Recordset)

    Dim found As Integer

    negociado = "" & mytablex.Fields("negociado")
    nrounico = "" & mytablex.Fields("nrounico")
    ochodia = "" & mytablex.Fields("ochodia")
    codigo = "" & mytablex.Fields("letra")
    fechai = "" & mytablex.Fields("fechai")
    fechaf = "" & mytablex.Fields("fechaf")
    importe = "" & mytablex.Fields("importe")
    moneda.ListIndex = 0

    If "" & mytablex.Fields("moneda") = "D" Then
        moneda.ListIndex = 1

    End If

    paridad = "" & mytablex.Fields("paridad")
    aceptante = "" & mytablex.Fields("aceptante")
    girador = "" & mytablex.Fields("girador")
    banco = "" & mytablex.Fields("banco")
    seccion = "" & mytablex.Fields("seccion")
    agencia = "" & mytablex.Fields("agencia")
    refactura = "" & mytablex.Fields("refactura")
    observa = "" & mytablex.Fields("observa")
    'estado = "" & mytablex.Fields("estado")
    estadop = "" & mytablex.Fields("estadop")
    estador = "" & mytablex.Fields("estador")
    estado = "" & mytablex.Fields("estado")
    fechavp = "" & mytablex.Fields("fechavp")
    found = busca_codigo("" & aceptante, 0)
    found = busca_codigo("" & girador, 1)
    found = busca_codigo("" & negociado, 3)
    abono = "" & mytablex.Fields("abono")
    saldo = "" & mytablex.Fields("saldo")
    found = busca_seccion()

End Sub

Sub grabando(mytablex As ADODB.Recordset)

    Dim sdx As Double

    sdx = Val(importe) - Val(abono)
    mytablex.Fields("abono") = Val(abono)
    mytablex.Fields("saldo") = Format(sdx, "0.00")
    mytablex.Fields("nrounico") = nrounico
    mytablex.Fields("ochodia") = ochodia
    mytablex.Fields("letra") = codigo
    mytablex.Fields("negociado") = negociado
    mytablex.Fields("fechai") = Format(fechai, "dd/mm/yyyy")
    mytablex.Fields("fechaf") = Format(fechaf, "dd/mm/yyyy")
    mytablex.Fields("importe") = Val(importe)
    mytablex.Fields("moneda") = moneda
    mytablex.Fields("paridad") = Val(paridad)
    mytablex.Fields("aceptante") = aceptante
    mytablex.Fields("girador") = girador
    mytablex.Fields("banco") = banco
    mytablex.Fields("seccion") = seccion
    mytablex.Fields("agencia") = agencia
    mytablex.Fields("refactura") = refactura
    mytablex.Fields("observa") = observa
    mytablex.Fields("estadop") = estadop
    mytablex.Fields("estador") = estador
    mytablex.Fields("estado") = estado
    mytablex.Fields("nombreg") = nombreg
    mytablex.Fields("nombrea") = nombrea

End Sub

Private Sub girador_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If Len(girador) = 0 Then Exit Sub
    found = busca_codigo("" & girador, 1)

    If found = 0 Then
        MsgBox "No existe Girador", 48, "Aviso"
        Exit Sub

    End If

    banco.SetFocus

End Sub

Private Sub girador_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        aceptante.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        consulta_girador

    End If

    If KeyCode = &H76 Then  'f7

        'tnclie.show 1
    End If

End Sub

Private Sub grba1_Click()

    Dim found As Integer

    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    found = grabar()

    If found = 0 Then Exit Sub
    If bandera = "MODIFICA" Then
        fechai.SetFocus

    End If

    If bandera = "NUEVO" Then
        codigo.SetFocus

    End If

    dlo132_Click

End Sub

Private Sub importe_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    moneda.SetFocus

End Sub

Private Sub importe_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        fechaf.SetFocus
        Exit Sub

    End If

End Sub

Private Sub importef_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    interes1f.SetFocus

End Sub

Private Sub importef_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        fechaff.SetFocus
        Exit Sub

    End If

End Sub

Private Sub interes1f_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    interes2f.SetFocus

End Sub

Private Sub interes1f_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        importef.SetFocus
        Exit Sub

    End If

End Sub

Private Sub interes2f_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    protestof.SetFocus

End Sub

Private Sub interes2f_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        interes1f.SetFocus
        Exit Sub

    End If

End Sub

Private Sub Label1_Click()
    cmdSort_Click

End Sub

Function grabar()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    found = valida()

    If found = 0 Then
        MsgBox "Campos invalidos", 48, "Aviso"
        Exit Function

    End If

    mytablex.Open "select * from " & xcuentaco & " where letra='" & codigo & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.AddNew
        grabando mytablex
        mytablex.Update
        grabar = 1
    Else

        If MsgBox("Desea Reescribir?", 1, "Aviso") = 1 Then

            'mytablex.Edit
            'mytablex.Fields("local") = "01"
            If tipox = "PROTESTO" Then
                mytablex.Fields("estadop") = "1"
                mytablex.Fields("fechavp") = Format(fechavp, "dd/mm/yyyy")
            Else
                grabando mytablex

            End If

            mytablex.Update
            grabar = 1

        End If

    End If

    mytablex.Close

End Function

Function valida()

    Dim found As Integer

    Dim sdx   As Double

    sdx = Val(importe) - Val(abono)
    saldo = Format(sdx, "0.00")

    If Len(codigo) = 0 Then
        codigo.SetFocus
        Exit Function

    End If

    found = valida_registro()

    If found = 1 Then
        If bandera = "NUEVO" Then
            MsgBox "Ya existe Numero de Letra,cambie por otro", 48, "Aviso"
            codigo.SetFocus
            Exit Function

        End If

    End If

    If valida_fecha("" & fechai) = 0 Then
        fechai = ""
        fechai.SetFocus
        Exit Function

    End If

    If valida_fecha("" & fechaf) = 0 Then
        fechaf = ""
        fechaf.SetFocus
        Exit Function

    End If

    ochodia = Format(CVDate(fechaf) + 8, "dd/mm/yyyy")
  
    If Len(aceptante) = 0 Then
        aceptante.SetFocus
        Exit Function

    End If

    found = busca_codigo("" & aceptante, 0)

    If found = 0 Then
        MsgBox "No existe aceptante", 48, "Aviso"
        aceptante.SetFocus
        Exit Function

    End If

    If Len(girador) = 0 Then
        girador.SetFocus
        Exit Function

    End If

    found = busca_codigo("" & girador, 1)

    If found = 0 Then
        MsgBox "No existe girador", 48, "Aviso"
        girador.SetFocus
        Exit Function

    End If

    If Len(banco) > 0 Then
        found = busca_banco()

        If found = 0 Then
            MsgBox "No existe Banco", 48, "Aviso"
            banco = ""
            banco.SetFocus
            Exit Function

        End If

    End If

    If Len(seccion) = 0 Then
        seccion.SetFocus
        Exit Function

    End If

    found = busca_seccion()

    If found = 0 Then
        seccion = ""
        seccion.SetFocus
        Exit Function

    End If

    valida = 1

End Function

Private Sub letraf_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If Len(letraf) = 0 Then Exit Sub
    found = busca_letraf()

    If found = 1 Then
        MsgBox "Letra Ya existe ", 48, "Aviso"
        letraf.SetFocus
        Exit Sub

    End If

    fechaif.SetFocus

End Sub

Private Sub letraf_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        amortizaf.SetFocus
        Exit Sub

    End If

End Sub

Private Sub moneda_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    paridad.SetFocus

End Sub

Private Sub moneda_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        importe.SetFocus
        Exit Sub

    End If

End Sub

Private Sub negociado_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If Len(negociado) > 0 Then
        found = busca_codigo("" & negociado, 3)

        If found = 0 Then
            negociado = ""
            negociadon = ""
            MsgBox "No existe negociado", 48, "Aviso"
            Exit Sub

        End If

    End If

    refactura.SetFocus

End Sub

Private Sub negociado_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_proveedor

    End If

End Sub

Private Sub nrounico_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    ochodia.SetFocus

End Sub

Private Sub nrounico_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        agencia.SetFocus
        Exit Sub

    End If

End Sub

Private Sub observa_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub observa_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        refactura.SetFocus
        Exit Sub

    End If

End Sub

Private Sub ochodia_KeyPress(KeyAscii As Integer)

    Dim xfecha As String

    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    negociado.SetFocus

End Sub

Private Sub ochodia_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        nrounico.SetFocus
        Exit Sub

    End If

End Sub

Private Sub otrosf_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    suma_letra

End Sub

Private Sub otrosf_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        protestof.SetFocus
        Exit Sub

    End If

End Sub

Private Sub paridad_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    aceptante.SetFocus

End Sub

Private Sub paridad_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        moneda.SetFocus
        Exit Sub

    End If

End Sub

Private Sub protestof_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    otrosf.SetFocus

End Sub

Private Sub protestof_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        interes2f.SetFocus
        Exit Sub

    End If

End Sub

Private Sub refactura_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    observa.SetFocus

End Sub

Private Sub refactura_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        ochodia.SetFocus
        Exit Sub

    End If

End Sub

Private Sub reloreno_Click()
    cmdHelp_Click

End Sub

Private Sub seccion_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If Len(seccion) = 0 Then Exit Sub
    found = busca_seccion()

    If found = 0 Then
        Exit Sub

    End If

    agencia.SetFocus

End Sub

Private Sub seccion_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        banco.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        consulta_seccion

    End If

End Sub

Function busca_codigo(buf As String, sw As Integer)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from " & xnameclie & " where codigo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_codigo = 1

        If sw = 0 Then
            nombrea = "" & mytablex.Fields("nombre")

        End If

        If sw = 1 Then
            nombreg = "" & mytablex.Fields("nombre")

        End If

        If sw = 3 Then
            negociadon = "" & mytablex.Fields("nombre")

        End If

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_banco()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from banco where banco='" & banco & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_banco = 1

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_seccion()

    Dim mytablex As New ADODB.Recordset

    nseccion = ""
    mytablex.Open "select * from carsec where carsec='" & seccion & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        nseccion = "" & mytablex.Fields("descripcio")
        busca_seccion = 1

    End If

    mytablex.Close

End Function

Sub consulta_proveedor()
    Frame1.Enabled = True
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "22"
    Command1_Click

End Sub

Sub consulta_aceptante()
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    Frame1.Enabled = True
    Frame1.Visible = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "2"
    Command1_Click

End Sub

Sub consulta_girador()

    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    Frame1.Enabled = True
    Frame1.Visible = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "3"
    Command1_Click

End Sub

Sub consulta_banco()

    Combo1.Clear
    Combo1.AddItem "Descripcio"
    Combo1.AddItem "Banco"
    Combo1.ListIndex = 0
    Frame1.Enabled = True
    Frame1.Visible = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "4"
    Command1_Click

End Sub

Sub consulta_seccion()

    Combo1.Clear
    Combo1.AddItem "Descripcio"
    Combo1.AddItem "Carsec"
    Combo1.ListIndex = 0
    Frame1.Enabled = True
    Frame1.Visible = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "5"
    Command1_Click

End Sub

Function busca_letraf()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from " & xcuentaco & " where letra='" & letraf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_letraf = 1

    End If

    mytablex.Close

End Function

Sub suma_letra()

    Dim sdx As Double

    sdx = Val(importef) + Val(interes1f) + Val(interes2f) + Val(protestof) + Val(otrosf)
    ntotal = Format(sdx, "0.00")

End Sub

Function graba_letraf()

    Dim sw       As Integer

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from " & xcuentaco & " where letra='" & letraf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        mytablex.Close
        Exit Function

    End If

    mytablex.AddNew
    mytablex.Fields("local") = "01"
    mytablex.Fields("letra") = letraf
    mytablex.Fields("letraant") = codigo
    mytablex.Fields("fechai") = Format(fechaif, "dd/mm/yyyy")
    mytablex.Fields("fechaf") = Format(fechaff, "dd/mm/yyyy")
    mytablex.Fields("moneda") = moneda
    mytablex.Fields("paridad") = Val(paridad)
    mytablex.Fields("aceptante") = aceptante
    mytablex.Fields("girador") = girador
    mytablex.Fields("banco") = banco
    mytablex.Fields("seccion") = seccion
    mytablex.Fields("agencia") = agencia
    mytablex.Fields("refactura") = refactura
    mytablex.Fields("observa") = observa
    mytablex.Fields("nombrea") = nombrea
    mytablex.Fields("nombreg") = nombreg
    mytablex.Fields("estadop") = estadop
    mytablex.Fields("estador") = estador

    If Len(estado) = 0 Then
        estado = "0"

    End If

    mytablex.Fields("estado") = estado
    mytablex.Fields("amortiza") = Val(amortizaf)
    mytablex.Fields("interes1") = Val(interes1f)
    mytablex.Fields("interes2") = Val(interes2f)
    mytablex.Fields("protesto") = Val(protestof)
    mytablex.Fields("otros") = Val(otrosf)
    mytablex.Fields("importe") = Val(ntotal)
    mytablex.Fields("abono") = 0
    mytablex.Fields("saldo") = Val(ntotal)
    mytablex.Update
    mytablex.Close
    mytablex.Open "select * from " & xcuentaco & " where letra='" & codigo & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        mytablex.Fields("letrarenov") = letraf
        mytablex.Fields("estado") = "1"
        mytablex.Update

    End If

    graba_letraf = 1
    mytablex.Close

End Function

Sub valida_campos()

End Sub

