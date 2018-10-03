VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tpactop 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Precios Pactado con Clientes"
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   12150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
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
      Height          =   8295
      Left            =   120
      TabIndex        =   108
      Top             =   120
      Visible         =   0   'False
      Width           =   11895
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
         Height          =   615
         Left            =   6120
         TabIndex        =   113
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
         TabIndex        =   112
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
         TabIndex        =   111
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Cerrar"
         Height          =   660
         Left            =   10200
         Picture         =   "tpactop.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   109
         ToolTipText     =   "Imprimir todo"
         Top             =   240
         Width           =   1470
      End
      Begin MSDataGridLib.DataGrid dbgrid1 
         Height          =   7095
         Left            =   120
         TabIndex        =   110
         Top             =   1080
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   12515
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   15
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
      Caption         =   "Frame1"
      Height          =   6495
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   9735
      Begin VB.TextBox pm10 
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
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   84
         Top             =   4920
         Width           =   855
      End
      Begin VB.TextBox pm9 
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
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   83
         Top             =   4560
         Width           =   855
      End
      Begin VB.TextBox pm8 
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
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   82
         Top             =   4200
         Width           =   855
      End
      Begin VB.TextBox pm7 
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
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   81
         Top             =   3840
         Width           =   855
      End
      Begin VB.TextBox pm6 
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
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   80
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox pm5 
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
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   79
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox pm4 
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
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   78
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox pm3 
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
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   77
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox pm2 
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
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   76
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox pm1 
         BackColor       =   &H00C0FFFF&
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
         MaxLength       =   10
         TabIndex        =   75
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox unidad1 
         BackColor       =   &H00C0FFFF&
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
         Left            =   240
         MaxLength       =   6
         TabIndex        =   74
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox factor1 
         BackColor       =   &H00C0FFFF&
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
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   73
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox pventa1 
         BackColor       =   &H00C0FFFF&
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   72
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox margen1 
         BackColor       =   &H00C0FFFF&
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
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   71
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox fechai11 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         MaxLength       =   10
         TabIndex        =   70
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox fechaf11 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         MaxLength       =   10
         TabIndex        =   69
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox fechaid 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         MaxLength       =   10
         TabIndex        =   68
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox fechafd 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         MaxLength       =   10
         TabIndex        =   67
         Top             =   4680
         Width           =   1455
      End
      Begin VB.TextBox margen11 
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
         Left            =   6720
         MaxLength       =   10
         TabIndex        =   66
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox pventa11 
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
         Left            =   5640
         MaxLength       =   10
         TabIndex        =   65
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox maximo11 
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
         Left            =   5160
         MaxLength       =   5
         TabIndex        =   64
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox minimo11 
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
         Left            =   4680
         MaxLength       =   5
         TabIndex        =   63
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox dscto 
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
         Left            =   6120
         MaxLength       =   10
         TabIndex        =   62
         Top             =   5040
         Width           =   735
      End
      Begin VB.TextBox margen2 
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
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   61
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox pventa2 
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   60
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox factor2 
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
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   59
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox unidad2 
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
         Left            =   240
         MaxLength       =   6
         TabIndex        =   58
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox margen3 
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
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   57
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox pventa3 
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   56
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox factor3 
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
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   55
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox unidad3 
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
         Left            =   240
         MaxLength       =   6
         TabIndex        =   54
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox margen4 
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
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   53
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox pventa4 
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   52
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox factor4 
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
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   51
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox unidad4 
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
         Left            =   240
         MaxLength       =   6
         TabIndex        =   50
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox margen5 
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
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   49
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox pventa5 
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   48
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox factor5 
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
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   47
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox unidad5 
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
         Left            =   240
         MaxLength       =   6
         TabIndex        =   46
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox margen6 
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
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   45
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox pventa6 
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   44
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox factor6 
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
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   43
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox unidad6 
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
         Left            =   240
         MaxLength       =   6
         TabIndex        =   42
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox margen7 
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
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   41
         Top             =   3840
         Width           =   855
      End
      Begin VB.TextBox pventa7 
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   40
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox factor7 
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
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   39
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox unidad7 
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
         Left            =   240
         MaxLength       =   6
         TabIndex        =   38
         Top             =   3840
         Width           =   855
      End
      Begin VB.TextBox margen8 
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
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   37
         Top             =   4200
         Width           =   855
      End
      Begin VB.TextBox pventa8 
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   36
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox factor8 
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
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   35
         Top             =   4200
         Width           =   735
      End
      Begin VB.TextBox unidad8 
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
         Left            =   240
         MaxLength       =   6
         TabIndex        =   34
         Top             =   4200
         Width           =   855
      End
      Begin VB.TextBox margen9 
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
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   33
         Top             =   4560
         Width           =   855
      End
      Begin VB.TextBox pventa9 
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   32
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox factor9 
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
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   31
         Top             =   4560
         Width           =   735
      End
      Begin VB.TextBox unidad9 
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
         Left            =   240
         MaxLength       =   6
         TabIndex        =   30
         Top             =   4560
         Width           =   855
      End
      Begin VB.TextBox margen10 
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
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   29
         Top             =   4920
         Width           =   855
      End
      Begin VB.TextBox pventa10 
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   28
         Top             =   4920
         Width           =   1215
      End
      Begin VB.TextBox factor10 
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
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   27
         Top             =   4920
         Width           =   735
      End
      Begin VB.TextBox unidad10 
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
         Left            =   240
         MaxLength       =   6
         TabIndex        =   26
         Top             =   4920
         Width           =   855
      End
      Begin VB.TextBox minimo12 
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
         Left            =   4680
         MaxLength       =   5
         TabIndex        =   25
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox maximo12 
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
         Left            =   5160
         MaxLength       =   5
         TabIndex        =   24
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox pventa12 
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
         Left            =   5640
         MaxLength       =   10
         TabIndex        =   23
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox margen12 
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
         Left            =   6720
         MaxLength       =   10
         TabIndex        =   22
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox minimo13 
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
         Left            =   4680
         MaxLength       =   5
         TabIndex        =   21
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox maximo13 
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
         Left            =   5160
         MaxLength       =   5
         TabIndex        =   20
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox pventa13 
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
         Left            =   5640
         MaxLength       =   10
         TabIndex        =   19
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox margen13 
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
         Left            =   6720
         MaxLength       =   10
         TabIndex        =   18
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox minimo14 
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
         Left            =   4680
         MaxLength       =   5
         TabIndex        =   17
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox maximo14 
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
         Left            =   5160
         MaxLength       =   5
         TabIndex        =   16
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox pventa14 
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
         Left            =   5640
         MaxLength       =   10
         TabIndex        =   15
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox margen14 
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
         Left            =   6720
         MaxLength       =   10
         TabIndex        =   14
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox minimo15 
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
         Left            =   4680
         MaxLength       =   5
         TabIndex        =   13
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox maximo15 
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
         Left            =   5160
         MaxLength       =   5
         TabIndex        =   12
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox pventa15 
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
         Left            =   5640
         MaxLength       =   10
         TabIndex        =   11
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox margen15 
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
         Left            =   6720
         MaxLength       =   10
         TabIndex        =   10
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox producto 
         Height          =   375
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   7920
         Picture         =   "tpactop.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   720
         Width           =   1470
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   7920
         Picture         =   "tpactop.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Imprimir todo"
         Top             =   1680
         Width           =   1470
      End
      Begin VB.Label Label61 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "%MiniPre"
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
         Left            =   3840
         TabIndex        =   107
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unidad"
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
         Left            =   240
         TabIndex        =   106
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Factor"
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
         Left            =   1080
         TabIndex        =   105
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pvta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1800
         TabIndex        =   104
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label40 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "%Margen"
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
         Left            =   3000
         TabIndex        =   103
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label41 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Min"
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
         Left            =   4680
         TabIndex        =   102
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label42 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Max"
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
         Left            =   5160
         TabIndex        =   101
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label43 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pvta"
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
         Left            =   5640
         TabIndex        =   100
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label44 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "%Margen"
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
         Left            =   6720
         TabIndex        =   99
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label45 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaInicio"
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
         Left            =   4680
         TabIndex        =   98
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label46 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaFinal"
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
         Left            =   4680
         TabIndex        =   97
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label47 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaInicio"
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
         Left            =   4680
         TabIndex        =   96
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label Label48 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaFinal"
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
         Left            =   4680
         TabIndex        =   95
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Label Label49 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dscto Autom."
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
         Left            =   4680
         TabIndex        =   94
         Top             =   5040
         Width           =   1575
      End
      Begin VB.Label descripcio 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1680
         TabIndex        =   93
         Top             =   720
         Width           =   5895
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion"
         Height          =   375
         Left            =   240
         TabIndex        =   92
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Producto"
         Height          =   375
         Left            =   240
         TabIndex        =   91
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label monedac 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4920
         TabIndex        =   90
         Top             =   360
         Width           =   615
      End
      Begin VB.Label costou 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6120
         TabIndex        =   89
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label factor 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4200
         TabIndex        =   88
         Top             =   360
         Width           =   735
      End
      Begin VB.Label monedav 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5520
         TabIndex        =   87
         Top             =   360
         Width           =   615
      End
      Begin VB.Label familia 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3480
         TabIndex        =   86
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Calcula Margen"
         Height          =   1095
         Left            =   7920
         TabIndex        =   85
         Top             =   2880
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Refrescar"
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
      Left            =   10320
      MaskColor       =   &H00FFFFFF&
      Picture         =   "tpactop.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid dbgrid2 
      Height          =   7455
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   13150
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   15
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
   Begin VB.Label buffer2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
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
      Height          =   300
      Left            =   6600
      TabIndex        =   114
      Top             =   480
      Width           =   105
   End
   Begin VB.Label nombre 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   4935
   End
   Begin VB.Label codigo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cliente"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cliente"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Menu fk993 
      Caption         =   "&Nuevo"
   End
   Begin VB.Menu dmi33 
      Caption         =   "&Modifica"
   End
   Begin VB.Menu fdj833 
      Caption         =   "&Borra"
   End
   Begin VB.Menu flo44 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tpactop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tabcliente As New ADODB.Recordset

Private Sub buffer_DblClick()
    buffer_KeyPress 13

End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        flo44_Click
        Exit Sub

    End If

    Command1_Click

End Sub

Private Sub cmdCerrar_Click()
    Frame1.Visible = False

End Sub

Private Sub cmdGuardar_Click()

    Dim mytablex As New ADODB.Recordset

    If Len(producto) = 0 Or Len(descripcio) = 0 Or Len(monedav) = 0 Or Val(pventa1) <= 0 Then
        producto.SetFocus
        Exit Sub

    End If

    If Frame1 = "NUEVO" Then
        mytablex.Open "select * from precio1  where codigo='" & codigo & "' and producto='" & producto & "' and local='01'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.AddNew
            grabando_precios mytablex
            mytablex.Update
        Else
            MsgBox "Ya existe ", 48, "Aviso"
            mytablex.Close
            Exit Sub

        End If

        mytablex.Close
        MsgBox "Adicion Proceso Realizado ", 48, "Aviso"
        Frame1.Visible = False
        sql_cabeza
        Exit Sub

    End If

    If Frame1 = "MODIFICA" Then
        mytablex.Open "select * from precio1  where codigo='" & codigo & "' and producto='" & producto & "' and local='01'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            grabando_precios mytablex
            mytablex.Update

        End If

        mytablex.Close
        MsgBox "Modificacion Proceso Realizado ", 48, "Aviso"
        Frame1.Visible = False
        sql_cabeza
        'sql_cabeza

        Exit Sub

    End If

End Sub

Private Sub Command1_Click()

    Dim buf       As String

    Dim sw        As Integer

    Dim rconsulta As New ADODB.Recordset

    If Len(buffer) = 0 Then
        buf = "select Descripcio,Producto,Marca,Monedac,Costou,Unidad,Factor,Monedav,Familia from producto "
    Else
        buf = "select Descripcio,Producto,Marca,Monedac,Costou,Unidad,Factor,Monedav,Familia from producto where " & Combo1 & " like '" & buffer & "%'"

    End If

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open buf, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True And rconsulta.BOF = True Then
        rconsulta.Close
        buffer.SetFocus
        Exit Sub

    End If

    Set dbGrid1.DataSource = rconsulta
    dbGrid1.columns(0).Width = 4000
    dbGrid1.columns(1).Width = 2000
    'If sw = 1 Then
    dbGrid1.SetFocus

    'End If
End Sub

Private Sub Command2_Click()
    flo44_Click

End Sub

Private Sub Command4_Click()
    sql_cabeza

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim found As Integer

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        '      If Val(dbGrid1.columns(4)) <= 0 Then
        '         MsgBox "Producto sin Costo ", 48, "Aviso"
        '         Exit Sub
        '      End If
        producto = Trim(dbGrid1.columns(1))
        descripcio = Trim(dbGrid1.columns(0))
        monedac = Trim(dbGrid1.columns(3))

        If Trim(dbGrid1.columns(4)) = "0" Then
            costou = 0
        Else
            costou = Trim(dbGrid1.columns(4))

        End If
    
        factor = Trim(dbGrid1.columns(6))
        monedav = Trim(dbGrid1.columns(7))
        familia = Trim(dbGrid1.columns(8))

        If Frame1 = "NUEVO" Then
            found = verifica_precio()

        End If

        Frame2.Visible = False
        Frame1.Enabled = True
        calcula_margenes
        pventa1.SetFocus
        Exit Sub

    End If

End Sub

Sub pone_margen()

    Dim mytablex As New ADODB.Recordset

    If Len(Trim(familia)) = 0 Then Exit Sub
    mytablex.Open "select * from familia where familia='" & Trim(familia) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
   
        If "" & mytablex.Fields("obliga") = "S" Then
      
            If Val("" & mytablex.Fields("margen1")) > 0 And Val(factor1) > 0 Then
                margen1 = "" & mytablex.Fields("margen1")

            End If

            If Val("" & mytablex.Fields("margen2")) > 0 And Val(factor2) > 0 Then
                margen2 = "" & mytablex.Fields("margen2")

            End If

            If Val("" & mytablex.Fields("margen3")) > 0 And Val(factor3) > 0 Then
                margen3 = "" & mytablex.Fields("margen3")

            End If

            If Val("" & mytablex.Fields("margen4")) > 0 And Val(factor4) > 0 Then
                margen4 = "" & mytablex.Fields("margen4")

            End If

            If Val("" & mytablex.Fields("margen5")) > 0 And Val(factor5) > 0 Then
                margen5 = "" & mytablex.Fields("margen5")

            End If

            If Val("" & mytablex.Fields("margen6")) > 0 And Val(factor6) > 0 Then
                margen6 = "" & mytablex.Fields("margen6")

            End If

            If Val("" & mytablex.Fields("margen7")) > 0 And Val(factor7) > 0 Then
                margen7 = "" & mytablex.Fields("margen7")

            End If

            If Val("" & mytablex.Fields("margen8")) > 0 And Val(factor8) > 0 Then
                margen8 = "" & mytablex.Fields("margen8")

            End If

            If Val("" & mytablex.Fields("margen9")) > 0 And Val(factor9) > 0 Then
                margen9 = "" & mytablex.Fields("margen9")

            End If

            If Val("" & mytablex.Fields("margen10")) > 0 And Val(factor10) > 0 Then
                margen10 = "" & mytablex.Fields("margen10")

            End If

        End If

    End If

    mytablex.Close

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
        Command1_Click
         
    End If

End Sub

Private Sub dj8333_Click()

    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub

End Sub

Private Sub dbgrid2_KeyPress(KeyAscii As Integer)

    Dim buf  As String

    Dim buf2 As String

    If KeyAscii <> 13 And KeyAscii <> 27 Then
        If KeyAscii = 8 Then
            If Len(buffer2) > 0 Then
                buf = Mid$(buffer2, 1, Len(buffer2) - 1)
                buffer2 = buf
                KeyAscii = 0
            Else
                KeyAscii = 0
                Exit Sub

            End If

        End If

        buf = Chr(KeyAscii)

        If Chr(KeyAscii) = "*" Then
            buf = ""
            buffer2 = buf

        End If

        If KeyAscii <> 13 Then
            buffer2 = buffer2 + buf

        End If

        buf = buffer2
        sql_cabeza

    End If

End Sub

Private Sub dbgrid2_Error(ByVal DataError As Integer, Response As Integer)
    Response = False

End Sub

Private Sub dmi33_Click()

    On Error GoTo cmd1_err

    Dim mytablex As New ADODB.Recordset

    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub

    producto = "" & tabcliente.Fields("producto")
    descripcio = "" & tabcliente.Fields("descripcio")

    mytablex.Open "select * from producto  where producto='" & producto & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        Exit Sub

    End If

    monedac = "" & mytablex.Fields("monedac")
    costou = "" & mytablex.Fields("costou")
    factor = "" & mytablex.Fields("factor")
    monedav = "" & mytablex.Fields("monedav")
    familia = "" & mytablex.Fields("familia")
   
    mytablex.Close

    mytablex.Open "select * from precio1  where codigo='" & codigo & "' and producto='" & producto & "' and local='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        pone_precios mytablex
        calcula_margenes
    Else
        Exit Sub

    End If

    Frame1.Visible = True
    Frame1.Caption = "MODIFICA"
    producto.Enabled = False
    pventa1.SetFocus
    Exit Sub
cmd1_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fdj833_Click()

    Dim buf1 As String

    Dim buf2 As String

    On Error GoTo cmd2_err

    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    buf1 = "" & tabcliente.Fields("producto")
    buf2 = "" '& tabcliente.Fields("codigo")
    cn.Execute ("delete from precio1 where local='01' and producto='" & buf1 & "' and codigo='" & codigo & "'")
    sql_cabeza
    Exit Sub
cmd2_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fk993_Click()

    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    producto = ""
    descripcio = ""
    Frame1.Visible = True
    inicializa_precios
    Frame1.Caption = "NUEVO"
    producto.Enabled = True
    producto.SetFocus

End Sub

Private Sub flo44_Click()

    If Frame2.Visible = True Then
        Frame2.Visible = False
        Frame1.Enabled = True
        Exit Sub

    End If

    If Frame1.Visible = True Then
        Frame1.Visible = False
        Exit Sub

    End If

    tpactop.Hide
    Unload tpactop

End Sub

Private Sub Form_Activate()
    sql_cabeza

End Sub

Function busca_productof(buf As String)

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim found    As Integer

    Dim buf1     As String

    Dim I        As Integer

    Dim ssw      As Integer

    Dim sw       As Integer

    I = 0

    found = 0

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM producto where producto='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        found = busca_equiva(buf) 'busca en la table codigo barras

        If found = 0 Then
            Exit Function

        End If

        mytablex.Open "SELECT * FROM producto where producto='" & buf & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.Close
            Exit Function

        End If

    End If

    producto = Trim("" & mytablex.Fields("producto"))
    descripcio = Trim("" & mytablex.Fields("descripcio"))
    monedac = Trim("" & mytablex.Fields("monedac"))
    costou = Trim("" & mytablex.Fields("costou"))
    factor = Trim("" & mytablex.Fields("factor"))
    monedav = Trim("" & mytablex.Fields("monedav"))
    familia = Trim("" & mytablex.Fields("familia"))
      
    mytablex.Close
    busca_productof = 1

End Function

Function busca_equiva(buf As String) As Integer

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM productb where barras='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        buf = "" & mytablex.Fields("producto")
        busca_equiva = 1
        mytablex.Close
        Exit Function

    End If

    mytablex.Close
   
    mytablex.Open "SELECT * FROM producto where barras='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        buf = "" & mytablex.Fields("producto")
        busca_equiva = 1

    End If

    mytablex.Close

End Function

Sub sql_cabeza()

    Dim buf  As String

    Dim buf1 As String

    buf1 = "select producto.descripcio,producto.producto,producto.marca,precio1.unidad1,precio1.factor1,precio1.pventa1,precio1.unidad2,precio1.factor2,precio1.pventa2,precio1.unidad3,precio1.factor3,precio1.pventa3,precio1.unidad4,precio1.factor4,precio1.pventa4,precio1.unidad5,precio1.factor5,precio1.pventa5,precio1.unidad6,precio1.factor6,precio1.pventa6,precio1.unidad7,precio1.factor7,precio1.pventa7,precio1.unidad8,precio1.factor8,precio1.pventa8,precio1.unidad9,precio1.factor9,precio1.pventa9,precio1.unidad10,precio1.factor10,precio1.pventa10 from producto  left join precio1 on producto.producto=precio1.producto  where precio1.local='01' and precio1.codigo='" & codigo & "'"

    If Len(buffer2) > 0 Then
        buf1 = buf1 & " and producto.descripcio like '" & buffer2 & "%'"

    End If

    'MsgBox buf1

    If tabcliente.State = 1 Then tabcliente.Close
    tabcliente.Open buf1, cn, adOpenStatic, adLockOptimistic
    Set DBGrid2.DataSource = tabcliente
    DBGrid2.refresh
   
    DBGrid2.columns(0).Width = 4000
    DBGrid2.columns(1).Width = 2000
   
    DBGrid2.columns(2).Width = 800
   
    DBGrid2.columns(3).Width = 800
    DBGrid2.columns(4).Width = 800
    DBGrid2.columns(5).Width = 1200
   
    DBGrid2.columns(6).Width = 800
    DBGrid2.columns(7).Width = 800
    DBGrid2.columns(8).Width = 1200
   
    DBGrid2.columns(9).Width = 800
    DBGrid2.columns(10).Width = 800
    DBGrid2.columns(11).Width = 1200
   
    DBGrid2.columns(12).Width = 800
    DBGrid2.columns(13).Width = 800
    DBGrid2.columns(14).Width = 1200
   
    DBGrid2.columns(15).Width = 800
    DBGrid2.columns(16).Width = 800
    DBGrid2.columns(17).Width = 1200
   
    DBGrid2.columns(18).Width = 800
    DBGrid2.columns(19).Width = 800
    DBGrid2.columns(20).Width = 1200
   
    DBGrid2.columns(21).Width = 800
    DBGrid2.columns(22).Width = 800
    DBGrid2.columns(23).Width = 1200
   
    DBGrid2.columns(24).Width = 800
    DBGrid2.columns(25).Width = 800
    DBGrid2.columns(26).Width = 1200
   
    DBGrid2.columns(27).Width = 800
    DBGrid2.columns(28).Width = 800
    DBGrid2.columns(29).Width = 1200
   
    DBGrid2.columns(30).Width = 800
    DBGrid2.columns(31).Width = 800
    DBGrid2.columns(32).Width = 1200

End Sub

Sub consulta_productos()
    Combo1.Clear
    Combo1.AddItem "Descripcio"
    Combo1.AddItem "Producto"
    Combo1.ListIndex = 0
    Frame2.Visible = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "1"
    Command1_Click

End Sub

Sub pone_precios(mytablex As ADODB.Recordset)
    'tpventa = Val("" & mytablex.Fields("pventa1"))
    unidad1 = "" & mytablex.Fields("unidad1")
    unidad2 = "" & mytablex.Fields("unidad2")
    unidad3 = "" & mytablex.Fields("unidad3")
    unidad4 = "" & mytablex.Fields("unidad4")
    unidad5 = "" & mytablex.Fields("unidad5")
    unidad6 = "" & mytablex.Fields("unidad6")
    unidad7 = "" & mytablex.Fields("unidad7")
    unidad8 = "" & mytablex.Fields("unidad8")
    unidad9 = "" & mytablex.Fields("unidad9")
    unidad10 = "" & mytablex.Fields("unidad10")
    factor1 = "" & mytablex.Fields("factor1")
    factor2 = "" & mytablex.Fields("factor2")
    factor3 = "" & mytablex.Fields("factor3")
    factor4 = "" & mytablex.Fields("factor4")
    factor5 = "" & mytablex.Fields("factor5")
    factor6 = "" & mytablex.Fields("factor6")
    factor7 = "" & mytablex.Fields("factor7")
    factor8 = "" & mytablex.Fields("factor8")
    factor9 = "" & mytablex.Fields("factor9")
    factor10 = "" & mytablex.Fields("factor10")
    pventa1 = "" & mytablex.Fields("pventa1")
    pventa2 = "" & mytablex.Fields("pventa2")
    pventa3 = "" & mytablex.Fields("pventa3")
    pventa4 = "" & mytablex.Fields("pventa4")
    pventa5 = "" & mytablex.Fields("pventa5")
    pventa6 = "" & mytablex.Fields("pventa6")
    pventa7 = "" & mytablex.Fields("pventa7")
    pventa8 = "" & mytablex.Fields("pventa8")
    pventa9 = "" & mytablex.Fields("pventa9")
    pventa10 = "" & mytablex.Fields("pventa10")
    margen1 = "" & mytablex.Fields("margen1")
    margen2 = "" & mytablex.Fields("margen2")
    margen3 = "" & mytablex.Fields("margen3")
    margen4 = "" & mytablex.Fields("margen4")
    margen5 = "" & mytablex.Fields("margen5")
    margen6 = "" & mytablex.Fields("margen6")
    margen7 = "" & mytablex.Fields("margen7")
    margen8 = "" & mytablex.Fields("margen8")
    margen9 = "" & mytablex.Fields("margen9")
    margen10 = "" & mytablex.Fields("margen10")
    minimo11 = "" & mytablex.Fields("minimo11")
    minimo12 = "" & mytablex.Fields("minimo12")
    minimo13 = "" & mytablex.Fields("minimo13")
    minimo14 = "" & mytablex.Fields("minimo14")
    minimo15 = "" & mytablex.Fields("minimo15")
    maximo11 = "" & mytablex.Fields("maximo11")
    maximo12 = "" & mytablex.Fields("maximo12")
    maximo13 = "" & mytablex.Fields("maximo13")
    maximo14 = "" & mytablex.Fields("maximo14")
    maximo15 = "" & mytablex.Fields("maximo15")
    pventa11 = "" & mytablex.Fields("pventa11")
    pventa12 = "" & mytablex.Fields("pventa12")
    pventa13 = "" & mytablex.Fields("pventa13")
    pventa14 = "" & mytablex.Fields("pventa14")
    pventa15 = "" & mytablex.Fields("pventa15")
    margen11 = "" & mytablex.Fields("margen11")
    margen12 = "" & mytablex.Fields("margen12")
    margen13 = "" & mytablex.Fields("margen13")
    margen14 = "" & mytablex.Fields("margen14")
    margen15 = "" & mytablex.Fields("margen15")
    fechai11 = "" & mytablex.Fields("fechai11")
    fechaf11 = "" & mytablex.Fields("fechaf11")
    fechaid = "" & mytablex.Fields("fechaid")
    fechafd = "" & mytablex.Fields("fechafd")
    dscto = "" & mytablex.Fields("dscto")
    'ccosto = "" & mytablex.Fields("ccosto")
    pm1 = "" & mytablex.Fields("pm1")
    pm2 = "" & mytablex.Fields("pm2")
    pm3 = "" & mytablex.Fields("pm3")
    pm4 = "" & mytablex.Fields("pm4")
    pm5 = "" & mytablex.Fields("pm5")
    pm6 = "" & mytablex.Fields("pm6")
    pm7 = "" & mytablex.Fields("pm7")
    pm8 = "" & mytablex.Fields("pm8")
    pm9 = "" & mytablex.Fields("pm9")
    pm10 = "" & mytablex.Fields("pm10")

    If Len(unidad1) = 0 Then
        unidad1 = "UND"

    End If

    If Len(factor1) = 0 Then
        factor1 = "1"

    End If

End Sub

Sub grabando_precios(mytablex As ADODB.Recordset)
    mytablex.Fields("local") = "01"
    mytablex.Fields("producto") = producto
    mytablex.Fields("codigo") = codigo
    mytablex.Fields("pm1") = Val(pm1)
    mytablex.Fields("pm2") = Val(pm2)
    mytablex.Fields("pm3") = Val(pm3)
    mytablex.Fields("pm4") = Val(pm4)
    mytablex.Fields("pm5") = Val(pm5)
    mytablex.Fields("pm6") = Val(pm6)
    mytablex.Fields("pm7") = Val(pm7)
    mytablex.Fields("pm8") = Val(pm8)
    mytablex.Fields("pm9") = Val(pm9)
    mytablex.Fields("pm10") = Val(pm10)

    'mytablex.Fields("ccosto") = ccosto
    mytablex.Fields("unidad1") = ""
    mytablex.Fields("unidad2") = ""
    mytablex.Fields("unidad3") = ""
    mytablex.Fields("unidad4") = ""
    mytablex.Fields("unidad5") = ""
    mytablex.Fields("unidad6") = ""
    mytablex.Fields("unidad7") = ""
    mytablex.Fields("unidad8") = ""
    mytablex.Fields("unidad9") = ""
    mytablex.Fields("unidad10") = ""
    mytablex.Fields("factor1") = 0
    mytablex.Fields("factor2") = 0
    mytablex.Fields("factor3") = 0
    mytablex.Fields("factor4") = 0
    mytablex.Fields("factor5") = 0
    mytablex.Fields("factor6") = 0
    mytablex.Fields("factor7") = 0
    mytablex.Fields("factor8") = 0
    mytablex.Fields("factor9") = 0
    mytablex.Fields("factor10") = 0
    mytablex.Fields("pventa1") = 0
    mytablex.Fields("pventa2") = 0
    mytablex.Fields("pventa3") = 0
    mytablex.Fields("pventa4") = 0
    mytablex.Fields("pventa5") = 0
    mytablex.Fields("pventa6") = 0
    mytablex.Fields("pventa7") = 0
    mytablex.Fields("pventa8") = 0
    mytablex.Fields("pventa9") = 0
    mytablex.Fields("pventa10") = 0
    mytablex.Fields("margen1") = 0
    mytablex.Fields("margen2") = 0
    mytablex.Fields("margen3") = 0
    mytablex.Fields("margen4") = 0
    mytablex.Fields("margen5") = 0
    mytablex.Fields("margen6") = 0
    mytablex.Fields("margen7") = 0
    mytablex.Fields("margen8") = 0
    mytablex.Fields("margen9") = 0
    mytablex.Fields("margen10") = 0

    If Val(factor1) > 0 And Len(unidad1) > 0 Then
        mytablex.Fields("unidad1") = unidad1
        mytablex.Fields("factor1") = Val(factor1)
        mytablex.Fields("pventa1") = Val(pventa1)
        mytablex.Fields("margen1") = Val(margen1)

    End If

    If Val(factor2) > 0 And Len(unidad2) > 0 Then
        mytablex.Fields("unidad2") = unidad2
        mytablex.Fields("factor2") = Val(factor2)
        mytablex.Fields("pventa2") = Val(pventa2)
        mytablex.Fields("margen2") = Val(margen2)

    End If

    If Val(factor3) > 0 And Len(unidad3) > 0 Then
        mytablex.Fields("unidad3") = unidad3
        mytablex.Fields("factor3") = Val(factor3)
        mytablex.Fields("pventa3") = Val(pventa3)
        mytablex.Fields("margen3") = Val(margen3)

    End If

    If Val(factor4) > 0 And Len(unidad4) > 0 Then
        mytablex.Fields("unidad4") = unidad4
        mytablex.Fields("factor4") = Val(factor4)
        mytablex.Fields("pventa4") = Val(pventa4)
        mytablex.Fields("margen1") = Val(margen1)

    End If

    If Val(factor5) > 0 And Len(unidad5) > 0 Then
        mytablex.Fields("unidad5") = unidad5
        mytablex.Fields("factor5") = Val(factor5)
        mytablex.Fields("pventa5") = Val(pventa5)
        mytablex.Fields("margen5") = Val(margen5)

    End If

    If Val(factor6) > 0 And Len(unidad6) > 0 Then
        mytablex.Fields("unidad6") = unidad6
        mytablex.Fields("factor6") = Val(factor6)
        mytablex.Fields("pventa6") = Val(pventa6)
        mytablex.Fields("margen1") = Val(margen1)

    End If

    If Val(factor7) > 0 And Len(unidad7) > 0 Then
        mytablex.Fields("unidad7") = unidad7
        mytablex.Fields("factor7") = Val(factor7)
        mytablex.Fields("pventa7") = Val(pventa7)
        mytablex.Fields("margen7") = Val(margen7)

    End If

    If Val(factor8) > 0 And Len(unidad8) > 0 Then
        mytablex.Fields("unidad8") = unidad8
        mytablex.Fields("factor8") = Val(factor8)
        mytablex.Fields("pventa8") = Val(pventa8)
        mytablex.Fields("margen8") = Val(margen8)

    End If

    If Val(factor9) > 0 And Len(unidad9) > 0 Then
        mytablex.Fields("unidad9") = unidad9
        mytablex.Fields("factor9") = Val(factor9)
        mytablex.Fields("pventa9") = Val(pventa9)
        mytablex.Fields("margen9") = Val(margen9)

    End If

    If Val(factor10) > 0 And Len(unidad10) > 0 Then
        mytablex.Fields("unidad10") = unidad10
        mytablex.Fields("factor10") = Val(factor10)
        mytablex.Fields("pventa10") = Val(pventa10)
        mytablex.Fields("margen10") = Val(margen10)

    End If

    mytablex.Fields("minimo11") = Val(minimo11)
    mytablex.Fields("minimo12") = Val(minimo12)
    mytablex.Fields("minimo13") = Val(minimo13)
    mytablex.Fields("minimo14") = Val(minimo14)
    mytablex.Fields("minimo15") = Val(minimo15)
    mytablex.Fields("maximo11") = Val(maximo11)
    mytablex.Fields("maximo12") = Val(maximo12)
    mytablex.Fields("maximo13") = Val(maximo13)
    mytablex.Fields("maximo14") = Val(maximo14)
    mytablex.Fields("maximo15") = Val(maximo15)
    mytablex.Fields("pventa11") = Val(pventa11)
    mytablex.Fields("pventa12") = Val(pventa12)
    mytablex.Fields("pventa13") = Val(pventa13)
    mytablex.Fields("pventa14") = Val(pventa14)
    mytablex.Fields("pventa15") = Val(pventa15)
    mytablex.Fields("margen11") = Val(margen11)
    mytablex.Fields("margen12") = Val(margen12)
    mytablex.Fields("margen13") = Val(margen13)
    mytablex.Fields("margen14") = Val(margen14)
    mytablex.Fields("margen15") = Val(margen15)

    If IsDate(fechai11) Then
        mytablex.Fields("fechai11") = fechai11

    End If

    If IsDate(fechaf11) Then
        mytablex.Fields("fechaf11") = fechaf11

    End If

    If IsDate(fechaid) Then
        mytablex.Fields("fechaid") = fechaid

    End If

    If IsDate(fechafd) Then
        mytablex.Fields("fechafd") = fechafd

    End If

    mytablex.Fields("dscto") = Val(dscto)

End Sub

Sub inicializa_precios()
    unidad1 = "UND"
    unidad2 = ""
    unidad3 = ""
    unidad4 = ""
    unidad5 = ""
    unidad6 = ""
    unidad7 = ""
    unidad8 = ""
    unidad9 = ""
    unidad10 = ""
    factor1 = "1"
    factor2 = ""
    factor3 = ""
    factor4 = ""
    factor5 = ""
    factor6 = ""
    factor7 = ""
    factor8 = ""
    factor9 = ""
    factor10 = ""

    pventa1 = ""
    pventa2 = ""
    pventa3 = ""
    pventa4 = ""
    pventa5 = ""
    pventa6 = ""
    pventa7 = ""
    pventa8 = ""
    pventa9 = ""
    pventa10 = ""

    margen1 = ""
    margen2 = ""
    margen3 = ""
    margen4 = ""
    margen5 = ""
    margen6 = ""
    margen7 = ""
    margen8 = ""
    margen9 = ""
    margen10 = ""
    minimo11 = ""
    minimo12 = ""
    minimo13 = ""
    minimo14 = ""
    minimo15 = ""

    maximo11 = ""
    maximo12 = ""
    maximo13 = ""
    maximo14 = ""
    maximo15 = ""

    pventa11 = ""
    pventa12 = ""
    pventa13 = ""
    pventa14 = ""
    pventa15 = ""

    margen11 = ""
    margen12 = ""
    margen13 = ""
    margen14 = ""
    margen15 = ""
    fechai11 = ""
    fechaf11 = ""
    fechaid = ""
    fechafd = ""
    dscto = ""

End Sub

Private Sub Label3_Click()
    calcula_margenes

End Sub

Private Sub producto_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        Frame1.Enabled = False
        consulta_productos

    End If

End Sub

Sub calcula_margenes()

    Dim sdx     As Double

    Dim sdx1    As Double

    Dim sdx2    As Double

    Dim acostou As String

    On Error GoTo cmd786_err

    sdx = Val(costou)
    acostou = "" & sdx

    If Val(factor) <= 0 Then
        factor = "1"

    End If

    If Val(factor1) <= 0 Then
        factor1 = "1"

    End If

    If Val(costou) = 0 Then
        margen1 = "0"
        margen2 = "0"
        margen3 = "0"
        margen4 = "0"
        margen5 = "0"
        margen6 = "0"
        margen7 = "0"
        margen8 = "0"
        margen9 = "0"
        margen10 = "0"
   
        margen11 = "0"
        margen12 = "0"
        margen13 = "0"
        margen14 = "0"
        margen15 = "0"
        pone_margen

        Exit Sub

    End If

    pone_margen

    If monedac = "S" Then
        If monedav = "D" Then
            sdx = Val(acostou) / busca_cambio()

            If sdx <= 0 Then
                sdx = 1

            End If

            acostou = "" & sdx

        End If

    End If

    If monedac = "D" Then
        If monedav = "S" Then
            sdx = Val(acostou) * busca_cambio()

            If sdx <= 0 Then
                sdx = 1

            End If

            acostou = "" & sdx

        End If

    End If

    If Val(acostou) > 0 And Val(pventa1) > 0 And Val(factor1) > 0 Then 'calculando margenes
        sdx = (Val(acostou))
        sdx = sdx * Val(factor1)
        sdx1 = Val(pventa1) '/ Val(factor1)
        sdx2 = (sdx1 - sdx) * 100 / sdx
        margen1 = Format(sdx2, "0.00")
        GoTo siguiente1

    End If

    If Val(margen1) > 0 And Val(pventa1) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen1) / 100
        sdx = sdx * Val(factor1)
        pventa1 = Format(sdx, "0.00")
        GoTo siguiente1

    End If

    If Val(acostou) <= 0 And Val(pventa1) > 0 And Val(margen1) > 0 Then
        sdx = Val(pventa1) / (1 + (Val(margen1) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente1

    End If
       
siguiente1:

    If Val(acostou) > 0 And Val(pventa2) > 0 And Val(factor2) > 0 Then 'calculando margenes
        sdx = (Val(acostou))
        sdx = sdx * Val(factor2)
        sdx1 = Val(pventa2) '/ Val(factor2)
        sdx2 = (Val(sdx1) - sdx) * 100 / sdx
        margen2 = Format(sdx2, "0.00")
        GoTo siguiente2

    End If

    If Val(margen2) > 0 And Val(pventa2) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen2) / 100
        sdx = sdx * Val(factor2)
        pventa2 = Format(sdx, "0.00")
        GoTo siguiente2

    End If

    If Val(acostou) <= 0 And Val(pventa2) > 0 And Val(margen2) > 0 Then
        sdx = Val(pventa2) / (1 + (Val(margen2) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente2

    End If

siguiente2:

    If Val(acostou) > 0 And Val(pventa3) > 0 And Val(factor3) > 0 Then 'calculando margenes
        sdx = (Val(acostou))  '/ Val(factor))
        sdx = sdx * Val(factor3)
        sdx1 = Val(pventa3) '/ Val(factor1)
        sdx2 = (sdx1 - sdx) * 100 / sdx
        margen3 = Format(sdx2, "0.00")
        GoTo siguiente3

    End If

    If Val(margen3) > 0 And Val(pventa3) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen3) / 100
        sdx = sdx * Val(factor3)
        pventa3 = Format(sdx, "0.00")
        GoTo siguiente3

    End If

    If Val(acostou) <= 0 And Val(pventa3) > 0 And Val(margen3) > 0 Then
        sdx = Val(pventa3) / (1 + (Val(margen3) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente2

    End If

siguiente3:

    If Val(acostou) > 0 And Val(pventa4) > 0 And Val(factor4) > 0 Then 'calculando margenes
        sdx = (Val(acostou)) '/ Val(factor))
        sdx = sdx * Val(factor4)
        sdx1 = Val(pventa4) '/ Val(factor1)
        sdx2 = (sdx1 - sdx) * 100 / sdx
        margen4 = Format(sdx2, "0.00")
        GoTo siguiente4

    End If

    If Val(margen4) > 0 And Val(pventa4) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen4) / 100
        sdx = sdx * Val(factor4)
        pventa4 = Format(sdx, "0.00")
        GoTo siguiente4

    End If

    If Val(acostou) <= 0 And Val(pventa4) > 0 And Val(margen4) > 0 Then
        sdx = Val(pventa4) / (1 + (Val(margen4) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente4

    End If

siguiente4:

    If Val(acostou) > 0 And Val(pventa5) > 0 And Val(factor5) > 0 Then 'calculando margenes
        sdx = (Val(acostou) / Val(factor))
        sdx = sdx * Val(factor5)
        sdx1 = Val(pventa5) '/ Val(factor1)
        sdx2 = (sdx1 - sdx) * 100 / sdx
        margen5 = Format(sdx2, "0.00")
        GoTo siguiente5

    End If

    If Val(margen5) > 0 And Val(pventa5) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen5) / 100
        sdx = sdx * Val(factor5)
        pventa5 = Format(sdx, "0.00")
        GoTo siguiente5

    End If

    If Val(acostou) <= 0 And Val(pventa5) > 0 And Val(margen5) > 0 Then
        sdx = Val(pventa5) / (1 + (Val(margen5) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente5

    End If

siguiente5:

    If Val(acostou) > 0 And Val(pventa6) > 0 And Val(factor6) > 0 Then 'calculando margenes
        sdx = (Val(acostou) / Val(factor))
        sdx = sdx * Val(factor6)
        sdx1 = Val(pventa6) '/ Val(factor1)
        sdx2 = (sdx1 - sdx) * 100 / sdx
        margen6 = Format(sdx2, "0.00")
        GoTo siguiente6

    End If

    If Val(margen6) > 0 And Val(pventa6) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen6) / 100
        sdx = sdx * Val(factor6)
        pventa6 = Format(sdx, "0.00")
        GoTo siguiente6

    End If

    If Val(acostou) <= 0 And Val(pventa6) > 0 And Val(margen6) > 0 Then
        sdx = Val(pventa6) / (1 + (Val(margen6) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente6

    End If

siguiente6:

    If Val(acostou) > 0 And Val(pventa7) > 0 And Val(factor7) > 0 Then 'calculando margenes
        sdx = (Val(acostou) / Val(factor))
        sdx = sdx * Val(factor7)
        sdx1 = Val(pventa7) '/ Val(factor1)
        sdx2 = (sdx1 - sdx) * 100 / sdx
        margen7 = Format(sdx2, "0.00")
        GoTo siguiente7

    End If

    If Val(margen7) > 0 And Val(pventa7) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen7) / 100
        sdx = sdx * Val(factor7)
        pventa7 = Format(sdx, "0.00")
        GoTo siguiente7

    End If

    If Val(acostou) <= 0 And Val(pventa7) > 0 And Val(margen7) > 0 Then
        sdx = Val(pventa7) / (1 + (Val(margen7) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente7

    End If

siguiente7:

    If Val(costou) > 0 And Val(pventa8) > 0 And Val(factor8) > 0 Then 'calculando margenes
        sdx = (Val(acostou) / Val(factor))
        sdx = sdx * Val(factor8)
        sdx1 = Val(pventa8) '/ Val(factor1)
        sdx2 = (sdx1 - sdx) * 100 / sdx
        margen8 = Format(sdx2, "0.00")
        GoTo siguiente8

    End If

    If Val(margen8) > 0 And Val(pventa8) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen8) / 100
        sdx = sdx * Val(factor8)
        pventa8 = Format(sdx, "0.00")
        GoTo siguiente8

    End If

    If Val(acostou) <= 0 And Val(pventa8) > 0 And Val(margen8) > 0 Then
        sdx = Val(pventa8) / (1 + (Val(margen8) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente8

    End If

siguiente8:

    If Val(acostou) > 0 And Val(pventa9) > 0 And Val(factor9) > 0 Then 'calculando margenes
        sdx = (Val(acostou) / Val(factor))
        sdx = sdx * Val(factor9)
        sdx1 = Val(pventa9) '/ Val(factor1)
        sdx2 = (sdx1 - sdx) * 100 / sdx
        margen9 = Format(sdx2, "0.00")
        GoTo siguiente9

    End If

    If Val(margen9) > 0 And Val(pventa9) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen9) / 100
        sdx = sdx * Val(factor9)
        pventa9 = Format(sdx, "0.00")
        GoTo siguiente9

    End If

    If Val(acostou) <= 0 And Val(pventa9) > 0 And Val(margen9) > 0 Then
        sdx = Val(pventa9) / (1 + (Val(margen9) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente9

    End If

siguiente9:

    If Val(acostou) > 0 And Val(pventa10) > 0 And Val(factor10) > 0 Then 'calculando margenes
        sdx = (Val(acostou) / Val(factor))
        sdx = sdx * Val(factor10)
        sdx1 = Val(pventa10) '/ Val(factor1)
        sdx2 = (sdx1 - sdx) * 100 / sdx
        margen10 = Format(sdx2, "0.00")
        GoTo siguiente10

    End If

    If Val(margen10) > 0 And Val(pventa10) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen10) / 100
        sdx = sdx * Val(factor10)
        pventa2 = Format(sdx, "0.00")
        GoTo siguiente10

    End If

    If Val(acostou) <= 0 And Val(pventa10) > 0 And Val(margen10) > 0 Then
        sdx = Val(pventa10) / (1 + (Val(margen10) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente10

    End If

siguiente10:

    If Val(acostou) > 0 And Val(pventa11) > 0 And Val(factor1) > 0 Then 'calculando margenes
        sdx = (Val(acostou) / Val(factor))
        sdx = sdx * Val(factor1)
        sdx1 = Val(pventa11) '/ Val(factor1)
        sdx2 = (sdx1 - sdx) * 100 / sdx
        margen11 = Format(sdx2, "0.00")
        GoTo siguiente11

    End If

    If Val(margen11) > 0 And Val(pventa11) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen11) / 100
        sdx = sdx * Val(factor1)
        pventa11 = Format(sdx, "0.00")
        GoTo siguiente11

    End If

    If Val(acostou) <= 0 And Val(pventa11) > 0 And Val(margen11) > 0 Then
        sdx = Val(pventa11) / (1 + (Val(margen11) / 100))
        sdx = sdx * Val(factor1)
        costou = Format(sdx, "0.0000")
        GoTo siguiente11

    End If

siguiente11:

    If Val(acostou) > 0 And Val(pventa12) > 0 And Val(factor1) > 0 Then 'calculando margenes
        sdx = (Val(acostou) / Val(factor))
        sdx = sdx * Val(factor1)
        sdx1 = Val(pventa12) '/ Val(factor1)
        sdx2 = (sdx1 - sdx) * 100 / sdx
        margen12 = Format(sdx2, "0.00")
        GoTo siguiente12

    End If

    If Val(margen12) > 0 And Val(pventa12) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen12) / 100
        sdx = sdx * Val(factor1)
        pventa12 = Format(sdx, "0.00")
        GoTo siguiente12

    End If

    If Val(acostou) <= 0 And Val(pventa12) > 0 And Val(margen12) > 0 Then
        sdx = Val(pventa12) / (1 + (Val(margen12) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente12

    End If

siguiente12:

    If Val(acostou) > 0 And Val(pventa13) > 0 And Val(factor1) > 0 Then 'calculando margenes
        sdx = (Val(acostou) / Val(factor))
        sdx = sdx * Val(factor1)
        sdx1 = Val(pventa13) '/ Val(factor1)
        sdx2 = (sdx1 - sdx) * 100 / sdx
        margen13 = Format(sdx2, "0.00")
        GoTo siguiente13

    End If

    If Val(margen13) > 0 And Val(pventa13) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen13) / 100
        sdx = sdx * Val(factor1)
        pventa13 = Format(sdx, "0.00")
        GoTo siguiente13

    End If

    If Val(acostou) <= 0 And Val(pventa13) > 0 And Val(margen13) > 0 Then
        sdx = Val(pventa13) / (1 + (Val(margen13) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente13

    End If

siguiente13:

    If Val(acostou) > 0 And Val(pventa14) > 0 And Val(factor1) > 0 Then 'calculando margenes
        sdx = (Val(acostou) / Val(factor))
        sdx = sdx * Val(factor1)
        sdx1 = Val(pventa14) '/ Val(factor1)
        sdx2 = (sdx1 - sdx) * 100 / sdx
        margen14 = Format(sdx2, "0.00")
        GoTo siguiente14

    End If

    If Val(margen14) > 0 And Val(pventa14) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen14) / 100
        sdx = sdx * Val(factor1)
        pventa14 = Format(sdx, "0.00")
        GoTo siguiente14

    End If

    If Val(acostou) <= 0 And Val(pventa14) > 0 And Val(margen14) > 0 Then
        sdx = Val(pventa14) / (1 + (Val(margen14) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente14

    End If

siguiente14:

    If Val(acostou) > 0 And Val(pventa15) > 0 And Val(factor1) > 0 Then 'calculando margenes
        sdx = (Val(acostou) / Val(factor))
        sdx = sdx * Val(factor1)
        sdx1 = Val(pventa15) '/ Val(factor1)
        sdx2 = (sdx1 - sdx) * 100 / sdx
        margen15 = Format(sdx2, "0.00")
        GoTo siguiente15

    End If

    If Val(margen15) > 0 And Val(pventa15) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen15) / 100
        sdx = sdx * Val(factor1)
        pventa15 = Format(sdx, "0.00")
        GoTo siguiente10

    End If

    If Val(acostou) <= 0 And Val(pventa15) > 0 And Val(margen15) > 0 Then
        sdx = Val(pventa15) / (1 + (Val(margen15) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente15

    End If

siguiente15:
    'cospaqu = Format(Val(costou) * Val(factor))
    'cospaqp = Format(Val(costop) * Val(factor))
    'cospaqi = Format(Val(costoini) * Val(factor))

    Exit Sub
cmd786_err:
    MsgBox "Error en calcula margenes", 48, "Aviso"
    Exit Sub

End Sub

Function busca_cambio() As Double

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Double

    sdx = 1
    mytablex.Open "SELECT * FROM parame where  codigo='01'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then

        sdx = Val("" & mytablex.Fields("paricomp"))

        If Val("" & mytablex.Fields("paricomp")) <= 0 Then
            sdx = 1

        End If

    End If

    busca_cambio = sdx
    mytablex.Close

End Function

Sub borra_campos()
    monedac = ""
    costou = ""
    factor = ""
    monedav = ""
    familia = ""

End Sub

Private Sub producto_KeyPress(KeyAscii As Integer)

    Dim mytablex As New ADODB.Recordset

    Dim found    As Integer

    If KeyAscii <> 13 Then Exit Sub
    If Len(producto) = 0 Then
        producto.SetFocus
        Exit Sub

    End If

    mytablex.Open "select * from precio1  where codigo='" & codigo & "' and producto='" & producto & "' and local='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        MsgBox "Ya existe ", 48, "Aviso"
        mytablex.Close
        borra_campos
        producto = ""
        descripcio = ""
        producto.SetFocus
        Exit Sub

    End If

    mytablex.Close

    found = busca_productof("" & producto)

    If found = 0 Then
        borra_campos
        producto = ""
        descripcio = ""
        producto.SetFocus
        Exit Sub

    End If

    If Val(costou) = 0 Then
        producto = ""
        descripcio = ""
        borra_campos
        producto.SetFocus
        Exit Sub

    End If

    found = verifica_precio()
    calcula_margenes
    pventa1.SetFocus

End Sub

Function verifica_precio()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from precios  where producto='" & producto & "' and local='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        pone_precios mytablex

    End If

    mytablex.Close

End Function

