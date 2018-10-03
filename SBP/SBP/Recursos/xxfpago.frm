VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form xxfpago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Forma Pago"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   13770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   13770
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data9 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   11760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Width           =   1140
   End
   Begin VB.Data Data10 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   11760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Framefp 
      BackColor       =   &H00FFFF00&
      Caption         =   "COBRANZAS"
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
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   11655
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFF00&
         Height          =   1215
         Left            =   7080
         ScaleHeight     =   1155
         ScaleWidth      =   4395
         TabIndex        =   4
         Top             =   240
         Width           =   4455
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Graba"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8640
         MaskColor       =   &H00E0E0E0&
         Picture         =   "xxfpago.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Grabar registro"
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
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
         Height          =   735
         Left            =   7200
         MaskColor       =   &H00E0E0E0&
         Picture         =   "xxfpago.frx":1212
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salir"
         Top             =   480
         Width           =   1335
      End
      Begin MSDBGrid.DBGrid DBGrid9 
         Bindings        =   "xxfpago.frx":2424
         Height          =   3375
         Left            =   120
         OleObjectBlob   =   "xxfpago.frx":2438
         TabIndex        =   1
         Top             =   2400
         Width           =   6855
      End
      Begin MSDBGrid.DBGrid DBGrid10 
         Bindings        =   "xxfpago.frx":330B
         Height          =   5175
         Left            =   7080
         OleObjectBlob   =   "xxfpago.frx":3320
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2400
         Width           =   4455
      End
      Begin VB.Label stxtotald 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   855
         Left            =   2640
         TabIndex        =   19
         Top             =   6720
         Width           =   4335
      End
      Begin VB.Label Label52 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  US$"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   2160
         TabIndex        =   18
         Top             =   6720
         Width           =   495
      End
      Begin VB.Label stxtotals 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   855
         Left            =   2640
         TabIndex        =   17
         Top             =   5880
         Width           =   4335
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   855
         Left            =   2160
         TabIndex        =   16
         Top             =   5880
         Width           =   495
      End
      Begin VB.Label paridadfp 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   8400
         TabIndex        =   15
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "T/Cambio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   7080
         TabIndex        =   14
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label ttxtotald 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   855
         Left            =   2640
         TabIndex        =   13
         Top             =   1080
         Width           =   4335
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  US$"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   2160
         TabIndex        =   12
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label ttxtotals 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   855
         Left            =   2640
         TabIndex        =   11
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   855
         Left            =   2160
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FALTA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1695
         Left            =   120
         TabIndex        =   9
         Top             =   5880
         Width           =   2055
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1695
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SELECCIONA TIPO DE PAGO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7080
         TabIndex        =   7
         Top             =   2040
         Width           =   4455
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "COMO PAGA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   6855
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   960
         Left            =   360
         Picture         =   "xxfpago.frx":41EC
         Stretch         =   -1  'True
         Top             =   840
         Width           =   1560
      End
      Begin VB.Image Image3 
         BorderStyle     =   1  'Fixed Single
         Height          =   975
         Left            =   360
         Picture         =   "xxfpago.frx":4AB2
         Stretch         =   -1  'True
         Top             =   6480
         Width           =   1680
      End
   End
   Begin VB.Menu lfo23 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "xxfpago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
