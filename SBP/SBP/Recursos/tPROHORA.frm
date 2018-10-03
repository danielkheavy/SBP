VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form tprohora 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Tabla de Programaciones Horarias"
   ClientHeight    =   6615
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   11490
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
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
      Height          =   6495
      Left            =   15
      TabIndex        =   68
      Top             =   645
      Visible         =   0   'False
      Width           =   11415
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
         TabIndex        =   71
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   240
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
         Left            =   8160
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "tPROHORA.frx":0000
         Height          =   5295
         Left            =   120
         OleObjectBlob   =   "tPROHORA.frx":0014
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   1080
         Width           =   11055
      End
   End
   Begin VB.TextBox tipo 
      BeginProperty Font 
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
      TabIndex        =   67
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox hora1 
      BeginProperty Font 
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
      MaxLength       =   6
      TabIndex        =   64
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox hora2 
      BeginProperty Font 
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
      MaxLength       =   6
      TabIndex        =   63
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox hora3 
      BeginProperty Font 
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
      MaxLength       =   6
      TabIndex        =   62
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox hora4 
      BeginProperty Font 
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
      MaxLength       =   6
      TabIndex        =   61
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox hora5 
      BeginProperty Font 
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
      MaxLength       =   6
      TabIndex        =   60
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox hora6 
      BeginProperty Font 
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
      MaxLength       =   6
      TabIndex        =   59
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox hora7 
      BeginProperty Font 
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
      MaxLength       =   6
      TabIndex        =   58
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox sdom1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      MaxLength       =   5
      TabIndex        =   54
      Text            =   "00:00"
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox edom1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      MaxLength       =   5
      TabIndex        =   53
      Text            =   "00:00"
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox ssab1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      MaxLength       =   5
      TabIndex        =   52
      Text            =   "00:00"
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox esab1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      MaxLength       =   5
      TabIndex        =   51
      Text            =   "00:00"
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox svie1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      MaxLength       =   5
      TabIndex        =   50
      Text            =   "00:00"
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox evie1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      MaxLength       =   5
      TabIndex        =   49
      Text            =   "00:00"
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox sjue1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      MaxLength       =   5
      TabIndex        =   48
      Text            =   "00:00"
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox ejue1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      MaxLength       =   5
      TabIndex        =   47
      Text            =   "00:00"
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox smie1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      MaxLength       =   5
      TabIndex        =   46
      Text            =   "00:00"
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox emie1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      MaxLength       =   5
      TabIndex        =   45
      Text            =   "00:00"
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox smar1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      MaxLength       =   5
      TabIndex        =   44
      Text            =   "00:00"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox emar1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      MaxLength       =   5
      TabIndex        =   43
      Text            =   "00:00"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox slun1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      MaxLength       =   5
      TabIndex        =   42
      Text            =   "00:00"
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox elun1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      MaxLength       =   5
      TabIndex        =   41
      Text            =   "00:00"
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox dome2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MaxLength       =   5
      TabIndex        =   40
      Text            =   "00:00"
      Top             =   5040
      Width           =   855
   End
   Begin VB.CheckBox cdom 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   840
      TabIndex        =   39
      Top             =   5040
      Width           =   375
   End
   Begin VB.TextBox dome1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   37
      Text            =   "00:00"
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox sabe2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MaxLength       =   5
      TabIndex        =   36
      Text            =   "00:00"
      Top             =   4680
      Width           =   855
   End
   Begin VB.CheckBox csab 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   840
      TabIndex        =   35
      Top             =   4680
      Width           =   375
   End
   Begin VB.TextBox sabe1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   33
      Text            =   "00:00"
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox viee2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MaxLength       =   5
      TabIndex        =   32
      Text            =   "00:00"
      Top             =   4320
      Width           =   855
   End
   Begin VB.CheckBox cvie 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   840
      TabIndex        =   31
      Top             =   4320
      Width           =   375
   End
   Begin VB.TextBox viee1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   29
      Text            =   "00:00"
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox juee2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MaxLength       =   5
      TabIndex        =   28
      Text            =   "00:00"
      Top             =   3960
      Width           =   855
   End
   Begin VB.CheckBox cjue 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   840
      TabIndex        =   27
      Top             =   3960
      Width           =   375
   End
   Begin VB.TextBox juee1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   25
      Text            =   "00:00"
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox miee2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MaxLength       =   5
      TabIndex        =   24
      Text            =   "00:00"
      Top             =   3600
      Width           =   855
   End
   Begin VB.CheckBox cmie 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   840
      TabIndex        =   23
      Top             =   3600
      Width           =   375
   End
   Begin VB.TextBox miee1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   21
      Text            =   "00:00"
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox mare2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MaxLength       =   5
      TabIndex        =   20
      Text            =   "00:00"
      Top             =   3240
      Width           =   855
   End
   Begin VB.CheckBox cmar 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   840
      TabIndex        =   19
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox mare1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   17
      Text            =   "00:00"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox lune2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MaxLength       =   5
      TabIndex        =   16
      Text            =   "00:00"
      Top             =   2880
      Width           =   855
   End
   Begin VB.CheckBox clun 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   840
      TabIndex        =   15
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox lune1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   13
      Text            =   "00:00"
      Top             =   2880
      Width           =   855
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox descripcio 
      BeginProperty Font 
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
      TabIndex        =   1
      Top             =   1200
      Width           =   5775
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
      MaxLength       =   6
      TabIndex        =   0
      Top             =   840
      Width           =   1335
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
      Picture         =   "tPROHORA.frx":09DF
      Style           =   1  'Graphical
      TabIndex        =   8
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
      Picture         =   "tPROHORA.frx":1BF1
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Nuevo registro"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdHelp 
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
      Left            =   3600
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tPROHORA.frx":2E03
      Style           =   1  'Graphical
      TabIndex        =   6
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
      Picture         =   "tPROHORA.frx":4015
      Style           =   1  'Graphical
      TabIndex        =   5
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
      Picture         =   "tPROHORA.frx":5227
      Style           =   1  'Graphical
      TabIndex        =   4
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
      Picture         =   "tPROHORA.frx":6439
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Borrar registro"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tPROHORA.frx":764B
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Consulta"
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2.Variable.El empleado debe cumplir un numero de Horas . Obligatorio Ingresar Nro.Horas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   66
      Top             =   1920
      Width           =   7815
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Horas"
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
      TabIndex        =   65
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Entrada Salida"
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
      TabIndex        =   57
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Entrada Salida"
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
      TabIndex        =   56
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripcion"
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
      Left            =   1200
      TabIndex        =   55
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Domingo"
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
      Left            =   1200
      TabIndex        =   38
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sabado"
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
      Left            =   1200
      TabIndex        =   34
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Viernes"
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
      Left            =   1200
      TabIndex        =   30
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Jueves"
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
      Left            =   1200
      TabIndex        =   26
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Miercoles"
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
      Left            =   1200
      TabIndex        =   22
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Martes"
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
      Left            =   1200
      TabIndex        =   18
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lunes"
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
      Left            =   1200
      TabIndex        =   14
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.Fijo. El empleado debe cumplir un horarario Entrada/Salida"
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
      TabIndex        =   12
      Top             =   1560
      Width           =   7335
   End
   Begin VB.Label Label2 
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
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripcion"
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
      TabIndex        =   10
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo.Prog.Hora"
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
      TabIndex        =   9
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
Attribute VB_Name = "tprohora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ajdu1_Click()

    If Frame1.Visible = True Then Exit Sub
    inicializa
    codigo = ""
    codigo.SetFocus

End Sub

Private Sub bo712_Click()

    Dim found As Integer

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

Private Sub buffer_DblClick()
    Command1_Click

End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        Frame1.Visible = False
        codigo.SetFocus
        Exit Sub

    End If

    Command1_Click

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

Private Sub cmdPrint_Click()
    djuer1_Click

End Sub

Private Sub cmdSave_Click()
    grba1_Click

End Sub

Private Sub cmdSort_Click()
    Frame1.Visible = True
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

    descripcio.SetFocus

End Sub

Private Sub Codigo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        cmdSort_Click

    End If

End Sub

Private Sub Command1_Click()

    Dim buf As String

    If Len(buffer) = 0 Then
        buf = "select Descripcio,tprohora from tprohora "
    Else
        buf = "select Descripcio,tprohora from tprohora where " & Combo1 & " like '" & buffer & "%'"

    End If

    Data1.Connect = "foxpro 2.5;"
    Data1.DatabaseName = globaldir
    Data1.RecordSource = buf
    Data1.refresh

    If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then
        Data1.Recordset.Close
        buffer.SetFocus
        Exit Sub

    End If

    dbGrid1.columns(0).Width = 4000
    dbGrid1.columns(1).Width = 2000
    dbGrid1.SetFocus

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
  
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        codigo = dbGrid1.columns(1)
        Frame1.Visible = False
        codigo.SetFocus
        codigo_KeyPress 13

    End If

End Sub

Private Sub descripcio_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    tipo.SetFocus

End Sub

Private Sub descripcio_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        codigo.SetFocus
        Exit Sub

    End If

End Sub

Private Sub djuer1_Click()

    If Frame1.Visible = True Then Exit Sub
    reporgen.NAMETABLA = "tprohora"
    reporgen.Show 1

End Sub

Private Sub dlo132_Click()

    If Frame1.Visible = True Then
        Frame1.Visible = False
        Exit Sub

    End If

    tprohora.Hide
    Unload tprohora

End Sub

Private Sub Form_Load()
    Combo1.Clear
    Combo1.AddItem "Descripcio"
    Combo1.AddItem "tprohora"
    Combo1.ListIndex = 0

End Sub

Sub inicializa()
    descripcio = ""
    tipo = ""
    clun.Value = 0
    cmar.Value = 0
    cmie.Value = 0
    cjue.Value = 0
    cvie.Value = 0
    csab.Value = 0
    cdom.Value = 0

    lune1 = "00:00"
    mare1 = "00:00"
    miee1 = "00:00"
    juee1 = "00:00"
    viee1 = "00:00"
    sabe1 = "00:00"
    dome1 = "00:00"

    lune2 = "00:00"
    mare2 = "00:00"
    miee2 = "00:00"
    juee2 = "00:00"
    viee2 = "00:00"
    sabe2 = "00:00"
    dome2 = "00:00"

    elun1 = "00:00"
    emar1 = "00:00"
    emie1 = "00:00"
    ejue1 = "00:00"
    evie1 = "00:00"
    esab1 = "00:00"
    edom1 = "00:00"

    slun1 = "00:00"
    smar1 = "00:00"
    smie1 = "00:00"
    sjue1 = "00:00"
    svie1 = "00:00"
    ssab1 = "00:00"
    sdom1 = "00:00"

    hora1 = ""
    hora2 = ""
    hora3 = ""
    hora4 = ""
    hora5 = ""
    hora6 = ""
    hora7 = ""

End Sub

Function borra_registro()

    Dim mytablex As Table

    Set mytablex = mydbxglo.OpenTable("tprohora")
    mytablex.Index = "tprohora"
    mytablex.Seek "=", codigo

    If Not mytablex.NoMatch Then
        If MsgBox("Desea Borra el registro", 1, "Aviso") = "1" Then
            mytablex.Delete
            borra_registro = 1

        End If

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_registro()

    Dim mytablex As Table

    Set mytablex = mydbxglo.OpenTable("tprohora")
    mytablex.Index = "tprohora"
    mytablex.Seek "=", codigo

    If Not mytablex.NoMatch Then
        pone_registro mytablex
        busca_registro = 1

    End If

    '------------------------------------- ------------
    mytablex.Close
 
End Function

Sub pone_registro(mytablex As Table)

    Dim found As Integer

    tipo = "" & mytablex.Fields("tipo")
    codigo = "" & mytablex.Fields("tprohora")
    descripcio = "" & mytablex.Fields("descripcio")

    clun.Value = Val("" & mytablex.Fields("clun"))
    cmar.Value = Val("" & mytablex.Fields("cmar"))
    cmie.Value = Val("" & mytablex.Fields("cmie"))
    cjue.Value = Val("" & mytablex.Fields("cjue"))
    cvie.Value = Val("" & mytablex.Fields("cvie"))
    csab.Value = Val("" & mytablex.Fields("csab"))
    cdom.Value = Val("" & mytablex.Fields("cdom"))

    lune1 = "" & mytablex.Fields("lune1")
    mare1 = "" & mytablex.Fields("mare1")
    miee1 = "" & mytablex.Fields("miee1")
    juee1 = "" & mytablex.Fields("juee1")
    viee1 = "" & mytablex.Fields("viee1")
    sabe1 = "" & mytablex.Fields("sabe1")
    dome1 = "" & mytablex.Fields("dome1")

    lune2 = "" & mytablex.Fields("lune2")
    mare2 = "" & mytablex.Fields("mare2")
    miee2 = "" & mytablex.Fields("miee2")
    juee2 = "" & mytablex.Fields("juee2")
    viee2 = "" & mytablex.Fields("viee2")
    sabe2 = "" & mytablex.Fields("sabe2")
    dome2 = "" & mytablex.Fields("dome2")

    elun1 = "" & mytablex.Fields("elun1")
    emar1 = "" & mytablex.Fields("emar1")
    emie1 = "" & mytablex.Fields("emie1")
    ejue1 = "" & mytablex.Fields("ejue1")
    evie1 = "" & mytablex.Fields("evie1")
    esab1 = "" & mytablex.Fields("esab1")
    edom1 = "" & mytablex.Fields("edom1")

    slun1 = "" & mytablex.Fields("slun1")
    smar1 = "" & mytablex.Fields("smar1")
    smie1 = "" & mytablex.Fields("smie1")
    sjue1 = "" & mytablex.Fields("sjue1")
    svie1 = "" & mytablex.Fields("svie1")
    ssab1 = "" & mytablex.Fields("ssab1")
    sdom1 = "" & mytablex.Fields("sdom1")

    hora1 = "" & mytablex.Fields("hora1")
    hora2 = "" & mytablex.Fields("hora2")
    hora3 = "" & mytablex.Fields("hora3")
    hora4 = "" & mytablex.Fields("hora4")
    hora5 = "" & mytablex.Fields("hora5")
    hora6 = "" & mytablex.Fields("hora6")
    hora7 = "" & mytablex.Fields("hora7")

    found = calcular_horas()

End Sub

Sub grabando(mytablex As Table)
    mytablex.Fields("tprohora") = codigo
    mytablex.Fields("descripcio") = descripcio
    mytablex.Fields("tipo") = tipo

    mytablex.Fields("clun") = clun.Value
    mytablex.Fields("cmar") = cmar.Value
    mytablex.Fields("cmie") = cmie.Value
    mytablex.Fields("cjue") = cjue.Value
    mytablex.Fields("csab") = csab.Value
    mytablex.Fields("cdom") = cdom.Value
    mytablex.Fields("cvie") = cvie.Value

    mytablex.Fields("lune1") = lune1
    mytablex.Fields("mare1") = mare1
    mytablex.Fields("miee1") = miee1
    mytablex.Fields("juee1") = juee1
    mytablex.Fields("viee1") = viee1
    mytablex.Fields("sabe1") = sabe1
    mytablex.Fields("dome1") = dome1

    mytablex.Fields("lune2") = lune2
    mytablex.Fields("mare2") = mare2
    mytablex.Fields("miee2") = miee2
    mytablex.Fields("juee2") = juee2
    mytablex.Fields("viee2") = viee2
    mytablex.Fields("sabe2") = sabe2
    mytablex.Fields("dome2") = dome2

    mytablex.Fields("elun1") = elun1
    mytablex.Fields("emar1") = emar1
    mytablex.Fields("emie1") = emie1
    mytablex.Fields("ejue1") = ejue1
    mytablex.Fields("evie1") = evie1
    mytablex.Fields("esab1") = esab1
    mytablex.Fields("edom1") = edom1

    mytablex.Fields("slun1") = slun1
    mytablex.Fields("smar1") = smar1
    mytablex.Fields("smie1") = smie1
    mytablex.Fields("sjue1") = sjue1
    mytablex.Fields("svie1") = svie1
    mytablex.Fields("ssab1") = ssab1
    mytablex.Fields("sdom1") = sdom1

    mytablex.Fields("hora1") = hora1
    mytablex.Fields("hora2") = hora2
    mytablex.Fields("hora3") = hora3
    mytablex.Fields("hora4") = hora4
    mytablex.Fields("hora5") = hora5
    mytablex.Fields("hora6") = hora6
    mytablex.Fields("hora7") = hora7

End Sub

Private Sub grba1_Click()

    Dim found As Integer

    If Frame1.Visible = True Then Exit Sub
    found = grabar()

    If found = 0 Then Exit Sub
    codigo.SetFocus

End Sub

Private Sub Label1_Click()
    cmdSort_Click

End Sub

Function grabar()

    Dim found    As Integer

    Dim mytablex As Table

    found = valida()

    If found = 0 Then
        MsgBox "Campos invalidos", 48, "Aviso"
        Exit Function

    End If

    Set mytablex = mydbxglo.OpenTable("tprohora")
    mytablex.Index = "tprohora"
    mytablex.Seek "=", codigo

    If mytablex.NoMatch Then
        mytablex.AddNew
        grabando mytablex
        mytablex.Update
        grabar = 1

    End If

    If Not mytablex.NoMatch Then
        If MsgBox("Desea Reescribir?", 1, "Aviso") = 1 Then
            mytablex.Edit
            grabando mytablex
            mytablex.Update
            grabar = 1

        End If

    End If

    '------------------------------------- ------------
    mytablex.Close
 
End Function

Function valida()

    Dim found As Integer

    If Len(codigo) = 0 Then
        codigo.SetFocus
        Exit Function

    End If

    If Len(descripcio) = 0 Then
        descripcio.SetFocus
        Exit Function

    End If

    found = calcular_horas()
    valida = 1

End Function

Function calcular_horas()

    Dim horax As String

    Dim horay As String

    On Error GoTo cmd345_err

    If tipo = "2" Then Exit Function
    If Val(Mid$(lune2, 1, 2)) >= 0 And Val(Mid$(lune2, 1, 2)) <= 24 And Val(Mid$(lune2, 4, 2)) >= 0 And Val(Mid$(lune2, 4, 2)) <= 60 Then
        If Val(Mid$(lune1, 1, 2)) >= 0 And Val(Mid$(lune1, 1, 2)) <= 24 And Val(Mid$(lune1, 4, 2)) >= 0 And Val(Mid$(lune1, 4, 2)) <= 60 Then
            If Val(Mid$(elun1, 1, 2)) >= 0 And Val(Mid$(elun1, 1, 2)) <= 24 And Val(Mid$(elun1, 4, 2)) >= 0 And Val(Mid$(elun1, 4, 2)) <= 60 Then
                If Val(Mid$(slun1, 1, 2)) >= 0 And Val(Mid$(slun1, 1, 2)) <= 24 And Val(Mid$(slun1, 4, 2)) >= 0 And Val(Mid$(slun1, 4, 2)) <= 60 Then
                    horax = Format(TimeValue(lune2) - TimeValue(lune1), "hh:mm")
                    horay = Format(TimeValue(slun1) - TimeValue(elun1), "hh:mm")
                    hora1 = Format(TimeValue(horax) + TimeValue(horay), "hh:mm")

                End If

            End If

        End If

    End If

    If Val(Mid$(mare2, 1, 2)) >= 0 And Val(Mid$(mare2, 1, 2)) <= 24 And Val(Mid$(mare2, 4, 2)) >= 0 And Val(Mid$(mare2, 4, 2)) <= 60 Then
        If Val(Mid$(mare1, 1, 2)) >= 0 And Val(Mid$(mare1, 1, 2)) <= 24 And Val(Mid$(mare1, 4, 2)) >= 0 And Val(Mid$(mare1, 4, 2)) <= 60 Then
            If Val(Mid$(emar1, 1, 2)) >= 0 And Val(Mid$(emar1, 1, 2)) <= 24 And Val(Mid$(emar1, 4, 2)) >= 0 And Val(Mid$(emar1, 4, 2)) <= 60 Then
                If Val(Mid$(smar1, 1, 2)) >= 0 And Val(Mid$(smar1, 1, 2)) <= 24 And Val(Mid$(smar1, 4, 2)) >= 0 And Val(Mid$(smar1, 4, 2)) <= 60 Then
                    horax = Format(TimeValue(mare2) - TimeValue(mare1), "hh:mm")
                    horay = Format(TimeValue(smar1) - TimeValue(emar1), "hh:mm")
                    hora2 = Format(TimeValue(horax) + TimeValue(horay), "hh:mm")

                End If

            End If

        End If

    End If

    If Val(Mid$(miee2, 1, 2)) >= 0 And Val(Mid$(miee2, 1, 2)) <= 24 And Val(Mid$(miee2, 4, 2)) >= 0 And Val(Mid$(miee2, 4, 2)) <= 60 Then
        If Val(Mid$(miee1, 1, 2)) >= 0 And Val(Mid$(miee1, 1, 2)) <= 24 And Val(Mid$(miee1, 4, 2)) >= 0 And Val(Mid$(miee1, 4, 2)) <= 60 Then
            If Val(Mid$(emie1, 1, 2)) >= 0 And Val(Mid$(emie1, 1, 2)) <= 24 And Val(Mid$(emie1, 4, 2)) >= 0 And Val(Mid$(emie1, 4, 2)) <= 60 Then
                If Val(Mid$(smie1, 1, 2)) >= 0 And Val(Mid$(smie1, 1, 2)) <= 24 And Val(Mid$(smie1, 4, 2)) >= 0 And Val(Mid$(smie1, 4, 2)) <= 60 Then
                    horax = Format(TimeValue(miee2) - TimeValue(miee1), "hh:mm")
                    horay = Format(TimeValue(smie1) - TimeValue(emie1), "hh:mm")
                    hora3 = Format(TimeValue(horax) + TimeValue(horay), "hh:mm")

                End If

            End If

        End If

    End If

    If Val(Mid$(juee2, 1, 2)) >= 0 And Val(Mid$(juee2, 1, 2)) <= 24 And Val(Mid$(juee2, 4, 2)) >= 0 And Val(Mid$(juee2, 4, 2)) <= 60 Then
        If Val(Mid$(juee1, 1, 2)) >= 0 And Val(Mid$(juee1, 1, 2)) <= 24 And Val(Mid$(juee1, 4, 2)) >= 0 And Val(Mid$(juee1, 4, 2)) <= 60 Then
            If Val(Mid$(juee1, 1, 2)) >= 0 And Val(Mid$(juee1, 1, 2)) <= 24 And Val(Mid$(juee1, 4, 2)) >= 0 And Val(Mid$(juee1, 4, 2)) <= 60 Then
                If Val(Mid$(sjue1, 1, 2)) >= 0 And Val(Mid$(sjue1, 1, 2)) <= 24 And Val(Mid$(sjue1, 4, 2)) >= 0 And Val(Mid$(sjue1, 4, 2)) <= 60 Then
                    horax = Format(TimeValue(juee2) - TimeValue(juee1), "hh:mm")
                    horay = Format(TimeValue(sjue1) - TimeValue(ejue1), "hh:mm")
                    hora4 = Format(TimeValue(horax) + TimeValue(horay), "hh:mm")

                End If

            End If

        End If

    End If

    If Val(Mid$(viee2, 1, 2)) >= 0 And Val(Mid$(viee2, 1, 2)) <= 24 And Val(Mid$(viee2, 4, 2)) >= 0 And Val(Mid$(viee2, 4, 2)) <= 60 Then
        If Val(Mid$(viee1, 1, 2)) >= 0 And Val(Mid$(viee1, 1, 2)) <= 24 And Val(Mid$(viee1, 4, 2)) >= 0 And Val(Mid$(viee1, 4, 2)) <= 60 Then
            If Val(Mid$(viee1, 1, 2)) >= 0 And Val(Mid$(viee1, 1, 2)) <= 24 And Val(Mid$(viee1, 4, 2)) >= 0 And Val(Mid$(viee1, 4, 2)) <= 60 Then
                If Val(Mid$(svie1, 1, 2)) >= 0 And Val(Mid$(svie1, 1, 2)) <= 24 And Val(Mid$(svie1, 4, 2)) >= 0 And Val(Mid$(svie1, 4, 2)) <= 60 Then
                    horax = Format(TimeValue(viee2) - TimeValue(viee1), "hh:mm")
                    horay = Format(TimeValue(svie1) - TimeValue(evie1), "hh:mm")
                    hora5 = Format(TimeValue(horax) + TimeValue(horay), "hh:mm")

                End If

            End If

        End If

    End If

    If Val(Mid$(sabe2, 1, 2)) >= 0 And Val(Mid$(sabe2, 1, 2)) <= 24 And Val(Mid$(sabe2, 4, 2)) >= 0 And Val(Mid$(sabe2, 4, 2)) <= 60 Then
        If Val(Mid$(sabe1, 1, 2)) >= 0 And Val(Mid$(sabe1, 1, 2)) <= 24 And Val(Mid$(sabe1, 4, 2)) >= 0 And Val(Mid$(sabe1, 4, 2)) <= 60 Then
            If Val(Mid$(sabe1, 1, 2)) >= 0 And Val(Mid$(sabe1, 1, 2)) <= 24 And Val(Mid$(sabe1, 4, 2)) >= 0 And Val(Mid$(sabe1, 4, 2)) <= 60 Then
                If Val(Mid$(ssab1, 1, 2)) >= 0 And Val(Mid$(ssab1, 1, 2)) <= 24 And Val(Mid$(ssab1, 4, 2)) >= 0 And Val(Mid$(ssab1, 4, 2)) <= 60 Then
                    horax = Format(TimeValue(sabe2) - TimeValue(sabe1), "hh:mm")
                    horay = Format(TimeValue(ssab1) - TimeValue(esab1), "hh:mm")
                    hora6 = Format(TimeValue(horax) + TimeValue(horay), "hh:mm")

                End If

            End If

        End If

    End If

    If Val(Mid$(dome2, 1, 2)) >= 0 And Val(Mid$(dome2, 1, 2)) <= 24 And Val(Mid$(dome2, 4, 2)) >= 0 And Val(Mid$(dome2, 4, 2)) <= 60 Then
        If Val(Mid$(dome1, 1, 2)) >= 0 And Val(Mid$(dome1, 1, 2)) <= 24 And Val(Mid$(dome1, 4, 2)) >= 0 And Val(Mid$(dome1, 4, 2)) <= 60 Then
            If Val(Mid$(dome1, 1, 2)) >= 0 And Val(Mid$(dome1, 1, 2)) <= 24 And Val(Mid$(dome1, 4, 2)) >= 0 And Val(Mid$(dome1, 4, 2)) <= 60 Then
                If Val(Mid$(sdom1, 1, 2)) >= 0 And Val(Mid$(sdom1, 1, 2)) <= 24 And Val(Mid$(sdom1, 4, 2)) >= 0 And Val(Mid$(sdom1, 4, 2)) <= 60 Then
                    horax = Format(TimeValue(dome2) - TimeValue(dome1), "hh:mm")
                    horay = Format(TimeValue(sdom1) - TimeValue(edom1), "hh:mm")
                    hora7 = Format(TimeValue(horax) + TimeValue(horay), "hh:mm")

                End If

            End If

        End If

    End If

    calcular_horas = 1
    Exit Function
cmd345_err:
    Exit Function

End Function

