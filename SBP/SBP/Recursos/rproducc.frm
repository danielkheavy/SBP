VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form rproducc 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Insumos+Mano Obra"
   ClientHeight    =   10845
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   14775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10845
   ScaleWidth      =   14775
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "Lineas"
      Height          =   2055
      Left            =   13440
      TabIndex        =   90
      Top             =   3720
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Label xt16 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2280
         TabIndex        =   106
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label xt15 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1560
         TabIndex        =   105
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label xt14 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   840
         TabIndex        =   104
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label xt13 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   103
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label xt12 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2280
         TabIndex        =   102
         Top             =   960
         Width           =   735
      End
      Begin VB.Label xt11 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1560
         TabIndex        =   101
         Top             =   960
         Width           =   735
      End
      Begin VB.Label xt10 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   840
         TabIndex        =   100
         Top             =   960
         Width           =   735
      End
      Begin VB.Label xt9 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   99
         Top             =   960
         Width           =   735
      End
      Begin VB.Label xt8 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2280
         TabIndex        =   98
         Top             =   600
         Width           =   735
      End
      Begin VB.Label xt7 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1560
         TabIndex        =   97
         Top             =   600
         Width           =   735
      End
      Begin VB.Label xt6 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   840
         TabIndex        =   96
         Top             =   600
         Width           =   735
      End
      Begin VB.Label xt5 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   95
         Top             =   600
         Width           =   735
      End
      Begin VB.Label xt4 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2280
         TabIndex        =   94
         Top             =   240
         Width           =   735
      End
      Begin VB.Label xt3 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1560
         TabIndex        =   93
         Top             =   240
         Width           =   735
      End
      Begin VB.Label xt2 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   840
         TabIndex        =   92
         Top             =   240
         Width           =   735
      End
      Begin VB.Label xt1 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   91
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
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
      Height          =   7215
      Left            =   11160
      TabIndex        =   81
      Top             =   1920
      Visible         =   0   'False
      Width           =   12255
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
         TabIndex        =   84
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
         TabIndex        =   83
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
         Left            =   7560
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "rproducc.frx":0000
         Height          =   6255
         Left            =   120
         OleObjectBlob   =   "rproducc.frx":0014
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   840
         Width           =   11895
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00E0E0E0&
      Height          =   680
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   14715
      TabIndex        =   79
      Top             =   0
      Width           =   14775
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
         Left            =   0
         MaskColor       =   &H00E0E0E0&
         Picture         =   "rproducc.frx":09DF
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Salir"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos Partes Personal"
      Height          =   3375
      Left            =   6960
      TabIndex        =   51
      Top             =   5760
      Visible         =   0   'False
      Width           =   4335
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
         Left            =   3000
         MaskColor       =   &H00E0E0E0&
         Picture         =   "rproducc.frx":1BF1
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Salir"
         Top             =   960
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
         Left            =   3000
         MaskColor       =   &H00E0E0E0&
         Picture         =   "rproducc.frx":2E03
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Grabar registro"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox horaf 
         Height          =   375
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   56
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox Horai 
         Height          =   375
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   55
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox operacion 
         Height          =   375
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   54
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox operario 
         Height          =   375
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   53
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox fecha 
         Height          =   375
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   52
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label solhora 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   71
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SolxHora"
         Height          =   375
         Left            =   120
         TabIndex        =   70
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nro.Horas"
         Height          =   375
         Left            =   120
         TabIndex        =   69
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label nrohora 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   68
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label seccion 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   65
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HoraFinal"
         Height          =   375
         Left            =   120
         TabIndex        =   64
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HoraInicio"
         Height          =   375
         Left            =   120
         TabIndex        =   63
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Operacion"
         Height          =   375
         Left            =   120
         TabIndex        =   62
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Seccion"
         Height          =   375
         Left            =   120
         TabIndex        =   61
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Operario"
         Height          =   375
         Left            =   120
         TabIndex        =   60
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   59
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7320
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
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7320
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Ingreso de Lineas"
      Height          =   3255
      Left            =   5520
      TabIndex        =   8
      Top             =   8400
      Visible         =   0   'False
      Width           =   7335
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
         Left            =   5400
         MaskColor       =   &H00E0E0E0&
         Picture         =   "rproducc.frx":4015
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Borrar registro"
         Top             =   2400
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
         Left            =   6240
         MaskColor       =   &H00E0E0E0&
         Picture         =   "rproducc.frx":5227
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Grabar registro"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox t16 
         Height          =   375
         Left            =   6240
         MaxLength       =   10
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox t15 
         Height          =   375
         Left            =   6240
         MaxLength       =   10
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox t14 
         Height          =   375
         Left            =   6240
         MaxLength       =   10
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox t13 
         Height          =   375
         Left            =   6240
         MaxLength       =   10
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox t12 
         Height          =   375
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox t11 
         Height          =   375
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox t10 
         Height          =   375
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox t9 
         Height          =   375
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox t8 
         Height          =   375
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox t7 
         Height          =   375
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox t6 
         Height          =   375
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox t5 
         Height          =   375
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox t4 
         Height          =   375
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox t3 
         Height          =   375
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox t2 
         Height          =   375
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox t1 
         Height          =   375
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.Label linea 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4440
         TabIndex        =   49
         Top             =   360
         Width           =   855
      End
      Begin VB.Label nt16 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5520
         TabIndex        =   48
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label nt15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5520
         TabIndex        =   47
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label nt14 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5520
         TabIndex        =   46
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label nt13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5520
         TabIndex        =   45
         Top             =   840
         Width           =   735
      End
      Begin VB.Label nt12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   44
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label nt11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   43
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label nt10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   42
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label nt9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   41
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label30 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   40
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label29 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   39
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label28 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   38
         Top             =   840
         Width           =   735
      End
      Begin VB.Label nt8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2640
         TabIndex        =   37
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label nt7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2640
         TabIndex        =   36
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label nt6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2640
         TabIndex        =   35
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label nt5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2640
         TabIndex        =   34
         Top             =   840
         Width           =   735
      End
      Begin VB.Label nt4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   33
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label nt3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   32
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label nt2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   31
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label nt1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   30
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tallas"
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   840
         Width           =   975
      End
      Begin VB.Label nlinea 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   28
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Linea"
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Data Data2 
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
      Top             =   3000
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "rproducc.frx":6439
      Height          =   2775
      Left            =   120
      OleObjectBlob   =   "rproducc.frx":644D
      TabIndex        =   0
      Top             =   1920
      Width           =   9135
   End
   Begin MSDBGrid.DBGrid DBGrid4 
      Bindings        =   "rproducc.frx":79B8
      Height          =   2415
      Left            =   120
      OleObjectBlob   =   "rproducc.frx":79CC
      TabIndex        =   50
      Top             =   4800
      Width           =   9135
   End
   Begin VB.Label VIENE 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   9360
      TabIndex        =   107
      Top             =   1440
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Linea"
      Height          =   375
      Left            =   6480
      TabIndex        =   89
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label xlinea 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7920
      TabIndex        =   88
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label cantidad 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7920
      TabIndex        =   87
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cantidad"
      Height          =   375
      Left            =   6480
      TabIndex        =   86
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Receta"
      Height          =   375
      Left            =   3480
      TabIndex        =   78
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripcio"
      Height          =   375
      Left            =   120
      TabIndex        =   77
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tarjeta"
      Height          =   375
      Left            =   120
      TabIndex        =   76
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label tarjeta 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1560
      TabIndex        =   75
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NumeroMapa"
      Height          =   375
      Left            =   120
      TabIndex        =   74
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label ntotal 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Height          =   495
      Left            =   9240
      TabIndex        =   73
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Costo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   72
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mano Obra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   67
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label total1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Height          =   495
      Left            =   9240
      TabIndex        =   66
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label numero 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label total 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Height          =   495
      Left            =   9240
      TabIndex        =   6
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Producto"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label nro 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label descripcio 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1440
      Width           =   4815
   End
   Begin VB.Label producto 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Materiales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   1
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Menu pami7823 
      Caption         =   "&Partes"
   End
   Begin VB.Menu ldo343 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "rproducc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddEntry_Click()
Frame3.Visible = True
fecha.SetFocus
End Sub

Private Sub buffer_DblClick()
Command1_Click
End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   ldo343_Click
   Exit Sub
End If
Command1_Click

End Sub

Private Sub cmdExit_Click()
Frame3.Visible = False
End Sub

Private Sub cmdSave_Click()
Dim found As Integer
found = valida_fecha("" & fecha)
If found = 0 Then
   fecha = ""
   fecha.SetFocus
   Exit Sub
End If
found = busca_operario("" & operario)
If found = 0 Then
   operario = ""
   operario.SetFocus
   Exit Sub
End If
found = busca_operacion("" & operacion)
If found = 0 Then
   operacion = ""
   operacion.SetFocus
   Exit Sub
End If
found = valida_hora("" & Horai)
If found = 0 Then
   Horai = ""
   Horai.SetFocus
   Exit Sub
End If
found = valida_hora("" & horaf)
If found = 0 Then
   horaf = ""
   horaf.SetFocus
   Exit Sub
End If
nrohora = Format(TimeValue(horaf) - TimeValue(Horai), "hh:mm")
'If Val(nrohora) <= 0 Then
'   Horai.SetFocus
'   Exit Sub
'End If
found = grabar()
cmdExit_Click
End Sub
Function grabar()
If opcion2 = "1" Then
Data3.Recordset.AddNew
Data3.Recordset.Fields("numero") = numero
Data3.Recordset.Fields("producto") = producto
Data3.Recordset.Fields("fecha") = fecha
Data3.Recordset.Fields("operario") = operario
Data3.Recordset.Fields("seccion") = seccion
Data3.Recordset.Fields("solhora") = Val(solhora)
Data3.Recordset.Fields("operacion") = operacion
Data3.Recordset.Fields("horai") = Horai
Data3.Recordset.Fields("horaf") = horaf
Data3.Recordset.Fields("nrohora") = nrohora
Data3.Recordset.Fields("total") = Val(nrohora) * Val(solhora)
Data3.Recordset.Update
tpartes.Data2.Recordset.Edit
tpartes.Data2.Recordset.Fields("seccion") = seccion
tpartes.Data2.Recordset.Update
End If
If opcion2 = "2" Then
Data3.Recordset.Edit
Data3.Recordset.Fields("producto") = producto
Data3.Recordset.Fields("numero") = numero
Data3.Recordset.Fields("fecha") = fecha
Data3.Recordset.Fields("operario") = operario
Data3.Recordset.Fields("seccion") = seccion
Data3.Recordset.Fields("solhora") = Val(solhora)
Data3.Recordset.Fields("operacion") = operacion
Data3.Recordset.Fields("horai") = Horai
Data3.Recordset.Fields("horaf") = horaf
Data3.Recordset.Fields("nrohora") = nrohora
Data3.Recordset.Fields("total") = Val(nrohora) * Val(solhora)
Data3.Recordset.Update
End If
End Function

Private Sub cmdSort_Click()

End Sub

Private Sub Command1_Click()
If opcion1 = "0" Then
   If Len(buffer) = 0 Then
   buf = "select Nombre,Codigo,Jornal,Seccion from vendedor "
   Else
   buf = "select Nombre,Codigo,Jornal,Seccion from vendedor where " & Combo1 & " like '" & buffer & "%'"
   End If
End If
If opcion1 = "1" Then
   If Len(buffer) = 0 Then
   buf = "select Descripcio,Operacion from toperaco "
   Else
   buf = "select Descripcio,Operacion from toperaco where " & Combo1 & " like '" & buffer & "%'"
   End If
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

Private Sub Command2_Click()
Dim sdx As Double
If opcion1 = "1" Then
Data2.Recordset.Edit
Data2.Recordset.Fields("t1") = Val("" & t1)
Data2.Recordset.Fields("t2") = Val("" & t2)
Data2.Recordset.Fields("t3") = Val("" & t3)
Data2.Recordset.Fields("t4") = Val("" & t4)
Data2.Recordset.Fields("t5") = Val("" & t5)
Data2.Recordset.Fields("t6") = Val("" & t6)
Data2.Recordset.Fields("t7") = Val("" & t7)
Data2.Recordset.Fields("t8") = Val("" & t8)
Data2.Recordset.Fields("t9") = Val("" & t9)
Data2.Recordset.Fields("t10") = Val("" & t10)
Data2.Recordset.Fields("t11") = Val("" & t11)
Data2.Recordset.Fields("t12") = Val("" & t12)
Data2.Recordset.Fields("t13") = Val("" & t13)
Data2.Recordset.Fields("t14") = Val("" & t14)
Data2.Recordset.Fields("t15") = Val("" & t15)
Data2.Recordset.Fields("t16") = Val("" & t16)
sdx = Val("" & t1) + Val("" & t2) + Val("" & t3) + Val("" & t4) + Val("" & t5) + Val("" & t6) + Val("" & t7) + Val("" & t8) + Val("" & t9) + Val("" & t10) + Val("" & t11) + Val("" & t12) + Val("" & t13) + Val("" & t14) + Val("" & t15) + Val("" & t16)
Data2.Recordset.Fields("cantidad") = sdx
Data2.Recordset.Update
End If
If opcion1 = "2" Then
Data2.Recordset.Edit
Data2.Recordset.Fields("tm1") = Val("" & t1)
Data2.Recordset.Fields("tm2") = Val("" & t2)
Data2.Recordset.Fields("tm3") = Val("" & t3)
Data2.Recordset.Fields("tm4") = Val("" & t4)
Data2.Recordset.Fields("tm5") = Val("" & t5)
Data2.Recordset.Fields("tm6") = Val("" & t6)
Data2.Recordset.Fields("tm7") = Val("" & t7)
Data2.Recordset.Fields("tm8") = Val("" & t8)
Data2.Recordset.Fields("tm9") = Val("" & t9)
Data2.Recordset.Fields("tm10") = Val("" & t10)
Data2.Recordset.Fields("tm11") = Val("" & t11)
Data2.Recordset.Fields("tm12") = Val("" & t12)
Data2.Recordset.Fields("tm13") = Val("" & t13)
Data2.Recordset.Fields("tm14") = Val("" & t14)
Data2.Recordset.Fields("tm15") = Val("" & t15)
Data2.Recordset.Fields("tm16") = Val("" & t16)
sdx = Val("" & t1) + Val("" & t2) + Val("" & t3) + Val("" & t4) + Val("" & t5) + Val("" & t6) + Val("" & t7) + Val("" & t8) + Val("" & t9) + Val("" & t10) + Val("" & t11) + Val("" & t12) + Val("" & t13) + Val("" & t14) + Val("" & t15) + Val("" & t16)
Data2.Recordset.Fields("merma") = sdx
Data2.Recordset.Update
End If
suma_linea
ldo343_Click

End Sub

Private Sub Command3_Click()
ldo343_Click
End Sub

Private Sub Command4_Click()
ldo343_Click
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim found As Integer
If KeyCode = 27 Then
   buffer.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
   If opcion1 = "0" Then
   operario = "" & Data1.Recordset.Fields("codigo")
   seccion = "" & Data1.Recordset.Fields("seccion")
   solhora = "" & Data1.Recordset.Fields("jornal")
   Frame1.Visible = False
   operario.Enabled = True
   operario.SetFocus
   operario_KeyPress 13
   End If
   If opcion1 = "1" Then
   operacion = "" & Data1.Recordset.Fields("operacion")
   Frame1.Visible = False
   operacion.SetFocus
   operacion_KeyPress 13
   End If

End If
   

End Sub

Private Sub DBGrid2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H71 Then  'f2
   If Len(DBGrid2.columns(0)) > 0 And Len(DBGrid2.columns(8)) > 0 And DBGrid2.Col = 4 Then
      opcion1 = "1"
      ingreso_tallas "" & DBGrid2.columns(8)
   End If
   If Len(DBGrid2.columns(0)) > 0 And Len(DBGrid2.columns(8)) > 0 And DBGrid2.Col = 5 Then
      opcion1 = "2"
      ingreso_tallas "" & DBGrid2.columns(8)
   End If
End If
End Sub
Private Sub DBGrid2_AfterColUpdate(ByVal ColIndex As Integer)
If ColIndex <> 4 And ColIndex <> 5 And ColIndex <> 6 Then
   Cancel = True
   Exit Sub
End If
Select Case ColIndex
       Case 4, 5, 6
          If Len("" & DBGrid2.columns(0)) = 0 Then
             Cancel = True
             Exit Sub
          End If
          
End Select
End Sub

Private Sub DBGrid2_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
Dim found As Integer
If ColIndex <> 4 And ColIndex <> 5 And ColIndex <> 6 Then
   Cancel = True
   Exit Sub
End If
Select Case ColIndex
     Case 4, 5
     If Len("" & DBGrid2.columns(0)) = 0 Then
             Cancel = True
             Exit Sub
          End If
    If Len("" & DBGrid2.columns(8)) >= 0 Then
             Cancel = True
             Exit Sub
          End If
     Case 6
          If Len("" & DBGrid2.columns(0)) = 0 Then
             Cancel = True
             Exit Sub
          End If
          
          
End Select
          
End Sub

Private Sub DBGrid2_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim found As Integer
Select Case ColIndex
     
   Case 4
        If Val("" & DBGrid2.columns(4)) = 0 Then
           Cancel = True
           Exit Sub
        End If
        If Not IsNumeric(DBGrid2.columns(4)) Then
           Cancel = True
           Exit Sub
        End If
        suma_linea
Case 5
        If Val("" & DBGrid2.columns(5)) = 0 Then
           Cancel = True
           Exit Sub
        End If
        If Not IsNumeric(DBGrid2.columns(5)) Then
           Cancel = True
           Exit Sub
        End If
        suma_linea
Case 6
        If Val("" & DBGrid2.columns(6)) = 0 Then
           Cancel = True
           Exit Sub
        End If
        If Not IsNumeric(DBGrid2.columns(6)) Then
           Cancel = True
           Exit Sub
        End If
        suma_linea
   
End Select

End Sub


Private Sub DBGrid4_DblClick()
On Error GoTo cmd23_err
opcion2 = "2"
fecha = "" & Data3.Recordset.Fields("fecha")
operario = "" & Data3.Recordset.Fields("operario")
seccion = "" & Data3.Recordset.Fields("seccion")
operacion = "" & Data3.Recordset.Fields("operacion")
Horai = "" & Data3.Recordset.Fields("horai")
horaf = "" & Data3.Recordset.Fields("horaf")
Frame3.Visible = True
fecha.SetFocus
Exit Sub
cmd23_err:
Exit Sub
End Sub

Private Sub DBGrid4_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo cmd46_err
If KeyCode = &H2E Then  'borrar linea
   Data3.Recordset.Delete
   Data3.refresh
   Exit Sub
End If
Exit Sub
cmd46_err:
Exit Sub

End Sub

Private Sub fecha_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If Len(fecha) = 0 Then
   fecha = Format(Now, "dd/mm/yyyy")
End If
If Len(fecha) <> 10 Then Exit Sub
If Not IsDate(fecha) Then Exit Sub
operario.SetFocus

End Sub

Private Sub Form_Activate()
sql_insumo

End Sub

Private Sub horaf_KeyPress(KeyAscii As Integer)
Dim found As Integer
On Error GoTo cmd32_err
If KeyAscii <> 13 Then Exit Sub
If Len(horaf) = 0 Then
   horaf = Format(Now, "hh:mm")
End If
If Len(horaf) <> 5 Then Exit Sub
found = valida_hora("" & horaf)
If found = 0 Then
   horaf = ""
   Exit Sub
End If
nrohora = Format(TimeValue(horaf) - TimeValue(Horai), "hh:mm")
Exit Sub
cmd32_err:
Exit Sub

End Sub

Private Sub horaf_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   Horai.SetFocus
   Exit Sub
End If

End Sub

Private Sub horai_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Len(Horai) = 0 Then
   Horai = Format(Now, "hh:mm")
End If
If Len(Horai) <> 5 Then Exit Sub
found = valida_hora("" & Horai)
If found = 0 Then
   Horai = ""
   Exit Sub
End If
horaf.SetFocus

End Sub

Private Sub horai_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   operacion.SetFocus
   Exit Sub
End If

End Sub

Private Sub Label15_Click()
Label2_Click
End Sub

Private Sub Label2_Click()
Dim sdx As Double

sumar_total
sumar_total1
sdx = (Val(total) + Val(total1))
ntotal = "" & sdx
If VIENE = "S" Then
   If Val("" & cantidad) > 0 Then
   tpartes.Data2.Recordset.Edit
   tpartes.Data2.Recordset.Fields("precio") = (Val(total) + Val(total1)) / Val("" & tpartes.Data2.Recordset.Fields("cantidad"))
   tpartes.Data2.Recordset.Fields("TOTAL") = Val("" & tpartes.Data2.Recordset.Fields("precio")) * Val("" & tpartes.Data2.Recordset.Fields("cantidad"))
   tpartes.Data2.Recordset.Update
End If
End If
End Sub

Private Sub Label4_Click()
Dim found As Integer
found = carga_receta("" & nro, "" & producto)
End Sub

Private Sub ldo343_Click()
If Frame2.Visible = True Then
   Frame2.Visible = False
   Exit Sub
End If
If Frame3.Visible = True Then
   Frame3.Visible = False
   DBGrid2.SetFocus
   Exit Sub
End If
If opcion1 = "0" Then
   If Frame3.Visible = True Then
      Frame3.Visible = False
      DBGrid2.SetFocus
      Exit Sub
   End If
End If

rproducc.Hide
Unload rproducc
End Sub
Function carga_receta(buf As String, buf1 As String)

Dim mytablex As Table
borra_anterior

Set mytablex = mydbxglo.OpenTable("receta")
mytablex.Index = "receta1"
mytablex.Seek "=", buf, buf1
If mytablex.NoMatch Then
   mytablex.Close
    
   Exit Function
End If
Do
If mytablex.EOF Then Exit Do
If "" & mytablex.Fields("nro") = buf And "" & mytablex.Fields("producto") = buf1 Then
   '------------------------------------------
   Data2.Recordset.AddNew
   Data2.Recordset.Fields("producto") = "" & producto
   Data2.Recordset.Fields("numero") = "" & numero
   Data2.Recordset.Fields("productor") = "" & mytablex.Fields("productoi")
   Data2.Recordset.Fields("descripcio") = "" & mytablex.Fields("descripcio")
   Data2.Recordset.Fields("unidad") = "" & mytablex.Fields("unidad")
   Data2.Recordset.Fields("factor") = Val("" & mytablex.Fields("factor"))
   Data2.Recordset.Fields("precio") = Val("" & mytablex.Fields("factor"))
   Data2.Recordset.Fields("cantidad") = Val("" & mytablex.Fields("cantidad")) * Val(cantidad)
   Data2.Recordset.Fields("precio") = Val("" & mytablex.Fields("precio"))
   Data2.Recordset.Fields("linea") = "" & xlinea
   If Len("" & xlinea) > 0 Then
   Data2.Recordset.Fields("T1") = Val("" & mytablex.Fields("t1")) * Val(xt1)
   Data2.Recordset.Fields("T2") = Val("" & mytablex.Fields("t2")) * Val(xt2)
   Data2.Recordset.Fields("T3") = Val("" & mytablex.Fields("t3")) * Val(xt3)
   Data2.Recordset.Fields("T4") = Val("" & mytablex.Fields("t4")) * Val(xt4)
   Data2.Recordset.Fields("T5") = Val("" & mytablex.Fields("t5")) * Val(xt5)
   Data2.Recordset.Fields("T6") = Val("" & mytablex.Fields("t6")) * Val(xt6)
   Data2.Recordset.Fields("T7") = Val("" & mytablex.Fields("t7")) * Val(xt7)
   Data2.Recordset.Fields("T8") = Val("" & mytablex.Fields("t8")) * Val(xt8)
   Data2.Recordset.Fields("T9") = Val("" & mytablex.Fields("t9")) * Val(xt9)
   Data2.Recordset.Fields("T10") = Val("" & mytablex.Fields("t10")) * Val(xt10)
   Data2.Recordset.Fields("T11") = Val("" & mytablex.Fields("t11")) * Val(xt11)
   Data2.Recordset.Fields("T12") = Val("" & mytablex.Fields("t12")) * Val(xt12)
   Data2.Recordset.Fields("T13") = Val("" & mytablex.Fields("t13")) * Val(xt13)
   Data2.Recordset.Fields("T14") = Val("" & mytablex.Fields("t14")) * Val(xt14)
   Data2.Recordset.Fields("T15") = Val("" & mytablex.Fields("t15")) * Val(xt15)
   Data2.Recordset.Fields("T16") = Val("" & mytablex.Fields("t16")) * Val(xt16)
   End If
   If Len("" & xlinea) > 0 Then
      sdx = Val("" & Data2.Recordset.Fields("t1")) + Val("" & Data2.Recordset.Fields("t2")) + Val("" & Data2.Recordset.Fields("t3")) + Val("" & Data2.Recordset.Fields("t4")) + Val("" & Data2.Recordset.Fields("t5")) + Val("" & Data2.Recordset.Fields("t6")) + Val("" & Data2.Recordset.Fields("t7")) + Val("" & Data2.Recordset.Fields("t8")) + Val("" & Data2.Recordset.Fields("t9")) + Val("" & Data2.Recordset.Fields("t10")) + Val("" & Data2.Recordset.Fields("t11")) + Val("" & Data2.Recordset.Fields("t12")) + Val("" & Data2.Recordset.Fields("t13")) + Val("" & Data2.Recordset.Fields("t14")) + Val("" & Data2.Recordset.Fields("t15")) + Val("" & Data2.Recordset.Fields("t16"))
      Data2.Recordset.Fields("cantidad") = sdx
   End If
   sdx = (Val("" & Data2.Recordset.Fields("cantidad")) + Val("" & Data2.Recordset.Fields("merma"))) * Val("" & Data2.Recordset.Fields("precio"))
   Data2.Recordset.Fields("total") = sdx
   Data2.Recordset.Update
   '------------------------------------------
   Else
   Exit Do
End If
mytablex.MoveNext
Loop
'------------------------------------- ------------
mytablex.Close
 
End Function
Sub borra_anterior()


mydbxglo.Execute "DELETE FROM rproducc where numero='" & numero & "' and producto='" & producto & "'"
 
Data2.refresh

End Sub
Sub sql_insumo()
Dim sdx As Double
   Data2.Connect = "foxpro 2.5;"
               Data2.DatabaseName = globaldir
               Data2.RecordSource = "Select * from rproducc where numero='" & numero & "' and producto='" & producto & "'"
               Data2.refresh
               
Data3.Connect = "foxpro 2.5;"
               Data3.DatabaseName = globaldir
               Data3.RecordSource = "Select * from partepro where numero='" & numero & "' and producto='" & producto & "'"
               Data3.refresh
               sumar_total
               sumar_total1
               sdx = (Val(total) + Val(total1))
ntotal = "" & sdx

End Sub
Sub ingreso_tallas(buf As String)
Dim found As Integer
linea = buf
found = busca_linea(buf)
If found = 0 Then Exit Sub
pone_tallas
Frame2.Visible = True
t1.SetFocus
End Sub
Function busca_linea(buf As String)

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("linea")
mytablex.Index = "linea"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   busca_linea = 1
   nlinea = "" & mytablex.Fields("descripcio")
   nt1 = "" & mytablex.Fields("t1")
   nt2 = "" & mytablex.Fields("t2")
   nt3 = "" & mytablex.Fields("t3")
   nt4 = "" & mytablex.Fields("t4")
   nt5 = "" & mytablex.Fields("t5")
   nt6 = "" & mytablex.Fields("t6")
   nt7 = "" & mytablex.Fields("t7")
   nt8 = "" & mytablex.Fields("t8")
   nt9 = "" & mytablex.Fields("t9")
   nt10 = "" & mytablex.Fields("t10")
   nt11 = "" & mytablex.Fields("t11")
   nt12 = "" & mytablex.Fields("t12")
   nt13 = "" & mytablex.Fields("t13")
   nt14 = "" & mytablex.Fields("t14")
   nt15 = "" & mytablex.Fields("t15")
   nt16 = "" & mytablex.Fields("t16")
End If
'------------------------------------- ------------
mytablex.Close
 
End Function
Sub pone_tallas()
If opcion1 = "1" Then
t1 = "" & Data2.Recordset.Fields("t1")
t2 = "" & Data2.Recordset.Fields("t2")
t3 = "" & Data2.Recordset.Fields("t3")
t4 = "" & Data2.Recordset.Fields("t4")
t5 = "" & Data2.Recordset.Fields("t5")
t6 = "" & Data2.Recordset.Fields("t6")
t7 = "" & Data2.Recordset.Fields("t7")
t8 = "" & Data2.Recordset.Fields("t8")
t9 = "" & Data2.Recordset.Fields("t9")
t10 = "" & Data2.Recordset.Fields("t10")
t11 = "" & Data2.Recordset.Fields("t11")
t12 = "" & Data2.Recordset.Fields("t12")
t13 = "" & Data2.Recordset.Fields("t13")
t14 = "" & Data2.Recordset.Fields("t14")
t15 = "" & Data2.Recordset.Fields("t15")
t16 = "" & Data2.Recordset.Fields("t16")
End If
If opcion1 = "2" Then
t1 = "" & Data2.Recordset.Fields("tm1")
t2 = "" & Data2.Recordset.Fields("tm2")
t3 = "" & Data2.Recordset.Fields("tm3")
t4 = "" & Data2.Recordset.Fields("tm4")
t5 = "" & Data2.Recordset.Fields("tm5")
t6 = "" & Data2.Recordset.Fields("tm6")
t7 = "" & Data2.Recordset.Fields("tm7")
t8 = "" & Data2.Recordset.Fields("tm8")
t9 = "" & Data2.Recordset.Fields("tm9")
t10 = "" & Data2.Recordset.Fields("tm10")
t11 = "" & Data2.Recordset.Fields("tm11")
t12 = "" & Data2.Recordset.Fields("tm12")
t13 = "" & Data2.Recordset.Fields("tm13")
t14 = "" & Data2.Recordset.Fields("tm14")
t15 = "" & Data2.Recordset.Fields("tm15")
t16 = "" & Data2.Recordset.Fields("tm16")
End If
End Sub
Sub suma_linea()
Dim sdx As Double
sdx = (Val("" & DBGrid2.columns(4)) + Val("" & DBGrid2.columns(5))) * Val("" & DBGrid2.columns(6))
DBGrid2.columns(7) = sdx
End Sub
Sub sumar_totalx()
Dim fila As Integer
Dim suma As Double
suma = 0
ir_inicio
For fila = 0 To Data2.Recordset.RecordCount - 1
DBGrid2.Row = fila    'El índice de la primera fila empieza en 0.
suma = suma + Val("" & DBGrid2.columns(7).Value)
Next
total = Format(suma, "0.00")
End Sub
Sub sumar_total1()
Dim fila As Integer
Dim suma As Double
suma = 0
ir_inicio1
Do
If Data3.Recordset.EOF Then Exit Do
suma = suma + Val("" & Data3.Recordset.Fields("total"))
Data3.Recordset.MoveNext
Loop
total1 = Format(suma, "0.00")

End Sub
Sub ir_inicio()
On Error GoTo cmd4_err
Data2.Recordset.MoveFirst
Exit Sub
cmd4_err:
Exit Sub
End Sub
Sub ir_inicio1()
On Error GoTo cmd42_err
Data3.Recordset.MoveFirst
Exit Sub
cmd42_err:
Exit Sub
End Sub




Private Sub operacion_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   KeyAscii = 0
   Exit Sub
End If
Horai.SetFocus
End Sub

Private Sub operacion_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   operario.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
consulta_operacion
End If

End Sub

Private Sub operario_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   KeyAscii = 0
   Exit Sub
End If
operacion.SetFocus
End Sub

Private Sub operario_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   fecha.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
consulta_operario
End If

End Sub

Private Sub pami7823_Click()
opcion2 = "1"
solhora = ""
nrohora = ""
fecha = ""
operario = ""
seccion = ""
operacion = ""
Horai = ""
horaf = ""
Frame3.Visible = True
fecha.SetFocus

End Sub
Sub consulta_operario()
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.AddItem "Codigo"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "0"
Command1_Click
End Sub
Sub consulta_operacion()
Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.AddItem "Operacion"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "1"
Command1_Click
End Sub
Function busca_operario(buf As String)

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("Vendedor")
mytablex.Index = "Codigo"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   busca_operario = 1
   seccion = "" & mytablex.Fields("seccion")
   solhora = "" & mytablex.Fields("jornal")
End If
mytablex.Close
 
End Function
Function busca_operacion(buf As String)

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("toperaco")
mytablex.Index = "toperaco"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   busca_operacion = 1
End If
mytablex.Close
 
End Function
Sub sumar_total()
Dim fila As Integer
Dim suma As Double
suma = 0
ir_inicio
Do
If Data2.Recordset.EOF Then Exit Do
suma = suma + Val("" & Data2.Recordset.Fields("total"))
Data2.Recordset.MoveNext
Loop
total = Format(suma, "0.00")
End Sub



