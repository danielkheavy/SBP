VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form texpadua 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Explorador de Documentos Importacion"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   14745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   14745
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Clave de Acceso"
      Height          =   5055
      Left            =   12000
      TabIndex        =   106
      Top             =   4800
      Visible         =   0   'False
      Width           =   5295
      Begin VB.TextBox clave 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   1200
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   109
         TabStop         =   0   'False
         Top             =   1800
         Width           =   3015
      End
      Begin VB.CommandButton Command4 
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
         Left            =   3480
         MaskColor       =   &H00FFFFFF&
         Picture         =   "TEXPADUA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   3240
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
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
         Left            =   3480
         MaskColor       =   &H00FFFFFF&
         Picture         =   "TEXPADUA.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   4080
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ingrese su clave para realizar esta Accion"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   600
         TabIndex        =   110
         Top             =   720
         Width           =   4215
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H0080FFFF&
      Caption         =   "Cambiar Documentos por Otro Documento"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5175
      Left            =   10680
      TabIndex        =   85
      Top             =   6120
      Visible         =   0   'False
      Width           =   10455
      Begin VB.TextBox decodigo 
         Height          =   375
         Left            =   2280
         MaxLength       =   11
         TabIndex        =   103
         Top             =   4080
         Width           =   1095
      End
      Begin VB.TextBox denombre 
         Height          =   375
         Left            =   2280
         MaxLength       =   60
         TabIndex        =   102
         Top             =   4440
         Width           =   5055
      End
      Begin VB.TextBox denumero 
         Height          =   375
         Left            =   2280
         MaxLength       =   11
         TabIndex        =   101
         Top             =   3720
         Width           =   2415
      End
      Begin VB.TextBox deserie 
         Height          =   375
         Left            =   2280
         MaxLength       =   3
         TabIndex        =   92
         Top             =   3360
         Width           =   1095
      End
      Begin VB.ComboBox detipo 
         BackColor       =   &H0000FFFF&
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   88
         Top             =   3000
         Width           =   2655
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Grabar"
         Height          =   495
         Left            =   8280
         TabIndex        =   87
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Salir"
         Height          =   495
         Left            =   8280
         TabIndex        =   86
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label29 
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
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
         Left            =   240
         TabIndex        =   105
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
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
         Left            =   240
         TabIndex        =   104
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Documento Fuente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   100
         Top             =   840
         Width           =   7215
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Documento Destino"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   99
         Top             =   2640
         Width           =   7215
      End
      Begin VB.Label Label35 
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero"
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
         Left            =   240
         TabIndex        =   98
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label sonumero 
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2280
         TabIndex        =   97
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label30 
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoDocumento"
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
         Left            =   240
         TabIndex        =   96
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label sotipo 
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2280
         TabIndex        =   95
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label28 
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie"
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
         Left            =   240
         TabIndex        =   94
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label soserie 
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2280
         TabIndex        =   93
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label33 
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Documento"
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
         Left            =   240
         TabIndex        =   91
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label Label32 
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie"
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
         Left            =   240
         TabIndex        =   90
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Label31 
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero"
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
         Left            =   240
         TabIndex        =   89
         Top             =   3720
         Width           =   2055
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Generar Traslado Automatico"
      ForeColor       =   &H00FF0000&
      Height          =   5175
      Left            =   13200
      TabIndex        =   53
      Top             =   3960
      Visible         =   0   'False
      Width           =   10455
      Begin VB.CommandButton Command7 
         Caption         =   "Salir"
         Height          =   495
         Left            =   8520
         TabIndex        =   71
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Grabar"
         Height          =   495
         Left            =   8520
         TabIndex        =   70
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox uvendedor 
         BackColor       =   &H0000FFFF&
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   2400
         Width           =   2655
      End
      Begin VB.ComboBox utipo 
         BackColor       =   &H0000FFFF&
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox ufecha 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   67
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox unumero 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Left            =   2280
         MaxLength       =   11
         TabIndex        =   64
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox userie 
         Height          =   375
         Left            =   2280
         MaxLength       =   3
         TabIndex        =   62
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox ubodegaf 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   59
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox ubodega 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   57
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox ulocal 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   55
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label unombre1 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3840
         TabIndex        =   73
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label unombre2 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3840
         TabIndex        =   72
         Top             =   960
         Width           =   3975
      End
      Begin VB.Label Label24 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha"
         Height          =   375
         Left            =   240
         TabIndex        =   66
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label23 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Responsable"
         Height          =   375
         Left            =   240
         TabIndex        =   65
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label22 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero"
         Height          =   375
         Left            =   240
         TabIndex        =   63
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label21 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie"
         Height          =   375
         Left            =   240
         TabIndex        =   61
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label20 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
         Height          =   375
         Left            =   240
         TabIndex        =   60
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label19 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Alm.Final"
         Height          =   375
         Left            =   240
         TabIndex        =   58
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label18 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Alm.Inicio"
         Height          =   375
         Left            =   240
         TabIndex        =   56
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LocaL"
         Height          =   375
         Left            =   240
         TabIndex        =   54
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8280
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFF00&
      Caption         =   "Visualizar Detalle"
      Height          =   8775
      Left            =   10080
      TabIndex        =   43
      Top             =   5040
      Visible         =   0   'False
      Width           =   14415
      Begin VB.CommandButton Command8 
         Caption         =   "Validar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   77
         Top             =   7200
         Width           =   3135
      End
      Begin VB.CheckBox yausado 
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   76
         Top             =   7320
         Width           =   255
      End
      Begin MSDBGrid.DBGrid DBGrid4 
         Bindings        =   "TEXPADUA.frx":0F5C
         Height          =   2175
         Left            =   120
         OleObjectBlob   =   "TEXPADUA.frx":0F70
         TabIndex        =   45
         Top             =   5640
         Width           =   14055
      End
      Begin MSDBGrid.DBGrid DBGrid3 
         Bindings        =   "TEXPADUA.frx":1FE3
         Height          =   5295
         Left            =   120
         OleObjectBlob   =   "TEXPADUA.frx":1FF7
         TabIndex        =   44
         Top             =   240
         Width           =   14055
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
      Height          =   7935
      Left            =   8040
      TabIndex        =   21
      Top             =   6480
      Visible         =   0   'False
      Width           =   14655
      Begin VB.TextBox buffer 
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
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   23
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
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "TEXPADUA.frx":79FE
         Height          =   6735
         Left            =   120
         OleObjectBlob   =   "TEXPADUA.frx":7A12
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   840
         Width           =   13575
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Consultas-Condiciones"
      Height          =   5175
      Left            =   2160
      TabIndex        =   15
      Top             =   2520
      Visible         =   0   'False
      Width           =   7695
      Begin VB.ComboBox moneda 
         Height          =   360
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox nombre 
         Height          =   375
         Left            =   2400
         MaxLength       =   11
         TabIndex        =   32
         Text            =   "*"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox numero 
         Height          =   375
         Left            =   2400
         MaxLength       =   11
         TabIndex        =   31
         Text            =   "*"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox serie 
         Height          =   375
         Left            =   2400
         MaxLength       =   11
         TabIndex        =   29
         Text            =   "*"
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
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
         Picture         =   "TEXPADUA.frx":83DD
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1440
         UseMaskColor    =   -1  'True
         Width           =   1695
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
         Height          =   735
         Left            =   5520
         MaskColor       =   &H00FFFFFF&
         Picture         =   "TEXPADUA.frx":8B8B
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.ComboBox estado 
         Height          =   360
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox codigo 
         Height          =   375
         Left            =   2400
         MaxLength       =   11
         TabIndex        =   16
         Text            =   "*"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Moneda"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   35
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estado"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   1440
         Width           =   2175
      End
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "TEXPADUA.frx":9339
      Height          =   6735
      Left            =   0
      OleObjectBlob   =   "TEXPADUA.frx":934D
      TabIndex        =   0
      Top             =   1320
      Width           =   14655
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFF00&
      Height          =   1275
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   14685
      TabIndex        =   9
      Top             =   0
      Width           =   14745
      Begin VB.TextBox dua 
         Height          =   375
         Left            =   11760
         MaxLength       =   11
         TabIndex        =   113
         Text            =   "*"
         Top             =   480
         Width           =   1695
      End
      Begin VB.ComboBox agencia 
         Height          =   315
         Left            =   11760
         Style           =   2  'Dropdown List
         TabIndex        =   111
         Top             =   120
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "No mostrar Tipo 5"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   960
         Width           =   2175
      End
      Begin VB.ComboBox cajero 
         Height          =   315
         Left            =   9720
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox vendedor 
         Height          =   315
         Left            =   7200
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   840
         Width           =   1695
      End
      Begin VB.ComboBox tipo 
         Height          =   315
         Left            =   9720
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox caja 
         Height          =   315
         Left            =   9720
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   120
         Width           =   1215
      End
      Begin VB.ComboBox bodega 
         Height          =   315
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Consul&Ta"
         Height          =   615
         Left            =   13680
         MaskColor       =   &H00FFFFFF&
         Picture         =   "TEXPADUA.frx":DF18
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox fechaf 
         Height          =   375
         Left            =   7200
         MaxLength       =   10
         TabIndex        =   41
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox fechai 
         Height          =   375
         Left            =   7200
         MaxLength       =   10
         TabIndex        =   39
         Top             =   120
         Width           =   1695
      End
      Begin VB.ComboBox local1 
         Height          =   315
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdSort 
         BackColor       =   &H00C0C0C0&
         Height          =   855
         Left            =   2160
         MaskColor       =   &H00E0E0E0&
         Picture         =   "TEXPADUA.frx":E6C6
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Consulta"
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
         Height          =   855
         Left            =   720
         MaskColor       =   &H00E0E0E0&
         Picture         =   "TEXPADUA.frx":F8D8
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Height          =   855
         Left            =   2880
         MaskColor       =   &H00E0E0E0&
         Picture         =   "TEXPADUA.frx":10AEA
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Height          =   855
         Left            =   1440
         MaskColor       =   &H00E0E0E0&
         Picture         =   "TEXPADUA.frx":11CFC
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Height          =   855
         Left            =   0
         Picture         =   "TEXPADUA.frx":12F0E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label38 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nro.Dua"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   11040
         TabIndex        =   114
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label34 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Agencia"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   11040
         TabIndex        =   112
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cajero"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   9000
         TabIndex        =   81
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Comprador"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6240
         TabIndex        =   79
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoDoc"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   9000
         TabIndex        =   52
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Caja"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   9000
         TabIndex        =   50
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Almacen"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3720
         TabIndex        =   48
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Final"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6240
         TabIndex        =   42
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6240
         TabIndex        =   40
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3720
         TabIndex        =   38
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7080
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
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6360
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label zooma 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   13800
      TabIndex        =   83
      Top             =   8280
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label YacaRGA 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   13800
      TabIndex        =   78
      Top             =   8040
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label nbodega1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   75
      Top             =   8400
      Width           =   3735
   End
   Begin VB.Label nbodega 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   74
      Top             =   8040
      Width           =   3735
   End
   Begin VB.Label tipoclie 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   360
      TabIndex        =   36
      Top             =   6960
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "US$."
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6480
      TabIndex        =   26
      Top             =   8400
      Width           =   2175
   End
   Begin VB.Label subtotald 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   8640
      TabIndex        =   8
      Top             =   8400
      Width           =   1695
   End
   Begin VB.Label subtotals 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   8640
      TabIndex        =   7
      Top             =   8040
      Width           =   1695
   End
   Begin VB.Label impuestod 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   10320
      TabIndex        =   6
      Top             =   8400
      Width           =   1695
   End
   Begin VB.Label impuestos 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   10320
      TabIndex        =   5
      Top             =   8040
      Width           =   1695
   End
   Begin VB.Label acu 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   13320
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label totald 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   12000
      TabIndex        =   3
      Top             =   8400
      Width           =   1695
   End
   Begin VB.Label totals 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   12000
      TabIndex        =   2
      Top             =   8040
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S/."
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   8040
      Width           =   2175
   End
   Begin VB.Menu djku232 
      Caption         =   "&Nuevo"
   End
   Begin VB.Menu agt62323 
      Caption         =   "&Borrar"
   End
   Begin VB.Menu modi343 
      Caption         =   "&Desmarca"
   End
   Begin VB.Menu mio8923 
      Caption         =   "&Modifica"
   End
   Begin VB.Menu anulier 
      Caption         =   "&Anular"
   End
   Begin VB.Menu dkiw232 
      Caption         =   "&Imprimir"
      Begin VB.Menu dkifor 
         Caption         =   "&1.FormatoDefinido"
      End
      Begin VB.Menu dkiewre 
         Caption         =   "&2.Reporteador"
      End
      Begin VB.Menu dl89er 
         Caption         =   "&3.Excell Impresion Total"
      End
      Begin VB.Menu dki889343 
         Caption         =   "&4.Excell Impresion solo seleccionado"
      End
      Begin VB.Menu impso02 
         Caption         =   "&5.Excell Impresion solo Documentos"
      End
   End
   Begin VB.Menu mit56232 
      Caption         =   "&Zoom"
   End
   Begin VB.Menu fdl89234 
      Caption         =   "&Validar"
   End
   Begin VB.Menu djbu232 
      Caption         =   "&Consulta"
   End
   Begin VB.Menu Hyu6723ge 
      Caption         =   "&Generar"
      Begin VB.Menu dj7823233 
         Caption         =   "&1.Generar Traslado Automatico"
      End
      Begin VB.Menu campo92 
         Caption         =   "&2.Cambio de Documento"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu xdki82 
         Caption         =   "&3.Documento "
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu ldo33 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "texpadua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub agt62323_Click()

Dim buf1 As String
On Error GoTo cmd6_err
If Frame4.Visible = True Then Exit Sub

If Frame2.Visible = True Then Exit Sub
If Frame3.Visible = True Then Exit Sub

If "" & Data2.Recordset.Fields("estado") <> "0" Then
   MsgBox "Para Borrar el documento debe estar en estado=0", 48, "Aviso"
   Exit Sub
End If
If MsgBox("Desea Borrar Documento", 1, "Aviso") <> 1 Then Exit Sub

buf1 = " and acu='" & "" & Data2.Recordset.Fields("acu") & "'"
mydbxglo.Execute "DELETE FROM  " & dgusuariog & "   where  local='" & "" & Data2.Recordset.Fields("local") & "' and tipo='" & "" & Data2.Recordset.Fields("tipo") & "' and serie='" & "" & Data2.Recordset.Fields("serie") & "' and  numero='" & "" & Data2.Recordset.Fields("numero") & "'" & buf1
mydbxglo.Execute "delete from  fpagov  where  local='" & "" & Data2.Recordset.Fields("local") & "' and tipo='" & "" & Data2.Recordset.Fields("tipo") & "' and serie='" & "" & Data2.Recordset.Fields("serie") & "' and  numero='" & "" & Data2.Recordset.Fields("numero") & "'" & buf1
Data2.Recordset.Delete
Data2.Refresh
MsgBox "Ok,Documento Borrado", 24, "Aviso"
Exit Sub
cmd6_err:
Exit Sub

End Sub

Private Sub anulier_Click()

Dim buf1 As String
Dim buf As String
Dim Msg As String
On Error GoTo cmd8_err
If Frame4.Visible = True Then Exit Sub

If Frame2.Visible = True Then Exit Sub
If Frame3.Visible = True Then Exit Sub
Msg = "Ojo.. Esta opcion de anular permite poner el documento en modo de anulacion ,luego de realizacion" + Chr$(10) + Chr$(13)
Msg = Msg + "No puede ya reversar.... " + Chr$(10) + Chr$(13)
If MsgBox(Msg, 1, "Aviso") <> 1 Then Exit Sub


If "" & Data2.Recordset.Fields("estado") = "2" Then
   MsgBox "Para anular el documento debe estar en estado=0 or estado=1", 48, "Aviso"
   Exit Sub
End If
If MsgBox("Desea Anular Documento,Quedara inmodificable ", 1, "Aviso") <> 1 Then Exit Sub
buf = "1"
If "" & Data2.Recordset.Fields("estado") = "1" Then
   buf = "0"
End If

Data2.Recordset.Edit
Data2.Recordset.Fields("estado") = buf
Data2.Recordset.Update
buf1 = " and acu='" & "" & Data2.Recordset.Fields("acu") & "'"
mydbxglo.Execute "update  " & dgusuariog & " set estado='" & buf & "'  where  local='" & "" & Data2.Recordset.Fields("local") & "' and tipo='" & "" & Data2.Recordset.Fields("tipo") & "' and serie='" & "" & Data2.Recordset.Fields("serie") & "' and  numero='" & "" & Data2.Recordset.Fields("numero") & "'" & buf1
mydbxglo.Execute "update  fpagov  set estado='" & buf & "'  where  local='" & "" & Data2.Recordset.Fields("local") & "' and tipo='" & "" & Data2.Recordset.Fields("tipo") & "' and serie='" & "" & Data2.Recordset.Fields("serie") & "' and  numero='" & "" & Data2.Recordset.Fields("numero") & "'" & buf1
MsgBox "Ok,Documento Anulado", 24, "Aviso"
sql_cabeza
Exit Sub
cmd8_err:
Exit Sub


End Sub

Private Sub buffer_DblClick()
Command1_Click
End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   ldo33_Click
   Exit Sub
End If
Command1_Click

End Sub

Private Sub cmdCancelar_Click()
Frame4.Visible = False
End Sub

Private Sub campo92_Click()
Dim mytablex As Table
On Error GoTo cmd7632_err
Frame6.Caption = "CambiaDatos"
sotipo = "" & Data2.Recordset.Fields("tipo")
soserie = "" & Data2.Recordset.Fields("serie")
sonumero = "" & Data2.Recordset.Fields("numero")
detipo.Clear
Set mytablex = mydbxglo.OpenTable("tipo")
Do
If mytablex.EOF Then Exit Do
If "" & mytablex.Fields("tipodoc") = "A" Or "" & mytablex.Fields("tipodoc") = "B" Or "" & mytablex.Fields("tipodoc") = "C" Or "" & mytablex.Fields("tipodoc") = "D" Or "" & mytablex.Fields("tipodoc") = "G" Or "" & mytablex.Fields("tipodoc") = "H" Or "" & mytablex.Fields("tipodoc") = "I" Or "" & mytablex.Fields("tipodoc") = "E" Or "" & mytablex.Fields("tipodoc") = "F" Or "" & mytablex.Fields("tipodoc") = "T" Then
   If "" & mytablex.Fields("tipodoc") <> acu Then
      detipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")
   End If
End If
mytablex.MoveNext
Loop
detipo.ListIndex = 0
mytablex.Close
Frame6.Visible = True
deserie = ""
denumero = ""
decodigo = ""
denumero = ""
detipo.SetFocus
Exit Sub
cmd7632_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub
End Sub

Private Sub clave_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub cmdAddEntry_Click()
djku232_Click
End Sub

Private Sub cmdExit_Click()
ldo33_Click
End Sub

Private Sub cmdGrabar_Click()
sql_cabeza
Frame3.Visible = False
End Sub

Private Sub cmdHelp_Click()

End Sub

Private Sub cmdPrint_Click()
dkifor_Click
End Sub

Private Sub cmdSave_Click()

End Sub

Private Sub cmdSort_Click()
djbu232_Click
End Sub

Private Sub Codigo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
consulta_codigo
End If

End Sub

Private Sub Command1_Click()
Dim buf As String
If opcion1 = "1" Then
If Len(buffer) = 0 Then
buf = "select Nombre,Codigo from clientes "
Else
buf = "select Nombre,Codigo from clientes where " & Combo1 & " like '" & buffer & "%'"
End If
End If
If opcion1 = "2" Then
   buf = "select Producto,Descripcio,Unidad as Und,Factor as Fac,Precio,Cantidad as Cant,Total,Local,Deslipo as Dscto from  " & dgusuariog & " where local='" & "" & Data2.Recordset.Fields("local") & "' and serie='" & "" & Data2.Recordset.Fields("serie") & "' and numero='" & "" & Data2.Recordset.Fields("numero") & "'"
End If

               Data1.Connect = "foxpro 2.5;"
               Data1.DatabaseName = globaldir
               Data1.RecordSource = buf
               Data1.Refresh
               If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then
                  Data1.Recordset.Close
                  buffer.SetFocus
                  Exit Sub
               End If
               If opcion1 = "1" Then
               dbGrid1.columns(0).Width = 4000
               dbGrid1.columns(1).Width = 2000
               End If
               If opcion1 = "2" Then
               dbGrid1.columns(0).Width = 1500
               dbGrid1.columns(1).Width = 5000
               dbGrid1.columns(2).Width = 900
               dbGrid1.columns(3).Width = 900
               dbGrid1.columns(4).Width = 900
               dbGrid1.columns(5).Width = 900
               dbGrid1.columns(6).Width = 1500
               dbGrid1.columns(7).Width = 900
               dbGrid1.columns(8).Width = 700
               End If
               dbGrid1.SetFocus

End Sub

Private Sub Command10_Click()
Dim mytablex As Table
Dim tmtipo As String
Dim tmserie As String
Dim tmnumero As String

If Len(deserie) = 0 Then
   deserie.SetFocus
   Exit Sub
End If
If Len(denumero) = 0 Then
   denumero.SetFocus
   Exit Sub
End If
tmtipo = "" & Data2.Recordset.Fields("tipo")
tmserie = "" & Data2.Recordset.Fields("serie")
tmnumero = "" & Data2.Recordset.Fields("numero")

Set mytablex = mydbxglo.OpenTable(cgusuario)
mytablex.Index = "tfacimp"
mytablex.Seek "=", Data2.Recordset.Fields("local"), extra_loquesea(detipo), deserie, denumero
If Not mytablex.NoMatch Then
   MsgBox "Ya existe el numero ", 48, "Aviso"
   mytablex.Close
   denumero.SetFocus
   Exit Sub
End If
mytablex.Close
mydbxglo.Execute "update  " & cgusuario & " set tipo='" & extra_loquesea(detipo) & "',serie='" & deserie & "',numero='" & denumero & "',codigo='" & decodigo & "',nombre='" & denombre & "'   where local='" & "" & Data2.Recordset.Fields("local") & "' and  tipo='" & "" & Data2.Recordset.Fields("tipo") & "' and serie='" & "" & Data2.Recordset.Fields("serie") & "' and  numero='" & "" & Data2.Recordset.Fields("numero") & "'"
mydbxglo.Execute "update  " & dgusuariog & " set tipo='" & extra_loquesea(detipo) & "',serie='" & deserie & "',numero='" & denumero & "',codigo='" & decodigo & "'  where local='" & "" & Data2.Recordset.Fields("local") & "' and  tipo='" & "" & tmtipo & "' and serie='" & "" & tmserie & "' and  numero='" & "" & tmnumero & "'"
mydbxglo.Execute "update  fpagov set tipo='" & extra_loquesea(detipo) & "',serie='" & deserie & "',numero='" & denumero & "',codigo='" & decodigo & "'  where local='" & "" & Data2.Recordset.Fields("local") & "' and  tipo='" & "" & tmtipo & "' and serie='" & "" & tmserie & "' and  numero='" & "" & tmnumero & "'"
MsgBox "Proceso Realizado ", 48, "Aviso"
ldo33_Click
sql_cabeza




End Sub

Private Sub Command2_Click()
ldo33_Click
End Sub

Private Sub Command3_Click()
ldo33_Click
End Sub

Private Sub Command4_Click()
On Error GoTo cmd7_err
Dim found As Integer
Dim buf As String
If Len(clave) = 0 Then
   clave.SetFocus
   Exit Sub
End If
found = valida_clave("" & clave)
If found = 0 Then
   MsgBox "Clave no valida para realizar este proceso ", 48, "Aviso"
   clave = ""
   clave.SetFocus
   Exit Sub
End If
If Frame2.Caption = "DESMARCA" Then
   If MsgBox("Desea Desmarca el Documento", 1, "Aviso") <> 1 Then Exit Sub
   If "" & Data2.Recordset.Fields("acu") = "A" Or "" & Data2.Recordset.Fields("acu") = "B" Or "" & Data2.Recordset.Fields("acu") = "C" Or "" & Data2.Recordset.Fields("acu") = "D" Or "" & Data2.Recordset.Fields("acu") = "G" Then  'ventas
      buf = "cuentacd"
      found = verificar_recibo(buf)
      If found = 1 Then
         MsgBox "Ya existe recibo ", 48, "Aviso"
         Exit Sub
      End If
   End If
   If "" & Data2.Recordset.Fields("acu") = "J" Or "" & Data2.Recordset.Fields("acu") = "K" Or "" & Data2.Recordset.Fields("acu") = "L" Or "" & Data2.Recordset.Fields("acu") = "M" Or "" & Data2.Recordset.Fields("acu") = "P" Then  'ventas
      buf = "cuentaPd"
      found = verificar_recibo(buf)
      If found = 1 Then
         MsgBox "Ya existe recibo ", 48, "Aviso"
         Exit Sub
      End If
   End If
   desmarca_documento
End If
Frame2.Visible = False
Exit Sub
cmd7_err:
MsgBox "Seleccione un dato " + error$, 48, "Aviso"
Frame2.Visible = False
Exit Sub
End Sub

Private Sub Command5_Click()
sql_cabeza
End Sub

Private Sub Command6_Click()
Dim found As Integer
If Len(userie) = 0 Then
   userie.SetFocus
   Exit Sub
End If
If Len(unumero) = 0 Then
   unumero.SetFocus
   Exit Sub
End If
If Len(ufecha) <> 10 Then
   ufecha = ""
   Exit Sub
End If
If Not IsDate(ufecha) Then
   ufecha = ""
   ufecha.SetFocus
   Exit Sub
End If
If utipo = "%" Then
   utipo.SetFocus
   Exit Sub
End If
If uvendedor = "%" Then
   uvendedor.SetFocus
   Exit Sub
End If
If valida_existe_despacho() = 0 Then
   MsgBox "No existe cantidad a Despachar ", 48, "Aviso"
   Exit Sub
End If
found = genera_traslado_automata()
If found = 1 Then
   found = graba_existe_despacho()
   MsgBox "Documento Generado ", 48, "Aviso"
   Frame5.Visible = False
   Exit Sub
End If


End Sub

Private Sub Command7_Click()
ldo33_Click
End Sub

Private Sub Command8_Click()
On Error GoTo cmd32412_err
'flag_clave1 = 0
'tconcla.X = "C"
'tconcla.Show 1
'If flag_clave1 <> 1 Then  'si es descongela
'   Exit Sub
'End If
If MsgBox("Esta seguro", 1, "Aviso") <> 1 Then Exit Sub

If yausado.Value = 1 Then
   Data2.Recordset.Edit
   Data2.Recordset.Fields("nombre") = Mid$("" & Data2.Recordset.Fields("nombre") + "-" + gusuario, 1, 60)
   Data2.Recordset.Fields("yausado") = "0"
   Data2.Recordset.Update
   yausado.Value = 0
   Exit Sub
End If
If yausado.Value = 0 Then
   Data2.Recordset.Edit
   Data2.Recordset.Fields("nombre") = Mid$("" & Data2.Recordset.Fields("nombre") + "-" + gusuario, 1, 60)
   Data2.Recordset.Fields("yausado") = "1"
   Data2.Recordset.Update
   yausado.Value = 1
   Exit Sub
End If
Exit Sub
cmd32412_err:
MsgBox "Aviso en Yausado " + error$, 48, "Aviso"
Exit Sub

End Sub

Private Sub Command9_Click()
ldo33_Click
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   buffer.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
   If opcion1 = "1" Then
   codigo = dbGrid1.columns(1)
   Frame1.Visible = False
   codigo.SetFocus
   End If
End If

End Sub

Private Sub DBGrid2_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If ColIndex <> 13 Then
   Cancel = True
   Exit Sub
End If
Select Case ColIndex
       Case 13
            If Len("" & DBGrid2.columns(2)) = 0 Then
               Cancel = True
               Exit Sub
            End If
End Select

End Sub

Private Sub DBGrid2_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim found As Integer
Select Case ColIndex
Case 13
     If Len(DBGrid2.columns(2)) = 0 Then
        Cancel = True
        Exit Sub
     End If
     If Len("" & DBGrid2.columns(13)) = 0 Then
        DBGrid2.columns(13) = Data2.Recordset.Fields("fecha")
        Exit Sub
     End If
     found = valida_fecha("" & DBGrid2.columns(13))
     If found = 0 Then
        Cancel = True
        Exit Sub
     End If
End Select
End Sub

Private Sub DBGrid2_DblClick()
NBODEGA = unombre_almacen(DBGrid2.columns(2))
nbodega1 = unombre_almacen(DBGrid2.columns(3))
End Sub

Private Sub DBGrid2_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = &H71 Then  'f2
If acu = "P" Or acu = "Z" Or acu = "Q" Then
   cambia_estado
End If
End If
If KeyCode = &H70 Then  'consultando
consulta_detalle
End If


End Sub

Private Sub dbgrid3_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If ColIndex <> 5 Then
   Cancel = True
   Exit Sub
End If
Select Case ColIndex
       Case 5
            If Len("" & dbgrid3.columns(0)) = 0 Then
               Cancel = True
               Exit Sub
            End If
           
End Select

End Sub

Private Sub DBGrid3_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Select Case ColIndex
       Case 5
            If Not IsNumeric("" & dbgrid3.columns(5)) Then
               Cancel = True
               Exit Sub
            End If
            If Val("" & dbgrid3.columns(3)) < (Val("" & dbgrid3.columns(4)) + Val("" & dbgrid3.columns(5))) Then
               MsgBox "Despacho excedido ", 48, "Aviso"
               Cancel = True
               Exit Sub
            End If
            
            '---------- validamos a donde va
            'valida_ingresado
End Select

End Sub

Private Sub dj7823233_Click()
If acu <> "Q" Then Exit Sub
If Frame4.Visible = True Then
   Frame4.Visible = False
End If
cerrar_data4
ir_menu_traslado

End Sub
Sub cerrar_data4()
On Error GoTo cmd45_err
Data4.Recordset.Close
Exit Sub
cmd45_err:
Exit Sub
End Sub

Private Sub djbu232_Click()
If Frame4.Visible = True Then Exit Sub

If Frame2.Visible = True Then Exit Sub
If Frame3.Visible = True Then Exit Sub
Frame3.Visible = True

End Sub

Private Sub djku232_Click()
Dim found As Integer
On Error GoTo cmd28_err
If Frame4.Visible = True Then Exit Sub

If Frame2.Visible = True Then Exit Sub
If Frame3.Visible = True Then Exit Sub
If acu = "Z" Then
   tfacimp.local1 = "01"
   tfacimp.codigo = "01"
   tfacimp.Label2.Caption = "Cod.Int"
   'tfacimp.Label14.Visible = True
   tfacimp.Label38.Visible = True
   'tfacimp.localf.Visible = True
   
End If
If acu = "V" Then
found = copiar_temporal()
If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   Exit Sub
End If
tfacimp.Label2 = "CodClie"
tfacimp.Caption = "Facturacion x Ventas"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
'tfacimp.caja = "00"
tfacimp.bandera = "Nuevo"
tfacimp.acu = "V"
tfacimp.tipoclie = tipoclie

'tfacimp.local1=local
tfacimp.Show 1
End If
If acu = "H" Then
found = copiar_temporal()
If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   Exit Sub
End If
tfacimp.Label2 = "CodClie"
tfacimp.Caption = "Cotizaciones Ventas"
cgusuario = "CCOTIZAV"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "dcotizav"
tfacimp.bandera = "Nuevo"
tfacimp.acu = "H"
tfacimp.tipoclie = tipoclie
tfacimp.Show 1
End If
If acu = "I" Then
found = copiar_temporal()
If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   Exit Sub
End If
tfacimp.Label2 = "CodClie"
tfacimp.Caption = "Cotizaciones Ventas"
cgusuario = "Cpedidov"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "dpedidov"
tfacimp.bandera = "Nuevo"
tfacimp.acu = "I"
tfacimp.tipoclie = tipoclie
tfacimp.Show 1

End If


If acu = "T" Then
found = copiar_temporal()
If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   Exit Sub
End If
tfacimp.Label2 = "CodClie"
tfacimp.Caption = "Guia Salida"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfacimp.bandera = "Nuevo"
tfacimp.acu = "T"
tfacimp.tipoclie = tipoclie
tfacimp.Show 1
End If
If acu = "E" Then
found = copiar_temporal()
If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   Exit Sub
End If
tfacimp.Label2 = "CodClie"
tfacimp.Caption = "Nota Credito Ventas"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfacimp.bandera = "Nuevo"
tfacimp.acu = "E"
tfacimp.tipoclie = tipoclie
tfacimp.Show 1
End If
If acu = "F" Then
found = copiar_temporal()
If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   Exit Sub
End If
tfacimp.Label2 = "CodClie"
tfacimp.Caption = "Nota debito Ventas"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
tfacimp.bandera = "Nuevo"
dgusuariog = "DETALLE"
tfacimp.acu = "F"
tfacimp.tipoclie = tipoclie
tfacimp.Show 1
End If
If acu = "R" Then
found = copiar_temporal()
If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   
   Exit Sub
End If
tfacimp.Label2 = "CodProv"
tfacimp.Caption = "Orden de Compra"
cgusuario = "CORDENC"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DORDENC"
tfacimp.acu = "R"
tfacimp.bandera = "Nuevo"
tfacimp.tipoclie = tipoclie
tfacimp.Show 1
End If
If acu = "S" Then
found = copiar_temporal()
If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   
   Exit Sub
End If
tfacimp.Label2 = "CodProv"
tfacimp.Caption = "Guia Remision Entrada"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfacimp.tipoclie = tipoclie
tfacimp.acu = "S"
tfacimp.bandera = "Nuevo"
tfacimp.Show 1
End If
If acu = "C" Then
found = copiar_temporal()
If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   
   Exit Sub
End If
tfacimp.Label2 = "CodProv"
tfacimp.Caption = "Factura de Importaciones"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfacimp.acu = "C"
tfacimp.bandera = "Nuevo"
tfacimp.tipoclie = tipoclie
tfacimp.Show 1
End If

If acu = "N" Then
found = copiar_temporal()
If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   Exit Sub
End If
tfacimp.Label2 = "CodProv"
tfacimp.Caption = "Nota Credito Compra"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfacimp.acu = "N"
tfacimp.bandera = "Nuevo"
tfacimp.tipoclie = tipoclie
tfacimp.Show 1
End If

If acu = "O" Then
found = copiar_temporal()
If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   Exit Sub
End If
tfacimp.Label2 = "CodProv"
tfacimp.Caption = "Nota debito de Compra"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfacimp.acu = "O"
tfacimp.bandera = "Nuevo"
tfacimp.tipoclie = tipoclie
tfacimp.Show 1
End If
If acu = "Q" Then
found = copiar_temporal()
If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   Exit Sub
End If
tfacimp.Label2 = "CodProv"
tfacimp.Caption = "Pedido Almacen"
cgusuario = "CREQUISA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DREQUISA"
tfacimp.acu = "Q"
tfacimp.bandera = "Nuevo"
tfacimp.tipoclie = tipoclie
tfacimp.Show 1

End If
If acu = "Z" Then
found = copiar_temporal()
If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   End
   Exit Sub
End If
'tfacimp.Label2 = "Cod.Inicio"
tfacimp.Caption = "Traslado entre almacen de un mismo establecimiento"
cgusuario = "CTRASLAD"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DTRASLAD"
tfacimp.acu = "Z"
tfacimp.bandera = "Nuevo"
tfacimp.tipoclie = tipoclie
tfacimp.Show 1
End If
sql_cabeza
Exit Sub
cmd28_err:
Exit Sub
End Sub

Private Sub dki889343_Click()
If Frame1.Visible = True Then Exit Sub
If Frame2.Visible = True Then Exit Sub
If Frame3.Visible = True Then Exit Sub
If MsgBox("Desea Exportar Excell", 1, "Aviso") <> 1 Then Exit Sub
conteo_excell_uno
End Sub

Private Sub dkiewre_Click()
If Frame4.Visible = True Then Exit Sub

If Frame2.Visible = True Then Exit Sub
If Frame3.Visible = True Then Exit Sub

reporgen.NAMETABLA = cgusuario
reporgen.Show 1
End Sub

Private Sub dkifor_Click()
If Frame4.Visible = True Then Exit Sub

If Frame2.Visible = True Then Exit Sub
If Frame3.Visible = True Then Exit Sub
proceso_impresion1
End Sub

Private Sub dl89er_Click()
If Frame1.Visible = True Then Exit Sub
If Frame2.Visible = True Then Exit Sub
If Frame3.Visible = True Then Exit Sub
If MsgBox("Desea Exportar Excell", 1, "Aviso") <> 1 Then Exit Sub
menu_excell
End Sub

Private Sub fdl89234_Click()
Dim buf As String
On Error GoTo cmd45112_err
buf = Data2.Recordset.Fields("local")
If "" & Data2.Recordset.Fields("estado") <> "2" Then
   MsgBox "Para este fin el estado debe estar en 2", 48, "Aviso"
   Exit Sub
End If
Select Case acu
       Case "Z", "S", "T"
       Case Else: Exit Sub
End Select
If Frame1.Visible = True Then Exit Sub
If Frame2.Visible = True Then Exit Sub
If Frame3.Visible = True Then Exit Sub
Frame4.Visible = True
sql_detalles
dbgrid3.SetFocus
Exit Sub
cmd45112_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub
End Sub

Private Sub Form_Activate()
Check1.Visible = False
'MsgBox acu
Select Case acu
       Case "V"
       Check1.Visible = True
End Select

If YacaRGA = "" Then
If acu = "Q" Then
   dbgrid3.AllowUpdate = True
   Hyu6723ge.Visible = True
   Hyu6723ge.Enabled = True
End If
carga_iniciales
'cmdGrabar_Click
YacaRGA = "S"
End If
'If YacaRGA <> "P" Then
If zooma = "Zomm" Then
   Frame3.Visible = False
   zooma = ""
   Exit Sub
End If
   zooma = ""
   cmdGrabar_Click
'   YacaRGA = "S"
'End If
'sql_cabeza
'color_cambio

End Sub
Sub color_cambio()
End Sub

Private Sub Form_Load()
moneda.Clear
moneda.AddItem "%"
moneda.AddItem "S"
moneda.AddItem "D"
moneda.ListIndex = 0
estado.Clear
estado.AddItem "%"
estado.AddItem "2"
estado.AddItem "1"
estado.AddItem "0"
estado.ListIndex = 0
End Sub
Sub carga_iniciales()
Dim mytablex As Table
cajero.Clear
cajero.AddItem "%"
vendedor.Clear
vendedor.AddItem "%"
caja.Clear
caja.AddItem "%"
tipo.Clear
tipo.AddItem "%"
bodega.Clear
bodega.AddItem "%"
local1.Clear
local1.AddItem "%"


Set mytablex = mydbxglo.OpenTable("vendedor")
Do
If mytablex.EOF Then Exit Do

   vendedor.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
   cajero.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")

mytablex.MoveNext
Loop
vendedor.ListIndex = 0
cajero.ListIndex = 0
mytablex.Close


Set mytablex = mydbxglo.OpenTable("tipo")
Do
If mytablex.EOF Then Exit Do
If "" & mytablex.Fields("grupo") = acu Then
   tipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")
End If
mytablex.MoveNext
Loop
tipo.ListIndex = 0
mytablex.Close
Set mytablex = mydbxglo.OpenTable("bodega")
Do
If mytablex.EOF Then Exit Do
bodega.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")

mytablex.MoveNext
Loop
bodega.ListIndex = 0
mytablex.Close
 
Set mytablex = mydbxglo.OpenTable("tlocal")
Do
If mytablex.EOF Then Exit Do
local1.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
mytablex.MoveNext
Loop
local1.ListIndex = 0
If local1.ListCount = 2 Then
local1.ListIndex = 1
End If

mytablex.Close

agencia.Clear
agencia.AddItem "%"
Set mytablex = mydbxglo.OpenTable("aduana")
Do
If mytablex.EOF Then Exit Do
agencia.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
mytablex.MoveNext
Loop
agencia.ListIndex = 0
mytablex.Close


Set mytablex = mydbxglo.OpenTable("parameca")
Do
If mytablex.EOF Then Exit Do
If "" & mytablex.Fields("terminal") = "C" Then
caja.AddItem "" & mytablex.Fields("caja") & "|" & mytablex.Fields("descripcio")
End If
mytablex.MoveNext
Loop
caja.ListIndex = 0
mytablex.Close


End Sub

Private Sub impso02_Click()
menu_excell1
End Sub

Private Sub ldo33_Click()
If Frame5.Visible = True Then
   Frame5.Visible = False
   DBGrid2.SetFocus
   Exit Sub
End If
If Frame6.Visible = True Then
   Frame6.Visible = False
   DBGrid2.SetFocus
   Exit Sub
End If

If Frame4.Visible = True Then
   Frame4.Visible = False
   DBGrid2.SetFocus
   Exit Sub
End If
If Frame3.Visible = True Then
   Frame3.Visible = False
   DBGrid2.SetFocus
   Exit Sub
End If

If Frame2.Visible = True Then
   Frame2.Visible = False
   DBGrid2.SetFocus
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
   DBGrid2.SetFocus
   Exit Sub
End If
End If

texpadua.Hide
Unload texpadua
End Sub
Sub sql_cabeza()
On Error GoTo cmd37_err
Dim buf As String
If Len(fechai) <> 10 Then Exit Sub
If Len(fechaf) <> 10 Then Exit Sub
If Not IsDate(fechai) Then Exit Sub
If Not IsDate(fechaf) Then Exit Sub
'MsgBox cgusuario
buf = "select * from " & cgusuario & " where "
buf = buf & "  fecha>=" & "DateValue('" & fechai & "'" & ")"
buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
'buf = buf & " and tipoclie='" & tipoclie & "'"
If tipo <> "%" Then
   buf = buf & " and tipo like '" & extra_loquesea(tipo) & "'"
End If
If caja <> "%" Then
   buf = buf & " and caja like '" & extra_loquesea(caja) & "'"
End If

If serie <> "%" Then
   buf = buf & " and serie like '" & serie & "'"
End If
If numero <> "%" Then
   buf = buf & " and numero like '" & numero & "'"
End If
If codigo <> "%" Then
   buf = buf & " and codigo like '" & codigo & "'"
End If
If nombre <> "%" Then
   buf = buf & " and nombre like '" & nombre & "'"
End If
If moneda <> "%" Then
   buf = buf & " and moneda like '" & moneda & "'"
End If
If estado <> "%" Then
   buf = buf & " and estado like '" & estado & "'"
End If
If local1 <> "%" Then
   buf = buf & " and local like '" & extra_loquesea(local1) & "'"
End If
If agencia <> "%" Then
   buf = buf & " and aduana like '" & extra_loquesea(agencia) & "'"
End If
If dua <> "%" Then
   buf = buf & " and dua like '" & dua & "'"
End If


If vendedor <> "%" Then
buf = buf & " and vendedor like '" & extra_loquesea(vendedor) & "'"
End If
If cajero <> "%" Then
buf = buf & " and usuario like '" & extra_loquesea(cajero) & "'"
End If

If bodega <> "%" Then
   buf = buf & " and bodega like '" & extra_loquesea(bodega) & "'"
End If

If acu <> "C" And acu <> "V" Then
   buf = buf & " and acu='" & acu & "'"
End If
If acu = "V" Then
   buf = buf & " and (acu='A' OR acu='B' OR acu='C' OR acu='D' OR acu='G' OR acu='E' OR acu='F')"
   If Check1.Value = 1 Then
      buf = buf & " and tipo<>'5'"
   End If

End If
If acu = "C" Then
   buf = buf & " and (acu='J' OR acu='K' OR acu='L' OR acu='M' OR acu='P' OR acu='N' OR acu='O')"
   'If Check1.Value = 1 Then
   '   buf = buf & " and tipo<>'5'"
   'End If
End If
buf = buf & " and importacio='S' "

buf = buf & " order by fecha,tipo,serie,str(numero)"
'MsgBox buf

               Data2.Connect = "foxpro 2.5;"
               Data2.DatabaseName = globaldir
               Data2.RecordSource = buf
               Data2.Refresh
               SUMAR_CABEZA
               ir_ultimo
               DBGrid2.SetFocus
Exit Sub
cmd37_err:
'MsgBox "Error en select " & error$, 48, "Aviso"
Exit Sub
End Sub
Sub SUMAR_CABEZA()
Dim sdx As Double
Dim sdx1 As Double
Dim sdx2 As Double
Dim sdx3 As Double

Dim sdx4 As Double
Dim sdx5 As Double
Dim sdx6 As Double

On Error GoTo cmd7812_err

sdx1 = 0
sdx2 = 0
sdx3 = 0

sdx4 = 0
sdx5 = 0
sdx6 = 0

ir_inicio
Do
If Data2.Recordset.EOF Then Exit Do
If "" & Data2.Recordset.Fields("estado") = "2" Then
If "" & Data2.Recordset.Fields("moneda") = "S" Then
   sdx1 = sdx1 + Val("" & Data2.Recordset.Fields("subtotal"))
   sdx2 = sdx2 + Val("" & Data2.Recordset.Fields("impuesto"))
   sdx3 = sdx3 + Val("" & Data2.Recordset.Fields("total"))
End If
If "" & Data2.Recordset.Fields("moneda") = "D" Then
   sdx4 = sdx4 + Val("" & Data2.Recordset.Fields("subtotal"))
   sdx5 = sdx5 + Val("" & Data2.Recordset.Fields("impuesto"))
   sdx6 = sdx6 + Val("" & Data2.Recordset.Fields("total"))
End If
End If
Data2.Recordset.MoveNext
Loop
subtotals = Format(sdx1, "0.00")
impuestos = Format(sdx2, "0.00")
totals = Format(sdx3, "0.00")

subtotald = Format(sdx4, "0.00")
impuestod = Format(sdx5, "0.00")
totald = Format(sdx6, "0.00")
Exit Sub
cmd7812_err:
MsgBox "Error en Suma" & error$, 48, "Aviso"
Exit Sub
End Sub
Sub ir_inicio()
On Error GoTo cmd4_err
Data2.Recordset.movefist
Exit Sub
cmd4_err:
Exit Sub
End Sub
Sub consulta_codigo()
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.AddItem "Telefono"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "1"
Command1_Click

End Sub
Sub consulta_detalle()
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.AddItem "Telefono"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "2"
Command1_Click
End Sub
Sub cambia_estado()
If "" & Data2.Recordset.Fields("yausado") = "1" Then
   Data2.Recordset.Edit
   Data2.Recordset.Fields("yausado") = "0"
   Data2.Recordset.Update
   Exit Sub
End If
If "" & Data2.Recordset.Fields("yausado") = "0" Or "" & Data2.Recordset.Fields("yausado") = "" Then
   Data2.Recordset.Edit
   Data2.Recordset.Fields("yausado") = "1"
   Data2.Recordset.Update
   Exit Sub
End If

End Sub
Sub ir_ultimo()
On Error GoTo cmd123_err
Data2.Recordset.MoveLast
Exit Sub
cmd123_err:
Exit Sub
End Sub
Sub proceso_impresion1()
Dim found As Integer
Dim archivot As String
Dim ttipo As String
Dim tserie As String
Dim local1 As String
Dim tnumero As String
On Error GoTo cmd6_err:
    local1 = "" & Data2.Recordset.Fields("local")
    ttipo = "" & Data2.Recordset.Fields("tipo")
    tserie = "" & Data2.Recordset.Fields("serie")
    tnumero = "" & Data2.Recordset.Fields("numero")
    cerrar_archivo
    factura_formato local1, "" & ttipo, "" & tserie, "" & tnumero, "", 0
    cerrar_archivo
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1
    Exit Sub
cmd6_err:
    MsgBox "Mensaje, Error al iniciar Impresion " & error$
    Exit Sub
End Sub
Sub desmarca_documento()
Dim mytablex As Table
Dim mytabley As Table
Dim mytablez As Table
Dim buf1 As String
Dim te As String
Dim ts As String
Dim found As Integer

Set mytabley = mydbxglo.OpenTable("almacen")
mytabley.Index = "almacen"
Set mytablex = mydbxglo.OpenTable(dgusuariog)
mytablex.Index = "tdetalle"
found = valida_flag("" & Data2.Recordset.Fields("acu"))
If found = 0 Then
End If
If found = 1 Or found = 2 Then
   If Len("" & Data2.Recordset.Fields("tipo1")) = 0 And Len("" & Data2.Recordset.Fields("serie1")) = 0 And Len("" & Data2.Recordset.Fields("numero1")) = 0 Then
      descarga_saldo mytablex, mytabley, "" & Data2.Recordset.Fields("local"), "" & Data2.Recordset.Fields("tipo"), "" & Data2.Recordset.Fields("serie"), "" & Data2.Recordset.Fields("numero"), 1
   End If
End If
If found = 3 Then  'si es traslado
 If Len("" & Data2.Recordset.Fields("tipo1")) = 0 And Len("" & Data2.Recordset.Fields("serie1")) = 0 And Len("" & Data2.Recordset.Fields("numero1")) = 0 Then
   Set mytablez = mydbxglo.OpenTable("detalle")
   mytablez.Index = "tdetalle"
   descarga_saldo mytablez, mytabley, "" & Data2.Recordset.Fields("local"), "TE", "" & Data2.Recordset.Fields("serie"), "" & Data2.Recordset.Fields("numero"), 1
   descarga_saldo mytablez, mytabley, "" & Data2.Recordset.Fields("local"), "TS", "" & Data2.Recordset.Fields("serie"), "" & Data2.Recordset.Fields("numero"), 1
   borra_detalle mytablez, "" & Data2.Recordset.Fields("local"), "TE", "" & Data2.Recordset.Fields("serie"), "" & Data2.Recordset.Fields("numero")
   borra_detalle mytablez, "" & Data2.Recordset.Fields("local"), "TS", "" & Data2.Recordset.Fields("serie"), "" & Data2.Recordset.Fields("numero")
   mytablez.Close
End If
End If
mytablex.Close
mytabley.Close

Data2.Recordset.Edit
Data2.Recordset.Fields("estado") = "0"
Data2.Recordset.Update

buf1 = " and acu='" & "" & Data2.Recordset.Fields("acu") & "'"
mydbxglo.Execute "update  " & dgusuariog & " set estado='0'  where local='" & "" & Data2.Recordset.Fields("local") & "' and  tipo='" & "" & Data2.Recordset.Fields("tipo") & "' and serie='" & "" & Data2.Recordset.Fields("serie") & "' and  numero='" & "" & Data2.Recordset.Fields("numero") & "'" & buf1
'adicionamos la desmarcacion de las guias
desmarca_yausado "" & Data2.Recordset.Fields("LOCAL"), "" & Data2.Recordset.Fields("tipo"), "" & Data2.Recordset.Fields("serie"), "" & Data2.Recordset.Fields("numero")
mydbxglo.Execute "update  fpagov  set estado='0'  where  local='" & "" & Data2.Recordset.Fields("local") & "' and tipo='" & "" & Data2.Recordset.Fields("tipo") & "' and serie='" & "" & Data2.Recordset.Fields("serie") & "' and  numero='" & "" & Data2.Recordset.Fields("numero") & "'" & buf1
mydbxglo.Execute "update  recibo  set usado='N'  where  local='" & "" & Data2.Recordset.Fields("local") & "' and tipo='" & "" & Data2.Recordset.Fields("retipo1") & "' and numero='" & "" & Data2.Recordset.Fields("renumero1") & "'"
mydbxglo.Execute "update  recibo  set usado='N'  where  local='" & "" & Data2.Recordset.Fields("local") & "' and tipo='" & "" & Data2.Recordset.Fields("retipo1") & "' and numero='" & "" & Data2.Recordset.Fields("renumero2") & "'"
mydbxglo.Execute "update  recibo  set usado='N'  where  local='" & "" & Data2.Recordset.Fields("local") & "' and tipo='" & "" & Data2.Recordset.Fields("retipo1") & "' and numero='" & "" & Data2.Recordset.Fields("renumero3") & "'"
If acu = "Z" Then
   mydbxglo.Execute "DELETE FROM detallE where local='" & "" & Data2.Recordset.Fields("local") & "' and tipo='" & te & "'and serie='" & "" & Data2.Recordset.Fields("serie") & "'  and numero='" & "" & Data2.Recordset.Fields("numero") & "TE" & "'"
   mydbxglo.Execute "DELETE FROM detallE where local='" & "" & Data2.Recordset.Fields("local") & "' and tipo='" & ts & "'and serie='" & "" & Data2.Recordset.Fields("serie") & "'  and numero='" & "" & Data2.Recordset.Fields("numero") & "TS" & "'"
End If
 
 
If valida_flag("" & "" & Data2.Recordset.Fields("acu")) = 1 Or valida_flag("" & "" & Data2.Recordset.Fields("acu")) = 2 Then  'compras o ventas
   found = desgraba_cuentac()
End If
End Sub
Function valida_flag(buf As String)
Select Case buf
       Case "Z"
       valida_flag = 3
       Case "T", "A", "B", "C", "D", "G", "E", "F"
       valida_flag = 1
       Case "S", "J", "K", "L", "M", "P", "N", "O"
       valida_flag = 2
End Select
End Function
Function busca_tipo1(sw As Integer) As String
Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable("tipo")
mytablex.Index = "tipo"
mytablex.Seek "=", "" & Data2.Recordset.Fields("tipo")
If Not mytablex.NoMatch Then
   If sw = 0 Then
      busca_tipo1 = "" & mytablex.Fields("te")
   End If
   If sw = 1 Then
      busca_tipo1 = "" & mytablex.Fields("ts")
   End If
   
   
End If
mytablex.Close
End Function
Sub descarga_saldo(mytablex As Table, mytabley As Table, xlocal As String, xtipo As String, xserie As String, xnumero As String, sw As Integer)
Dim sdx As Double
Dim signo As Double
'AHORA HAY QUE VALIDAR QUE no existe ya cruzado el documento----


mytablex.Seek "=", xlocal, xtipo, xserie, xnumero
If mytablex.NoMatch Then Exit Sub
 'If permite_entrada_salida("" & mytablex.Fields("acu1")) = 1 Then 'si existe acu1 no descontar
 '   Exit Sub
 'End If
   Do
   If mytablex.EOF Then Exit Do
   If "" & mytablex.Fields("local") = xlocal And "" & mytablex.Fields("tipo") = xtipo And "" & mytablex.Fields("serie") = xserie And "" & mytablex.Fields("numero") = xnumero Then
      '-------------------------------------------------
      signo = 1
      'MsgBox "" & mytablex.Fields("acu")
      Select Case "" & mytablex.Fields("acu")
             Case "S", "J", "K", "L", "M", "P"
             signo = 1
             Case "T", "A", "B", "C", "D", "G"
             signo = -1
      End Select
      'If sw = 1 Then
      '   signo = signo * (-1)
      'End If
      'signo = signo * signo1
   '-------------------------------------------------

   mytabley.Seek "=", "" & mytablex.Fields("local"), "" & mytablex.Fields("producto"), "" & mytablex.Fields("bodega")
   If mytabley.NoMatch Then
      mytabley.AddNew
      mytabley.Fields("local") = "" & mytablex.Fields("local")
      mytabley.Fields("producto") = "" & mytablex.Fields("producto")
      mytabley.Fields("bodega") = "" & mytablex.Fields("bodega")
      sdx = Val("" & mytabley.Fields("saldo")) + signo * Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
      mytabley.Fields("saldo") = sdx
      mytabley.Update
      'GoTo busden
   End If
   If Not mytabley.NoMatch Then
      '-------------------------------
      If sw = 0 Then
         mytabley.Edit
         sdx = Val("" & mytabley.Fields("saldo")) + signo * Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
         mytabley.Fields("saldo") = sdx
         decarga_saldo_talla mytabley, mytablex, signo
         mytabley.Update
      End If
      If sw = 1 Then
         mytabley.Edit
         sdx = Val("" & mytabley.Fields("saldo")) - signo * Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
         mytabley.Fields("saldo") = sdx
         decarga_saldo_talla mytabley, mytablex, signo
        
         mytabley.Update
      End If
      '-------------------------------
   End If

   '-------------------------------------------------
   Else
   Exit Do
End If
mytablex.MoveNext
Loop
End Sub

Sub borra_detalle(mytablex As Table, xlocal As String, xtipo As String, xserie As String, xnumero As String)
aimbi1:
mytablex.Seek "=", xlocal, xtipo, xserie, xnumero
If Not mytablex.NoMatch Then
   mytablex.Delete
  GoTo aimbi1
End If
End Sub
Sub desmarca_yausado(buf0 As String, buf1 As String, buf2 As String, buf3 As String)
Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable(cgusuario)
mytablex.Index = "tfactura"
mytablex.Seek "=", buf0, buf1, buf2, buf3
If Not mytablex.NoMatch Then
   descarga_el_uso "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo1"), "" & mytablex.Fields("serie1"), "" & mytablex.Fields("numero1"), "0"
   descarga_el_uso "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo1"), "" & mytablex.Fields("serie2"), "" & mytablex.Fields("numero2"), "0"
   descarga_el_uso "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo1"), "" & mytablex.Fields("serie3"), "" & mytablex.Fields("numero3"), "0"
   descarga_el_uso "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo1"), "" & mytablex.Fields("serie4"), "" & mytablex.Fields("numero4"), "0"
   descarga_el_uso "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo1"), "" & mytablex.Fields("serie5"), "" & mytablex.Fields("numero5"), "0"
   descarga_el_uso "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo1"), "" & mytablex.Fields("serie6"), "" & mytablex.Fields("numero6"), "0"
   descarga_el_uso "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo1"), "" & mytablex.Fields("serie7"), "" & mytablex.Fields("numero7"), "0"
End If
'------------------------------------- ------------
mytablex.Close
 
End Sub
Sub descarga_el_uso(buf0 As String, buf1 As String, buf2 As String, buf3 As String, xsw As String)
Dim mytablex As Table
If Len(buf1) = 0 Then Exit Sub
If Len(buf2) = 0 Then Exit Sub
If Len(buf3) = 0 Then Exit Sub
Set mytablex = mydbxglo.OpenTable(cgusuario)
mytablex.Index = "tfacimp"
mytablex.Seek "=", buf0, buf1, buf2, buf3
If Not mytablex.NoMatch Then
   mytablex.Edit
   mytablex.Fields("yausado") = xsw
   mytablex.Update
End If
'------------------------------------- ------------
mytablex.Close
 
End Sub
Sub decarga_saldo_talla(mytablex As Table, mytabley As Table, signo As Double)
Dim sdx As Double
sdx = Val("" & mytablex.Fields("t1")) + signo * Val("" & mytabley.Fields("t1"))
mytablex.Fields("t1") = sdx
sdx = Val("" & mytablex.Fields("t2")) + signo * Val("" & mytabley.Fields("t2"))
mytablex.Fields("t2") = sdx
sdx = Val("" & mytablex.Fields("t3")) + signo * Val("" & mytabley.Fields("t3"))
mytablex.Fields("t3") = sdx
sdx = Val("" & mytablex.Fields("t4")) + signo * Val("" & mytabley.Fields("t4"))
mytablex.Fields("t4") = sdx
sdx = Val("" & mytablex.Fields("t5")) + signo * Val("" & mytabley.Fields("t5"))
mytablex.Fields("t5") = sdx
sdx = Val("" & mytablex.Fields("t6")) + signo * Val("" & mytabley.Fields("t6"))
mytablex.Fields("t6") = sdx
sdx = Val("" & mytablex.Fields("t7")) + signo * Val("" & mytabley.Fields("t7"))
mytablex.Fields("t7") = sdx
sdx = Val("" & mytablex.Fields("t8")) + signo * Val("" & mytabley.Fields("t8"))
mytablex.Fields("t8") = sdx
sdx = Val("" & mytablex.Fields("t9")) + signo * Val("" & mytabley.Fields("t9"))
mytablex.Fields("t9") = sdx
sdx = Val("" & mytablex.Fields("t10")) + signo * Val("" & mytabley.Fields("t10"))
mytablex.Fields("t10") = sdx
sdx = Val("" & mytablex.Fields("t11")) + signo * Val("" & mytabley.Fields("t11"))
mytablex.Fields("t11") = sdx
sdx = Val("" & mytablex.Fields("t12")) + signo * Val("" & mytabley.Fields("t12"))
mytablex.Fields("t12") = sdx
sdx = Val("" & mytablex.Fields("t13")) + signo * Val("" & mytabley.Fields("t13"))
mytablex.Fields("t13") = sdx
sdx = Val("" & mytablex.Fields("t14")) + signo * Val("" & mytabley.Fields("t14"))
mytablex.Fields("t14") = sdx
sdx = Val("" & mytablex.Fields("t15")) + signo * Val("" & mytabley.Fields("t15"))
mytablex.Fields("t15") = sdx
sdx = Val("" & mytablex.Fields("t16")) + signo * Val("" & mytabley.Fields("t16"))
mytablex.Fields("t16") = sdx
End Sub

Private Sub local1_Click()
sql_cabeza
End Sub

Private Sub local1_KeyPress(KeyAscii As Integer)
sql_cabeza
End Sub

Private Sub mio8923_Click()
Dim found As Integer
On Error GoTo cmd27_err
If Frame4.Visible = True Then Exit Sub

If Frame2.Visible = True Then Exit Sub
If Frame3.Visible = True Then Exit Sub

If "" & Data2.Recordset.Fields("estado") <> "0" Then
   MsgBox "Estado debe estar =0", 48, "Aviso"
   Exit Sub
End If
found = copiar_temporal()
If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   Exit Sub
End If
If acu = "Z" Then
   'tfacimp.Label14.Visible = True
   tfacimp.Label38.Visible = True
   'tfacimp.localf.Visible = True
   
   tfacimp.Label2.Caption = "Cod.Int."
End If

If acu = "V" Then
tfacimp.Label2 = "CodClie"
tfacimp.Caption = "Facturacion x Ventas"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfacimp.bandera = "Modifica"
tfacimp.cmdAddEntry.Enabled = False
tfacimp.dnu834.Enabled = False
tfacimp.acu = "V"
tfacimp.tipoclie = tipoclie
tfacimp.zlocal = "" & Data2.Recordset.Fields("local")
tfacimp.ztipo = "" & Data2.Recordset.Fields("tipo")
tfacimp.zserie = "" & Data2.Recordset.Fields("serie")
tfacimp.znumero = "" & Data2.Recordset.Fields("numero")
tfacimp.Show 1
sql_cabeza

End If
If acu = "H" Then
tfacimp.Label2 = "CodClie"
tfacimp.Caption = "Cotizacion x Ventas"
cgusuario = "ccotizav"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "dcotizav"
tfacimp.bandera = "Modifica"
tfacimp.cmdAddEntry.Enabled = False
tfacimp.dnu834.Enabled = False
tfacimp.acu = "H"
tfacimp.tipoclie = tipoclie
tfacimp.zlocal = "" & Data2.Recordset.Fields("local")
tfacimp.ztipo = "" & Data2.Recordset.Fields("tipo")
tfacimp.zserie = "" & Data2.Recordset.Fields("serie")
tfacimp.znumero = "" & Data2.Recordset.Fields("numero")
tfacimp.Show 1
End If
If acu = "I" Then
tfacimp.Label2 = "CodClie"
tfacimp.Caption = "Pedidos x Ventas"
cgusuario = "cpedidov"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "dpedidov"
tfacimp.bandera = "Modifica"
tfacimp.cmdAddEntry.Enabled = False
tfacimp.dnu834.Enabled = False
tfacimp.acu = "I"
tfacimp.tipoclie = tipoclie
tfacimp.zlocal = "" & Data2.Recordset.Fields("local")
tfacimp.ztipo = "" & Data2.Recordset.Fields("tipo")
tfacimp.zserie = "" & Data2.Recordset.Fields("serie")
tfacimp.znumero = "" & Data2.Recordset.Fields("numero")
tfacimp.Show 1
End If

If acu = "T" Then
tfacimp.Label2 = "CodClie"
tfacimp.Caption = "Guia Remision x Ventas"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfacimp.bandera = "Modifica"
tfacimp.cmdAddEntry.Enabled = False
tfacimp.dnu834.Enabled = False
tfacimp.acu = "T"
tfacimp.tipoclie = tipoclie
tfacimp.zlocal = "" & Data2.Recordset.Fields("local")
tfacimp.ztipo = "" & Data2.Recordset.Fields("tipo")
tfacimp.zserie = "" & Data2.Recordset.Fields("serie")
tfacimp.znumero = "" & Data2.Recordset.Fields("numero")
tfacimp.Show 1
End If

If acu = "E" Then
tfacimp.Label2 = "CodClie"
tfacimp.Caption = "Nota Credito x Ventas"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfacimp.bandera = "Modifica"
tfacimp.cmdAddEntry.Enabled = False
tfacimp.dnu834.Enabled = False
tfacimp.acu = "E"
tfacimp.tipoclie = tipoclie
tfacimp.zlocal = "" & Data2.Recordset.Fields("local")
tfacimp.ztipo = "" & Data2.Recordset.Fields("tipo")
tfacimp.zserie = "" & Data2.Recordset.Fields("serie")
tfacimp.znumero = "" & Data2.Recordset.Fields("numero")
tfacimp.Show 1
End If
If acu = "R" Then
tfacimp.Label2 = "CodProv"
tfacimp.Caption = "Orden Compra"
cgusuario = "CORDENC"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DORDENC"
tfacimp.bandera = "Modifica"
tfacimp.cmdAddEntry.Enabled = False
tfacimp.dnu834.Enabled = False
tfacimp.acu = "R"
tfacimp.tipoclie = tipoclie
tfacimp.zlocal = "" & Data2.Recordset.Fields("local")
tfacimp.ztipo = "" & Data2.Recordset.Fields("tipo")
tfacimp.zserie = "" & Data2.Recordset.Fields("serie")
tfacimp.znumero = "" & Data2.Recordset.Fields("numero")
tfacimp.Show 1
End If
If acu = "F" Then
tfacimp.Label2 = "CodClie"
tfacimp.Caption = "Nota Debito x Ventas"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"

tfacimp.acu = "F"
tfacimp.tipoclie = tipoclie
tfacimp.cmdAddEntry.Enabled = False
tfacimp.dnu834.Enabled = False
tfacimp.bandera = "Modifica"
tfacimp.zlocal = "" & Data2.Recordset.Fields("local")
tfacimp.ztipo = "" & Data2.Recordset.Fields("tipo")
tfacimp.zserie = "" & Data2.Recordset.Fields("serie")
tfacimp.znumero = "" & Data2.Recordset.Fields("numero")
tfacimp.Show 1
End If

If acu = "S" Then
tfacimp.Label2 = "CodProv"
tfacimp.Caption = "Guia Remision Entrada"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfacimp.tipoclie = tipoclie
tfacimp.acu = "S"
tfacimp.cmdAddEntry.Enabled = False
tfacimp.dnu834.Enabled = False
tfacimp.bandera = "Modifica"
tfacimp.zlocal = "" & Data2.Recordset.Fields("local")
tfacimp.ztipo = "" & Data2.Recordset.Fields("tipo")
tfacimp.zserie = "" & Data2.Recordset.Fields("serie")
tfacimp.znumero = "" & Data2.Recordset.Fields("numero")
tfacimp.Show 1
End If
If acu = "C" Then
tfacimp.Label2 = "CodProv"
tfacimp.Caption = "Factura de Compra"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfacimp.acu = "C"
tfacimp.tipoclie = tipoclie
tfacimp.cmdAddEntry.Enabled = False
tfacimp.dnu834.Enabled = False
tfacimp.bandera = "Modifica"
tfacimp.zlocal = "" & Data2.Recordset.Fields("local")
tfacimp.ztipo = "" & Data2.Recordset.Fields("tipo")
tfacimp.zserie = "" & Data2.Recordset.Fields("serie")
tfacimp.znumero = "" & Data2.Recordset.Fields("numero")
tfacimp.Show 1
End If

If acu = "N" Then
tfacimp.Label2 = "CodProv"
tfacimp.Caption = "Nota Credito Compra"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfacimp.acu = "N"
tfacimp.tipoclie = tipoclie
tfacimp.cmdAddEntry.Enabled = False
tfacimp.dnu834.Enabled = False
tfacimp.bandera = "Modifica"
tfacimp.zlocal = "" & Data2.Recordset.Fields("local")
tfacimp.ztipo = "" & Data2.Recordset.Fields("tipo")
tfacimp.zserie = "" & Data2.Recordset.Fields("serie")
tfacimp.znumero = "" & Data2.Recordset.Fields("numero")
tfacimp.Show 1
End If

If acu = "O" Then
tfacimp.Label2 = "CodProv"
tfacimp.Caption = "Nota debito de Compra"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfacimp.acu = "O"
tfacimp.tipoclie = tipoclie
tfacimp.cmdAddEntry.Enabled = False
tfacimp.dnu834.Enabled = False
tfacimp.bandera = "Modifica"
tfacimp.zlocal = "" & Data2.Recordset.Fields("local")
tfacimp.ztipo = "" & Data2.Recordset.Fields("tipo")
tfacimp.zserie = "" & Data2.Recordset.Fields("serie")
tfacimp.znumero = "" & Data2.Recordset.Fields("numero")
tfacimp.Show 1
End If
If acu = "Q" Then
tfacimp.Label2 = "CodProv"
tfacimp.Caption = "Pedido Almacen"
cgusuario = "CREQUISA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DREQUISA"
tfacimp.acu = "Q"
tfacimp.tipoclie = tipoclie
tfacimp.cmdAddEntry.Enabled = False
tfacimp.dnu834.Enabled = False
tfacimp.bandera = "Modifica"
tfacimp.zlocal = "" & Data2.Recordset.Fields("local")
tfacimp.ztipo = "" & Data2.Recordset.Fields("tipo")
tfacimp.zserie = "" & Data2.Recordset.Fields("serie")
tfacimp.znumero = "" & Data2.Recordset.Fields("numero")
tfacimp.Show 1
End If
If acu = "Z" Then
tfacimp.Caption = "Traslado entre almacen de un mismo establecimiento"
cgusuario = "CTRASLAD"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DTRASLAD"
tfacimp.acu = "Z"
tfacimp.tipoclie = tipoclie
tfacimp.cmdAddEntry.Enabled = False
tfacimp.dnu834.Enabled = False
tfacimp.bandera = "Modifica"
tfacimp.zlocal = "" & Data2.Recordset.Fields("local")
tfacimp.ztipo = "" & Data2.Recordset.Fields("tipo")
tfacimp.zserie = "" & Data2.Recordset.Fields("serie")
tfacimp.znumero = "" & Data2.Recordset.Fields("numero")
tfacimp.Show 1
End If


Exit Sub
cmd27_err:
Exit Sub
End Sub
Sub pone_registro()

End Sub

Private Sub mit56232_Click()
If Frame1.Visible = True Then Exit Sub
If Frame2.Visible = True Then Exit Sub
If Frame3.Visible = True Then Exit Sub
zooma = "Zomm"
visualizar_zoom
Exit Sub

Frame4.Visible = True
sql_detalles
dbgrid3.SetFocus

End Sub
Sub visualizar_zoom()
Dim found As Integer
On Error GoTo cmd278_err
If Frame4.Visible = True Then Exit Sub

If Frame2.Visible = True Then Exit Sub
If Frame3.Visible = True Then Exit Sub

'If "" & Data2.Recordset.Fields("estado") <> "0" Then
'   MsgBox "Estado debe estar =0", 48, "Aviso"
'   Exit Sub
'End If
found = copiar_temporal()
If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   Exit Sub
End If
If acu = "Z" Then
   'tfacimp.Label14.Visible = True
   tfacimp.Label38.Visible = True
   'tfacimp.localf.Visible = True
   'tfacimp.bodegaf.Visible = True
   tfacimp.Label2.Caption = "Cod.Int."
End If

If acu = "V" Then
tfacimp.Label2 = "CodClie"
tfacimp.Caption = "Facturacion x Ventas"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfacimp.bandera = "Ver"
tfacimp.cmdAddEntry.Enabled = False
tfacimp.dnu834.Enabled = False
tfacimp.acu = "V"
tfacimp.tipoclie = tipoclie
tfacimp.zlocal = "" & Data2.Recordset.Fields("local")
tfacimp.ztipo = "" & Data2.Recordset.Fields("tipo")
tfacimp.zserie = "" & Data2.Recordset.Fields("serie")
tfacimp.znumero = "" & Data2.Recordset.Fields("numero")
tfacimp.Show 1
'sql_cabeza

End If
If acu = "H" Then
tfacimp.Label2 = "CodClie"
tfacimp.Caption = "Cotizacion x Ventas"
cgusuario = "ccotizav"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "dcotizav"
tfacimp.bandera = "Ver"
tfacimp.cmdAddEntry.Enabled = False
tfacimp.dnu834.Enabled = False
tfacimp.acu = "H"
tfacimp.tipoclie = tipoclie
tfacimp.zlocal = "" & Data2.Recordset.Fields("local")
tfacimp.ztipo = "" & Data2.Recordset.Fields("tipo")
tfacimp.zserie = "" & Data2.Recordset.Fields("serie")
tfacimp.znumero = "" & Data2.Recordset.Fields("numero")
tfacimp.Show 1
End If
If acu = "I" Then
tfacimp.Label2 = "CodClie"
tfacimp.Caption = "Pedidos x Ventas"
cgusuario = "cpedidov"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "dpedidov"
tfacimp.bandera = "Ver"
tfacimp.cmdAddEntry.Enabled = False
tfacimp.dnu834.Enabled = False
tfacimp.acu = "I"
tfacimp.tipoclie = tipoclie
tfacimp.zlocal = "" & Data2.Recordset.Fields("local")
tfacimp.ztipo = "" & Data2.Recordset.Fields("tipo")
tfacimp.zserie = "" & Data2.Recordset.Fields("serie")
tfacimp.znumero = "" & Data2.Recordset.Fields("numero")
tfacimp.Show 1
End If

If acu = "T" Then
tfacimp.Label2 = "CodClie"
tfacimp.Caption = "Guia Remision x Ventas"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfacimp.bandera = "Ver"
tfacimp.cmdAddEntry.Enabled = False
tfacimp.dnu834.Enabled = False
tfacimp.acu = "T"
tfacimp.tipoclie = tipoclie
tfacimp.zlocal = "" & Data2.Recordset.Fields("local")
tfacimp.ztipo = "" & Data2.Recordset.Fields("tipo")
tfacimp.zserie = "" & Data2.Recordset.Fields("serie")
tfacimp.znumero = "" & Data2.Recordset.Fields("numero")
tfacimp.Show 1
End If

If acu = "E" Then
tfacimp.Label2 = "CodClie"
tfacimp.Caption = "Nota Credito x Ventas"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfacimp.bandera = "Ver"
tfacimp.cmdAddEntry.Enabled = False
tfacimp.dnu834.Enabled = False
tfacimp.acu = "E"
tfacimp.tipoclie = tipoclie
tfacimp.zlocal = "" & Data2.Recordset.Fields("local")
tfacimp.ztipo = "" & Data2.Recordset.Fields("tipo")
tfacimp.zserie = "" & Data2.Recordset.Fields("serie")
tfacimp.znumero = "" & Data2.Recordset.Fields("numero")
tfacimp.Show 1
End If
If acu = "R" Then
tfacimp.Label2 = "CodProv"
tfacimp.Caption = "Orden Compra"
cgusuario = "CORDENC"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DORDENC"
tfacimp.bandera = "Ver"
tfacimp.cmdAddEntry.Enabled = False
tfacimp.dnu834.Enabled = False
tfacimp.acu = "R"
tfacimp.tipoclie = tipoclie
tfacimp.zlocal = "" & Data2.Recordset.Fields("local")
tfacimp.ztipo = "" & Data2.Recordset.Fields("tipo")
tfacimp.zserie = "" & Data2.Recordset.Fields("serie")
tfacimp.znumero = "" & Data2.Recordset.Fields("numero")
tfacimp.Show 1
End If
If acu = "F" Then
tfacimp.Label2 = "CodClie"
tfacimp.Caption = "Nota Debito x Ventas"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"

tfacimp.acu = "F"
tfacimp.tipoclie = tipoclie
tfacimp.cmdAddEntry.Enabled = False
tfacimp.dnu834.Enabled = False
tfacimp.bandera = "Ver"
tfacimp.zlocal = "" & Data2.Recordset.Fields("local")
tfacimp.ztipo = "" & Data2.Recordset.Fields("tipo")
tfacimp.zserie = "" & Data2.Recordset.Fields("serie")
tfacimp.znumero = "" & Data2.Recordset.Fields("numero")
tfacimp.Show 1
End If

If acu = "S" Then
tfacimp.Label2 = "CodProv"
tfacimp.Caption = "Guia Remision Entrada"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfacimp.tipoclie = tipoclie
tfacimp.acu = "S"
tfacimp.cmdAddEntry.Enabled = False
tfacimp.dnu834.Enabled = False
tfacimp.bandera = "Ver"
tfacimp.zlocal = "" & Data2.Recordset.Fields("local")
tfacimp.ztipo = "" & Data2.Recordset.Fields("tipo")
tfacimp.zserie = "" & Data2.Recordset.Fields("serie")
tfacimp.znumero = "" & Data2.Recordset.Fields("numero")
tfacimp.Show 1
End If
If acu = "C" Then
tfacimp.Label2 = "CodProv"
tfacimp.Caption = "Factura de Compra"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfacimp.acu = "C"
tfacimp.tipoclie = tipoclie
tfacimp.cmdAddEntry.Enabled = False
tfacimp.dnu834.Enabled = False
tfacimp.bandera = "Ver"
tfacimp.zlocal = "" & Data2.Recordset.Fields("local")
tfacimp.ztipo = "" & Data2.Recordset.Fields("tipo")
tfacimp.zserie = "" & Data2.Recordset.Fields("serie")
tfacimp.znumero = "" & Data2.Recordset.Fields("numero")
tfacimp.Show 1
End If

If acu = "N" Then
tfacimp.Label2 = "CodProv"
tfacimp.Caption = "Nota Credito Compra"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfacimp.acu = "N"
tfacimp.tipoclie = tipoclie
tfacimp.cmdAddEntry.Enabled = False
tfacimp.dnu834.Enabled = False
tfacimp.bandera = "Ver"
tfacimp.zlocal = "" & Data2.Recordset.Fields("local")
tfacimp.ztipo = "" & Data2.Recordset.Fields("tipo")
tfacimp.zserie = "" & Data2.Recordset.Fields("serie")
tfacimp.znumero = "" & Data2.Recordset.Fields("numero")
tfacimp.Show 1
End If

If acu = "O" Then
tfacimp.Label2 = "CodProv"
tfacimp.Caption = "Nota debito de Compra"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfacimp.acu = "O"
tfacimp.tipoclie = tipoclie
tfacimp.cmdAddEntry.Enabled = False
tfacimp.dnu834.Enabled = False
tfacimp.bandera = "Ver"
tfacimp.zlocal = "" & Data2.Recordset.Fields("local")
tfacimp.ztipo = "" & Data2.Recordset.Fields("tipo")
tfacimp.zserie = "" & Data2.Recordset.Fields("serie")
tfacimp.znumero = "" & Data2.Recordset.Fields("numero")
tfacimp.Show 1
End If
If acu = "Q" Then
tfacimp.Label2 = "CodProv"
tfacimp.Caption = "Pedido Almacen"
cgusuario = "CREQUISA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DREQUISA"
tfacimp.acu = "Q"
tfacimp.tipoclie = tipoclie
tfacimp.cmdAddEntry.Enabled = False
tfacimp.dnu834.Enabled = False
tfacimp.bandera = "Ver"
tfacimp.zlocal = "" & Data2.Recordset.Fields("local")
tfacimp.ztipo = "" & Data2.Recordset.Fields("tipo")
tfacimp.zserie = "" & Data2.Recordset.Fields("serie")
tfacimp.znumero = "" & Data2.Recordset.Fields("numero")
tfacimp.Show 1
End If
If acu = "Z" Then
tfacimp.Caption = "Traslado entre almacen de un mismo establecimiento"
cgusuario = "CTRASLAD"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DTRASLAD"
tfacimp.acu = "Z"
tfacimp.tipoclie = tipoclie
tfacimp.cmdAddEntry.Enabled = False
tfacimp.dnu834.Enabled = False
tfacimp.bandera = "Ver"
tfacimp.zlocal = "" & Data2.Recordset.Fields("local")
tfacimp.ztipo = "" & Data2.Recordset.Fields("tipo")
tfacimp.zserie = "" & Data2.Recordset.Fields("serie")
tfacimp.znumero = "" & Data2.Recordset.Fields("numero")
tfacimp.Show 1
End If


Exit Sub
cmd278_err:
Exit Sub

End Sub
Sub sql_detalles()
Dim buf As String
Dim sdx As Double
On Error GoTo cmd321_err
sdx = 0
buf = "select * from " & dgusuariog & " where "
buf = buf & " local='" & "" & Data2.Recordset.Fields("local") & "'"
buf = buf & " and tipo='" & "" & Data2.Recordset.Fields("tipo") & "'"
buf = buf & " and serie='" & "" & Data2.Recordset.Fields("serie") & "'"
buf = buf & " and numero='" & "" & Data2.Recordset.Fields("numero") & "'"
Data4.Connect = "foxpro 2.5;"
Data4.DatabaseName = globaldir
Data4.RecordSource = buf
Data4.Refresh
Do
If Data4.Recordset.EOF Then Exit Do
sdx = sdx + Val("" & Data4.Recordset.Fields("cantidad")) * Val("" & Data4.Recordset.Fields("factor"))
Data4.Recordset.MoveNext
Loop
Command8.Caption = "Valida:" & sdx
buf = "select * from fpagov where "
buf = buf & " local='" & "" & Data2.Recordset.Fields("local") & "'"
buf = buf & " and tipo='" & "" & Data2.Recordset.Fields("tipo") & "'"
buf = buf & " and serie='" & "" & Data2.Recordset.Fields("serie") & "'"
buf = buf & " and numero='" & "" & Data2.Recordset.Fields("numero") & "'"
Data5.Connect = "foxpro 2.5;"
Data5.DatabaseName = globaldir
Data5.RecordSource = buf
Data5.Refresh
yausado.Value = 0
If "" & Data2.Recordset.Fields("yausado") = "1" Then
   yausado.Value = 1
End If

Exit Sub
cmd321_err:
MsgBox "Aviso sql Detalles " + error, 48, "Aviso"
Exit Sub
End Sub

Private Sub modi343_Click()
On Error GoTo cmd117_err
Dim found As Integer
Dim buf As String
If Frame4.Visible = True Then Exit Sub
If Frame2.Visible = True Then Exit Sub
If Frame3.Visible = True Then Exit Sub
buf = "" & Data2.Recordset.Fields("estado")
If "" & Data2.Recordset.Fields("estado") <> "2" Then Exit Sub
Frame2.Visible = True
Frame2.Caption = "DESMARCA"
clave = ""
clave.SetFocus
Exit Sub
cmd117_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub
End Sub
Function verificar_recibo(buf As String)
Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable(buf)
mytablex.Index = "tmpcta1"
mytablex.Seek "=", "" & Data2.Recordset.Fields("local"), "" & Data2.Recordset.Fields("tipo"), "" & Data2.Recordset.Fields("serie"), "" & Data2.Recordset.Fields("numero")
If Not mytablex.NoMatch Then
   verificar_recibo = 1
End If
mytablex.Close
End Function
Function desgraba_cuentac()

Dim mytabley As Table
Dim i As Integer
'---------- validando si es cuenta corriente

If valida_flag("" & Data2.Recordset.Fields("acu")) = 2 Then   'compras
Set mytabley = mydbxglo.OpenTable("cuentap")
End If
If valida_flag("" & Data2.Recordset.Fields("acu")) = 1 Then   'ventas
Set mytabley = mydbxglo.OpenTable("cuentac")
End If
mytabley.Index = "cuentac"
amk1:
mytabley.Seek "=", "" & Data2.Recordset.Fields("local"), "" & Data2.Recordset.Fields("tipo"), "" & Data2.Recordset.Fields("serie"), "" & Data2.Recordset.Fields("numero"), "1"
If Not mytabley.NoMatch Then
   mytabley.Delete
   GoTo amk1
End If
mytabley.Close
 
End Function
Function genera_traslado_automata()
Dim found As Integer
Dim sdx As Double
Dim mytablez As Table
Dim mytablex As Table
Dim mytabley As Table
Dim mytablew As Table
Set mytablex = mydbxglo.OpenTable("ctraslad")
mytablex.Index = "tfacimp"
mytablex.Seek "=", ulocal, extra_loquesea(utipo), userie, unumero
If Not mytablex.NoMatch Then
   MsgBox "Documento ya existe", 48, "Aviso"
   mytablex.Close
   Exit Function
End If
If mytablex.NoMatch Then
   mytablex.AddNew
   grabado_cabecera_tra mytablex
   mytablex.Update
End If
mytablex.Close
'detalle
Set mytablew = mydbxglo.OpenTable("almacen")
mytablew.Index = "almacen"

Set mytablez = mydbxglo.OpenTable("detalle")
mytablez.Index = "Tdetalle"
Set mytabley = mydbxglo.OpenTable("dtraslad")
Set mytablex = mydbxglo.OpenTable("Drequisa")
mytablex.Index = "Tdetalle"
mytablex.Seek "=", "" & Data2.Recordset.Fields("local"), "" & Data2.Recordset.Fields("tipo"), "" & Data2.Recordset.Fields("serie"), "" & Data2.Recordset.Fields("numero")
If Not mytablex.NoMatch Then
   Do
   If mytablex.EOF Then Exit Do
      If "" & mytablex.Fields("local") = "" & Data2.Recordset.Fields("local") And "" & mytablex.Fields("tipo") = "" & Data2.Recordset.Fields("tipo") And "" & mytablex.Fields("serie") = "" & Data2.Recordset.Fields("serie") And "" & mytablex.Fields("numero") = "" & Data2.Recordset.Fields("numero") Then
         If Val("" & mytablex.Fields("t15")) > 0 Then 'si tiene cantidad para despachaar se genera
            mytabley.AddNew
            grabando_detalle_tra mytabley, mytablex
            mytabley.Update
            adiciona_traslado_es mytablez, mytablex
         End If
      Else
      Exit Do
      End If
      mytablex.MoveNext
   Loop
End If
   descarga_saldo mytablez, mytablew, ulocal, "TE", userie, unumero, 1
   descarga_saldo mytablez, mytablew, ulocal, "TS", userie, unumero, 0
'descarga saldo
mytablex.Close
mytabley.Close
mytablez.Close
mytablew.Close


found = busca_utipo(extra_loquesea(utipo), 1)
genera_traslado_automata = 1
End Function
Sub grabado_cabecera_tra(mytablex As Table)
mytablex.Fields("tipo1") = "" & Data2.Recordset.Fields("tipo")
mytablex.Fields("serie1") = "" & Data2.Recordset.Fields("serie")
mytablex.Fields("numero1") = "" & Data2.Recordset.Fields("numero")
mytablex.Fields("acu1") = "" & Data2.Recordset.Fields("acu")

mytablex.Fields("observa") = ""
mytablex.Fields("adetotal") = 0
mytablex.Fields("acuenta") = 0
mytablex.Fields("retipo1") = ""
mytablex.Fields("renumero1") = ""
mytablex.Fields("renumero2") = ""
mytablex.Fields("renumero3") = ""
mytablex.Fields("retotal1") = 0
mytablex.Fields("retotal2") = 0
mytablex.Fields("retotal3") = 0
mytablex.Fields("retotal") = 0
mytablex.Fields("zona") = ""
mytablex.Fields("nombre") = "TRASLADO"
mytablex.Fields("estado") = "2"
mytablex.Fields("tipoclie") = "V"
mytablex.Fields("tipo") = extra_loquesea(utipo)
mytablex.Fields("serie") = userie
mytablex.Fields("numero") = unumero
mytablex.Fields("codigo") = "01"
mytablex.Fields("partida") = ""
mytablex.Fields("destino") = ""
mytablex.Fields("yausado") = "0"
mytablex.Fields("nro_items") = 0
mytablex.Fields("fecha") = Format(ufecha, "dd/mm/yyyy")
mytablex.Fields("fechae") = Format(Now, "dd/mm/yyyy")
mytablex.Fields("moneda") = "S"
mytablex.Fields("vendedor") = extra_loquesea(uvendedor)
mytablex.Fields("fpago") = "1"
mytablex.Fields("transporte") = ""
mytablex.Fields("paridad") = 2.8
mytablex.Fields("dias") = 1
mytablex.Fields("bodega") = ubodega
mytablex.Fields("bodegaf") = ubodegaf
'mytablex.Fields("observa") = ""
mytablex.Fields("usuario") = ""
mytablex.Fields("caja") = ""
mytablex.Fields("turno") = ""
mytablex.Fields("acu") = "Z"
mytablex.Fields("acu1") = ""
mytablex.Fields("flage") = ""
mytablex.Fields("telefono") = ""
mytablex.Fields("hora") = Format(Now, "hh:MM")
mytablex.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
mytablex.Fields("gravado") = 0
mytablex.Fields("total") = 0
mytablex.Fields("redondeo") = 0
mytablex.Fields("descuento") = 0
mytablex.Fields("neto") = 0
mytablex.Fields("impuesto") = 0
mytablex.Fields("subtotal") = 0

mytablex.Fields("tipo1") = ""
mytablex.Fields("serie1") = ""
mytablex.Fields("serie2") = ""
mytablex.Fields("serie3") = ""
mytablex.Fields("serie4") = ""
mytablex.Fields("serie5") = ""
mytablex.Fields("serie6") = ""
mytablex.Fields("serie7") = ""

mytablex.Fields("numero1") = ""
mytablex.Fields("numero2") = ""
mytablex.Fields("numero3") = ""
mytablex.Fields("numero4") = ""
mytablex.Fields("numero5") = ""
mytablex.Fields("numero6") = ""
mytablex.Fields("numero7") = ""
mytablex.Fields("c1") = 0
mytablex.Fields("c2") = 0
mytablex.Fields("c3") = 0
mytablex.Fields("c4") = 0
mytablex.Fields("c5") = 0
mytablex.Fields("c6") = 0
mytablex.Fields("c7") = 0
mytablex.Fields("c8") = 0
mytablex.Fields("c9") = 0
mytablex.Fields("local") = "01" '& Data2.Recordset.Fields("local")
mytablex.Fields("localf") = "01" '& Data2.Recordset.Fields("local")
mytablex.Fields("montopagar") = 0
mytablex.Fields("ruc") = ""
mytablex.Fields("TDOCDELI") = ""

End Sub
Sub grabando_detalle_tra(mytablex As Table, mytabley As Table)
Dim sdx As Double
Dim i As Integer
On Error GoTo cmd3218912_err
    For i = 0 To mytabley.Fields.count - 1
        mytablex.Fields(i) = mytabley.Fields(i)
    Next i
    mytablex.Fields("cantidad") = Val("" & mytablex.Fields("t15"))
    mytablex.Fields("t15") = 0
    mytablex.Fields("t16") = 0
    mytablex.Fields("local") = "" & ulocal
    mytablex.Fields("tipo") = "" & extra_loquesea(utipo)
    mytablex.Fields("serie") = "" & userie
    mytablex.Fields("numero") = "" & unumero
    mytablex.Fields("vendedor") = extra_loquesea(uvendedor)
    mytablex.Fields("tipoclie") = "V"
    mytablex.Fields("moneda") = "" & mytabley.Fields("moneda")
    mytablex.Fields("bodega") = "" & ubodega
    mytablex.Fields("bodegaf") = "" & ubodegaf
    mytablex.Fields("acu") = "Z"
    mytablex.Fields("localf") = "" & ulocal '& codigo  'si no es traslado
    mytablex.Fields("flage") = ""
    mytablex.Fields("codigo") = "01"
    mytablex.Fields("caja") = ""
    mytablex.Fields("turno") = ""
    mytablex.Fields("usuario") = ""
    mytablex.Fields("fecha") = Format(ufecha, "dd/mm/yyyy")
    mytablex.Fields("hora") = Format(Now, "hh:MM")
    mytablex.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
    mytablex.Fields("estado") = "2"
      'mytabley.Edit
      'sdx = Val("" & mytabley.Fields("t16")) + Val("" & mytabley.Fields("t15"))
      'mytabley.Fields("t16") = sdx
      'mytabley.Fields("t15") = 0
      'mytabley.Update
    
Exit Sub
cmd3218912_err:
MsgBox "Error en Grabando detalle ttra" + error$, 48, "Aviso"
Exit Sub
End Sub
Sub ir_menu_traslado()
On Error GoTo cmd89121_err
Dim mytablex As Table
utipo.Clear
uvendedor.Clear
Set mytablex = mydbxglo.OpenTable("tipo")
utipo.AddItem "%"
Do
If mytablex.EOF Then Exit Do
If "" & mytablex.Fields("tipodoc") = "Z" Then
   utipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")
End If
mytablex.MoveNext
Loop
utipo.ListIndex = 0
mytablex.Close

Set mytablex = mydbxglo.OpenTable("vendedor")
uvendedor.AddItem "%"
Do
If mytablex.EOF Then Exit Do
uvendedor.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
mytablex.MoveNext
Loop
uvendedor.ListIndex = 0
mytablex.Close
ulocal = "01"
ubodega = "01"
ubodegaf = "" & Data2.Recordset.Fields("bodega")
userie = ""
unumero = ""
ufecha = Format(Now, "dd/mm/yyyy")
Frame5.Visible = True

unombre1 = unombre_almacen(ubodega)
unombre2 = unombre_almacen(ubodegaf)
utipo.SetFocus
Exit Sub
cmd89121_err:
MsgBox "Seleccione un Numero ", 48, "Aviso"
Exit Sub


End Sub

Function busca_utipo(buf As String, sw As Integer) As Integer
Dim mytablex As Table
Dim sdx As Double

Set mytablex = mydbxglo.OpenTable("tipo")
mytablex.Index = "tipo"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   If sw = 0 Then
   If Len(userie) = 0 Then
      userie = "" & mytablex.Fields("serie")
   End If
   If Len(unumero) = 0 Then
      sdx = Val("" & mytablex.Fields("numero")) + 1
      unumero = "" & sdx
   End If
   End If
   If sw = 1 Then
      mytablex.Edit
      mytablex.Fields("numero") = unumero
      mytablex.Update
   End If
End If
mytablex.Close
End Function

Private Sub ufecha_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If Len(ufecha) = 0 Then
   ufecha = Format(Now, "dd/mm/yyyy")
End If
End Sub

Private Sub utipo_Change()
utipo_Click
End Sub

Private Sub utipo_Click()
Dim found As Integer
found = busca_utipo(extra_loquesea(utipo), 0)


End Sub
Function valida_existe_despacho()
Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable("Drequisa")
mytablex.Index = "Tdetalle"
mytablex.Seek "=", "" & Data2.Recordset.Fields("local"), "" & Data2.Recordset.Fields("tipo"), "" & Data2.Recordset.Fields("serie"), "" & Data2.Recordset.Fields("numero")
If Not mytablex.NoMatch Then
   Do
   If mytablex.EOF Then Exit Do
      If "" & mytablex.Fields("local") = "" & Data2.Recordset.Fields("local") And "" & mytablex.Fields("tipo") = "" & Data2.Recordset.Fields("tipo") And "" & mytablex.Fields("serie") = "" & Data2.Recordset.Fields("serie") And "" & mytablex.Fields("numero") = "" & Data2.Recordset.Fields("numero") Then
         If Val("" & mytablex.Fields("t15")) > 0 Then 'si tiene cantidad para despachaar se genera
            valida_existe_despacho = 1
         End If
      Else: Exit Do
      End If
      mytablex.MoveNext
   Loop
End If
mytablex.Close
End Function
Function graba_existe_despacho()
On Error GoTo cmd981211_err
Dim mytablex As Table
Dim sw As Integer
sw = 0
Set mytablex = mydbxglo.OpenTable("Drequisa")
mytablex.Index = "Tdetalle"
mytablex.Seek "=", "" & Data2.Recordset.Fields("local"), "" & Data2.Recordset.Fields("tipo"), "" & Data2.Recordset.Fields("serie"), "" & Data2.Recordset.Fields("numero")
If Not mytablex.NoMatch Then
   Do
   If mytablex.EOF Then Exit Do
      If "" & mytablex.Fields("local") = "" & Data2.Recordset.Fields("local") And "" & mytablex.Fields("tipo") = "" & Data2.Recordset.Fields("tipo") And "" & mytablex.Fields("serie") = "" & Data2.Recordset.Fields("serie") And "" & mytablex.Fields("numero") = "" & Data2.Recordset.Fields("numero") Then
         If Val("" & mytablex.Fields("t15")) > 0 Then 'si tiene cantidad para despachaar se genera
            'MsgBox "Hola"
            mytablex.Edit
            mytablex.Fields("t16") = Val("" & mytablex.Fields("t16")) + Val("" & mytablex.Fields("t15"))
            mytablex.Fields("t15") = 0
            mytablex.Update
            If Val("" & mytablex.Fields("cantidad")) <= Val(Val("" & mytablex.Fields("t16"))) Then
               If sw <> 2 Then
                  sw = 1
               End If
               Else
               sw = 2
            End If
         End If
      Else: Exit Do
      End If
      mytablex.MoveNext
   Loop
End If
mytablex.Close
If sw = 1 Then
   Data2.Recordset.Edit
   Data2.Recordset.Fields("yausado") = "1"
   Data2.Recordset.Update
End If
Exit Function
cmd981211_err:
MsgBox "Error en graba existe despacho " + error$, 48, "Aviso"
Exit Function

End Function
Sub adiciona_traslado_es(mytablex As Table, mytabley As Table)
Dim sdx As Double
Dim i As Integer
    mytablex.AddNew
    For i = 0 To mytabley.Fields.count - 1
        mytablex.Fields(i) = mytabley.Fields(i)
    Next i
        
    mytablex.Fields("cantidad") = Val("" & mytablex.Fields("t15"))
    mytablex.Fields("t15") = 0
    mytablex.Fields("t16") = 0
    mytablex.Fields("local") = "" & ulocal
    mytablex.Fields("tipo") = "TS" '& extra_loquesea(utipo)
    mytablex.Fields("serie") = "" & userie
    mytablex.Fields("numero") = "" & unumero
    mytablex.Fields("vendedor") = extra_loquesea(uvendedor)
    mytablex.Fields("tipoclie") = "V"
    mytablex.Fields("moneda") = "" & mytabley.Fields("moneda")
    mytablex.Fields("bodega") = "" & ubodega
    mytablex.Fields("bodegaf") = "" & ubodegaf
    mytablex.Fields("acu") = "T"
    mytablex.Fields("localf") = "" & ulocal '& codigo  'si no es traslado
    mytablex.Fields("flage") = ""
    mytablex.Fields("codigo") = "01"
    mytablex.Fields("caja") = ""
    mytablex.Fields("turno") = ""
    mytablex.Fields("usuario") = ""
    mytablex.Fields("fecha") = Format(ufecha, "dd/mm/yyyy")
    mytablex.Fields("hora") = Format(Now, "hh:MM")
    mytablex.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
    mytablex.Fields("estado") = "2"
    mytablex.Update
       
    mytablex.AddNew
    For i = 0 To mytabley.Fields.count - 1
        mytablex.Fields(i) = mytabley.Fields(i)
    Next i
       
    mytablex.Fields("cantidad") = Val("" & mytablex.Fields("t15"))
    mytablex.Fields("t15") = 0
    mytablex.Fields("t16") = 0
    mytablex.Fields("local") = "" & ulocal
    mytablex.Fields("tipo") = "TE" '& extra_loquesea(utipo)
    mytablex.Fields("serie") = "" & userie
    mytablex.Fields("numero") = "" & unumero
    mytablex.Fields("vendedor") = extra_loquesea(uvendedor)
    mytablex.Fields("tipoclie") = "V"
    mytablex.Fields("moneda") = "" & mytabley.Fields("moneda")
    mytablex.Fields("bodega") = "" & ubodegaf
    mytablex.Fields("bodegaf") = "" & ubodega
    mytablex.Fields("acu") = "S"
    mytablex.Fields("localf") = "" & ulocal '& codigo  'si no es traslado
    mytablex.Fields("flage") = ""
    mytablex.Fields("codigo") = "01"
    mytablex.Fields("caja") = ""
    mytablex.Fields("turno") = ""
    mytablex.Fields("usuario") = ""
    mytablex.Fields("fecha") = Format(ufecha, "dd/mm/yyyy")
    mytablex.Fields("hora") = Format(Now, "hh:MM")
    mytablex.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
    mytablex.Fields("estado") = "2"
    mytablex.Update


End Sub
Function unombre_almacen(buf As String) As String
Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable("bodega")
mytablex.Index = "codigo"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   unombre_almacen = "" & mytablex.Fields("nombre")
End If
mytablex.Close
End Function
Sub menu_excell1()
If acu = "V" Or acu = "T" Or acu = "E" Or acu = "F" Or acu = "S" Or acu = "C" Or acu = "N" Or acu = "O" Then
   cgusuario = "FACTURA"
   dgusuariog = "DETALLE"
End If
If acu = "H" Then
   cgusuario = "CCOTIZAV"
   dgusuariog = "DCOTIZAV"
End If
If acu = "I" Then
   cgusuario = "CPEDIDOV"
   dgusuariog = "DPEDIDOV"
End If
If acu = "R" Then
   cgusuario = "CORDENC"
   dgusuariog = "DORDENC"
End If
If acu = "Q" Then
   cgusuario = "CREQUISA"
   dgusuariog = "DREQUISA"
End If
If acu = "Z" Then
   cgusuario = "CTRASLAD"
   dgusuariog = "DTRASLAD"
End If
excel_paso1

End Sub
Sub menu_excell()
If acu = "V" Or acu = "T" Or acu = "E" Or acu = "F" Or acu = "S" Or acu = "C" Or acu = "N" Or acu = "O" Then
   cgusuario = "FACTURA"
   dgusuariog = "DETALLE"
End If
If acu = "H" Then
   cgusuario = "CCOTIZAV"
   dgusuariog = "DCOTIZAV"
End If
If acu = "I" Then
   cgusuario = "CPEDIDOV"
   dgusuariog = "DPEDIDOV"
End If
If acu = "R" Then
   cgusuario = "CORDENC"
   dgusuariog = "DORDENC"
End If
If acu = "Q" Then
   cgusuario = "CREQUISA"
   dgusuariog = "DREQUISA"
End If
If acu = "Z" Then
   cgusuario = "CTRASLAD"
   dgusuariog = "DTRASLAD"
End If
excel_paso

End Sub
Sub excel_paso1()
Dim sdx As String
On Error GoTo cmd813_err
sdx = "" & Data2.Recordset.Fields("numero")
conteo_excell1
Exit Sub
cmd813_err:
MsgBox "Elegir un dato ", 48, "Aviso"
Exit Sub

End Sub
Sub excel_paso()
Dim sdx As String
On Error GoTo cmd81_err
sdx = "" & Data2.Recordset.Fields("numero")
conteo_excell
Exit Sub
cmd81_err:
MsgBox "Elegir un dato ", 48, "Aviso"
Exit Sub

End Sub
Sub conteo_excell1()
Dim mytablex As Table
 Dim v, h As Integer
 Dim found As Integer
 Dim i As Integer
 Dim sdx As Double
 Dim sdx1 As Double
 Dim sdx2 As Double
 Dim vprecios(11) As String
    Dim Heading(12) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion
    On Error GoTo cmd56124_err
   'Data1.Refresh
   
   
    Heading(1) = "Codigo"
    Heading(2) = "Nombre"
    Heading(3) = "Local"
    Heading(4) = "Tipo"
    Heading(5) = "Serie"
    Heading(6) = "Numero"
    Heading(7) = "M"
    Heading(8) = "Fecha"
    Heading(9) = "Total"
    Heading(10) = "impuesto"
    Heading(11) = "Subtotal"
    
   
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_Excel(11, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
v = 4
h = 1
sdx = 0
sdx1 = 0
sdx2 = 0
Data2.Refresh
Do
If Data2.Recordset.EOF Then Exit Do
            objExcel.ActiveSheet.Cells(v, h + 0) = "" & Data2.Recordset.Fields("codigo")
            objExcel.ActiveSheet.Cells(v, h + 1) = "" & Data2.Recordset.Fields("nombre")
            objExcel.ActiveSheet.Cells(v, h + 2) = "" & Data2.Recordset.Fields("local")
            objExcel.ActiveSheet.Cells(v, h + 3) = "" & Data2.Recordset.Fields("tipo")
            objExcel.ActiveSheet.Cells(v, h + 4) = "" & Data2.Recordset.Fields("serie")
            objExcel.ActiveSheet.Cells(v, h + 5) = "" & Data2.Recordset.Fields("numero")
            objExcel.ActiveSheet.Cells(v, h + 6) = "" & Data2.Recordset.Fields("moneda")
            objExcel.ActiveSheet.Cells(v, h + 7) = "" & Data2.Recordset.Fields("fecha")
            objExcel.ActiveSheet.Cells(v, h + 8) = Val("" & Data2.Recordset.Fields("Total"))
            objExcel.ActiveSheet.Cells(v, h + 9) = Val("" & Data2.Recordset.Fields("Impuesto"))
            objExcel.ActiveSheet.Cells(v, h + 10) = Val("" & Data2.Recordset.Fields("subtotal"))
            v = v + 1
            sdx = sdx + Val("" & Data2.Recordset.Fields("total"))
            sdx1 = sdx1 + Val("" & Data2.Recordset.Fields("impuesto"))
            sdx2 = sdx2 + Val("" & Data2.Recordset.Fields("subtotal"))
  Data2.Recordset.MoveNext
  Loop
  v = v + 1
            objExcel.ActiveSheet.Cells(v, h + 0) = ""
            objExcel.ActiveSheet.Cells(v, h + 1) = ""
            objExcel.ActiveSheet.Cells(v, h + 2) = ""
            objExcel.ActiveSheet.Cells(v, h + 3) = ""
            objExcel.ActiveSheet.Cells(v, h + 4) = ""
            objExcel.ActiveSheet.Cells(v, h + 5) = ""
            objExcel.ActiveSheet.Cells(v, h + 6) = ""
            objExcel.ActiveSheet.Cells(v, h + 7) = ""
            objExcel.ActiveSheet.Cells(v, h + 8) = "" & sdx
            objExcel.ActiveSheet.Cells(v, h + 9) = "" & sdx1
            objExcel.ActiveSheet.Cells(v, h + 10) = "" & sdx2
 Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
 Exit Sub
cmd56124_err:
MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
Exit Sub

End Sub

Sub conteo_excell()
Dim mytablex As Table
 Dim v, h As Integer
 Dim found As Integer
 Dim i As Integer
 Dim sdx As Double
 Dim sdx1 As Double
 Dim vprecios(7) As String
    Dim Heading(8) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion
    On Error GoTo cmd5612_err
   'Data1.Refresh
   
   
    Heading(1) = "Producto"
    Heading(2) = "Descripcio"
    Heading(3) = "Und"
    Heading(4) = "Factor"
    Heading(5) = "cantidad"
    Heading(6) = "Precio"
    Heading(7) = "Total"
   
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_Excel(7, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
    
    
Set mytablex = mydbxglo.OpenTable(dgusuariog)
mytablex.Index = "tdetalle"
v = 5
h = 1

Data2.Refresh
Do
If Data2.Recordset.EOF Then Exit Do
sdx = 0
sdx1 = 0
objExcel.ActiveSheet.Cells(v, h) = ""
            objExcel.ActiveSheet.Cells(v, h + 0) = "Tip:" & Data2.Recordset.Fields("tipo")
            objExcel.ActiveSheet.Cells(v, h + 1) = "Ser:" & Data2.Recordset.Fields("serie")
            objExcel.ActiveSheet.Cells(v, h + 2) = "Num:" & Data2.Recordset.Fields("numero")
            objExcel.ActiveSheet.Cells(v, h + 3) = "Vend:" & Data2.Recordset.Fields("vendedor")
            objExcel.ActiveSheet.Cells(v, h + 4) = "Caja:" & Data2.Recordset.Fields("caja")
            objExcel.ActiveSheet.Cells(v, h + 5) = "Alm:" & Data2.Recordset.Fields("bodega") & " AlmF:" & Data2.Recordset.Fields("bodegaf")
            v = v + 1
mytablex.Seek "=", "" & Data2.Recordset.Fields("local"), "" & Data2.Recordset.Fields("tipo"), "" & Data2.Recordset.Fields("serie"), "" & Data2.Recordset.Fields("numero")
If Not mytablex.NoMatch Then
     Do
     If mytablex.EOF Then Exit Do
     
     If "" & mytablex.Fields("local") = "" & Data2.Recordset.Fields("local") And "" & mytablex.Fields("tipo") = "" & Data2.Recordset.Fields("tipo") And "" & mytablex.Fields("serie") = "" & Data2.Recordset.Fields("serie") And "" & mytablex.Fields("numero") = "" & Data2.Recordset.Fields("numero") Then
            sdx = sdx + Val("" & mytablex.Fields("cantidad"))
            sdx1 = sdx1 + Val("" & mytablex.Fields("total"))
            objExcel.ActiveSheet.Cells(v, h) = "" & mytablex.Fields("producto")
            objExcel.ActiveSheet.Cells(v, h + 1) = "" & mytablex.Fields("descripcio")
            objExcel.ActiveSheet.Cells(v, h + 2) = "" & mytablex.Fields("unidad")
            objExcel.ActiveSheet.Cells(v, h + 3) = "" & mytablex.Fields("factor")
            objExcel.ActiveSheet.Cells(v, h + 4) = "" & mytablex.Fields("cantidad")
            objExcel.ActiveSheet.Cells(v, h + 5) = "" & mytablex.Fields("precio")
            objExcel.ActiveSheet.Cells(v, h + 6) = "" & mytablex.Fields("total")
            v = v + 1
            Else: Exit Do
     End If
     mytablex.MoveNext
     Loop
 End If
            objExcel.ActiveSheet.Cells(v, h) = ""
            objExcel.ActiveSheet.Cells(v, h + 1) = ""
            objExcel.ActiveSheet.Cells(v, h + 2) = ""
            objExcel.ActiveSheet.Cells(v, h + 3) = ""
            objExcel.ActiveSheet.Cells(v, h + 4) = "" & sdx
            objExcel.ActiveSheet.Cells(v, h + 5) = ""
            objExcel.ActiveSheet.Cells(v, h + 6) = "" & sdx1
            v = v + 1
  Data2.Recordset.MoveNext
  Loop
 Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
 mytablex.Close
 Exit Sub
cmd5612_err:
MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
Exit Sub
End Sub
Sub conteo_excell_uno()
Dim mytablex As Table
 Dim v, h As Integer
 Dim found As Integer
 Dim i As Integer
 Dim sdx As Double
 Dim sdx1 As Double
 Dim vprecios(7) As String
    Dim Heading(8) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion
    On Error GoTo cmd561212_err
   'Data1.Refresh
   
   
    Heading(1) = "Producto"
    Heading(2) = "Descripcio"
    Heading(3) = "Und"
    Heading(4) = "Factor"
    Heading(5) = "cantidad"
    Heading(6) = "Precio"
    Heading(7) = "Total"
   
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_Excel(7, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
    
    
Set mytablex = mydbxglo.OpenTable(dgusuariog)
mytablex.Index = "tdetalle"
v = 5
h = 1
sdx = 0
sdx1 = 0
objExcel.ActiveSheet.Cells(v, h) = ""
            objExcel.ActiveSheet.Cells(v, h + 0) = "Tip:" & Data2.Recordset.Fields("tipo")
            objExcel.ActiveSheet.Cells(v, h + 1) = "Ser:" & Data2.Recordset.Fields("serie")
            objExcel.ActiveSheet.Cells(v, h + 2) = "Num:" & Data2.Recordset.Fields("numero")
            objExcel.ActiveSheet.Cells(v, h + 3) = "Vend:" & Data2.Recordset.Fields("vendedor")
            objExcel.ActiveSheet.Cells(v, h + 4) = "Caja:" & Data2.Recordset.Fields("caja")
            objExcel.ActiveSheet.Cells(v, h + 5) = "Alm:" & Data2.Recordset.Fields("bodega") & " AlmF:" & Data2.Recordset.Fields("bodegaf")
            v = v + 1
mytablex.Seek "=", "" & Data2.Recordset.Fields("local"), "" & Data2.Recordset.Fields("tipo"), "" & Data2.Recordset.Fields("serie"), "" & Data2.Recordset.Fields("numero")
If Not mytablex.NoMatch Then
     Do
     If mytablex.EOF Then Exit Do
     
     If "" & mytablex.Fields("local") = "" & Data2.Recordset.Fields("local") And "" & mytablex.Fields("tipo") = "" & Data2.Recordset.Fields("tipo") And "" & mytablex.Fields("serie") = "" & Data2.Recordset.Fields("serie") And "" & mytablex.Fields("numero") = "" & Data2.Recordset.Fields("numero") Then
            sdx = sdx + Val("" & mytablex.Fields("cantidad"))
            sdx1 = sdx1 + Val("" & mytablex.Fields("total"))
            objExcel.ActiveSheet.Cells(v, h) = "" & mytablex.Fields("producto")
            objExcel.ActiveSheet.Cells(v, h + 1) = "" & mytablex.Fields("descripcio")
            objExcel.ActiveSheet.Cells(v, h + 2) = "" & mytablex.Fields("unidad")
            objExcel.ActiveSheet.Cells(v, h + 3) = "" & mytablex.Fields("factor")
            objExcel.ActiveSheet.Cells(v, h + 4) = "" & mytablex.Fields("cantidad")
            objExcel.ActiveSheet.Cells(v, h + 5) = "" & mytablex.Fields("precio")
            objExcel.ActiveSheet.Cells(v, h + 6) = "" & mytablex.Fields("total")
            v = v + 1
            Else: Exit Do
     End If
     mytablex.MoveNext
     Loop
 End If
            objExcel.ActiveSheet.Cells(v, h) = ""
            objExcel.ActiveSheet.Cells(v, h + 1) = ""
            objExcel.ActiveSheet.Cells(v, h + 2) = ""
            objExcel.ActiveSheet.Cells(v, h + 3) = ""
            objExcel.ActiveSheet.Cells(v, h + 4) = "" & sdx
            objExcel.ActiveSheet.Cells(v, h + 5) = ""
            objExcel.ActiveSheet.Cells(v, h + 6) = "" & sdx1
            v = v + 1
 Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
 mytablex.Close
 Exit Sub
cmd561212_err:
MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
Exit Sub

End Sub


Private Sub xdki82_Click()
Dim mytablex As Table
On Error GoTo cmd76324_err
Frame6.Caption = "NuevoDocumento"
sotipo = "" & Data2.Recordset.Fields("tipo")
soserie = "" & Data2.Recordset.Fields("serie")
sonumero = "" & Data2.Recordset.Fields("numero")
detipo.Clear
Set mytablex = mydbxglo.OpenTable("tipo")
Do
If mytablex.EOF Then Exit Do
If "" & mytablex.Fields("tipodoc") = "A" Or "" & mytablex.Fields("tipodoc") = "B" Or "" & mytablex.Fields("tipodoc") = "C" Or "" & mytablex.Fields("tipodoc") = "D" Or "" & mytablex.Fields("tipodoc") = "G" Then
   detipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")
End If
mytablex.MoveNext
Loop
detipo.ListIndex = 0
mytablex.Close
Frame6.Visible = True
deserie = ""
denumero = ""
decodigo = ""
denumero = ""
detipo.SetFocus
Exit Sub
cmd76324_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub

End Sub

'AQUI VAMOS A GENERAR LE documento automaticamente
Sub generar_documentos(buf0 As String, buf1 As String, buf2 As String, buf3 As String)
Dim mytablex As Table
Dim mytabley As Table
Dim i As Integer
If Len(buf0) = 0 Then Exit Sub
If Len(buf1) = 0 Then Exit Sub
If Len(buf2) = 0 Then Exit Sub
If Len(buf3) = 0 Then Exit Sub


Set mytabley = mydbxglo.OpenTable(cgusuario)
mytabley.Index = "tfacimp"

Set mytablex = mydbxglo.OpenTable(cgusuario)
mytablex.Index = "tfacimp"
mytablex.Seek "=", buf0, buf1, buf2, buf3
If Not mytablex.NoMatch Then
   mytabley.Seek "=", buf0, buf1, buf2, buf3
   If Not mytabley.NoMatch Then
      mytabley.AddNew
      For i = 0 To mytablex.Fields.count - 1
          mytabley.Fields(i) = mytablex.Fields(i)
      Next i
      mytabley.Fields("local") = ""
      mytabley.Fields("tipo") = ""
      mytabley.Fields("serie") = ""
      mytabley.Fields("numero") = ""
      mytabley.Fields("acu") = ""
      mytabley.Fields("estado") = ""
      mytabley.Update
   End If
End If
mytablex.Close
mytabley.Close

'detalle
Set mytabley = mydbxglo.OpenTable(cgusuario)
mytabley.Index = "tfacimp"

Set mytablex = mydbxglo.OpenTable(cgusuario)
mytablex.Index = "tfacimp"

mytablex.Seek "=", buf0, buf1, buf2, buf3
If Not mytablex.NoMatch Then
   Do
   If mytablex.EOF Then Exit Do
   If "" & mytablex.Fields("local") = buf0 And "" & mytablex.Fields("tipo") = buf1 And "" & mytablex.Fields("serie") = buf2 And "" & mytablex.Fields("numero") = buf3 Then
      mytabley.AddNew
      For i = 0 To mytablex.Fields.count - 1
          mytabley.Fields(i) = mytablex.Fields(i)
      Next i
      mytabley.Fields("local") = ""
      mytabley.Fields("tipo") = ""
      mytabley.Fields("serie") = ""
      mytabley.Fields("numero") = ""
      mytabley.Fields("acu") = ""
      mytabley.Fields("estado") = ""
      mytabley.Update
      Else: Exit Do
   End If
   mytablex.MoveNext
   Loop
End If
mytablex.Close
mytabley.Close
End Sub
Function valida_clave(buf As String)
Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable("vendedor")
mytablex.Index = "clave"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   valida_clave = 1
End If
mytablex.Close
End Function


