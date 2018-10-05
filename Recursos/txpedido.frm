VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form txpedido 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Explorador de Documentos"
   ClientHeight    =   8955
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   13995
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Generar Traslado Automatico"
      ForeColor       =   &H00FF0000&
      Height          =   5175
      Left            =   2640
      TabIndex        =   60
      Top             =   1080
      Visible         =   0   'False
      Width           =   10455
      Begin VB.CommandButton Command7 
         Caption         =   "Salir"
         Height          =   495
         Left            =   8520
         TabIndex        =   78
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Grabar"
         Height          =   495
         Left            =   8520
         TabIndex        =   77
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox uvendedor 
         BackColor       =   &H0000FFFF&
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   2400
         Width           =   2655
      End
      Begin VB.ComboBox utipo 
         BackColor       =   &H0000FFFF&
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   75
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox ufecha 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   74
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox unumero 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Left            =   2280
         MaxLength       =   11
         TabIndex        =   71
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox userie 
         Height          =   375
         Left            =   2280
         MaxLength       =   3
         TabIndex        =   69
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox ubodegaf 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   66
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox ubodega 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   64
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox ulocal 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   62
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label unombre1 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3840
         TabIndex        =   80
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label unombre2 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3840
         TabIndex        =   79
         Top             =   960
         Width           =   3975
      End
      Begin VB.Label Label24 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha"
         Height          =   375
         Left            =   240
         TabIndex        =   73
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label23 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Responsable"
         Height          =   375
         Left            =   240
         TabIndex        =   72
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label22 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero"
         Height          =   375
         Left            =   240
         TabIndex        =   70
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label21 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie"
         Height          =   375
         Left            =   240
         TabIndex        =   68
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label20 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
         Height          =   375
         Left            =   240
         TabIndex        =   67
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label19 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Alm.Final"
         Height          =   375
         Left            =   240
         TabIndex        =   65
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label18 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Alm.Inicio"
         Height          =   375
         Left            =   240
         TabIndex        =   63
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LocaL"
         Height          =   375
         Left            =   240
         TabIndex        =   61
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
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8040
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
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8040
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFF00&
      Caption         =   "Visualizar Detalle"
      Height          =   8775
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Visible         =   0   'False
      Width           =   13695
      Begin MSDBGrid.DBGrid DBGrid4 
         Bindings        =   "explorap.frx":0000
         Height          =   2415
         Left            =   120
         OleObjectBlob   =   "explorap.frx":0014
         TabIndex        =   52
         Top             =   6240
         Width           =   13335
      End
      Begin MSDBGrid.DBGrid DBGrid3 
         Bindings        =   "explorap.frx":1087
         Height          =   5895
         Left            =   120
         OleObjectBlob   =   "explorap.frx":109B
         TabIndex        =   51
         Top             =   240
         Width           =   13335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clave de Acceso"
      Height          =   5055
      Left            =   3720
      TabIndex        =   36
      Top             =   1440
      Visible         =   0   'False
      Width           =   5295
      Begin VB.TextBox clave 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   1200
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   40
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
         Picture         =   "explorap.frx":6AA2
         Style           =   1  'Graphical
         TabIndex        =   38
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
         Picture         =   "explorap.frx":7250
         Style           =   1  'Graphical
         TabIndex        =   37
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
         TabIndex        =   39
         Top             =   720
         Width           =   4215
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
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   13815
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
         Bindings        =   "explorap.frx":79FE
         Height          =   6735
         Left            =   120
         OleObjectBlob   =   "explorap.frx":7A12
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
      Left            =   2400
      TabIndex        =   15
      Top             =   1440
      Visible         =   0   'False
      Width           =   7695
      Begin VB.TextBox vendedor 
         Height          =   375
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   48
         Text            =   "*"
         Top             =   2880
         Width           =   1935
      End
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
         Picture         =   "explorap.frx":83DD
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
         Picture         =   "explorap.frx":8B8B
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
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vendedor"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   49
         Top             =   2880
         Width           =   2175
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
      Bindings        =   "explorap.frx":9339
      Height          =   6735
      Left            =   0
      OleObjectBlob   =   "explorap.frx":934D
      TabIndex        =   0
      Top             =   1200
      Width           =   13815
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00E0E0E0&
      Height          =   915
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   13935
      TabIndex        =   9
      Top             =   0
      Width           =   13995
      Begin VB.ComboBox tipo 
         Height          =   315
         Left            =   10080
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   480
         Width           =   1695
      End
      Begin VB.ComboBox caja 
         Height          =   315
         Left            =   10080
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   120
         Width           =   1695
      End
      Begin VB.ComboBox bodega 
         Height          =   315
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Consul&Tar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   12000
         MaskColor       =   &H00FFFFFF&
         Picture         =   "explorap.frx":C7FC
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.TextBox fechaf 
         Height          =   375
         Left            =   7440
         MaxLength       =   10
         TabIndex        =   46
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox fechai 
         Height          =   375
         Left            =   7440
         MaxLength       =   10
         TabIndex        =   44
         Top             =   120
         Width           =   1695
      End
      Begin VB.ComboBox local1 
         Height          =   315
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdSort 
         BackColor       =   &H00C0C0C0&
         Height          =   855
         Left            =   2160
         MaskColor       =   &H00E0E0E0&
         Picture         =   "explorap.frx":CFAA
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
         Picture         =   "explorap.frx":E1BC
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
         Picture         =   "explorap.frx":F3CE
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
         Picture         =   "explorap.frx":105E0
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
         Picture         =   "explorap.frx":117F2
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoDoc"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   9120
         TabIndex        =   59
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Caja"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   9120
         TabIndex        =   57
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Almacen"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3840
         TabIndex        =   55
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Final"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6480
         TabIndex        =   47
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6480
         TabIndex        =   45
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3840
         TabIndex        =   43
         Top             =   120
         Width           =   615
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
   Begin VB.Label tipoclie 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   360
      TabIndex        =   41
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
   End
   Begin VB.Menu mit56232 
      Caption         =   "&Zoom"
   End
   Begin VB.Menu djbu232 
      Caption         =   "&Consulta"
   End
   Begin VB.Menu Hyu6723ge 
      Caption         =   "&Generar"
      Begin VB.Menu dj7823233 
         Caption         =   "&1.Generar Traslado Automatico"
      End
   End
   Begin VB.Menu ldo33 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "txpedido"
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
On Error GoTo cmd8_err
If Frame4.Visible = True Then Exit Sub

If Frame2.Visible = True Then Exit Sub
If Frame3.Visible = True Then Exit Sub

If "" & Data2.Recordset.Fields("estado") <> "0" Then
   MsgBox "Para anular el documento debe estar en estado=0", 48, "Aviso"
   Exit Sub
End If
If MsgBox("Desea Anular Documento", 1, "Aviso") <> 1 Then Exit Sub

Data2.Recordset.Edit
Data2.Recordset.Fields("estado") = "1"
Data2.Recordset.Update
buf1 = " and acu='" & "" & Data2.Recordset.Fields("acu") & "'"
mydbxglo.Execute "update  " & dgusuariog & " set estado='1'  where  local='" & "" & Data2.Recordset.Fields("local") & "' and tipo='" & "" & Data2.Recordset.Fields("tipo") & "' and serie='" & "" & Data2.Recordset.Fields("serie") & "' and  numero='" & "" & Data2.Recordset.Fields("numero") & "'" & buf1
mydbxglo.Execute "update  fpagov  set estado='1'  where  local='" & "" & Data2.Recordset.Fields("local") & "' and tipo='" & "" & Data2.Recordset.Fields("tipo") & "' and serie='" & "" & Data2.Recordset.Fields("serie") & "' and  numero='" & "" & Data2.Recordset.Fields("numero") & "'" & buf1
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

Private Sub clave_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub cmdAddEntry_Click()
djku232_Click
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

Private Sub codigo_KeyUp(KeyCode As Integer, Shift As Integer)
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
buf = "select Nombre,Codigo from clientes where " & Combo1 & " like '" & buffer & "*'"
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
               DBGrid1.Columns(0).Width = 4000
               DBGrid1.Columns(1).Width = 2000
               End If
               If opcion1 = "2" Then
               DBGrid1.Columns(0).Width = 1500
               DBGrid1.Columns(1).Width = 5000
               DBGrid1.Columns(2).Width = 900
               DBGrid1.Columns(3).Width = 900
               DBGrid1.Columns(4).Width = 900
               DBGrid1.Columns(5).Width = 900
               DBGrid1.Columns(6).Width = 1500
               DBGrid1.Columns(7).Width = 900
               DBGrid1.Columns(8).Width = 700
               End If
               DBGrid1.SetFocus

End Sub

Private Sub Command2_Click()
ldo33_Click
End Sub

Private Sub Command3_Click()
ldo33_Click
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
If utipo = "*" Then
   utipo.SetFocus
   Exit Sub
End If
If uvendedor = "*" Then
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

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   buffer.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
   If opcion1 = "1" Then
   codigo = DBGrid1.Columns(1)
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
            If Len("" & DBGrid2.Columns(2)) = 0 Then
               Cancel = True
               Exit Sub
            End If
End Select

End Sub

Private Sub DBGrid2_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim found As Integer
Select Case ColIndex
Case 13
     If Len(DBGrid2.Columns(2)) = 0 Then
        Cancel = True
        Exit Sub
     End If
     If Len("" & DBGrid2.Columns(13)) = 0 Then
        DBGrid2.Columns(13) = Data2.Recordset.Fields("fecha")
        Exit Sub
     End If
     found = valida_fecha("" & DBGrid2.Columns(13))
     If found = 0 Then
        Cancel = True
        Exit Sub
     End If
End Select
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

Private Sub DBGrid3_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If ColIndex <> 5 Then
   Cancel = True
   Exit Sub
End If
Select Case ColIndex
       Case 5
            If Len("" & DBGrid3.Columns(0)) = 0 Then
               Cancel = True
               Exit Sub
            End If
           
End Select

End Sub

Private Sub DBGrid3_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Select Case ColIndex
       Case 5
            If Not IsNumeric("" & DBGrid3.Columns(5)) Then
               Cancel = True
               Exit Sub
            End If
            If Val("" & DBGrid3.Columns(3)) < (Val("" & DBGrid3.Columns(4)) + Val("" & DBGrid3.Columns(5))) Then
               MsgBox "Despacho excedido ", 48, "Aviso"
               Cancel = True
               Exit Sub
            End If
            
            '---------- validamos a donde va
            'valida_ingresado
End Select

End Sub

Private Sub dj7823233_Click()
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
   tfactura.local1 = "01"
   tfactura.codigo = "01"
   tfactura.Label2.Caption = "Cod.Int"
   'tfactura.Label14.Visible = True
   tfactura.Label38.Visible = True
   'tfactura.localf.Visible = True
   tfactura.bodegaf.Visible = True
End If
If acu = "V" Then
found = copiar_temporal()
If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   Exit Sub
End If
tfactura.Label2 = "CodClie"
tfactura.Caption = "Facturacion x Ventas"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfactura.bandera = "Nuevo"
tfactura.acu = "V"
tfactura.tipoclie = tipoclie
'tfactura.local1=local
tfactura.Show 1
End If
If acu = "H" Then
found = copiar_temporal()
If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   Exit Sub
End If
tfactura.Label2 = "CodClie"
tfactura.Caption = "Cotizaciones Ventas"
cgusuario = "CCOTIZAV"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "dcotizav"
tfactura.bandera = "Nuevo"
tfactura.acu = "H"
tfactura.tipoclie = tipoclie
tfactura.Show 1
End If
If acu = "I" Then
found = copiar_temporal()
If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   Exit Sub
End If
tfactura.Label2 = "CodClie"
tfactura.Caption = "Cotizaciones Ventas"
cgusuario = "Cpedidov"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "dpedidov"
tfactura.bandera = "Nuevo"
tfactura.acu = "I"
tfactura.tipoclie = tipoclie
tfactura.Show 1

End If


If acu = "T" Then
found = copiar_temporal()
If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   Exit Sub
End If
tfactura.Label2 = "CodClie"
tfactura.Caption = "Guia Salida"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfactura.bandera = "Nuevo"
tfactura.acu = "T"
tfactura.tipoclie = tipoclie
tfactura.Show 1
End If
If acu = "E" Then
found = copiar_temporal()
If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   Exit Sub
End If
tfactura.Label2 = "CodClie"
tfactura.Caption = "Nota Credito Ventas"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfactura.bandera = "Nuevo"
tfactura.acu = "E"
tfactura.tipoclie = tipoclie
tfactura.Show 1
End If
If acu = "F" Then
found = copiar_temporal()
If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   Exit Sub
End If
tfactura.Label2 = "CodClie"
tfactura.Caption = "Nota debito Ventas"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
tfactura.bandera = "Nuevo"
dgusuariog = "DETALLE"
tfactura.acu = "F"
tfactura.tipoclie = tipoclie
tfactura.Show 1
End If
If acu = "R" Then
found = copiar_temporal()
If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   
   Exit Sub
End If
tfactura.Label2 = "CodProv"
tfactura.Caption = "Orden de Compra"
cgusuario = "CORDENC"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DORDENC"
tfactura.acu = "R"
tfactura.bandera = "Nuevo"
tfactura.tipoclie = tipoclie
tfactura.Show 1
End If
If acu = "S" Then
found = copiar_temporal()
If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   
   Exit Sub
End If
tfactura.Label2 = "CodProv"
tfactura.Caption = "Guia Remision Entrada"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfactura.tipoclie = tipoclie
tfactura.acu = "S"
tfactura.bandera = "Nuevo"
tfactura.Show 1
End If
If acu = "C" Then
found = copiar_temporal()
If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   
   Exit Sub
End If
tfactura.Label2 = "CodProv"
tfactura.Caption = "Factura de Compra"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfactura.acu = "C"
tfactura.bandera = "Nuevo"
tfactura.tipoclie = tipoclie
tfactura.Show 1
End If

If acu = "N" Then
found = copiar_temporal()
If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   Exit Sub
End If
tfactura.Label2 = "CodProv"
tfactura.Caption = "Nota Credito Compra"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfactura.acu = "N"
tfactura.bandera = "Nuevo"
tfactura.tipoclie = tipoclie
tfactura.Show 1
End If

If acu = "O" Then
found = copiar_temporal()
If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   Exit Sub
End If
tfactura.Label2 = "CodProv"
tfactura.Caption = "Nota debito de Compra"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfactura.acu = "O"
tfactura.bandera = "Nuevo"
tfactura.tipoclie = tipoclie
tfactura.Show 1
End If
If acu = "Q" Then
found = copiar_temporal()
If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   Exit Sub
End If
tfactura.Label2 = "CodProv"
tfactura.Caption = "Pedido Almacen"
cgusuario = "CREQUISA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DREQUISA"
tfactura.acu = "Q"
tfactura.bandera = "Nuevo"
tfactura.tipoclie = tipoclie
tfactura.Show 1

End If
If acu = "Z" Then
found = copiar_temporal()
If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   End
   Exit Sub
End If
'tfactura.Label2 = "Cod.Inicio"
tfactura.Caption = "Traslado entre almacen de un mismo establecimiento"
cgusuario = "CTRASLAD"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DTRASLAD"
tfactura.acu = "Z"
tfactura.bandera = "Nuevo"
tfactura.tipoclie = tipoclie
tfactura.Show 1
End If
sql_cabeza
Exit Sub
cmd28_err:
Exit Sub
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

Private Sub Form_Activate()

If acu = "Q" Then
   DBGrid3.AllowUpdate = True
   Hyu6723ge.Visible = True
   Hyu6723ge.Enabled = True

End If
carga_iniciales
cmdGrabar_Click
'sql_cabeza
End Sub

Private Sub Form_Load()
moneda.Clear
moneda.AddItem "*"
moneda.AddItem "S"
moneda.AddItem "D"
moneda.ListIndex = 0
estado.Clear
estado.AddItem "*"
estado.AddItem "2"
estado.AddItem "1"
estado.AddItem "0"
estado.ListIndex = 0
End Sub
Sub carga_iniciales()
Dim mytablex As Table
caja.Clear
caja.AddItem "*"
tipo.Clear
tipo.AddItem "*"
bodega.Clear
bodega.AddItem "*"
local1.Clear
local1.AddItem "*"
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
mytablex.Close


Set mytablex = mydbxglo.OpenTable("parameca")
Do
If mytablex.EOF Then Exit Do
caja.AddItem "" & mytablex.Fields("caja") & "|" & mytablex.Fields("descripcio")
mytablex.MoveNext
Loop
caja.ListIndex = 0
mytablex.Close


End Sub

Private Sub ldo33_Click()
If Frame5.Visible = True Then
   Frame5.Visible = False
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

explorap.Hide
Unload explorap
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
If tipo <> "*" Then
   buf = buf & " and tipo like '" & extra_loquesea(tipo) & "'"
End If
If caja <> "*" Then
   buf = buf & " and caja like '" & extra_loquesea(caja) & "'"
End If

If serie <> "*" Then
   buf = buf & " and serie like '" & serie & "'"
End If
If numero <> "*" Then
   buf = buf & " and numero like '" & numero & "'"
End If
If codigo <> "*" Then
   buf = buf & " and codigo like '" & codigo & "'"
End If
If nombre <> "*" Then
   buf = buf & " and nombre like '" & nombre & "'"
End If
If moneda <> "*" Then
   buf = buf & " and moneda like '" & moneda & "'"
End If
If estado <> "*" Then
   buf = buf & " and estado like '" & estado & "'"
End If
If local1 <> "*" Then
   buf = buf & " and local like '" & extra_loquesea(local1) & "'"
End If
If vendedor <> "*" Then
buf = buf & " and vendedor like '" & vendedor & "'"
End If
If bodega <> "*" Then
   buf = buf & " and bodega like '" & extra_loquesea(bodega) & "'"
End If
If acu <> "C" And acu <> "V" Then
   buf = buf & " and acu='" & acu & "'"
End If
If acu = "V" Then
   buf = buf & " and (acu='A' OR acu='B' OR acu='C' OR acu='D' OR acu='G' OR acu='E' OR acu='F')"
End If
If acu = "C" Then
   buf = buf & " and (acu='J' OR acu='K' OR acu='L' OR acu='M' OR acu='P' OR acu='N' OR acu='O')"
End If
buf = buf & " order by fecha,tipo,serie,val(numero)"
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
MsgBox "Error en select " & error$, 48, "Aviso"
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
    factura_formato local1, "" & ttipo, "" & tserie, "" & tnumero, ""
    cerrar_archivo
    genver.File = globaldir & "\temporal\" & gusuario & ".txt"
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
mytablex.Index = "tfactura"
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
   'tfactura.Label14.Visible = True
   tfactura.Label38.Visible = True
   'tfactura.localf.Visible = True
   tfactura.bodegaf.Visible = True
   tfactura.Label2.Caption = "Cod.Int."
End If

If acu = "V" Then
tfactura.Label2 = "CodClie"
tfactura.Caption = "Facturacion x Ventas"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfactura.bandera = "Modifica"
tfactura.cmdAddEntry.Enabled = False
tfactura.dnu834.Enabled = False
tfactura.acu = "V"
tfactura.tipoclie = tipoclie
tfactura.zlocal = "" & Data2.Recordset.Fields("local")
tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
tfactura.zserie = "" & Data2.Recordset.Fields("serie")
tfactura.znumero = "" & Data2.Recordset.Fields("numero")
tfactura.Show 1
sql_cabeza

End If
If acu = "H" Then
tfactura.Label2 = "CodClie"
tfactura.Caption = "Cotizacion x Ventas"
cgusuario = "ccotizav"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "dcotizav"
tfactura.bandera = "Modifica"
tfactura.cmdAddEntry.Enabled = False
tfactura.dnu834.Enabled = False
tfactura.acu = "H"
tfactura.tipoclie = tipoclie
tfactura.zlocal = "" & Data2.Recordset.Fields("local")
tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
tfactura.zserie = "" & Data2.Recordset.Fields("serie")
tfactura.znumero = "" & Data2.Recordset.Fields("numero")
tfactura.Show 1
End If
If acu = "I" Then
tfactura.Label2 = "CodClie"
tfactura.Caption = "Pedidos x Ventas"
cgusuario = "cpedidov"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "dpedidov"
tfactura.bandera = "Modifica"
tfactura.cmdAddEntry.Enabled = False
tfactura.dnu834.Enabled = False
tfactura.acu = "I"
tfactura.tipoclie = tipoclie
tfactura.zlocal = "" & Data2.Recordset.Fields("local")
tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
tfactura.zserie = "" & Data2.Recordset.Fields("serie")
tfactura.znumero = "" & Data2.Recordset.Fields("numero")
tfactura.Show 1
End If

If acu = "T" Then
tfactura.Label2 = "CodClie"
tfactura.Caption = "Guia Remision x Ventas"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfactura.bandera = "Modifica"
tfactura.cmdAddEntry.Enabled = False
tfactura.dnu834.Enabled = False
tfactura.acu = "T"
tfactura.tipoclie = tipoclie
tfactura.zlocal = "" & Data2.Recordset.Fields("local")
tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
tfactura.zserie = "" & Data2.Recordset.Fields("serie")
tfactura.znumero = "" & Data2.Recordset.Fields("numero")
tfactura.Show 1
End If

If acu = "E" Then
tfactura.Label2 = "CodClie"
tfactura.Caption = "Nota Credito x Ventas"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfactura.bandera = "Modifica"
tfactura.cmdAddEntry.Enabled = False
tfactura.dnu834.Enabled = False
tfactura.acu = "E"
tfactura.tipoclie = tipoclie
tfactura.zlocal = "" & Data2.Recordset.Fields("local")
tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
tfactura.zserie = "" & Data2.Recordset.Fields("serie")
tfactura.znumero = "" & Data2.Recordset.Fields("numero")
tfactura.Show 1
End If
If acu = "R" Then
tfactura.Label2 = "CodProv"
tfactura.Caption = "Orden Compra"
cgusuario = "CORDENC"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DORDENC"
tfactura.bandera = "Modifica"
tfactura.cmdAddEntry.Enabled = False
tfactura.dnu834.Enabled = False
tfactura.acu = "R"
tfactura.tipoclie = tipoclie
tfactura.zlocal = "" & Data2.Recordset.Fields("local")
tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
tfactura.zserie = "" & Data2.Recordset.Fields("serie")
tfactura.znumero = "" & Data2.Recordset.Fields("numero")
tfactura.Show 1
End If
If acu = "F" Then
tfactura.Label2 = "CodClie"
tfactura.Caption = "Nota Debito x Ventas"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"

tfactura.acu = "F"
tfactura.tipoclie = tipoclie
tfactura.cmdAddEntry.Enabled = False
tfactura.dnu834.Enabled = False
tfactura.bandera = "Modifica"
tfactura.zlocal = "" & Data2.Recordset.Fields("local")
tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
tfactura.zserie = "" & Data2.Recordset.Fields("serie")
tfactura.znumero = "" & Data2.Recordset.Fields("numero")
tfactura.Show 1
End If

If acu = "S" Then
tfactura.Label2 = "CodProv"
tfactura.Caption = "Guia Remision Entrada"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfactura.tipoclie = tipoclie
tfactura.acu = "S"
tfactura.cmdAddEntry.Enabled = False
tfactura.dnu834.Enabled = False
tfactura.bandera = "Modifica"
tfactura.zlocal = "" & Data2.Recordset.Fields("local")
tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
tfactura.zserie = "" & Data2.Recordset.Fields("serie")
tfactura.znumero = "" & Data2.Recordset.Fields("numero")
tfactura.Show 1
End If
If acu = "C" Then
tfactura.Label2 = "CodProv"
tfactura.Caption = "Factura de Compra"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfactura.acu = "C"
tfactura.tipoclie = tipoclie
tfactura.cmdAddEntry.Enabled = False
tfactura.dnu834.Enabled = False
tfactura.bandera = "Modifica"
tfactura.zlocal = "" & Data2.Recordset.Fields("local")
tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
tfactura.zserie = "" & Data2.Recordset.Fields("serie")
tfactura.znumero = "" & Data2.Recordset.Fields("numero")
tfactura.Show 1
End If

If acu = "N" Then
tfactura.Label2 = "CodProv"
tfactura.Caption = "Nota Credito Compra"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfactura.acu = "N"
tfactura.tipoclie = tipoclie
tfactura.cmdAddEntry.Enabled = False
tfactura.dnu834.Enabled = False
tfactura.bandera = "Modifica"
tfactura.zlocal = "" & Data2.Recordset.Fields("local")
tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
tfactura.zserie = "" & Data2.Recordset.Fields("serie")
tfactura.znumero = "" & Data2.Recordset.Fields("numero")
tfactura.Show 1
End If

If acu = "O" Then
tfactura.Label2 = "CodProv"
tfactura.Caption = "Nota debito de Compra"
cgusuario = "FACTURA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DETALLE"
tfactura.acu = "O"
tfactura.tipoclie = tipoclie
tfactura.cmdAddEntry.Enabled = False
tfactura.dnu834.Enabled = False
tfactura.bandera = "Modifica"
tfactura.zlocal = "" & Data2.Recordset.Fields("local")
tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
tfactura.zserie = "" & Data2.Recordset.Fields("serie")
tfactura.znumero = "" & Data2.Recordset.Fields("numero")
tfactura.Show 1
End If
If acu = "Q" Then
tfactura.Label2 = "CodProv"
tfactura.Caption = "Pedido Almacen"
cgusuario = "CREQUISA"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DREQUISA"
tfactura.acu = "Q"
tfactura.tipoclie = tipoclie
tfactura.cmdAddEntry.Enabled = False
tfactura.dnu834.Enabled = False
tfactura.bandera = "Modifica"
tfactura.zlocal = "" & Data2.Recordset.Fields("local")
tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
tfactura.zserie = "" & Data2.Recordset.Fields("serie")
tfactura.znumero = "" & Data2.Recordset.Fields("numero")
tfactura.Show 1
End If
If acu = "Z" Then
tfactura.Caption = "Traslado entre almacen de un mismo establecimiento"
cgusuario = "CTRASLAD"
dgusuario = "_d" & gusuario
fgusuario = "_f" & gusuario
dgusuariog = "DTRASLAD"
tfactura.acu = "Z"
tfactura.tipoclie = tipoclie
tfactura.cmdAddEntry.Enabled = False
tfactura.dnu834.Enabled = False
tfactura.bandera = "Modifica"
tfactura.zlocal = "" & Data2.Recordset.Fields("local")
tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
tfactura.zserie = "" & Data2.Recordset.Fields("serie")
tfactura.znumero = "" & Data2.Recordset.Fields("numero")
tfactura.Show 1
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
Frame4.Visible = True
sql_detalles
DBGrid3.SetFocus

End Sub
Sub sql_detalles()
Dim buf As String
On Error GoTo cmd321_err
buf = "select * from " & dgusuariog & " where "
buf = buf & " local='" & "" & Data2.Recordset.Fields("local") & "'"
buf = buf & " and tipo='" & "" & Data2.Recordset.Fields("tipo") & "'"
buf = buf & " and serie='" & "" & Data2.Recordset.Fields("serie") & "'"
buf = buf & " and numero='" & "" & Data2.Recordset.Fields("numero") & "'"
Data4.Connect = "foxpro 2.5;"
Data4.DatabaseName = globaldir
Data4.RecordSource = buf
Data4.Refresh

buf = "select * from fpagov where "
buf = buf & " local='" & "" & Data2.Recordset.Fields("local") & "'"
buf = buf & " and tipo='" & "" & Data2.Recordset.Fields("tipo") & "'"
buf = buf & " and serie='" & "" & Data2.Recordset.Fields("serie") & "'"
buf = buf & " and numero='" & "" & Data2.Recordset.Fields("numero") & "'"
Data5.Connect = "foxpro 2.5;"
Data5.DatabaseName = globaldir
Data5.RecordSource = buf
Data5.Refresh
Exit Sub
cmd321_err:
Exit Sub
End Sub

Private Sub modi343_Click()
On Error GoTo cmd7_err
If Frame4.Visible = True Then Exit Sub

If Frame2.Visible = True Then Exit Sub
If Frame3.Visible = True Then Exit Sub
If "" & Data2.Recordset.Fields("estado") <> "2" Then Exit Sub
If MsgBox("Desea Desmarca el Documento", 1, "Aviso") <> 1 Then Exit Sub
desmarca_documento
Exit Sub
cmd7_err:
Exit Sub
End Sub
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
mytablex.Index = "TFACTURA"
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
mytablex.Fields("tipoclie") = "I"
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
    For i = 0 To mytabley.Fields.Count - 1
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
    mytablex.Fields("tipoclie") = "I"
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
utipo.AddItem "*"
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
uvendedor.AddItem "*"
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
    For i = 0 To mytabley.Fields.Count - 1
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
    mytablex.Fields("tipoclie") = "I"
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
    For i = 0 To mytabley.Fields.Count - 1
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
    mytablex.Fields("tipoclie") = "I"
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
