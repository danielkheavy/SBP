VERSION 5.00
Object = "{D4D26F6B-6564-44F4-A913-03C91CE37740}#2.1#0"; "reportman.ocx"
Begin VB.Form repinv 
   BackColor       =   &H00808080&
   Caption         =   "Reportes"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   10440
   Begin reportman.ReportManX Rep 
      Height          =   300
      Left            =   10560
      TabIndex        =   69
      Top             =   7800
      Width           =   1125
      filename        =   ""
      Preview         =   0   'False
      ShowProgress    =   0   'False
      ShowPrintDialog =   0   'False
      Title           =   ""
      Language        =   0
      DoubleBuffered  =   0   'False
      Enabled         =   -1  'True
      Object.Visible         =   -1  'True
      Cursor          =   0
      HelpType        =   0
      HelpKeyword     =   ""
      DefaultPrinter  =   "CAJA"
      AsyncExecution  =   0   'False
   End
   Begin VB.CommandButton ppp 
      Caption         =   "Precio Promedio Ponderado"
      Height          =   720
      Left            =   8400
      TabIndex        =   80
      Top             =   4200
      Width           =   1590
   End
   Begin VB.TextBox FechaInicial 
      BackColor       =   &H0080FFFF&
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
      Left            =   5520
      MaxLength       =   10
      TabIndex        =   78
      Top             =   1320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CheckBox ChkSaldoInicial 
      BackColor       =   &H0080FFFF&
      Caption         =   "Saldo Inicial"
      Height          =   375
      Left            =   6000
      TabIndex        =   77
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdReporteDe 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reporte de productos con Stock Minimo"
      Height          =   930
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   225
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.CommandButton cmdREPSalMinAgrupadoPorFam 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reporte de productos con Stock Minimo Agrupado por Familia"
      Height          =   930
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   1200
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.CommandButton cmdREPSalMinAgrupado 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reporte de productos con Stock Minimo Agrupado por Familia"
      Height          =   930
      Left            =   7980
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   7440
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click para Cancelar..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   1320
      TabIndex        =   64
      Top             =   2520
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.ComboBox quecosto 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   62
      Top             =   2400
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox fechari 
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
      Left            =   5520
      MaxLength       =   10
      TabIndex        =   58
      Text            =   "%"
      Top             =   5640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox fecharf 
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
      Left            =   5520
      MaxLength       =   10
      TabIndex        =   57
      Text            =   "%"
      Top             =   6000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox fechavpf 
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
      Left            =   5520
      MaxLength       =   10
      TabIndex        =   53
      Text            =   "%"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox fechavpi 
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
      Left            =   5520
      MaxLength       =   10
      TabIndex        =   52
      Text            =   "%"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox fechavi 
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
      Left            =   5520
      MaxLength       =   10
      TabIndex        =   48
      Text            =   "%"
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox fechavf 
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
      Left            =   5520
      MaxLength       =   10
      TabIndex        =   47
      Text            =   "%"
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox proveedor 
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
      MaxLength       =   11
      TabIndex        =   45
      Text            =   "%"
      Top             =   5280
      Width           =   1935
   End
   Begin VB.TextBox gcanti 
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
      Left            =   5520
      MaxLength       =   1
      TabIndex        =   43
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox excell 
      BackColor       =   &H0080FFFF&
      Caption         =   "Excell"
      Height          =   375
      Left            =   3840
      TabIndex        =   42
      Top             =   7080
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.ComboBox conteo 
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
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   40
      Top             =   7800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox local1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   5640
      Width           =   1935
   End
   Begin VB.ComboBox monedac 
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
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   6360
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "repinv.frx":0000
      Left            =   5520
      List            =   "repinv.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   6720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox titulo 
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
      MaxLength       =   30
      TabIndex        =   33
      Top             =   7440
      Width           =   4935
   End
   Begin VB.TextBox nrolineas 
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
      TabIndex        =   31
      Text            =   "44"
      Top             =   7080
      Width           =   1095
   End
   Begin VB.ComboBox bodega 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   6000
      Width           =   1935
   End
   Begin VB.ComboBox igv 
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
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   4440
      Width           =   1935
   End
   Begin VB.TextBox fechaf 
      BackColor       =   &H0080FFFF&
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
      Left            =   5520
      MaxLength       =   10
      TabIndex        =   26
      Text            =   "%"
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox fechai 
      BackColor       =   &H0080FFFF&
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
      Left            =   5520
      MaxLength       =   10
      TabIndex        =   24
      Text            =   "%"
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
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
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   4800
      Width           =   1935
   End
   Begin VB.ComboBox marca 
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
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   3720
      Width           =   1935
   End
   Begin VB.ComboBox color 
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
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3240
      Width           =   1935
   End
   Begin VB.ComboBox linea 
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
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2880
      Width           =   1935
   End
   Begin VB.ComboBox categoria 
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
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2400
      Width           =   1935
   End
   Begin VB.ComboBox seccion 
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
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2040
      Width           =   1935
   End
   Begin VB.ComboBox subfamilia 
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
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1680
      Width           =   1935
   End
   Begin VB.ComboBox familia 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox descripcio 
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
      MaxLength       =   60
      TabIndex        =   5
      Text            =   "%"
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox barras 
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
      MaxLength       =   15
      TabIndex        =   3
      Text            =   "%"
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox producto 
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
      MaxLength       =   15
      TabIndex        =   1
      Text            =   "%"
      Top             =   120
      Width           =   1935
   End
   Begin VB.ComboBox conigv 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   66
      Top             =   1680
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "repinv.frx":0004
      Left            =   5520
      List            =   "repinv.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   74
      Top             =   6360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ComboBox vesubfamilia 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6680
      Style           =   2  'Dropdown List
      TabIndex        =   75
      Top             =   7080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label37 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicial Sistema"
      Height          =   375
      Left            =   3840
      TabIndex        =   79
      Top             =   1320
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label36 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VerSubfamilia"
      Height          =   345
      Left            =   5520
      TabIndex        =   76
      Top             =   7080
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Opcion "
      Height          =   375
      Left            =   3840
      TabIndex        =   73
      Top             =   6345
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label xbasedatos 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7560
      TabIndex        =   68
      Top             =   120
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label34 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Con Igv"
      Height          =   375
      Left            =   3840
      TabIndex        =   67
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label VENTANEGRA 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3840
      TabIndex        =   65
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label33 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Costo/PrecioVta"
      Height          =   375
      Left            =   3840
      TabIndex        =   63
      Top             =   2400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label32 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Días Sin Rotación"
      Height          =   375
      Left            =   3840
      TabIndex        =   61
      Top             =   5280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label31 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicio"
      Height          =   375
      Left            =   3840
      TabIndex        =   60
      Top             =   5640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label30 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Final"
      Height          =   375
      Left            =   3840
      TabIndex        =   59
      Top             =   6000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label29 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Final"
      Height          =   375
      Left            =   3840
      TabIndex        =   56
      Top             =   4680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label28 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicio"
      Height          =   375
      Left            =   3840
      TabIndex        =   55
      Top             =   4320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label27 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Variacion Precios Vta"
      Height          =   375
      Left            =   3840
      TabIndex        =   54
      Top             =   3960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label26 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vencimientos"
      Height          =   375
      Left            =   3840
      TabIndex        =   51
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label24 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicio"
      Height          =   375
      Left            =   3840
      TabIndex        =   50
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Final"
      Height          =   375
      Left            =   3840
      TabIndex        =   49
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Proveedor"
      Height          =   375
      Left            =   120
      TabIndex        =   46
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label Label25 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SoloCantidad"
      Height          =   375
      Left            =   3840
      TabIndex        =   44
      Top             =   2040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label23 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Comparativo Conteos"
      Height          =   375
      Left            =   120
      TabIndex        =   41
      Top             =   7800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label22 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Local"
      Height          =   375
      Left            =   120
      TabIndex        =   39
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label Label21 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MonedaConversion"
      Height          =   375
      Left            =   120
      TabIndex        =   37
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Condicion Saldo"
      Height          =   375
      Left            =   3840
      TabIndex        =   35
      Top             =   6720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label20 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Titulo Reporte"
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nro.Lineas Reporte"
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label Label18 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Almacen"
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Final"
      Height          =   375
      Left            =   3840
      TabIndex        =   25
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicio"
      Height          =   375
      Left            =   3840
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Moneda"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Impuesto"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Marca"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Color"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Linea Tallas"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Categoria"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Seccion"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SubFamilia"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Familia"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripcio"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo Barras"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Menu dlkoiewr 
      Caption         =   "&Procesa"
   End
   Begin VB.Menu ldso8912 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "repinv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim cbuf(15) As String

Dim vbuf(15) As Double

Dim dfile    As String

Public Function EstadoDeArchivo(ByVal archivo As String) As Boolean

    Dim fso

    Set fso = CreateObject("Scripting.FileSystemObject")

    If (fso.FileExists(archivo)) Then
        EstadoDeArchivo = True
    Else
        EstadoDeArchivo = False

    End If

End Function

Private Sub bodega_Click()
    'fechai = Format(busca_paramed(extra_loquesea(bodega)), "dd/mm/yyyy")

    ''05/06/2017 kenyo no borra fecha al sleecccionar almacen
    'If bodega <> "%" Then
    '   fechai = Format(busca_paramed(extra_loquesea(bodega)), "dd/mm/yyyy")
    'End If

End Sub

Private Sub bodega_DblClick()
    'fechai = Format(busca_paramed(extra_loquesea(bodega)), "dd/mm/yyyy")

End Sub

Private Sub bodega_KeyDown(KeyCode As Integer, Shift As Integer)
    'fechai = Format(busca_paramed(extra_loquesea(bodega)), "dd/mm/yyyy")

End Sub

Private Sub bodega_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then

        'fechai = Format(busca_paramed(extra_loquesea(bodega)), "dd/mm/yyyy")
    End If

End Sub

Private Sub bodega_KeyUp(KeyCode As Integer, Shift As Integer)
    'fechai = Format(busca_paramed(extra_loquesea(bodega)), "dd/mm/yyyy")

End Sub

Private Sub cmdCommand2_Click()
    envio_correosReportes

End Sub

Private Sub ChkSaldoInicial_Click()

    ''' 16/01/2018 Stock Inicial en Kardex Producto y Sunat
    If ChkSaldoInicial.Value = 1 Then
        fechainicial.Visible = True
        Label37.Visible = True
    Else
        fechainicial.Visible = False
        Label37.Visible = False

    End If

    ''' 16/01/2018 Stock Inicial en Kardex Producto y Sunat
End Sub

Private Sub cmdREPSalMinAgrupado_Click()
    Rep.Preview = True
    'Establecemos la cadena de conexion de manera dinamica.
    'Rep.SetDatabaseConnectionString "SERVIDOR", "Persist Security Info=True;Driver=MySQL ODBC 5.1 Driver;SERVER=" & cServidor & ";UID=" & cUserName & ";PWD=" & cPassword & ";DATABASE=" & cBDatos & ";PORT=3306"

    'Rep.FileName = ruta & "\informes\Rpt_ProductosStockMinimo.rep"
    Rep.FileName = App.path & "\Reportes\Rpt_ProductosStockMinimoAgrupado.rep"
    'cn.Open "Driver={SQL Server};Server=" & menup.vservidor & ";Database=" & buf & " ;Uid=sa;pwd=" & clave_servidor & ""
    'Rep.SetDatabaseConnectionString ("RESTAURANT6", "Driver={SQL Server};Server=TESTING\SQL2008;Uid=sa;pwd=mastercard;")
    'ok
    'Rep.SetDatabaseConnectionString "VISIORION", "Provider=SQLOLEDB.1;Integrated Security=SSPI;Server=TESTING\SQL2008;Database=RESTAURANT6;Uid=sa;pwd=mastercard;Persist Security Info=False"
    Rep.SetDatabaseConnectionString "VISIORION", "Provider=SQLOLEDB.1;Integrated Security=SSPI;Server=" & menup.vservidor & ";Database=" & basedatos & ";Uid=sa;pwd=" & clave_servidor & ";Persist Security Info=False"

    'Provider=SQLOLEDB.1;Integrated Security=SSPI;Server=TESTING\SQL2008;Database=RESTAURANT6;Uid=sa;pwd=mastercard;Persist Security Info=False
    'Rep.SetDatabaseConnectionString "RESTAURANT6", "Provider=SQLOLEDB.1;Integrated Security=SSPI;Server=TESTING\SQL2008;Uid=sa;pwd=mastercard;Persist Security Info=False"

    Rep.ShowPrintDialog = False
    ' LLamamos al reporte seleccionado.

    ' Ejecutamos el informe
    'IF REP.CalcReport
    On Error Resume Next ' SI NO TIENE DATOS

    Rep.Execute

End Sub

Private Sub cmdREPSalMinAgrupadoPorFam_Click()

    On Error GoTo UPS_err

    Dim Splitter   As String

    Dim DataPart() As String

    Splitter = "|"
    DataPart = Split(familia.Text, Splitter)
    Rep.Preview = True

    If Trim(familia.Text) = "%" Then 'osea todos
        Rep.FileName = App.path & "\Reportes\Rpt_ProductosStockMinimoAgrupado.rep"
        Rep.SetDatabaseConnectionString "VISIORION", "Provider=SQLOLEDB.1;Integrated Security=SSPI;Server=" & menup.vservidor & ";Database=" & basedatos & ";Uid=sa;pwd=" & clave_servidor & ";Persist Security Info=False"

        '    Rep.ShowPrintDialog = False
        '    Rep.Execute
    End If

    If familia.Text <> "%" Then
        Rep.FileName = App.path & "\Reportes\Rpt_ProductosStockMinimoAgrupadoPorFam.rep"
        Rep.SetDatabaseConnectionString "VISIORION", "Provider=SQLOLEDB.1;Integrated Security=SSPI;Server=" & menup.vservidor & ";Database=" & basedatos & ";Uid=sa;pwd=" & clave_servidor & ";Persist Security Info=False"
        Rep.SetParamValue "FAMILIA", DataPart(0) 'busca_familia("" & mytablex.Fields("familia")) ' familia.Text  '"SOPAS"

        'MsgBox "Seleccione una familia"
    End If

    Rep.ShowPrintDialog = False
    Rep.Execute
    Exit Sub
UPS_err:
    MsgBox "No hay registros para mostrar, ó seleccione una familia...", 48, "Aviso"
    Exit Sub

End Sub

Private Sub cmdSaldoMenorMinimo_Click()
    Rep.Preview = True
    'Establecemos la cadena de conexion de manera dinamica.
    'Rep.SetDatabaseConnectionString "SERVIDOR", "Persist Security Info=True;Driver=MySQL ODBC 5.1 Driver;SERVER=" & cServidor & ";UID=" & cUserName & ";PWD=" & cPassword & ";DATABASE=" & cBDatos & ";PORT=3306"

    'Rep.FileName = ruta & "\informes\Rpt_ProductosStockMinimo.rep"
    Rep.FileName = App.path & "\Reportes\Rpt_ProductosStockMinimo.rep"
    'cn.Open "Driver={SQL Server};Server=" & menup.vservidor & ";Database=" & buf & " ;Uid=sa;pwd=" & clave_servidor & ""
    'Rep.SetDatabaseConnectionString ("RESTAURANT6", "Driver={SQL Server};Server=TESTING\SQL2008;Uid=sa;pwd=mastercard;")
    'ok
    'Rep.SetDatabaseConnectionString "VISIORION", "Provider=SQLOLEDB.1;Integrated Security=SSPI;Server=TESTING\SQL2008;Database=RESTAURANT6;Uid=sa;pwd=mastercard;Persist Security Info=False"
    Rep.SetDatabaseConnectionString "VISIORION", "Provider=SQLOLEDB.1;Integrated Security=SSPI;Server=" & menup.vservidor & ";Database=" & basedatos & ";Uid=sa;pwd=" & clave_servidor & ";Persist Security Info=False"

    'Provider=SQLOLEDB.1;Integrated Security=SSPI;Server=TESTING\SQL2008;Database=RESTAURANT6;Uid=sa;pwd=mastercard;Persist Security Info=False
    'Rep.SetDatabaseConnectionString "RESTAURANT6", "Provider=SQLOLEDB.1;Integrated Security=SSPI;Server=TESTING\SQL2008;Uid=sa;pwd=mastercard;Persist Security Info=False"

    'Rep.SetDatabaseConnectionString REPORTEDATOS, REPORTE1
    ' No activamos el selector de impresora
    Rep.ShowPrintDialog = False
    ' LLamamos al reporte seleccionado.

    ' Ejecutamos el informe
    'IF REP.CalcReport
    On Error Resume Next 'SI NO TIENE DATOS

    Rep.Execute

    'posible solucion con dns
    '> Da codice poi ho fatto:
    '> ReportManX1.FileName = App.Path & "\repReport.rep"
    '> ReportManX1.SetDatabaseConnectionString "TT3", "DSN=" & strNomeODBC &
    '> ";PWD=pippo"
    '> ReportManX1.SetDatasetSQL "CLARABELLA", "SELECT * FROM tblMiaTabella"
    '> PreviewControl1.SetReport ReportManX1.Report
End Sub

Private Sub Command1_Click()
    Command1.Visible = False

End Sub

Public Sub dlkoiewr_Click()

    Dim mytablex As New ADODB.Recordset

    Dim I        As Integer

    If Command1.Visible = True Then Exit Sub
    If fechai.Visible = True Or fechaf.Visible = True Then
        If Len(fechai) = 0 Then
            MsgBox "Fechai No valida", 48, "Aviso"
            Exit Sub

        End If

        If Len(fechai) <> 10 Then
            MsgBox "Fechai No valida", 48, "Aviso"
            Exit Sub

        End If

        If Not IsDate(fechai) Then
            MsgBox "Fechai No valida", 48, "Aviso"
            Exit Sub

        End If

        If Len(fechaf) = 0 Then
            MsgBox "Fechaf No valida", 48, "Aviso"
            fechaf.SetFocus
            Exit Sub

        End If

        If Len(fechaf) <> 10 Then
            MsgBox "Fechaf No valida", 48, "Aviso"
            fechaf.SetFocus
            Exit Sub

        End If

        If Not IsDate(fechaf) Then
            MsgBox "Fechaf No valida", 48, "Aviso"
            fechaf.SetFocus
            Exit Sub

        End If
    
        ''' 16/01/2018 Stock Inicial en Kardex Producto y Sunat
        If Len(fechainicial) = 0 Then
            MsgBox "FechaInicial No valida", 48, "Aviso"
            fechainicial.SetFocus
            Exit Sub

        End If

        If Len(fechaf) <> 10 Then
            MsgBox "FechaInicial No valida", 48, "Aviso"
            fechainicial.SetFocus
            Exit Sub

        End If

        If Not IsDate(fechainicial) Then
            MsgBox "FechaInicial No valida", 48, "Aviso"
            fechainicial.SetFocus
            Exit Sub

        End If

        ''' 16/01/2018 Stock Inicial en Kardex Producto y Sunat
    
    End If

    'MsgBox opcion2
    If opcion2 = "1" Then 'reporte de kardex
        If bodega = "%" Then
            MsgBox "Debe Seleccionar un Almacen", 48, "Aviso"
            Exit Sub

        End If

        If excell.Value = 1 Then
            reporte_kardex_excell
            Exit Sub

        End If

        reporte_kardex
        Exit Sub

    End If

    If opcion2 = "100" Then 'reporte de kardex
        If bodega = "%" Then
            MsgBox "Debe Seleccionar un Almacen", 48, "Aviso"
            Exit Sub

        End If
   
        If excell.Value = 1 Then
            kardex_sunat_excell
            Exit Sub

        End If

        reporte_kardex_sunat
        Exit Sub

    End If

    If opcion2 = "3" Then 'Analisis de lineas
        If bodega = "%" Then
            MsgBox "Debe Seleccionar un Almacen", 48, "Aviso"
            Exit Sub

        End If

        reporte_lineas
        Exit Sub

    End If

    If opcion2 = "2" Then 'reporte de saldo inicial valorizado

        'MsgBox conteo
        If conteo = "S" Then
            If bodega = "%" Then
                MsgBox "Debe Seleccionar un Almacen", 48, "Aviso"
                Exit Sub

            End If

            reporte_conteo
            Exit Sub

        End If

        If bodega = "%" Then
            MsgBox "Debe Seleccionar un Almacen", 48, "Aviso"
            Exit Sub

        End If

        'MsgBox excell.Value
        If excell.Value = 1 Then
            reporte_excellini
            Exit Sub

        End If

        reporte_saldoini
        Exit Sub

    End If

    If opcion2 = "4" Then  'reporte de saldo actual
        If bodega = "%" Then
            MsgBox "Debe Seleccionar un Almacen", 48, "Aviso"
            Exit Sub

        End If

        If excell.Value = 1 Then
            reporte_saldoexcell
            Exit Sub

        End If

        reporte_saldo
        Exit Sub

    End If

    If opcion2 = "44" Then  'reporte de entradas Salidas
        If bodega = "%" Then
            MsgBox "Debe Seleccionar un Almacen", 48, "Aviso"
            Exit Sub

        End If

        excel_entrada_salida
        Exit Sub

    End If

    If opcion2 = "948" Then 'reporte de PRODUCTOS SIN ROTACION
        If bodega = "%" Then
            MsgBox "Debe Seleccionar un Almacen", 48, "Aviso"
            Exit Sub

        End If

        If excell.Value = 1 Then
            reporte_sinrotacionexcell
            Exit Sub

        End If

        reporte_saldo_rotacion
        Exit Sub

    End If

    If opcion2 = "1948" Then 'reporte de PRODUCTOS SIN ROTACION
        'If bodega = "%" Then
        '   MsgBox "Debe Seleccionar un Almacen", 48, "Aviso"
        '   Exit Sub
        'End If
        '  If excell.Value = 1 Then
        '     'reporte_saldoexcell
        '     Exit Sub
        '  End If
        reporte_receta

        '  Exit Sub
    End If

    If opcion2 = "94" Then 'reporte de MARGENES
        If bodega = "%" Then
            MsgBox "Debe Seleccionar un Almacen", 48, "Aviso"
            Exit Sub

        End If

        If excell.Value = 1 Then
            'reporte_saldoexcell
            Exit Sub

        End If

        reporte_saldo_margen
        Exit Sub

    End If

    If opcion2 = "6" Then 'reporte todos los almacenes
        'If bodega = "%" Then
        '   MsgBox "Debe Seleccionar un Almacen", 48, "Aviso"
        '   Exit Sub
        'End If
  
        mytablex.Open "select * from bodega ", cn, adOpenStatic, adLockOptimistic

        For I = 1 To 14
            cbuf(I) = ""
        Next I

        I = 0
        Do

            If mytablex.EOF Then Exit Do
            I = I + 1
            cbuf(I) = "" & mytablex.Fields("codigo")
            mytablex.MoveNext
        Loop
        mytablex.Close
        reporte_saldo1
        Exit Sub

    End If

    If opcion2 = "7" Then 'Lista de Precios
        If bodega = "%" Then
            MsgBox "Debe Seleccionar un Almacen", 48, "Aviso"
            Exit Sub

        End If

        reporte_saldo2
        Exit Sub

    End If

    If opcion2 = "72" Then 'lista para conteo fisico
        If bodega = "%" Then
            MsgBox "Debe Seleccionar un Almacen", 48, "Aviso"
            Exit Sub

        End If

        reporte_saldocf
        Exit Sub

    End If

    If opcion2 = "8" Then 'REPORTE DE SALDOS A UN PERIODO
        If bodega = "%" Then
            MsgBox "Debe Seleccionar un Almacen", 48, "Aviso"
            Exit Sub

        End If

        If excell.Value = 1 Then
            reporte_saldoex8
            Command1.Visible = False
            Exit Sub

        End If

        reporte_saldo8
        Exit Sub

    End If

    Command1.Visible = False

End Sub

Sub reporte_excellini()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    found = sql_producto(mytablex)

    If found = 0 Then
        mytablex.Close
        Exit Sub

    End If

    cuerpo_programa_excellini mytablex
    Command1.Visible = False

End Sub

Sub cuerpo_programa_saldo2(mytablex As ADODB.Recordset)

    Dim I           As Integer

    Dim mytabley    As New ADODB.Recordset

    Dim mytableyy   As New ADODB.Recordset

    Dim tipo_cambio As Double

    Dim sw          As Integer

    Dim temp        As String

    Dim buf         As String

    Dim sw1         As Integer

    Dim temp1       As String

    Dim buff        As String

    Dim buf1        As String

    Dim bufx        As String

    Dim saldoini    As Double

    Dim saldoindx   As Double

    Dim sdx         As Double

    Dim sdx1        As Double

    Dim sdx2        As Double

    Dim sdx3        As Double

    Dim found       As Integer

    Dim pventa1     As String

    Dim pventa2     As String

    Dim pventa3     As String

    Dim pventa4     As String

    Dim pventa5     As String

    Dim unidad1     As String

    Dim unidad2     As String

    Dim unidad3     As String

    Dim unidad4     As String

    Dim factor1     As String

    Dim factor2     As String

    Dim factor3     As String

    Dim factor4     As String

    Dim sbuf        As String

    Dim dbuf        As String

    Dim vr

    sbuf = ""
    saldoindx = 0
    sdx1 = 0
    sdx2 = 0
    suma1 = 0
    sw1 = 0
    tipo_cambio = busca_paridad()
    Command1.Visible = True
    Do

        If mytablex.EOF Then Exit Do
        vr = DoEvents()
        'found = verifica_precios("" & mytablex.Fields("producto"))
        'If found = 0 Then
        '   GoTo siguy12
        'End If

        pventa1 = ""
        pventa2 = ""
        pventa3 = ""
        pventa4 = ""
   
        unidad1 = ""
        unidad2 = ""
        unidad3 = ""
        unidad4 = ""
   
        factor1 = ""
        factor2 = ""
        factor3 = ""
        factor4 = ""

        dbuf = "select * from precios  where producto='" & "" & mytablex.Fields("producto") & "' and local='" & extra_loquesea(local1) & "'"

        If IsDate(fechavpi) And IsDate(fechavpf) Then
            dbuf = dbuf & "  and fechavp>='" & Format(fechavpi, "YYYYMMDD") & "'"
            dbuf = dbuf & " and fechavp<='" & Format(fechavpf, "YYYYMMDD") & "' "

        End If

        If mytableyy.State = 1 Then mytableyy.Close
        mytableyy.Open dbuf, cn, adOpenStatic, adLockOptimistic

        If mytableyy.RecordCount = 0 Then
            mytableyy.Close
            GoTo seguy12

        End If

        If Command1.Visible = False Then Exit Do

        '----------------------------------------------
        If proveedor <> "%" Then
            found = ver_proveedor("" & mytablex.Fields("producto"))

            If found = 0 Then GoTo seguy12

        End If

        '----------------------------------------------

        '------------- verificamos la condicion
        buf = "" & mytablex.Fields("familia")

        If sw = 0 Then
            sw = 1
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            buf = busca_familia("" & mytablex.Fields("familia"))
            found = formateaa(buf, 30, 0, 0)
            found = formateaa("*", 1, 2, 0)
            temp = "" & mytablex.Fields("familia")
            nlineas

        End If

        If "" & mytablex.Fields("familia") <> temp Then
            temp = "" & mytablex.Fields("familia")
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            buf = busca_familia("" & mytablex.Fields("familia"))
            found = formateaa(buf, 30, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas

        End If

        If sw1 = 0 Then
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("-", 1, 0, 0)
            found = formateaa("" & mytablex.Fields("subfamilia"), 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")
            sw1 = 1

        End If

        If temp1 <> "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia") Then
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("-", 1, 0, 0)
            found = formateaa("" & mytablex.Fields("subfamilia"), 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")

        End If

        found = formateaa("" & mytablex.Fields("producto"), 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("descripcio"), 40, 0, 0)
        found = formateaa("", 1, 0, 0)
        sbuf = "" & mytablex.Fields("monedav")

        If monedac = "D" Then
            sbuf = "D"

        End If

        If monedac = "S" Then
            sbuf = "S"

        End If

        found = formateaa(sbuf, 1, 0, 0)
        found = formateaa("", 1, 0, 0)

        found = formateaa("" & mytablex.Fields("unidad"), 4, 0, 0)
        found = formateaa("", 1, 0, 0)
        'buscando la lista de precios que hemos elejido
   
        'If mytableyy.State = 1 Then mytableyy.Close
        'mytableyy.Open "select * from precios  where producto='" & "" & mytablex.Fields("producto") & "' and local='" & extra_loquesea(local1) & "'", cn, adOpenStatic, adLockOptimistic
        'If mytableyy.RecordCount > 0 Then
        pventa1 = "" & mytableyy.Fields("pventa1")
        pventa2 = "" & mytableyy.Fields("pventa2")
        pventa3 = "" & mytableyy.Fields("pventa3")
        pventa4 = "" & mytableyy.Fields("pventa4")
   
        unidad1 = "" & mytableyy.Fields("unidad1")
        unidad2 = "" & mytableyy.Fields("unidad2")
        unidad3 = "" & mytableyy.Fields("unidad3")
        unidad4 = "" & mytableyy.Fields("unidad4")
   
        factor1 = "" & mytableyy.Fields("factor1")
        factor2 = "" & mytableyy.Fields("factor2")
        factor3 = "" & mytableyy.Fields("factor3")
        factor4 = "" & mytableyy.Fields("factor4")

        'End If
        mytableyy.Close

        If monedac = "S" Then
            If "" & mytablex.Fields("monedav") = "D" Then
                pventa1 = Format(Val(pventa1) * tipo_cambio, "0.00")
                pventa2 = Format(Val(pventa2) * tipo_cambio, "0.00")
                pventa3 = Format(Val(pventa3) * tipo_cambio, "0.00")
                pventa4 = Format(Val(pventa4) * tipo_cambio, "0.00")

            End If

        End If

        If monedac = "D" Then
            If "" & mytablex.Fields("monedav") = "S" Then
                pventa1 = Format(Val(pventa1) / tipo_cambio, "0.00")
                pventa2 = Format(Val(pventa2) / tipo_cambio, "0.00")
                pventa3 = Format(Val(pventa3) / tipo_cambio, "0.00")
                pventa4 = Format(Val(pventa4) / tipo_cambio, "0.00")

            End If

        End If

        found = formateaa(pventa1, 7, 0, 1)
        found = formateaa("", 1, 0, 0)
        found = formateaa(unidad2, 4, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa(pventa2, 7, 0, 1)
        found = formateaa("", 1, 0, 0)
        found = formateaa(unidad3, 4, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa(pventa3, 7, 0, 1)
        found = formateaa("", 1, 0, 0)
        found = formateaa(unidad4, 4, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa(pventa4, 7, 0, 1)
        found = formateaa("", 1, 2, 0)
        nlineas
seguy12:
        mytablex.MoveNext
    Loop

End Sub

Sub cabecera_saldo2()

    Dim mytablex As Table

    Dim buf      As String

    Dim I        As Integer

    Dim found    As Integer

    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("ALMACEN : " & bodega, 25, 2, 0)
    found = formateaa("FECHA INICIO :" & fechai, 40, 2, 0)
    found = formateaa("FECHA FINAL  :" & fechaf, 40, 2, 0)
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)
    found = formateaa("Producto ", 11, 0, 0)
    found = formateaa("Descripcio ", 41, 0, 0)
    found = formateaa("M", 2, 0, 0)
    found = formateaa("Unid", 5, 0, 0)
    found = formateaa("Pventa1", 8, 0, 0)
    found = formateaa("Unid", 5, 0, 0)
    found = formateaa("Pventa2", 8, 0, 0)
    found = formateaa("Unid", 5, 0, 0)
    found = formateaa("Pventa3", 8, 0, 0)
    found = formateaa("Unid", 5, 0, 0)
    found = formateaa("Pventa4", 8, 2, 0)
    '--------------------------------------
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)

End Sub

Sub reporte_saldo2()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    found = sql_producto(mytablex)

    If found = 0 Then
        mytablex.Close
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_saldo2
    cuerpo_programa_saldo2 mytablex
    Command1.Visible = False
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
    found = valida_wordpad(FileName)

End Sub

Sub cuerpo_programa_saldo1(mytablex As ADODB.Recordset)

    Dim I         As Integer

    Dim mytabley  As New ADODB.Recordset

    Dim mytablez  As New ADODB.Recordset

    Dim sw        As Integer

    Dim temp      As String

    Dim buf       As String

    Dim sw1       As Integer

    Dim temp1     As String

    Dim buf1      As String

    Dim bufx      As String

    Dim saldoini  As Double

    Dim saldoindx As Double

    Dim sdx       As Double

    Dim sdx1      As Double

    Dim sdx2      As Double

    Dim sdx3      As Double

    Dim found     As Integer

    saldoindx = 0
    sdx1 = 0
    sdx2 = 0
    suma1 = 0
    sw1 = 0

    Dim vr

    Command1.Visible = True
    Do

        If mytablex.EOF Then Exit Do
        vr = DoEvents()

        If Command1.Visible = False Then Exit Do

        '------------- verificamos la condicion
        If proveedor <> "%" Then
            found = ver_proveedor("" & mytablex.Fields("producto"))

            If found = 0 Then GoTo seguy12

        End If

        buf = "" & mytablex.Fields("familia")

        If sw = 0 Then
            sw = 1
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            buf = busca_familia("" & mytablex.Fields("familia"))
            found = formateaa(buf, 30, 0, 0)
            found = formateaa("*", 1, 2, 0)
            temp = "" & mytablex.Fields("familia")
            nlineas

        End If

        If "" & mytablex.Fields("familia") <> temp Then
            temp = "" & mytablex.Fields("familia")
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            buf = busca_familia("" & mytablex.Fields("familia"))
            found = formateaa(buf, 30, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas

        End If

        If sw1 = 0 Then
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("-", 1, 0, 0)
            found = formateaa("" & mytablex.Fields("subfamilia"), 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")
            sw1 = 1

        End If

        If temp1 <> "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia") Then
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("-", 1, 0, 0)
            found = formateaa("" & mytablex.Fields("subfamilia"), 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")

        End If

        found = formateaa("" & mytablex.Fields("producto"), 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("descripcio"), 40, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("unidad"), 6, 0, 0)
        found = formateaa("x", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("factor"), 4, 0, 0)
        found = formateaa("", 1, 0, 0)

        '------- IMPRIMIR LOS ALMACENES -----------------
        For I = 1 To 7
            vbuf(I) = 0
        Next I

        saldoini = 0

        If mytablez.State = 1 Then mytablez.Close
        mytablez.Open "Select * from almacen where local='" & extra_loquesea(local1) & "' and producto='" & "" & mytablex.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic

        If mytablez.RecordCount > 0 Then
            Do

                If mytablez.EOF Then Exit Do

                '----------------------------------------------------------
                For I = 1 To 7

                    If "" & cbuf(I) = "" & mytablez.Fields("bodega") Then
                        vbuf(I) = vbuf(I) + Val("" & mytablez.Fields("saldo"))

                    End If

                Next I

                '----------------------------------------------------------
   
                mytablez.MoveNext
            Loop

        End If

        sdx3 = 0

        For I = 1 To 7
            buf = "" & vbuf(I)

            If Val(buf) = 0 Then
                buf = ""

            End If

            bufx = calcula_saldo(Val(buf), Val("" & mytablex.Fields("factor")))
            found = formateaa(bufx, 7, 0, 0)
            found = formateaa("", 1, 0, 0)
            sdx3 = sdx3 + Val("" & vbuf(I))
        Next I

        buf = "" & sdx3

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 7, 2, 0)
        nlineas
seguy12:
        mytablex.MoveNext
    Loop
    bufx = "" & suma1
    found = formateaa("", 86, 0, 0)
    found = formateaa(bufx, 10, 0, 1)
    found = formateaa("", 1, 2, 0)

End Sub

Sub cabecera_saldo1()

    Dim mytablex As Table

    Dim buf      As String

    Dim I        As Integer

    Dim found    As Integer

    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("ALMACEN : " & bodega, 25, 2, 0)
    found = formateaa("FECHA INICIO :" & fechai, 40, 2, 0)
    found = formateaa("FECHA FINAL  :" & fechaf, 40, 2, 0)
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)
    found = formateaa("Producto ", 11, 0, 0)
    found = formateaa("Descripcio ", 41, 0, 0)
    found = formateaa("Unid", 7, 0, 0)
    found = formateaa("Fx", 5, 0, 0)

    '--------------------------------------
    For I = 1 To 7
        found = formateaa("" & cbuf(I), 7, 0, 0)
        found = formateaa("", 1, 0, 0)
    Next I

    found = formateaa("Total ", 7, 0, 0)
    found = formateaa("", 1, 2, 0)
    '--------------------------------------
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)

End Sub

Sub reporte_saldo1()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    found = sql_producto(mytablex)

    If found = 0 Then
        mytablex.Close
    
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_saldo1
    cuerpo_programa_saldo1 mytablex
    Command1.Visible = False
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
     
    'genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1
    found = valida_wordpad(FileName)

End Sub

Sub reporte_saldo()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    ''''21/09/2017 kenyo Reporte de Stock minimo Ticket
    'found = sql_producto(mytablex)
    If Combo2 = "Debajo de Stock minimo" Then
        found = sql_productominimo(mytablex)
    Else
        found = sql_producto(mytablex)

    End If

    ''''21/09/2017 kenyo Reporte de Stock minimo Ticket

    If found = 0 Then
        mytablex.Close
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_saldo
    cuerpo_programa_saldo mytablex
    Command1.Visible = False
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
    'genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1
    found = valida_wordpad(FileName)

End Sub

Sub reporte_saldo_rotacion()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    found = sql_producto2(mytablex)

    If found = 0 Then
        mytablex.Close
   
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_rotacion
    cuerpo_programa_rotacion mytablex
    Command1.Visible = False
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
     
    'genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1
    found = valida_wordpad(FileName)

End Sub

Sub reporte_receta()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    found = sql_producto(mytablex)

    If found = 0 Then
        mytablex.Close
   
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_receta
    cuerpo_programa_receta mytablex
    Command1.Visible = False
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
     
    'genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1
    found = valida_wordpad(FileName)

End Sub

Sub reporte_saldoini()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    found = sql_producto(mytablex)

    If found = 0 Then
        mytablex.Close

        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_saldoini
    cuerpo_programa_saldoini mytablex
    Command1.Visible = False
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
     
    'genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1
    found = valida_wordpad(FileName)

End Sub

Sub reporte_conteo()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    found = sql_producto(mytablex)

    If found = 0 Then
        mytablex.Close
    
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_conteo
    cuerpo_programa_conteo mytablex
    Command1.Visible = False
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
    'genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1
    found = valida_wordpad(FileName)

End Sub

Sub reporte_kardex()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    found = sql_producto(mytablex)

    If found = 0 Then
        mytablex.Close
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_kardex
    cuerpo_programa_kardex mytablex
    Command1.Visible = False
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
    found = valida_wordpad(FileName)

End Sub

Sub reporte_kardex_sunat()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    found = sql_producto(mytablex)

    If found = 0 Then
        mytablex.Close
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    'cabecera_kardex_sunat
    cuerpo_programa_kardex_sunat mytablex
    Command1.Visible = False
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
    'genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1
    found = valida_wordpad(FileName)

End Sub

Private Sub Form_Activate()
    Command1.Top = 2500: Command1.Left = 1000

    If Len(Trim(xbasedatos)) = 0 Then
        xbasedatos = "detalle"

    End If

    'fechai = Format(busca_paramed(extra_loquesea(bodega)), "dd/mm/yyyy")

End Sub

Private Sub Form_Load()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    'listap.Clear
    'listap.AddItem "01"
    'listap.AddItem "02"
    'listap.AddItem "03"
    'listap.AddItem "04"
    'listap.AddItem "05"
    'listap.ListIndex = 0

    ''''21/09/2017 kenyo Reporte de Stock minimo Ticket
    Combo2.Clear
    Combo2.AddItem "%"
    Combo2.AddItem "Debajo de Stock minimo"
    Combo2.ListIndex = 0
    ''''21/09/2017 kenyo Reporte de Stock minimo Ticket

    ''''21/09/2017 kenyo Mejora reporte entradas salidas, kardex
    vesubfamilia.Clear
    vesubfamilia.AddItem "N"
    vesubfamilia.AddItem "S"
    vesubfamilia.ListIndex = 0
    ''''21/09/2017 kenyo Mejora reporte entradas salidas, kardex

    conigv.Clear
    conigv.AddItem ""
    conigv.AddItem "S"
    conigv.AddItem "N"
    conigv.ListIndex = 0

    quecosto.AddItem "COSTOULTIMO"
    quecosto.AddItem "COSTOPROMEDIO"
    quecosto.AddItem "PRECIOVENTA"
    quecosto.ListIndex = 0

    conteo.Clear
    conteo.AddItem "N"
    conteo.AddItem "S"
    conteo.ListIndex = 0

    Combo1.AddItem "SALDO>0"
    Combo1.AddItem "SALDO<0"
    Combo1.AddItem "SALDO<=0"
    Combo1.AddItem "SALDO>=0"
    Combo1.AddItem "SALDO=0"
    Combo1.AddItem "TODOS"
    Combo1.ListIndex = 0

    moneda.Clear
    moneda.AddItem "%"
    moneda.AddItem "S"
    moneda.AddItem "D"
    moneda.ListIndex = 0

    monedac.Clear
    monedac.AddItem "%"
    monedac.AddItem "S"
    monedac.AddItem "D"
    monedac.ListIndex = 0

    igv.Clear
    igv.AddItem "%"
    igv.AddItem "GRAVADO"
    igv.AddItem "EXENTO"
    igv.ListIndex = 0

    mytablex.Open "Select * from familia ORDER by familia", cn, adOpenStatic, adLockOptimistic
    familia.AddItem "%"
    Do

        If mytablex.EOF Then Exit Do
        familia.AddItem "" & mytablex.Fields("familia") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    familia.ListIndex = 0
    mytablex.Close
    mytablex.Open "Select * from subfamil ", cn, adOpenStatic, adLockOptimistic
    subfamilia.AddItem "%"
    Do

        If mytablex.EOF Then Exit Do
        subfamilia.AddItem "" & mytablex.Fields("subfamilia")
        mytablex.MoveNext
    Loop
    subfamilia.ListIndex = 0
    mytablex.Close

    mytablex.Open "Select * from seccion", cn, adOpenStatic, adLockOptimistic
    seccion.AddItem "%"
    Do

        If mytablex.EOF Then Exit Do
        seccion.AddItem "" & mytablex.Fields("seccion")
        mytablex.MoveNext
    Loop
    seccion.ListIndex = 0
    mytablex.Close
    mytablex.Open "Select * from categori", cn, adOpenStatic, adLockOptimistic
    categoria.AddItem "%"
    Do

        If mytablex.EOF Then Exit Do
        categoria.AddItem "" & mytablex.Fields("categoria")
        mytablex.MoveNext
    Loop
    categoria.ListIndex = 0
    mytablex.Close
    mytablex.Open "Select * from color", cn, adOpenStatic, adLockOptimistic
    color.AddItem "%"
    Do

        If mytablex.EOF Then Exit Do
        color.AddItem "" & mytablex.Fields("color")
        mytablex.MoveNext
    Loop
    color.ListIndex = 0
    mytablex.Close
    mytablex.Open "Select * from marca ORDER by marca ", cn, adOpenStatic, adLockOptimistic
    marca.AddItem "%"
    Do

        If mytablex.EOF Then Exit Do
        marca.AddItem "" & mytablex.Fields("marca")
        mytablex.MoveNext
    Loop
    marca.ListIndex = 0
    mytablex.Close
    mytablex.Open "Select * from talla", cn, adOpenStatic, adLockOptimistic
    linea.AddItem "%"
    Do

        If mytablex.EOF Then Exit Do
        linea.AddItem "" & mytablex.Fields("talla")
        mytablex.MoveNext
    Loop
    linea.ListIndex = 0
    mytablex.Close
    bodega.Clear
    mytablex.Open "Select * from bodega", cn, adOpenStatic, adLockOptimistic
    bodega.AddItem "%"
    Do

        If mytablex.EOF Then Exit Do
        bodega.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("Nombre")
        mytablex.MoveNext
    Loop
    bodega.ListIndex = 1
    mytablex.Close
 
    mytablex.Open "Select * from tlocal", cn, adOpenStatic, adLockOptimistic
    local1.Clear
    Do

        If mytablex.EOF Then Exit Do
        local1.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("Nombre")
        mytablex.MoveNext
    Loop
    local1.ListIndex = 0

    If local1.ListCount = 2 Then
        local1.ListIndex = 1

    End If

    mytablex.Close
    fechai = Format(Now, "dd/mm/yyyy")

    If opcion2 = "1" Or opcion2 = "2" Then

        'found = busca_parame(3)
        'fechai.Enabled = False
        'fechai = Format(busca_paramed(extra_loquesea(bodega)), "dd/mm/yyyy")
        'Else: fechai = Format(Now, "dd/mm/yyyy")
    End If

    fechaf = Format(Now, "dd/mm/yyyy")

    ''' 16/01/2018 Stock Inicial en Kardex Producto y Sunat
    fechainicial = busca_fechainicialXAlmacen
    ''' 16/01/2018 Stock Inicial en Kardex Producto y Sunat

End Sub

''' 16/01/2018 Stock Inicial en Kardex Producto y Sunat
Function busca_fechainicialXAlmacen() As String

    Dim buf      As String

    Dim found    As Integer

    'MsgBox buvendedor
    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT fecha FROM bodega where  codigo='" & extra_loquesea(bodega) & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_fechainicialXAlmacen = "" & mytablex.Fields("fecha")
      
        If mytablex.Fields("fecha") = "" Then
            busca_fechainicialXAlmacen = Format(Now, "dd/mm/yyyy")

        End If
      
    End If

    mytablex.Close
    Exit Function

End Function

''' 16/01/2018 Stock Inicial en Kardex Producto y Sunat

Private Sub ldso8912_Click()

    If Command1.Visible = True Then
        Command1.Visible = False
        Exit Sub

    End If

    repinv.Hide
    Unload repinv

End Sub

Function busca_parame(sw As Integer)

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "Select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If sw = 3 Then
            fechai = "" & mytablex.Fields("saldoini")

            'fechai = Mid$(fechai, 1, 2) & "/" & Mid$(fechai, 3, 2) & "/" & Mid$(fechai, 5, 4)
        End If

        busca_parame = 1

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function sql_producto2(mytablex As ADODB.Recordset)

    Dim buf As String

    '19/06/2017 kenyo Correción  reporte productos sin rotacion. '''reporte de PRODUCTOS SIN ROTACION
    'buf = "select * from producto where producto not in (select producto from detalle where (tipo='1' or tipo='2' or tipo='FC') and estado='2' and fecha>='" & Format(fechari, "YYYYMMDD") & "' and fecha<='" & Format(fecharf, "YYYYMMDD") & "') and producto like '" & producto & "'"
    buf = "select * from producto where producto not in (select producto from detalle where (ACU='1' or ACU='A' or ACU='B' OR ACU='C' or ACU='G' or ACU='N') and estado='2' and fecha>='" & Format(fechari, "YYYYMMDD") & "' and fecha<='" & Format(fecharf, "YYYYMMDD") & "') and producto like '" & producto & "'"

    '19/06/2017 kenyo Correción  reporte productos sin rotacion. '''reporte de PRODUCTOS SIN ROTACION

    If Barras <> "%" Then
        buf = buf & " and barras like '" & Barras & "'"

    End If

    If descripcio <> "%" Then
        buf = buf & " and descripcio like '" & descripcio & "'"

    End If

    If familia <> "%" Then
        buf = buf & " and familia like '" & extra_loquesea(familia) & "'"

    End If

    If subfamilia <> "%" Then
        buf = buf & " and subfamilia like '" & subfamilia & "'"

    End If

    If seccion <> "%" Then
        buf = buf & " and seccion like '" & seccion & "'"

    End If

    If categoria <> "%" Then
        buf = buf & " and categoria like '" & categoria & "'"

    End If

    If linea <> "%" Then
        buf = buf & " and linea like '" & linea & "'"

    End If

    If color <> "%" Then
        buf = buf & " and color like '" & color & "'"

    End If

    If marca <> "%" Then
        buf = buf & " and marca like '" & marca & "'"

    End If

    If fechavi <> "%" And fechavf <> "%" Then
        If IsDate(fechavi) And IsDate(fechavf) Then
            buf = buf & "  fechavence>='" & Format(fechavi, "YYYYMMDD") & "'"
            buf = buf & " and fechavence<='" & Format(fechavf, "YYYYMMDD") & "' "

        End If

    End If

    If igv = "EXENTO" Then
        buf = buf & " and igv=0"

    End If

    If igv = "GRAVADO" Then
        buf = buf & " and igv>0"

    End If

    buf = buf & " order by familia,Subfamilia,descripcio"

    'If Combo1.Text = "Con Stock Minimo" Then
    '    buf = "select P.PRODUCTO,P.DESCRIPCIO,P.UNIDAD,P.FACTOR,P.MINIMO,A.SALDO AS SALDO_ACTUAL, P.COSTOU, P.COSTOP, A.SALDO * P.COSTOU AS TOTAL, P.FAMILIA, P.SUBFAMILIA   from producto AS P, ALMACEN AS A WHERE P.PRODUCTO=A.PRODUCTO AND A.SALDO < P.MINIMO "
    'End If
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    sql_producto2 = 1

End Function

''''21/09/2017 kenyo Reporte de Stock minimo Ticket
Function sql_productominimo(mytablex As ADODB.Recordset)

    Dim buf As String

    buf = "select p.producto,p.descripcio,p.familia,p.subfamilia,p.categoria,p.unidad,p.factor,p.costou,p.costop,p.igv,p.minimo from producto AS P, ALMACEN AS A WHERE P.PRODUCTO=A.PRODUCTO AND A.SALDO <= P.MINIMO AND P.MINIMO>0  and p.producto like '" & producto & "'"

    If Barras <> "%" Then
        buf = buf & " and barras like '" & Barras & "'"

    End If

    If descripcio <> "%" Then
        buf = buf & " and descripcio like '" & descripcio & "'"

    End If

    If familia <> "%" Then
        buf = buf & " and familia like '" & extra_loquesea(familia) & "'"

    End If

    If subfamilia <> "%" Then
        buf = buf & " and subfamilia like '" & subfamilia & "'"

    End If

    If seccion <> "%" Then
        buf = buf & " and seccion like '" & seccion & "'"

    End If

    If categoria <> "%" Then
        buf = buf & " and categoria like '" & categoria & "'"

    End If

    If linea <> "%" Then
        buf = buf & " and linea like '" & linea & "'"

    End If

    If color <> "%" Then
        buf = buf & " and color like '" & color & "'"

    End If

    If marca <> "%" Then
        buf = buf & " and marca like '" & marca & "'"

    End If

    If fechavi <> "%" And fechavf <> "%" Then
        If IsDate(fechavi) And IsDate(fechavf) Then
            buf = buf & "  fechavence>='" & Format(fechavi, "YYYYMMDD") & "'"
            buf = buf & " and fechavence<='" & Format(fechavf, "YYYYMMDD") & "' "

        End If

    End If

    If igv = "EXENTO" Then
        buf = buf & " and igv=0"

    End If

    If igv = "GRAVADO" Then
        buf = buf & " and igv>0"

    End If

    buf = buf & " order by p.familia,p.Subfamilia,p.descripcio"
    'If Combo1.Text = "Con Stock Minimo" Then
    '    buf = "select P.PRODUCTO,P.DESCRIPCIO,P.UNIDAD,P.FACTOR,P.MINIMO,A.SALDO AS SALDO_ACTUAL, P.COSTOU, P.COSTOP, A.SALDO * P.COSTOU AS TOTAL, P.FAMILIA, P.SUBFAMILIA   from producto AS P, ALMACEN AS A WHERE P.PRODUCTO=A.PRODUCTO AND A.SALDO < P.MINIMO "
    'End If
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    sql_productominimo = 1

End Function

''''21/09/2017 kenyo Reporte de Stock minimo Ticket

''''21/09/2017 kenyo Mejora reporte entradas salidas, kardex
Function sql_productoConMovimiento(mytablex As ADODB.Recordset)

    'Dim buf As String
    'buf = "select fechavence,p.igv,p.barras,p.producto,p.DESCRIPCIO,p.familia,p.Subfamilia,p.factor,p.unidad,p.seccion,p.categoria,p.linea,p.color,p.marca from producto p inner join detalle d on p.producto=d.producto  and p.producto like '" & producto & "'"
    '
    '
    'buf = buf & " and local='" & extra_loquesea(local1) & "' and bodega='" & extra_loquesea(bodega) & "'"
    '
    'If Barras <> "%" Then
    '    buf = buf & " and barras like '" & Barras & "'"
    'End If
    'If descripcio <> "%" Then
    '    buf = buf & " and p.descripcio like '" & descripcio & "'"
    'End If
    'If familia <> "%" Then
    '    buf = buf & " and p.familia like '" & extra_loquesea(familia) & "'"
    'End If
    'If subfamilia <> "%" Then
    '    buf = buf & " and p.subfamilia like '" & subfamilia & "'"
    'End If
    'If seccion <> "%" Then
    '    buf = buf & " and p.seccion like '" & seccion & "'"
    'End If
    'If categoria <> "%" Then
    '    buf = buf & " and p.categoria like '" & categoria & "'"
    'End If
    'If linea <> "%" Then
    '    buf = buf & " and linea like '" & linea & "'"
    'End If
    'If color <> "%" Then
    '    buf = buf & " and color like '" & color & "'"
    'End If
    'If marca <> "%" Then
    '    buf = buf & " and p.marca like '" & marca & "'"
    'End If
    'If fechavi <> "%" And fechavf <> "%" Then
    '   If IsDate(fechavi) And IsDate(fechavf) Then
    '      buf = buf & "  fechavence>='" & Format(fechavi, "YYYYMMDD") & "'"
    '      buf = buf & " and fechavence<='" & Format(fechavf, "YYYYMMDD") & "' "
    '   End If
    'End If
    '
    'buf = buf & "  and d.fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    'buf = buf & " and d.fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    '
    'If igv = "EXENTO" Then
    '   buf = buf & " and igv=0"
    'End If
    'If igv = "GRAVADO" Then
    '   buf = buf & " and igv>0"
    'End If
    '
    '
    'buf = buf & "  group by fechavence,p.igv,p.barras,p.producto,p.DESCRIPCIO,p.familia,p.Subfamilia,p.factor,p.unidad,p.seccion,p.categoria,p.linea,p.color,p.marca  order by p.familia,p.Subfamilia,p.descripcio"
    ''If Combo1.Text = "Con Stock Minimo" Then
    ''    buf = "select P.PRODUCTO,P.DESCRIPCIO,P.UNIDAD,P.FACTOR,P.MINIMO,A.SALDO AS SALDO_ACTUAL, P.COSTOU, P.COSTOP, A.SALDO * P.COSTOU AS TOTAL, P.FAMILIA, P.SUBFAMILIA   from producto AS P, ALMACEN AS A WHERE P.PRODUCTO=A.PRODUCTO AND A.SALDO < P.MINIMO "
    ''End If
    'mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    '
    'sql_productoConMovimiento = 1
    Dim buf As String

    buf = "select * from producto where producto.producto in (select producto from detalle d where ESTADO='2' AND  d.fecha>='" & Format(fechai, "YYYYMMDD") & "' and d.fecha<='" & Format(fechaf, "YYYYMMDD") & "' and local='" & extra_loquesea(local1) & "' and bodega='" & extra_loquesea(bodega) & "' ) and producto like '" & producto & "'"

    If Barras <> "%" Then
        buf = buf & " and barras like '" & Barras & "'"

    End If

    If descripcio <> "%" Then
        buf = buf & " and descripcio like '" & descripcio & "'"

    End If

    If familia <> "%" Then
        buf = buf & " and familia like '" & extra_loquesea(familia) & "'"

    End If

    If subfamilia <> "%" Then
        buf = buf & " and subfamilia like '" & subfamilia & "'"

    End If

    If seccion <> "%" Then
        buf = buf & " and seccion like '" & seccion & "'"

    End If

    If categoria <> "%" Then
        buf = buf & " and categoria like '" & categoria & "'"

    End If

    If linea <> "%" Then
        buf = buf & " and linea like '" & linea & "'"

    End If

    If color <> "%" Then
        buf = buf & " and color like '" & color & "'"

    End If

    If marca <> "%" Then
        buf = buf & " and marca like '" & marca & "'"

    End If

    If fechavi <> "%" And fechavf <> "%" Then
        If IsDate(fechavi) And IsDate(fechavf) Then
            buf = buf & "  fechavence>='" & Format(fechavi, "YYYYMMDD") & "'"
            buf = buf & " and fechavence<='" & Format(fechavf, "YYYYMMDD") & "' "

        End If

    End If

    If igv = "EXENTO" Then
        buf = buf & " and igv=0"

    End If

    If igv = "GRAVADO" Then
        buf = buf & " and igv>0"

    End If

    buf = buf & " order by familia,Subfamilia,descripcio"
    'If Combo1.Text = "Con Stock Minimo" Then
    '    buf = "select P.PRODUCTO,P.DESCRIPCIO,P.UNIDAD,P.FACTOR,P.MINIMO,A.SALDO AS SALDO_ACTUAL, P.COSTOU, P.COSTOP, A.SALDO * P.COSTOU AS TOTAL, P.FAMILIA, P.SUBFAMILIA   from producto AS P, ALMACEN AS A WHERE P.PRODUCTO=A.PRODUCTO AND A.SALDO < P.MINIMO "
    'End If
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    sql_productoConMovimiento = 1

End Function

''''21/09/2017 kenyo Mejora reporte entradas salidas, kardex

Function sql_producto(mytablex As ADODB.Recordset)

    Dim buf As String

    buf = "select * from producto where producto like '" & producto & "'"

    If Barras <> "%" Then
        buf = buf & " and barras like '" & Barras & "'"

    End If

    If descripcio <> "%" Then
        buf = buf & " and descripcio like '" & descripcio & "'"

    End If

    If familia <> "%" Then
        buf = buf & " and familia like '" & extra_loquesea(familia) & "'"

    End If

    If subfamilia <> "%" Then
        buf = buf & " and subfamilia like '" & subfamilia & "'"

    End If

    If seccion <> "%" Then
        buf = buf & " and seccion like '" & seccion & "'"

    End If

    If categoria <> "%" Then
        buf = buf & " and categoria like '" & categoria & "'"

    End If

    If linea <> "%" Then
        buf = buf & " and linea like '" & linea & "'"

    End If

    If color <> "%" Then
        buf = buf & " and color like '" & color & "'"

    End If

    If marca <> "%" Then
        buf = buf & " and marca like '" & marca & "'"

    End If

    If fechavi <> "%" And fechavf <> "%" Then
        If IsDate(fechavi) And IsDate(fechavf) Then
            buf = buf & "  fechavence>='" & Format(fechavi, "YYYYMMDD") & "'"
            buf = buf & " and fechavence<='" & Format(fechavf, "YYYYMMDD") & "' "

        End If

    End If

    If igv = "EXENTO" Then
        buf = buf & " and igv=0"

    End If

    If igv = "GRAVADO" Then
        buf = buf & " and igv>0"

    End If

    buf = buf & " order by familia,Subfamilia,descripcio"
    'If Combo1.Text = "Con Stock Minimo" Then
    '    buf = "select P.PRODUCTO,P.DESCRIPCIO,P.UNIDAD,P.FACTOR,P.MINIMO,A.SALDO AS SALDO_ACTUAL, P.COSTOU, P.COSTOP, A.SALDO * P.COSTOU AS TOTAL, P.FAMILIA, P.SUBFAMILIA   from producto AS P, ALMACEN AS A WHERE P.PRODUCTO=A.PRODUCTO AND A.SALDO < P.MINIMO "
    'End If
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    sql_producto = 1

End Function

Sub cuerpo_programa_kardex(mytablex As ADODB.Recordset)

    Dim mytabley  As New ADODB.Recordset

    Dim sw        As Integer

    Dim xbuf      As String

    Dim temp      As String

    Dim buf       As String

    Dim mytablez  As New ADODB.Recordset

    Dim sw1       As Integer

    Dim temp1     As String

    Dim buf1      As String

    Dim bufx      As String

    Dim saldoini  As Double

    Dim saldoindx As Double

    Dim sdx       As Double

    Dim sdx1      As Double

    Dim sdx2      As Double

    'Dim xbuf As String
    Dim xcosto    As Double

    Dim mytablera As New ADODB.Recordset

    Dim found     As Integer

    Dim vr

    Dim nsw As Integer

    saldoindx = 0
    sdx1 = 0
    sdx2 = 0
    sw1 = 0
    xcosto = 0
    Command1.Visible = True
    Do

        If mytablex.EOF Then Exit Do

        If proveedor <> "%" Then
            found = ver_proveedor("" & mytablex.Fields("producto"))

            If found = 0 Then GoTo seguy2

        End If

        vr = DoEvents()

        If Command1.Visible = False Then Exit Do

        buf = "" & mytablex.Fields("familia")

        If sw = 0 Then
            sw = 1
            temp = "" & mytablex.Fields("familia")
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("*", 1, 2, 0)
            nlineas

        End If

        If "" & mytablex.Fields("familia") <> temp Then
            temp = "" & mytablex.Fields("familia")
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas

        End If

        If sw1 = 0 Then
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("-", 1, 0, 0)
            found = formateaa("" & mytablex.Fields("subfamilia"), 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")
            sw1 = 1

        End If

        If temp1 <> "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia") Then
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("-", 1, 0, 0)
            found = formateaa("" & mytablex.Fields("subfamilia"), 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")

        End If

        'xcosto = Val(Format(Val("" & mytablex.Fields("costou")), "0.000000"))
        found = formateaa("" & mytablex.Fields("producto"), 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("descripcio"), 40, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("unidad"), 6, 0, 0)
        found = formateaa("x", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("factor"), 4, 0, 0)
        found = formateaa("", 1, 0, 0)
        'SALDO INICIAL
        saldoini = 0
        xcosto = 0

        If mytablez.State = 1 Then mytablez.Close
        mytablez.Open "Select * from dsaldoini where local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "' and fecha='" & fechai & "'", cn, adOpenStatic, adLockOptimistic
        Do

            If mytablez.EOF Then Exit Do
            saldoini = saldoini + Val("" & mytablez.Fields("cantidad")) * Val("" & mytablez.Fields("factor"))
            xcosto = Val(Format(Val("" & mytablez.Fields("precio")), "0.000000"))
            mytablez.MoveNext
        Loop
        mytablez.Close
        xcosto = Val(Format(Val("" & mytablex.Fields("costou")), "0.000000"))

        If conigv = "" Or conigv = "S" Then
            xcosto = xcosto

        End If

        If conigv = "N" Then
            If Val("" & mytablex.Fields("igv")) > 0 Then
                xcosto = xcosto / (1 + (Val("" & mytablex.Fields("igv")) / 100))
                xcosto = Val(Format(xcosto, "0.00000"))

            End If

        End If

        bufx = "" & saldoini
        found = formateaa("", 5, 0, 0)
        saldoindx = saldoini
        xbuf = "" & saldoini 'calcula_saldo(saldoini, Val("" & mytablex.Fields("factor")))
        found = formateaa(xbuf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & xcosto
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        sdx = Format(xcosto * Val(bufx), "0.00")
        buf = Format(sdx, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 2, 0)
        nlineas
        sdx2 = 0
        sdx1 = 0
        nsw = 0
        '-------ahora las transacciones------------
        buf = "select * from " & xbasedatos & " where "
        buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
        buf = buf & " and local='" & extra_loquesea(local1) & "'"
        buf = buf & " and producto='" & "" & mytablex.Fields("producto") & "'"
        buf = buf & " and bodega='" & extra_loquesea(bodega) & "'"

        If VENTANEGRA = "S" Then
            buf = buf & " and acu<>'G' AND acu<>'P' "

        End If

        buf = buf & " and (acu='1' or acu='S' or acu='T' or acu='C' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='E' or acu='N')"
        buf = buf & " and acu1=''"
        buf = buf & " and estado='2'"
        buf = buf & " order by fecha,hora"

        'MsgBox buf
        If mytabley.State = 1 Then mytabley.Close
        mytabley.Open buf, cn, adOpenStatic, adLockOptimistic
        Do

            If mytabley.EOF Then Exit Do
            vr = DoEvents()

            If Command1.Visible = False Then Exit Do
            Command1.Caption = " " & mytabley.Fields("local") & " " & mytabley.Fields("tipo") & " " & mytabley.Fields("serie") & " " & mytabley.Fields("numero") & " " & mytabley.Fields("fecha")
            'If mytablera.State = 1 Then mytablera.Close
            'mytablera.Open "select tipo1 from factura where local='" & "" & mytabley.Fields("local") & "' and tipo='" & "" & mytabley.Fields("tipo") & "' and serie='" & "" & mytabley.Fields("serie") & "'  and numero='" & "" & mytabley.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic
            'If mytablera.RecordCount > 0 Then
            '   found = ve_descarga("" & mytablera.Fields("tipo1"))
            '   If found = 1 Then 'qe no se descarge
            '      mytablera.Close
            '      GoTo siguiente_busca
            '   End If
            'End If
            'mytablera.Close

            'MsgBox "" & mytabley.Fields("acu")

            '19/06/2017 kenyo NOTA DE CREDITO
            'If "" & mytabley.Fields("acu") = "1" Or "" & mytabley.Fields("acu") = "T" Or "" & mytabley.Fields("acu") = "A" Or "" & mytabley.Fields("acu") = "B" Or "" & mytabley.Fields("acu") = "C" Or "" & mytabley.Fields("acu") = "D" Or "" & mytabley.Fields("acu") = "G" Or "" & mytabley.Fields("acu") = "N" Then
            If "" & mytabley.Fields("acu") = "1" Or "" & mytabley.Fields("acu") = "T" Or "" & mytabley.Fields("acu") = "A" Or "" & mytabley.Fields("acu") = "B" Or "" & mytabley.Fields("acu") = "C" Or "" & mytabley.Fields("acu") = "D" Or "" & mytabley.Fields("acu") = "G" Or "" & mytabley.Fields("acu") = "N" Or "" & mytabley.Fields("acu") = "E" Then
                '19/06/2017 kenyo NOTA DE CREDITO
                found = formateaa("" & mytabley.Fields("tipo"), 3, 0, 0)
                found = formateaa("", 1, 0, 0)
                found = formateaa("" & mytabley.Fields("serie"), 4, 0, 0)
                found = formateaa("", 1, 0, 0)
                found = formateaa("" & mytabley.Fields("numero"), 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                xbuf = Format("" & mytabley.Fields("Fecha"), "dd/mm/yyyy")
                found = formateaa(xbuf, 10, 0, 0)
                found = formateaa("", 1, 0, 0)
                found = formateaa("" & mytabley.Fields("bodega"), 2, 0, 0)
                found = formateaa("", 1, 0, 0)
                found = formateaa("" & mytabley.Fields("unidad"), 6, 0, 0)
                found = formateaa("", 1, 0, 0)
                found = formateaa("" & mytabley.Fields("Factor"), 4, 0, 0)
                found = formateaa("", 1, 0, 0)
                found = formateaa("", 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                found = formateaa("" & mytabley.Fields("cantidad"), 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                'xcosto = 0
                'If Val("" & mytabley.Fields("precio")) = 0 Then
                '   xcosto = mytablex.Fields("costou")
                'End If
   
                saldoindx = saldoindx - Val("" & mytabley.Fields("cantidad")) * Val("" & mytabley.Fields("factor"))
                buf = "" & saldoindx 'calcula_saldo(saldoindx, Val("" & mytabley.Fields("factor")))
                sdx2 = sdx2 + Val("" & mytabley.Fields("cantidad")) * Val("" & mytabley.Fields("factor"))
                found = formateaa(buf, 10, 0, 1)   'saldo
                found = formateaa("", 1, 0, 0)
                found = formateaa("" & xcosto, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                sdx = xcosto * Val("" & mytabley.Fields("cantidad"))
                buf = Format(sdx, "0.00")
                found = formateaa(buf, 10, 0, 1)   'costo
                found = formateaa("", 1, 0, 0)
                found = formateaa("" & mytabley.Fields("hora"), 10, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_nombre(mytabley)
                found = formateaa(buf, 10, 2, 0)
                nlineas

            End If
   
            '19/06/2017 kenyo NOTA DE CREDITO
            'If "" & mytabley.Fields("acu") = "S" Or "" & mytabley.Fields("acu") = "J" Or "" & mytabley.Fields("acu") = "K" Or "" & mytabley.Fields("acu") = "L" Or "" & mytabley.Fields("acu") = "M" Or "" & mytabley.Fields("acu") = "P" Or "" & mytabley.Fields("acu") = "E" Then
            If "" & mytabley.Fields("acu") = "S" Or "" & mytabley.Fields("acu") = "J" Or "" & mytabley.Fields("acu") = "K" Or "" & mytabley.Fields("acu") = "L" Or "" & mytabley.Fields("acu") = "M" Or "" & mytabley.Fields("acu") = "P" Then
                '19/06/2017 kenyo NOTA DE CREDITO
  
                'found = formateaa("", 44, 0, 0)
                found = formateaa("" & mytabley.Fields("tipo"), 3, 0, 0)
                found = formateaa("", 1, 0, 0)
                found = formateaa("" & mytabley.Fields("serie"), 4, 0, 0)
                found = formateaa("", 1, 0, 0)
                found = formateaa("" & mytabley.Fields("numero"), 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                xbuf = Format("" & mytabley.Fields("Fecha"), "dd/mm/yyyy")
                found = formateaa(xbuf, 10, 0, 0)
                found = formateaa("", 1, 0, 0)
                found = formateaa("" & mytabley.Fields("bodega"), 2, 0, 0)
                found = formateaa("", 1, 0, 0)
                found = formateaa("" & mytabley.Fields("unidad"), 6, 0, 0)
                found = formateaa("", 1, 0, 0)
                found = formateaa("" & mytabley.Fields("Factor"), 4, 0, 0)
                found = formateaa("", 1, 0, 0)
                found = formateaa("" & mytabley.Fields("cantidad"), 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                found = formateaa("", 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                'xcosto = 0
                'If Val("" & mytabley.Fields("precio")) = 0 Then
                '   xcosto = mytablex.Fields("costou")
                'End If
   
                saldoindx = saldoindx + Val("" & mytabley.Fields("cantidad")) * Val("" & mytabley.Fields("factor"))
                'buf = saldoindx
                buf = "" & saldoindx 'calcula_saldo(saldoindx, Val("" & mytabley.Fields("factor")))
                sdx1 = sdx1 + Val("" & mytabley.Fields("cantidad")) * Val("" & mytabley.Fields("factor"))
                found = formateaa(buf, 10, 0, 1)   'saldo
                found = formateaa("", 1, 0, 0)
                found = formateaa("" & xcosto, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                sdx = xcosto * saldoindx
                buf = Format(sdx, "0.00")
                found = formateaa(buf, 10, 0, 1)   'saldo
                found = formateaa("", 1, 0, 0)
                found = formateaa("" & mytabley.Fields("hora"), 10, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_nombre(mytabley)
                found = formateaa(buf, 10, 0, 0)
                found = formateaa("", 1, 2, 0)
                'MsgBox "ab"
                nlineas

            End If

siguiente_busca:
            mytabley.MoveNext
        Loop
        found = formateaa("", 46, 0, 0)
        buf = "" & sdx1
        found = formateaa(buf, 10, 0, 1)   'saldo
        found = formateaa("", 1, 0, 0)
        buf = "" & sdx2
        found = formateaa(buf, 10, 0, 1)   'saldo
        found = formateaa("", 1, 2, 0)
        nlineas
        sdx1 = 0
        sdx2 = 0
        '---------------------------------------
seguy2:
        mytablex.MoveNext
    Loop

End Sub

Sub carga_lineas(mytablex As ADODB.Recordset)

    Dim found As Integer

    found = formateaa("" & mytablex.Fields("fecha"), 10, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("" & mytablex.Fields("tipo"), 5, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("" & mytablex.Fields("serie"), 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("" & mytablex.Fields("numero"), 10, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("" & mytablex.Fields("tipope"), 5, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("" & mytablex.Fields("ecantidad"), 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("" & mytablex.Fields("ecosto"), 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("" & mytablex.Fields("ecostot"), 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("" & mytablex.Fields("scantidad"), 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("" & mytablex.Fields("scosto"), 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("" & mytablex.Fields("scostot"), 10, 0, 1)
    found = formateaa("", 1, 0, 0)
   
    found = formateaa("" & mytablex.Fields("tcantidad"), 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("" & mytablex.Fields("tcosto"), 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("" & mytablex.Fields("tcostot"), 10, 2, 1)

End Sub

Sub cuerpo_programa_kardex_sunat(mytablex As ADODB.Recordset)

    Dim mytabley As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset

    Dim sw       As Integer

    Dim vr

    Dim xentrada  As Double

    Dim xsalida   As Double

    Dim txentrada As Double

    Dim txsalida  As Double

    Dim xsaldo    As Double

    Dim saldoini  As Double

    Dim xcosto    As Double

    Dim xcostot   As Double

    Dim txsaldot  As Double

    Dim txcostot  As Double

    Dim xsaldot   As Double

    Dim buf       As String

    Dim found     As Integer

    Dim xsw       As Integer

    Dim xprecio   As Double

    Command1.Visible = True
    sw = 0
    Do

        If mytablex.EOF Then Exit Do
        If proveedor <> "%" Then
            found = ver_proveedor("" & mytablex.Fields("producto"))

            If found = 0 Then GoTo seguy3

        End If

        vr = DoEvents()

        If Command1.Visible = False Then Exit Do

        'xcosto = Val(Format(Val("" & mytablex.Fields("costou")), "0.000000"))
        'If conigv = "" Or conigv = "S" Then
        '   xcosto = xcosto
        'End If
        'If conigv = "N" Then
        '   If Val("" & mytablex.Fields("igv")) > 0 Then
        '      xcosto = xcosto / (1 + (Val("" & mytablex.Fields("igv")) / 100))
        '      xcosto = Val(Format(xcosto, "0.00000"))
        '   End If
        'End If

        xsw = 0
        'saldo inicial

        'AQUI VEMOS SI EXISTE MOVIMIENTO
        buf = "select * from " & xbasedatos & " where "
        buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
        buf = buf & " and local='" & extra_loquesea(local1) & "'"
        buf = buf & " and producto='" & "" & mytablex.Fields("producto") & "'"
        buf = buf & " and bodega='" & extra_loquesea(bodega) & "'"
        buf = buf & " and (acu='S' or acu='T' or acu='C' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='E' or acu='N')"
        buf = buf & " and estado='2'"
        buf = buf & " order by fecha,tipo,serie,numero"

        If mytabley.State = 1 Then mytabley.Close
        mytabley.Open buf, cn, adOpenStatic, adLockOptimistic

        If mytabley.RecordCount = 0 Then
            xsw = 1
            GoTo nomueve

        End If

        mytabley.Close
        '-----la cabecera de formato sunat
        'xcosto = 0
        contlin = 0
        contpag = 0
        cabecera_tipico "", "", "" & "" & gusuario
        found = formateaa("Periodo:", 45, 2, 0)
        found = formateaa("Ruc    :", 45, 2, 0)
        found = formateaa("Apellidos y Nombre:", 45, 2, 0)
        found = formateaa("Establecimiento:", 45, 2, 0)
        found = formateaa("Codigo Existencia:" + "" & mytablex.Fields("producto"), 45, 2, 0)
        found = formateaa("Tipo:", 45, 2, 0)
        found = formateaa("Descripcio:" + "" & mytablex.Fields("descripcio"), 45, 2, 0)
        found = formateaa("Codigo Und:", 45, 2, 0)
        found = formateaa("Metodo:", 45, 2, 0)
        contlin = 0
        contpag = 0
        cabecera_kardex_sunat
        'SALDO INICIAL
        xentrada = 0
        xsalida = 0
        xsaldo = 0
        'saldoini = 0
        'xcosto = 0
        xsaldot = 0
        txsaldot = 0
        txcostot = 0
        txentrada = 0
        txsalida = 0

        saldoini = 0
        xcosto = 0

        If mytablez.State = 1 Then mytablez.Close
        mytablez.Open "Select * from dsaldoini where local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "' and fecha='" & fechai & "'", cn, adOpenStatic, adLockOptimistic
        Do

            If mytablez.EOF Then Exit Do
            saldoini = saldoini + Val("" & mytablez.Fields("cantidad")) * Val("" & mytablez.Fields("factor"))
            xcosto = Val(Format(Val("" & mytablez.Fields("precio")), "0.000000"))
            mytablez.MoveNext
        Loop
        mytablez.Close
        xcosto = Val(Format(Val("" & mytablex.Fields("costou")), "0.000000"))

        If conigv = "" Or conigv = "S" Then
            xcosto = xcosto

        End If

        If conigv = "N" Then
            If Val("" & mytablex.Fields("igv")) > 0 Then
                xcosto = xcosto / (1 + (Val("" & mytablex.Fields("igv")) / 100))
                xcosto = Val(Format(xcosto, "0.00000"))

            End If

        End If

        xcostot = saldoini * xcosto
        xentrada = saldoini
        txsaldot = xentrada
        txcostot = txsaldot * xcosto
        txentrada = txentrada + xentrada
        'mytablez.Close
        xentrada = saldoini
        found = formateaa("" & fechai, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("", 5, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("", 5, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("", 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("", 5, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & xentrada, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & xcosto, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & xcostot, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
   
        found = formateaa("", 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        found = formateaa("", 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        found = formateaa("", 10, 0, 1)
        found = formateaa("", 1, 0, 0)
   
        found = formateaa("" & txsaldot, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & xcosto, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & txcostot, 10, 0, 1)
        found = formateaa("", 1, 2, 0)
   
        nlineas
        '-------ahora las transacciones------------
        buf = "select * from " & xbasedatos & " where "
        buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
        buf = buf & " and local='" & extra_loquesea(local1) & "'"
        buf = buf & " and producto='" & "" & mytablex.Fields("producto") & "'"
        buf = buf & " and bodega='" & extra_loquesea(bodega) & "'"
        buf = buf & " and (acu='S' or acu='T' or acu='C' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='E' or acu='N')"
        buf = buf & " and estado='2'"
        buf = buf & " order by fecha,tipo,serie,numero"

        If mytabley.State = 1 Then mytabley.Close
        mytabley.Open buf, cn, adOpenStatic, adLockOptimistic
        Do

            If mytabley.EOF Then Exit Do

            vr = DoEvents()

            If Command1.Visible = False Then Exit Do
            Command1.Caption = " " & mytabley.Fields("local") & " " & mytabley.Fields("tipo") & " " & mytabley.Fields("serie") & " " & mytabley.Fields("numero") & " " & mytabley.Fields("fecha")

            found = formateaa("" & mytabley.Fields("fecha"), 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            found = formateaa("" & mytabley.Fields("tipo"), 5, 0, 0)
            found = formateaa("", 1, 0, 0)
            found = formateaa("" & mytabley.Fields("serie"), 5, 0, 0)
            found = formateaa("", 1, 0, 0)
            found = formateaa("" & mytabley.Fields("numero"), 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            found = formateaa("" & mytabley.Fields("tipo"), 5, 0, 0)
            found = formateaa("", 1, 0, 0)
   
            '19/06/2017 kenyo NOTA DE CREDITO
            ' If "" & mytabley.Fields("acu") = "T" Or "" & mytabley.Fields("acu") = "A" Or "" & mytabley.Fields("acu") = "B" Or "" & mytabley.Fields("acu") = "C" Or "" & mytabley.Fields("acu") = "D" Or "" & mytabley.Fields("acu") = "G" Or "" & mytabley.Fields("acu") = "N"  Then
            If "" & mytabley.Fields("acu") = "T" Or "" & mytabley.Fields("acu") = "A" Or "" & mytabley.Fields("acu") = "B" Or "" & mytabley.Fields("acu") = "C" Or "" & mytabley.Fields("acu") = "D" Or "" & mytabley.Fields("acu") = "G" Or "" & mytabley.Fields("acu") = "N" Or "" & mytabley.Fields("acu") = "E" Then
  
                '19/06/2017 kenyo NOTA DE CREDITO
         
                '------------------------------------
                xsalida = Val("" & mytabley.Fields("cantidad")) * Val("" & mytabley.Fields("factor"))
                txsaldot = txsaldot - xsalida
                xcostot = xsalida * xcosto
                txcostot = txsaldot * xcosto
   
                txsalida = txsalida + xsalida
                found = formateaa("", 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                found = formateaa("", 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                found = formateaa("", 10, 0, 1)
                found = formateaa("", 1, 0, 0)
   
                found = formateaa("" & xsalida, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                found = formateaa("" & xcosto, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                found = formateaa("" & xcostot, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
   
                found = formateaa("" & txsaldot, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                found = formateaa("" & xcosto, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                found = formateaa("" & txcostot, 10, 0, 1)
                found = formateaa("", 1, 2, 0)
                nlineas
                '------------------------------------
                'xsalida = xsalida + Val("" & mytabley.Fields("cantidad")) * Val("" & mytabley.Fields("factor"))
                xsalida = Val("" & mytabley.Fields("cantidad")) * Val("" & mytabley.Fields("factor"))

            End If
   
            '19/06/2017 kenyo NOTA DE CREDITO
            'If "" & mytabley.Fields("acu") = "S" Or "" & mytabley.Fields("acu") = "J" Or "" & mytabley.Fields("acu") = "K" Or "" & mytabley.Fields("acu") = "L" Or "" & mytabley.Fields("acu") = "M" Or "" & mytabley.Fields("acu") = "P" Or "" & mytabley.Fields("acu") = "E" Then
            If "" & mytabley.Fields("acu") = "S" Or "" & mytabley.Fields("acu") = "J" Or "" & mytabley.Fields("acu") = "K" Or "" & mytabley.Fields("acu") = "L" Or "" & mytabley.Fields("acu") = "M" Or "" & mytabley.Fields("acu") = "P" Then
      
                '19/06/2017 kenyo NOTA DE CREDITO
   
                'xentrada = xentrada + Val("" & mytabley.Fields("cantidad")) * Val("" & mytabley.Fields("factor"))
                'xentrada = Val("" & mytabley.Fields("cantidad")) * Val("" & mytabley.Fields("factor"))
                '------------------------------------
                'xcosto = Val("" & mytabley.Fields("precio"))
                xentrada = Val("" & mytabley.Fields("cantidad")) * Val("" & mytabley.Fields("factor"))
                txsaldot = txsaldot + xentrada
                xcostot = xentrada * xcosto
                txcostot = txsaldot * xcosto
                txentrada = txentrada + xentrada
                found = formateaa("" & xentrada, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                found = formateaa("" & xcosto, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                found = formateaa("" & xcostot, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
   
                found = formateaa("", 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                found = formateaa("", 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                found = formateaa("", 10, 0, 1)
                found = formateaa("", 1, 0, 0)
   
                found = formateaa("" & txsaldot, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                found = formateaa("" & xcosto, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                found = formateaa("" & txcostot, 10, 0, 1)
                found = formateaa("", 1, 2, 0)
                nlineas

                '------------------------------------
            End If
   
            'nomueve:
            mytabley.MoveNext
        Loop
        'total del producto
        found = formateaa("Total", 40, 0, 0)
        found = formateaa("" & txentrada, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        found = formateaa("", 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        found = formateaa("", 10, 0, 1)
        found = formateaa("", 1, 0, 0)
   
        found = formateaa("" & txsalida, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        found = formateaa("", 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        found = formateaa("", 10, 0, 1)
        found = formateaa("", 1, 2, 0)
        'nlineas
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)
   
seguy3:
nomueve:

        mytablex.MoveNext
        '-------------------------

    Loop
    '------------- aqui debe imprimir el total de entradas /salidas

End Sub

Sub cabecera_kardex()

    Dim buf   As String

    Dim I     As Integer

    Dim found As Integer

    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("ALMACEN : " & bodega, 25, 2, 0)
    found = formateaa("FECHA INICIO :" & fechai, 40, 2, 0)
    found = formateaa("FECHA FINAL  :" & fechaf, 40, 2, 0)
    buf = String(130, "-")
   
    found = formateaa(buf, 130, 2, 0)
    found = formateaa("Tip", 3, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Ser", 4, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Numero", 11, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Fecha", 10, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Bo", 2, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Unidad", 7, 0, 0)
    found = formateaa("Fac", 5, 0, 0)
    found = formateaa("Entrada ", 11, 0, 1)
    found = formateaa("Salida ", 11, 0, 1)
    found = formateaa("Saldo ", 11, 0, 1)
    found = formateaa("Valor ", 11, 0, 1)
    found = formateaa("Total ", 11, 0, 1)
    'found = formateaa("Descripcio ", 11, 2, 0)
    found = formateaa("Doc.Cruc ", 11, 0, 0)
    found = formateaa("Nombre ", 11, 2, 0)
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)

End Sub

Sub cabecera_kardex_sunat()

    Dim buf   As String

    Dim I     As Integer

    Dim found As Integer

    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    'cabecera_tipico "", "", "" & "" & gusuario
    contpag = contpag + 1
    contlin = 0
    buf = String(140, "-")
    found = formateaa(buf, 140, 2, 0)
   
    'found = formateaa("Documento Traslado,Pago,Interno", 34, 0, 0)
    found = formateaa("", 35, 0, 0)
    found = formateaa("", 4, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("[---------", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Entradas", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("--------]", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("[----------", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Salidas", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("---------]", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("[---------", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Saldo", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("---------]", 10, 2, 1)
   
    found = formateaa("Fecha", 10, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Tipo", 5, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Serie", 5, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Numero", 10, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Tipo", 5, 0, 0)
    found = formateaa("", 1, 0, 0)
   
    found = formateaa("[Cantidad", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Costo", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Costot]", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("[Cantidad ", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Costo ", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Costot]", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("[Cantidad ", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Costo ", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Costot]", 10, 2, 1)
   
    buf = String(140, "-")
    found = formateaa(buf, 140, 2, 0)

End Sub

Sub cabecera_saldo()

    Dim buf   As String

    Dim I     As Integer

    Dim found As Integer

    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("ALMACEN : " & bodega, 25, 2, 0)
    found = formateaa("FECHA INICIO :" & fechai, 40, 2, 0)
    found = formateaa("FECHA FINAL  :" & fechaf, 40, 2, 0)
   
    '''24/08/2017  Kenyo descripcion larga en reportes ticket
    'buf = String(130, "-")
    'found = formateaa(buf, 130, 2, 0)
    If gcanti = "S" Then
        buf = String(75, "-")
        found = formateaa(buf, 75, 2, 0)
    Else
        buf = String(130, "-")
        found = formateaa(buf, 130, 2, 0)

    End If

    '''24/08/2017  Kenyo descripcion larga en reportes ticket
  
    found = formateaa("Producto ", 11, 0, 0)
    found = formateaa("Descripcio ", 41, 0, 0)
    found = formateaa("Unid", 7, 0, 0)
    found = formateaa("Fx", 5, 0, 0)
   
    found = formateaa("Cantid ", 11, 0, 1)

    If gcanti <> "S" Then
        found = formateaa("Costo ", 11, 0, 1)
        found = formateaa("Total ", 11, 2, 1)
    Else
        found = formateaa("", 1, 2, 0)

    End If
   
    '''24/08/2017  Kenyo descripcion larga en reportes ticket
    'buf = String(130, "-")
    'FOund = formateaa(buf, 130, 2, 0)
    If gcanti = "S" Then
        buf = String(75, "-")
        found = formateaa(buf, 75, 2, 0)
    Else
        buf = String(130, "-")
        found = formateaa(buf, 130, 2, 0)

    End If

    '''24/08/2017  Kenyo descripcion larga en reportes ticket

End Sub

Sub cabecera_rotacion()

    Dim buf   As String

    Dim I     As Integer

    Dim found As Integer

    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("ALMACEN : " & bodega, 25, 2, 0)
    found = formateaa("FECHA INICIO :" & fechai, 40, 2, 0)
    found = formateaa("FECHA FINAL  :" & fechaf, 40, 2, 0)
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)
    found = formateaa("Producto ", 11, 0, 0)
    found = formateaa("Descripcio ", 41, 0, 0)
    found = formateaa("Unid", 7, 0, 0)
    found = formateaa("Fx", 5, 0, 0)
   
    found = formateaa("Cantid ", 11, 0, 1)

    If gcanti <> "S" Then
        found = formateaa("Costo ", 11, 0, 1)
        found = formateaa("Total ", 11, 2, 1)
    Else
        found = formateaa("", 1, 2, 0)

    End If

    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)

End Sub

Sub cabecera_receta()

    Dim buf   As String

    Dim I     As Integer

    Dim found As Integer

    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("FECHA INICIO :" & fechai, 40, 2, 0)
    found = formateaa("FECHA FINAL  :" & fechaf, 40, 2, 0)
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)
    found = formateaa("Producto ", 11, 0, 0)
    found = formateaa("Descripcio ", 41, 0, 0)
    found = formateaa("Unid", 7, 0, 0)
    found = formateaa("Fx", 5, 0, 0)
    found = formateaa("Costo ", 11, 0, 1)
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)

End Sub

Sub cabecera_saldoini()

    Dim buf   As String

    Dim I     As Integer

    Dim found As Integer

    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("ALMACEN : " & bodega, 25, 2, 0)
    found = formateaa("FECHA INICIO :" & fechai, 40, 2, 0)
    found = formateaa("FECHA FINAL  :" & fechaf, 40, 2, 0)
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)
    found = formateaa("Producto ", 11, 0, 0)
    found = formateaa("Descripcio ", 41, 0, 0)
    found = formateaa("Unid", 7, 0, 0)
    found = formateaa("Fx", 5, 0, 0)
   
    found = formateaa("Cantid ", 11, 0, 1)
    found = formateaa("Costo ", 11, 0, 1)
    found = formateaa("Total ", 11, 2, 1)
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)

End Sub

Sub cabecera_conteo()

    Dim buf   As String

    Dim I     As Integer

    Dim found As Integer

    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("ALMACEN : " & bodega, 25, 2, 0)
    found = formateaa("FECHA INICIO :" & fechai, 40, 2, 0)
    found = formateaa("FECHA FINAL  :" & fechaf, 40, 2, 0)
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)
    found = formateaa("Producto ", 11, 0, 0)
    found = formateaa("Descripcio ", 41, 0, 0)
    found = formateaa("Unid", 7, 0, 0)
    found = formateaa("Fx", 5, 0, 0)
   
    found = formateaa("Cantid ", 11, 0, 1)
    found = formateaa("Conteo ", 11, 0, 1)
    found = formateaa("Costo ", 11, 0, 1)
    found = formateaa("Faltante ", 11, 0, 1)
    found = formateaa("Sobrante ", 11, 2, 1)
    'found = formateaa("Total ", 11, 2, 1)
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)

End Sub

Sub cabecera_lineas()

    Dim buf   As String

    Dim I     As Integer

    Dim found As Integer

    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("ALMACEN : " & bodega, 25, 2, 0)
    found = formateaa("FECHA INICIO :" & fechai, 40, 2, 0)
    found = formateaa("FECHA FINAL  :" & fechaf, 40, 2, 0)
    buf = String(140, "-")
    found = formateaa(buf, 140, 2, 0)
   
    found = formateaa("Producto ", 11, 0, 0)
    found = formateaa("Descripcio ", 51, 0, 0)
    found = formateaa("Unid", 7, 0, 0)
    found = formateaa("Fx", 5, 0, 0)
    found = formateaa("Linea", 7, 0, 0)
    found = formateaa("Stock", 9, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Tallas", 20, 0, 0)
    found = formateaa("", 1, 2, 0)

    buf = String(140, "-")
    found = formateaa(buf, 140, 2, 0)
    imprime_lineas

End Sub

Sub cuerpo_programa_lineas(mytablex As ADODB.Recordset)

    Dim mytabley  As New ADODB.Recordset

    Dim mytableyy As New ADODB.Recordset

    Dim sw        As Integer

    Dim temp      As String

    Dim xsaldo    As Double

    Dim buf       As String

    Dim sw1       As Integer

    Dim temp1     As String

    Dim buf1      As String

    Dim bufx      As String

    Dim sdx       As Double

    Dim sdx1      As Double

    Dim sdx2      As Double

    Dim sdx3      As Double

    Dim found     As Integer

    Dim vr

    sdx1 = 0
    sdx2 = 0
    sdx3 = 0
    suma1 = 0
    sw1 = 0
    'MsgBox "" & mytablex.RecordCount

    Do

        If mytablex.EOF Then Exit Do
        If proveedor <> "%" Then
            found = ver_proveedor("" & mytablex.Fields("producto"))

            If found = 0 Then GoTo seguy4

        End If

        vr = DoEvents()

        If Command1.Visible = False Then Exit Do

        xsaldo = 0
        buf = "" & mytablex.Fields("familia")

        If mytableyy.State = 1 Then mytableyy.Close
        mytableyy.Open "Select * from almacen where local='" & Trim(extra_loquesea(local1)) & "' and  producto='" & Trim("" & mytablex.Fields("producto")) & "' and bodega='" & Trim(extra_loquesea(bodega)) & "'", cn, adOpenStatic, adLockOptimistic

        If mytableyy.RecordCount > 0 Then
            sdx = Val("" & mytableyy.Fields("T1")) + Val("" & mytableyy.Fields("T2")) + Val("" & mytableyy.Fields("T3")) + Val("" & mytableyy.Fields("T4")) + Val("" & mytableyy.Fields("T5")) + Val("" & mytableyy.Fields("T6")) + Val("" & mytableyy.Fields("T7")) + Val("" & mytableyy.Fields("T8")) + Val("" & mytableyy.Fields("T9")) + Val("" & mytableyy.Fields("T10")) + Val("" & mytableyy.Fields("T11")) + Val("" & mytableyy.Fields("T12")) + Val("" & mytableyy.Fields("T13")) + Val("" & mytableyy.Fields("T14")) + Val("" & mytableyy.Fields("T15")) + Val("" & mytableyy.Fields("T16"))
            sdx3 = sdx3 + sdx
            xsaldo = Val("" & mytableyy.Fields("saldo"))

            If Combo1.Text = "TODOS" Then
                GoTo necesita

            End If

            If Combo1.Text = "SALDO>0" Then
                If sdx > 0 Then
                    GoTo necesita
                    Else: GoTo siguente

                End If

            End If

            If Combo1.Text = "SALDO<0" Then
                If sdx < 0 Then
                    GoTo necesita
                    Else: GoTo siguente

                End If

            End If

            If Combo1 = "SALDO<=0" Then
                If sdx <= 0 Then
                    GoTo necesita
                    Else: GoTo siguente

                End If

            End If

            If Combo1 = "SALDO>=0" Then
                If sdx >= 0 Then
                    GoTo necesita
                    Else: GoTo siguente

                End If

            End If

            If Combo1 = "SALDO=0" Then
                If sdx = 0 Then
                    GoTo necesita
                    Else: GoTo siguente

                End If

            End If

        End If

necesita:

        'MsgBox "abc"
        If sw = 0 Then
            sw = 1
            temp = "" & mytablex.Fields("familia")
            'found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            'found = formateaa("", 1, 1, 0)
            buf = busca_familia("" & mytablex.Fields("familia"))
            found = formateaa(buf, 30, 0, 0)
            found = formateaa("*", 1, 2, 0)
            nlineas

        End If

        If "" & mytablex.Fields("familia") <> temp Then
            temp = "" & mytablex.Fields("familia")
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("", 1, 1, 0)
            'buf = busca_familia("" & mytablex.Fields("familia"))
            'found = formateaa(buf, 30, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas

        End If

        If sw1 = 0 Then
            'found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            'found = formateaa("", 1, 1, 0)
            buf = busca_familia("" & mytablex.Fields("familia"))
            found = formateaa(buf, 30, 0, 0)
            found = formateaa("-", 1, 0, 0)
            'found = formateaa("" & mytablex.Fields("subfamilia"), 6, 0, 0)
            'found = formateaa("", 1, 1, 0)
            buf = busca_subfamilia("" & mytablex.Fields("familia"), "" & mytablex.Fields("subfamilia"))
            found = formateaa(buf, 30, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")
            sw1 = 1

        End If

        If temp1 <> "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia") Then
            'found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            'found = formateaa("", 1, 1, 0)
            buf = busca_familia("" & mytablex.Fields("familia"))
            found = formateaa(buf, 30, 0, 0)
            found = formateaa("-", 1, 0, 0)
            'found = formateaa("" & mytablex.Fields("subfamilia"), 6, 0, 0)
            'found = formateaa("", 1, 1, 0)
            buf = busca_subfamilia("" & mytablex.Fields("familia"), "" & mytablex.Fields("subfamilia"))
            found = formateaa(buf, 30, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")

        End If

        found = formateaa("" & mytablex.Fields("producto"), 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("descripcio"), 50, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("unidad"), 6, 0, 0)
        found = formateaa("x", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("factor"), 4, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("linea"), 5, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor")))
        'buf = "" & xsaldo
        found = formateaa(buf, 9, 0, 0)
        found = formateaa("", 1, 0, 0)

        'producto saldo
        If mytableyy.State = 1 Then mytableyy.Close
        mytableyy.Open "Select * from almacen where local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "'", cn, adOpenStatic, adLockOptimistic

        If mytableyy.RecordCount > 0 Then
            pone_tallas mytableyy

        End If

        mytableyy.Close
        found = formateaa("", 1, 2, 0)
        nlineas
        '---------------------------------------
siguente:
seguy4:
        mytablex.MoveNext
    Loop
    '----------------
    buf = "" & sdx3
    found = formateaa("", 79, 0, 0)
    found = formateaa(buf, 10, 2, 0)
    nlineas

End Sub

Sub pone_tallas(mytablex As ADODB.Recordset)

    Dim found As Integer

    Dim buf   As String

    buf = "" & mytablex.Fields("T1")

    If Val(buf) = 0 Then
        buf = ""

    End If

    found = formateaa(buf, 5, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "" & mytablex.Fields("T2")

    If Val(buf) = 0 Then
        buf = ""

    End If

    found = formateaa(buf, 5, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "" & mytablex.Fields("T3")

    If Val(buf) = 0 Then
        buf = ""

    End If

    found = formateaa(buf, 5, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "" & mytablex.Fields("T4")

    If Val(buf) = 0 Then
        buf = ""

    End If

    found = formateaa(buf, 5, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "" & mytablex.Fields("T5")

    If Val(buf) = 0 Then
        buf = ""

    End If

    found = formateaa(buf, 5, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "" & mytablex.Fields("T6")

    If Val(buf) = 0 Then
        buf = ""

    End If

    found = formateaa(buf, 5, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "" & mytablex.Fields("T7")

    If Val(buf) = 0 Then
        buf = ""

    End If

    found = formateaa(buf, 5, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "" & mytablex.Fields("T8")

    If Val(buf) = 0 Then
        buf = ""

    End If

    found = formateaa(buf, 5, 0, 0)
    found = formateaa("", 1, 0, 0)

    '------------------------------------
    buf = "" & mytablex.Fields("T9")

    If Val(buf) = 0 Then
        buf = ""

    End If

    found = formateaa(buf, 5, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "" & mytablex.Fields("T10")

    If Val(buf) = 0 Then
        buf = ""

    End If

    found = formateaa(buf, 5, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "" & mytablex.Fields("T11")

    If Val(buf) = 0 Then
        buf = ""

    End If

    found = formateaa(buf, 5, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "" & mytablex.Fields("T12")

    If Val(buf) = 0 Then
        buf = ""

    End If

    found = formateaa(buf, 5, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "" & mytablex.Fields("T13")

    If Val(buf) = 0 Then
        buf = ""

    End If

    found = formateaa(buf, 5, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "" & mytablex.Fields("T14")

    If Val(buf) = 0 Then
        buf = ""

    End If

    found = formateaa(buf, 5, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "" & mytablex.Fields("T15")

    If Val(buf) = 0 Then
        buf = ""

    End If

    found = formateaa(buf, 5, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "" & mytablex.Fields("T16")

    If Val(buf) = 0 Then
        buf = ""

    End If

    found = formateaa(buf, 5, 0, 0)
    found = formateaa("", 1, 0, 0)

End Sub

Sub cuerpo_programa_saldo(mytablex As ADODB.Recordset)

    Dim mytabley  As New ADODB.Recordset

    Dim mytablez  As New ADODB.Recordset

    Dim sw        As Integer

    Dim xnroitem  As Double

    Dim temp      As String

    Dim buf       As String

    Dim sw1       As Integer

    Dim temp1     As String

    Dim buf1      As String

    Dim bufx      As String

    Dim saldoini  As Double

    Dim saldoindx As Double

    Dim sdx       As Double

    Dim sdx1      As Double

    Dim sdx2      As Double

    Dim found     As Integer

    Dim xprecio   As Double

    Dim vr

    Dim buf2 As String

    saldoindx = 0
    sdx1 = 0
    sdx2 = 0
    suma1 = 0
    suma2 = 0
    sw1 = 0
    saldoini = 0
    xnroitem = 0
    Command1.Visible = True
    Do

        If mytablex.EOF Then Exit Do
        If proveedor <> "%" Then
            found = ver_proveedor("" & mytablex.Fields("producto"))

            If found = 0 Then GoTo seguy5

        End If

        vr = DoEvents()

        If Command1.Visible = False Then Exit Do

        xprecio = Val(Format(Val("" & mytablex.Fields("costop")), "0.00"))

        'verificamos que tipo de costeo
        If mytablez.State = 1 Then mytablez.Close
        If Trim(quecosto) = "COSTOULTIMO" Then
            xprecio = Val(Format(Val("" & mytablex.Fields("costou")), "0.00"))

        End If

        If Trim(quecosto) = "COSTOPROMEDIO" Then
            xprecio = Val(Format(Val("" & mytablex.Fields("costop")), "0.00"))

        End If

        If Trim(quecosto) = "PRECIOVENTA" Then
            mytablez.Open "Select * from precios where local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic

            If mytablez.RecordCount > 0 Then
                xprecio = "" & mytablez.Fields("pventa1")

            End If

            mytablez.Close

        End If

        If conigv = "" Or conigv = "S" Then
            xprecio = xprecio

        End If

        If conigv = "N" Then
            If Val("" & mytablex.Fields("igv")) > 0 Then
                xprecio = xprecio / (1 + Val("" & mytablex.Fields("igv")) / 100)
                xprecio = Val(Format(xprecio, "0.00"))

            End If

        End If

        '------------- verificamos la condicion
        'If Combo1 <> "TODOS" Then
        saldoini = 0
        buf2 = "Select * from almacen where local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "'"

        If mytablez.State = 1 Then mytablez.Close
        mytablez.Open buf2, cn, adOpenStatic, adLockOptimistic

        If mytablez.RecordCount > 0 Then
            saldoini = Val("" & mytablez.Fields("saldo"))

            If Combo1 = "SALDO>0" Then
                If saldoini > 0 Then GoTo sigueme

            End If

            If Combo1 = "SALDO>=0" Then
                If saldoini >= 0 Then GoTo sigueme

            End If

            If Combo1 = "SALDO<0" Then
                If saldoini < 0 Then GoTo sigueme

            End If

            If Combo1 = "SALDO<=0" Then
                If saldoini <= 0 Then GoTo sigueme

            End If

            If Combo1 = "SALDO=0" Then
                If saldoini = 0 Then GoTo sigueme

            End If
   
            If Combo1 = "TODOS" Then
                'MsgBox "Hola"
                'End
                GoTo sigueme

            End If

            GoTo sigueme1
            Else: GoTo sigueme1

        End If

        'End If
sigueme:
        'mytablez.Close

        buf = "" & mytablex.Fields("familia")

        If sw = 0 Then
            sw = 1
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            buf = busca_familia("" & mytablex.Fields("familia"))
            found = formateaa(buf, 30, 0, 0)
            found = formateaa("*", 1, 2, 0)
            temp = "" & mytablex.Fields("familia")
            nlineas

        End If

        If "" & mytablex.Fields("familia") <> temp Then
            temp = "" & mytablex.Fields("familia")
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            buf = busca_familia("" & mytablex.Fields("familia"))
            found = formateaa(buf, 30, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas

        End If

        If sw1 = 0 Then
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("-", 1, 0, 0)
            found = formateaa("" & mytablex.Fields("subfamilia"), 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")
            sw1 = 1

        End If

        If temp1 <> "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia") Then
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("-", 1, 0, 0)
            found = formateaa("" & mytablex.Fields("subfamilia"), 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")

        End If

        found = formateaa("" & mytablex.Fields("producto"), 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("descripcio"), 40, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("unidad"), 6, 0, 0)
        found = formateaa("x", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("factor"), 4, 0, 0)
        found = formateaa("", 1, 0, 0)
        'SALDO ALMACEN
        'saldoini = 0
        'mytablez.Seek "=", local1, "" & mytablex.Fields("producto"), extra_loquesea(bodega)
        'If Not mytablez.NoMatch Then
        '   saldoini = Val("" & mytablez.Fields("saldo"))
        'End If
        saldoindx = saldoini
        bufx = "" & saldoini

        If Val(bufx) = 0 Then
            bufx = ""

        End If

        bufx = "" & saldoindx 'calcula_saldo(saldoini, Val("" & mytablex.Fields("factor")))
        found = formateaa(bufx, 10, 0, 1)
        found = formateaa("", 1, 0, 0)

        If gcanti <> "S" Then
            buf = Format(xprecio, "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            sdx = xprecio * saldoindx
            buf = Format(sdx, "0.00")
            suma1 = suma1 + sdx
            suma2 = suma2 + saldoini
            found = formateaa(buf, 10, 0, 1)

        End If

        found = formateaa("", 1, 2, 0)
        nlineas
        xnroitem = xnroitem + 1
        '---------------------------------------
sigueme1:
seguy5:
        'mytablez.Close

        mytablex.MoveNext
    Loop
    buf = "" & xnroitem
    found = formateaa("Nro Productos " + buf, 64, 2, 0)
    nlineas
    bufx = Format(suma2, "0.00")
    found = formateaa("", 64, 0, 0)
    found = formateaa(bufx, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("", 10, 0, 1)
    found = formateaa("", 1, 0, 0)

    If gcanti <> "S" Then
        bufx = Format(suma1, "0.00")
        found = formateaa(bufx, 10, 0, 1)
        found = formateaa("", 1, 2, 0)
    Else
        found = formateaa("", 1, 2, 0)

    End If

End Sub

Sub cuerpo_programa_rotacion(mytablex As ADODB.Recordset)
    '19/06/2017 kenyo Correción  reporte productos sin rotacion. '''reporte de PRODUCTOS SIN ROTACION

    'Dim mytabley As New ADODB.Recordset
    'Dim mytablez As New ADODB.Recordset
    'Dim mytabler As New ADODB.Recordset
    '
    'Dim sw As Integer
    'Dim temp As String
    'Dim buf As String
    'Dim sw1 As Integer
    'Dim temp1 As String
    'Dim buf1 As String
    'Dim bufx As String
    'Dim saldoini As Double
    'Dim saldoindx As Double
    'Dim sdx As Double
    'Dim sdx1 As Double
    'Dim sdx2 As Double
    'Dim found As Integer
    'Dim vr
    'Dim bufecha As String
    'saldoindx = 0
    'sdx1 = 0
    'sdx2 = 0
    'suma1 = 0
    'suma2 = 0
    'sw1 = 0
    'saldoini = 0
    'Command1.Visible = True
    'Do
    'If mytablex.EOF Then Exit Do
    'If proveedor <> "%" Then
    '   found = ver_proveedor("" & mytablex.Fields("producto"))
    '   If found = 0 Then GoTo seguy5
    'End If
    '
    ''ver si existe algun movimiento
    ''entre un rango de fechas
    'If IsDate(fechari) And IsDate(fecharf) Then
    'bufecha = "Select Producto from " & xbasedatos & " where  local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "'"
    'bufecha = bufecha & " and  fecha>='" & Format(fechari, "YYYYMMDD") & "'"
    'bufecha = bufecha & " and fecha<='" & Format(fecharf, "YYYYMMDD") & "' "
    '
    'mytabler.Open bufecha, cn, adOpenStatic, adLockOptimistic
    'If mytabler.RecordCount > 0 Then
    '   mytabler.Close
    '   GoTo seguy5
    'End If
    'mytabler.Close
    'End If
    '
    '
    '
    'vr = DoEvents()
    'If Command1.Visible = False Then Exit Do
    '
    ''------------- verificamos la condicion
    ''If Combo1 <> "TODOS" Then
    'saldoini = 0
    'If mytablez.State = 1 Then mytablez.Close
    'mytablez.Open "Select * from almacen where local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "'", cn, adOpenStatic, adLockOptimistic
    'If mytablez.RecordCount > 0 Then
    '   saldoini = Val("" & mytablez.Fields("saldo"))
    '   If Combo1 = "SALDO>0" Then
    '      If saldoini > 0 Then GoTo sigueme
    '   End If
    '   If Combo1 = "SALDO>=0" Then
    '      If saldoini >= 0 Then GoTo sigueme
    '   End If
    '   If Combo1 = "SALDO<0" Then
    '      If saldoini < 0 Then GoTo sigueme
    '   End If
    '   If Combo1 = "SALDO<=0" Then
    '      If saldoini <= 0 Then GoTo sigueme
    '   End If
    ''   If Combo1 = "Con Stock Minimo" Then
    ''      If saldoini <= 0 Then GoTo sigueme
    ''   End If
    '   If Combo1 = "SALDO=0" Then
    '      If saldoini = 0 Then GoTo sigueme
    '   End If
    '   If Combo1 = "TODOS" Then
    '      'MsgBox "Hola"
    '      'End
    '      GoTo sigueme
    '   End If
    '   GoTo sigueme1
    '   Else: GoTo sigueme1
    'End If
    ''End If
    'sigueme:
    ''mytablez.Close
    '
    'buf = "" & mytablex.Fields("familia")
    'If sw = 0 Then
    '   sw = 1
    '   found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
    '   buf = busca_familia("" & mytablex.Fields("familia"))
    '   found = formateaa(buf, 30, 0, 0)
    '   found = formateaa("*", 1, 2, 0)
    '   temp = "" & mytablex.Fields("familia")
    '   nlineas
    'End If
    'If "" & mytablex.Fields("familia") <> temp Then
    '   temp = "" & mytablex.Fields("familia")
    '   found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
    '   buf = busca_familia("" & mytablex.Fields("familia"))
    '   found = formateaa(buf, 30, 0, 0)
    '   found = formateaa("", 1, 2, 0)
    '   nlineas
    'End If
    'If sw1 = 0 Then
    '   found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
    '   found = formateaa("-", 1, 0, 0)
    '   found = formateaa("" & mytablex.Fields("subfamilia"), 6, 0, 0)
    '   found = formateaa("", 1, 2, 0)
    '   nlineas
    '   temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")
    '   sw1 = 1
    'End If
    'If temp1 <> "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia") Then
    '   found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
    '   found = formateaa("-", 1, 0, 0)
    '   found = formateaa("" & mytablex.Fields("subfamilia"), 6, 0, 0)
    '   found = formateaa("", 1, 2, 0)
    '   nlineas
    '   temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")
    'End If
    'found = formateaa("" & mytablex.Fields("producto"), 10, 0, 0)
    'found = formateaa("", 1, 0, 0)
    'found = formateaa("" & mytablex.Fields("descripcio"), 40, 0, 0)
    'found = formateaa("", 1, 0, 0)
    'found = formateaa("" & mytablex.Fields("unidad"), 6, 0, 0)
    'found = formateaa("x", 1, 0, 0)
    'found = formateaa("" & mytablex.Fields("factor"), 4, 0, 0)
    'found = formateaa("", 1, 0, 0)
    ''SALDO ALMACEN
    ''saldoini = 0
    ''mytablez.Seek "=", local1, "" & mytablex.Fields("producto"), extra_loquesea(bodega)
    ''If Not mytablez.NoMatch Then
    ''   saldoini = Val("" & mytablez.Fields("saldo"))
    ''End If
    'saldoindx = saldoini
    'bufx = "" & saldoini
    'If Val(bufx) = 0 Then
    '   bufx = ""
    'End If
    'bufx = calcula_saldo(saldoini, Val("" & mytablex.Fields("factor")))
    'found = formateaa(bufx, 10, 0, 1)
    'found = formateaa("", 1, 0, 0)
    '
    'If gcanti <> "S" Then
    '   buf = "" & mytablex.Fields("costop")
    '   found = formateaa(buf, 10, 0, 1)
    '   found = formateaa("", 1, 0, 0)
    '   sdx = Val("" & mytablex.Fields("costop")) * Val(bufx)
    '   buf = "" & sdx
    '   suma1 = suma1 + sdx
    '   suma2 = suma2 + saldoini
    '   found = formateaa(buf, 10, 0, 1)
    'End If
    'found = formateaa("", 1, 2, 0)
    'nlineas
    ''---------------------------------------
    'sigueme1:
    'seguy5:
    ''mytablez.Close
    'mytablex.MoveNext
    'Loop
    'bufx = "" & suma2
    'found = formateaa("", 64, 0, 0)
    'found = formateaa(bufx, 10, 0, 1)
    'found = formateaa("", 1, 0, 0)
    'found = formateaa("", 10, 0, 1)
    'found = formateaa("", 1, 0, 0)
    'If gcanti <> "S" Then
    '   bufx = "" & suma1
    '   found = formateaa(bufx, 10, 0, 1)
    '   found = formateaa("", 1, 2, 0)
    '   Else
    '   found = formateaa("", 1, 2, 0)
    'End If

    Dim mytabley  As New ADODB.Recordset

    Dim mytablez  As New ADODB.Recordset

    Dim sw        As Integer

    Dim xnroitem  As Double

    Dim temp      As String

    Dim buf       As String

    Dim sw1       As Integer

    Dim temp1     As String

    Dim buf1      As String

    Dim bufx      As String

    Dim saldoini  As Double

    Dim saldoindx As Double

    Dim sdx       As Double

    Dim sdx1      As Double

    Dim sdx2      As Double

    Dim found     As Integer

    Dim xprecio   As Double

    Dim vr

    Dim buf2 As String

    saldoindx = 0
    sdx1 = 0
    sdx2 = 0
    suma1 = 0
    suma2 = 0
    sw1 = 0
    saldoini = 0
    xnroitem = 0
    Command1.Visible = True
    Do

        If mytablex.EOF Then Exit Do
        If proveedor <> "%" Then
            found = ver_proveedor("" & mytablex.Fields("producto"))

            If found = 0 Then GoTo seguy5

        End If

        vr = DoEvents()

        If Command1.Visible = False Then Exit Do

        xprecio = Val(Format(Val("" & mytablex.Fields("costop")), "0.00"))

        'verificamos que tipo de costeo
        If mytablez.State = 1 Then mytablez.Close
        If Trim(quecosto) = "COSTOULTIMO" Then
            xprecio = Val(Format(Val("" & mytablex.Fields("costou")), "0.00"))

        End If

        If Trim(quecosto) = "COSTOPROMEDIO" Then
            xprecio = Val(Format(Val("" & mytablex.Fields("costop")), "0.00"))

        End If

        If Trim(quecosto) = "PRECIOVENTA" Then
            mytablez.Open "Select * from precios where local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic

            If mytablez.RecordCount > 0 Then
                xprecio = "" & mytablez.Fields("pventa1")

            End If

            mytablez.Close

        End If

        If conigv = "" Or conigv = "S" Then
            xprecio = xprecio

        End If

        If conigv = "N" Then
            If Val("" & mytablex.Fields("igv")) > 0 Then
                xprecio = xprecio / (1 + Val("" & mytablex.Fields("igv")) / 100)
                xprecio = Val(Format(xprecio, "0.00"))

            End If

        End If

        '------------- verificamos la condicion
        'If Combo1 <> "TODOS" Then
        saldoini = 0
        buf2 = "Select * from almacen where local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "'"

        If mytablez.State = 1 Then mytablez.Close
        mytablez.Open buf2, cn, adOpenStatic, adLockOptimistic

        If mytablez.RecordCount > 0 Then
            saldoini = Val("" & mytablez.Fields("saldo"))

            If Combo1 = "SALDO>0" Then
                If saldoini > 0 Then GoTo sigueme

            End If

            If Combo1 = "SALDO>=0" Then
                If saldoini >= 0 Then GoTo sigueme

            End If

            If Combo1 = "SALDO<0" Then
                If saldoini < 0 Then GoTo sigueme

            End If

            If Combo1 = "SALDO<=0" Then
                If saldoini <= 0 Then GoTo sigueme

            End If

            If Combo1 = "SALDO=0" Then
                If saldoini = 0 Then GoTo sigueme

            End If
   
            If Combo1 = "TODOS" Then
                'MsgBox "Hola"
                'End
                GoTo sigueme

            End If

            GoTo sigueme1
            Else: GoTo sigueme1

        End If

        'End If
sigueme:
        'mytablez.Close

        buf = "" & mytablex.Fields("familia")

        If sw = 0 Then
            sw = 1
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            buf = busca_familia("" & mytablex.Fields("familia"))
            found = formateaa(buf, 30, 0, 0)
            found = formateaa("*", 1, 2, 0)
            temp = "" & mytablex.Fields("familia")
            nlineas

        End If

        If "" & mytablex.Fields("familia") <> temp Then
            temp = "" & mytablex.Fields("familia")
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            buf = busca_familia("" & mytablex.Fields("familia"))
            found = formateaa(buf, 30, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas

        End If

        If sw1 = 0 Then
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("-", 1, 0, 0)
            found = formateaa("" & mytablex.Fields("subfamilia"), 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")
            sw1 = 1

        End If

        If temp1 <> "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia") Then
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("-", 1, 0, 0)
            found = formateaa("" & mytablex.Fields("subfamilia"), 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")

        End If

        found = formateaa("" & mytablex.Fields("producto"), 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("descripcio"), 40, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("unidad"), 6, 0, 0)
        found = formateaa("x", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("factor"), 4, 0, 0)
        found = formateaa("", 1, 0, 0)
        'SALDO ALMACEN
        'saldoini = 0
        'mytablez.Seek "=", local1, "" & mytablex.Fields("producto"), extra_loquesea(bodega)
        'If Not mytablez.NoMatch Then
        '   saldoini = Val("" & mytablez.Fields("saldo"))
        'End If
        saldoindx = saldoini
        bufx = "" & saldoini

        If Val(bufx) = 0 Then
            bufx = ""

        End If

        bufx = "" & saldoindx 'calcula_saldo(saldoini, Val("" & mytablex.Fields("factor")))
        found = formateaa(bufx, 10, 0, 1)
        found = formateaa("", 1, 0, 0)

        If gcanti <> "S" Then
            buf = Format(xprecio, "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            sdx = xprecio * saldoindx
            buf = Format(sdx, "0.00")
            suma1 = suma1 + sdx
            suma2 = suma2 + saldoini
            found = formateaa(buf, 10, 0, 1)

        End If

        found = formateaa("", 1, 2, 0)
        nlineas
        xnroitem = xnroitem + 1
        '---------------------------------------
sigueme1:
seguy5:
        'mytablez.Close

        mytablex.MoveNext
    Loop
    buf = "" & xnroitem
    found = formateaa("Nro Productos " + buf, 64, 2, 0)
    nlineas
    bufx = Format(suma2, "0.00")
    found = formateaa("", 64, 0, 0)
    found = formateaa(bufx, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("", 10, 0, 1)
    found = formateaa("", 1, 0, 0)

    If gcanti <> "S" Then
        bufx = Format(suma1, "0.00")
        found = formateaa(bufx, 10, 0, 1)
        found = formateaa("", 1, 2, 0)
    Else
        found = formateaa("", 1, 2, 0)

    End If

    '19/06/2017 kenyo Correción  reporte productos sin rotacion. '''reporte de PRODUCTOS SIN ROTACION

End Sub

Sub cuerpo_programa_receta(mytablex As ADODB.Recordset)

    Dim mytabley  As New ADODB.Recordset

    Dim mytablez  As New ADODB.Recordset

    Dim mytabler  As New ADODB.Recordset

    Dim sw        As Integer

    Dim temp      As String

    Dim buf       As String

    Dim sw1       As Integer

    Dim temp1     As String

    Dim buf1      As String

    Dim bufx      As String

    Dim saldoini  As Double

    Dim saldoindx As Double

    Dim sdx       As Double

    Dim sdx1      As Double

    Dim sdx2      As Double

    Dim found     As Integer

    Dim vr

    Dim bufecha As String

    saldoindx = 0
    sdx1 = 0
    sdx2 = 0
    suma1 = 0
    suma2 = 0
    sw1 = 0
    saldoini = 0
    Command1.Visible = True
    Do

        If mytablex.EOF Then Exit Do
        If mytabler.State = 1 Then
            mytabler.Close
            Set mytabler = Nothing

        End If

        mytabler.Open "Select * from receta where producto='" & "" & mytablex.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic

        If mytabler.RecordCount = 0 Then
            mytabler.Close
            GoTo seguy58
            Exit Sub

        End If

        If proveedor <> "%" Then
            found = ver_proveedor("" & mytablex.Fields("producto"))

            If found = 0 Then GoTo seguy58

        End If

        'ver si existe algun movimiento
        'entre un rango de fechas
        vr = DoEvents()

        If Command1.Visible = False Then Exit Do

        '------------- verificamos la condicion
        'If Combo1 <> "TODOS" Then
        saldoini = 0

        If mytablez.State = 1 Then mytablez.Close
        mytablez.Open "Select * from almacen where local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "'", cn, adOpenStatic, adLockOptimistic

        If mytablez.RecordCount > 0 Then
            saldoini = Val("" & mytablez.Fields("saldo"))

            If Combo1 = "SALDO>0" Then
                If saldoini > 0 Then GoTo ssigueme

            End If

            If Combo1 = "SALDO>=0" Then
                If saldoini >= 0 Then GoTo ssigueme

            End If

            If Combo1 = "SALDO<0" Then
                If saldoini < 0 Then GoTo ssigueme

            End If

            If Combo1 = "SALDO<=0" Then
                If saldoini <= 0 Then GoTo ssigueme

            End If

            If Combo1 = "SALDO=0" Then
                If saldoini = 0 Then GoTo ssigueme

            End If

            If Combo1 = "TODOS" Then
                'MsgBox "Hola"
                'End
                GoTo ssigueme

            End If

            GoTo ssigueme1
            Else: GoTo ssigueme1

        End If

        'End If
ssigueme:
        'mytablez.Close

        buf = "" & mytablex.Fields("familia")

        If sw = 0 Then
            sw = 1
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            buf = busca_familia("" & mytablex.Fields("familia"))
            found = formateaa(buf, 30, 0, 0)
            found = formateaa("*", 1, 2, 0)
            temp = "" & mytablex.Fields("familia")
            nlineas

        End If

        If "" & mytablex.Fields("familia") <> temp Then
            temp = "" & mytablex.Fields("familia")
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            buf = busca_familia("" & mytablex.Fields("familia"))
            found = formateaa(buf, 30, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas

        End If

        If sw1 = 0 Then
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("-", 1, 0, 0)
            found = formateaa("" & mytablex.Fields("subfamilia"), 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")
            sw1 = 1

        End If

        If temp1 <> "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia") Then
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("-", 1, 0, 0)
            found = formateaa("" & mytablex.Fields("subfamilia"), 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")

        End If

        found = formateaa("" & mytablex.Fields("producto"), 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("descripcio"), 40, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("unidad"), 6, 0, 0)
        found = formateaa("x", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("factor"), 4, 0, 0)
        found = formateaa("", 1, 2, 0)
        nlineas

        If mytabler.State = 1 Then
            mytabler.Close
            Set mytabler = Nothing

        End If

        mytabler.Open "Select * from receta where producto='" & "" & mytablex.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic
        'If mytabler.RecordCount > 0 Then
        Do

            If mytabler.EOF Then Exit Do
            found = formateaa("*" & mytabler.Fields("productoi"), 4, 0, 0)
            found = formateaa("", 1, 0, 0)
            found = formateaa("" & mytabler.Fields("descripcio"), 20, 0, 0)
            found = formateaa("", 1, 0, 0)
            found = formateaa("" & mytabler.Fields("cantidad"), 7, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            mytabler.MoveNext
        Loop
        'End If
        mytabler.Close
        found = formateaa("", 1, 2, 0)
        nlineas
ssigueme1:
seguy58:
        mytablex.MoveNext
    Loop

End Sub

Sub cuerpo_programa_saldoini(mytablex As ADODB.Recordset)

    Dim mytabley  As New ADODB.Recordset

    Dim mytablez  As New ADODB.Recordset

    Dim indx      As Double

    Dim mcanti    As Double

    Dim mmx       As Double

    Dim sw        As Integer

    Dim temp      As String

    Dim buf       As String

    Dim sw1       As Integer

    Dim temp1     As String

    Dim buf1      As String

    Dim bufx      As String

    Dim saldoini  As Double

    Dim saldoindx As Double

    Dim sdx       As Double

    Dim sdx1      As Double

    Dim sdx2      As Double

    Dim sdx3      As Double

    Dim found     As Integer

    Dim vr

    On Error GoTo cmd7878_err

    saldoindx = 0
    sdx1 = 0
    sdx2 = 0
    sdx3 = 0
    suma1 = 0
    suma2 = 0
    sw1 = 0
    mcanti = 0
    Command1.Visible = True
    indx = 0
    Do

        If mytablex.EOF Then Exit Do
        If proveedor <> "%" Then
            found = ver_proveedor("" & mytablex.Fields("producto"))

            If found = 0 Then GoTo seguy6

        End If

        vr = DoEvents()

        If Command1.Visible = False Then Exit Do
        Command1.Caption = "" & indx
        indx = indx + 1

        mmx = 0

        If mytablez.State = 1 Then mytablez.Close
        mytablez.Open "Select * from " & xbasedatos & " where local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "' and fecha='" & Format(fechai, "YYYYMMDD") & "' and l1='S' and acu='S' ", cn, adOpenStatic, adLockOptimistic

        If mytablez.RecordCount > 0 Then
            Do

                If mytablez.EOF Then Exit Do
                mmx = mmx + Val("" & mytablez.Fields("cantidad")) * Val("" & mytablez.Fields("factor"))
                mytablez.MoveNext
            Loop

        End If

        mytablez.Close
   
        If Combo1 = "SALDO>0" Then
            If mmx > 0 Then GoTo sigueme11

        End If

        If Combo1 = "SALDO>=0" Then
            If mmx >= 0 Then GoTo sigueme11

        End If

        If Combo1 = "SALDO<0" Then
            If mmx < 0 Then GoTo sigueme11

        End If

        If Combo1 = "SALDO<=0" Then
            If mmx <= 0 Then GoTo sigueme11

        End If

        If Combo1 = "SALDO=0" Then
            If mmx = 0 Then GoTo sigueme11

        End If

        If Combo1 = "TODOS" Then
            GoTo sigueme11

        End If

        GoTo sigueme21
   
sigueme11:
        'mytablez.Close

        buf = "" & mytablex.Fields("familia")

        If sw = 0 Then
            sw = 1
            temp = "" & mytablex.Fields("familia")
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("*", 1, 2, 0)
            nlineas

        End If

        If "" & mytablex.Fields("familia") <> temp Then
            temp = "" & mytablex.Fields("familia")
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas

        End If

        If sw1 = 0 Then
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("-", 1, 0, 0)
            found = formateaa("" & mytablex.Fields("subfamilia"), 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")
            sw1 = 1

        End If

        If temp1 <> "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia") Then
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("-", 1, 0, 0)
            found = formateaa("" & mytablex.Fields("subfamilia"), 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")

        End If

        found = formateaa("" & mytablex.Fields("producto"), 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("descripcio"), 40, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("unidad"), 6, 0, 0)
        found = formateaa("x", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("factor"), 4, 0, 0)
        found = formateaa("", 1, 0, 0)
        'SALDO INICIAL
        saldoini = mmx
        'saldoini = 0
        'If mytablez.State = 1 Then mytablez.Close
        'mytablez.Open "Select * from saldoini where local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "' and fecha='" & fechai & "'", cn, adOpenStatic, adLockOptimistic
        'If mytablez.RecordCount > 0 Then
        '   saldoini = Val("" & mytablez.Fields("cantidad")) * Val("" & mytablez.Fields("factor")) + Val("" & mytablez.Fields("cantidad1"))
        'End If
        'mytablez.Close
        suma2 = suma2 + saldoini
        saldoindx = saldoini
        bufx = "" & saldoini
        buf = calcula_saldo(saldoindx, Val("" & mytablex.Fields("factor")))

        If Val(bufx) = 0 Then
            bufx = ""

        End If

        'MsgBox saldoindx
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)

        buf = "" & mytablex.Fields("costop")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        sdx = Val("" & mytablex.Fields("costop")) * Val(bufx)
        buf = "" & sdx
        suma1 = suma1 + sdx
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 2, 0)
        nlineas
        '---------------------------------------
sigueme21:
seguy6:
        mytablex.MoveNext
    Loop
    found = formateaa(" Total Inventario ", 64, 0, 0)
    bufx = "" & suma2
    found = formateaa(bufx, 10, 0, 1)
    found = formateaa("", 12, 0, 0)
    bufx = Format(Val("" & suma1), "0.00")
    found = formateaa(bufx, 10, 2, 1)

    Exit Sub
cmd7878_err:
    MsgBox "Aviso en cuerpo programa saldo ini ", 48, "Aviso"
    Exit Sub

End Sub

Sub cuerpo_programa_conteo(mytablex As ADODB.Recordset)

    Dim mytabley  As New ADODB.Recordset

    Dim mytablez  As New ADODB.Recordset

    Dim mcanti    As Double

    Dim mmx       As Double

    Dim sw        As Integer

    Dim temp      As String

    Dim buf       As String

    Dim sw1       As Integer

    Dim temp1     As String

    Dim buf1      As String

    Dim bufx      As String

    Dim saldoini  As Double

    Dim saldoant  As Double

    Dim saldoindx As Double

    Dim sdx       As Double

    Dim sdx1      As Double

    Dim sdx2      As Double

    Dim sdx3      As Double

    Dim found     As Integer

    Dim sobrante  As Double

    Dim faltante  As Double

    Dim vr

    Dim xcosto As Double

    saldoindx = 0
    sdx1 = 0
    sdx2 = 0
    sdx3 = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0
    sw1 = 0
    mcanti = 0
    Command1.Visible = True
    Do

        If mytablex.EOF Then Exit Do
        If proveedor <> "%" Then
            found = ver_proveedor("" & mytablex.Fields("producto"))

            If found = 0 Then GoTo seguy7

        End If

        vr = DoEvents()

        If Command1.Visible = False Then Exit Do

        mmx = 0

        If mytablez.State = 1 Then mytablez.Close
        mytablez.Open "Select * from dsaldoini where local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "' and fecha='" & Format(fechai, "YYYYMMDD") & "'", cn, adOpenStatic, adLockOptimistic

        If mytablez.RecordCount > 0 Then
            mmx = Val("" & mytablez.Fields("cantidad")) * Val("" & mytablez.Fields("factor"))

            If Combo1 = "SALDO>0" Then
                If mmx > 0 Then GoTo sigueme111

            End If

            If Combo1 = "SALDO>=0" Then
                If mmx >= 0 Then GoTo sigueme111

            End If

            If Combo1 = "SALDO<0" Then
                If mmx < 0 Then GoTo sigueme111

            End If

            If Combo1 = "SALDO<=0" Then
                If mmx <= 0 Then GoTo sigueme111

            End If

            If Combo1 = "SALDO=0" Then
                If mmx = 0 Then GoTo sigueme111

            End If

            GoTo sigueme121

        End If

sigueme111:
        'mytablez.Close
        buf = "" & mytablex.Fields("familia")

        If sw = 0 Then
            sw = 1
            temp = "" & mytablex.Fields("familia")
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("*", 1, 2, 0)
            nlineas

        End If

        If "" & mytablex.Fields("familia") <> temp Then
            temp = "" & mytablex.Fields("familia")
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas

        End If

        If sw1 = 0 Then
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("-", 1, 0, 0)
            found = formateaa("" & mytablex.Fields("subfamilia"), 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")
            sw1 = 1

        End If

        If temp1 <> "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia") Then
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("-", 1, 0, 0)
            found = formateaa("" & mytablex.Fields("subfamilia"), 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")

        End If

        found = formateaa("" & mytablex.Fields("producto"), 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("descripcio"), 40, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("unidad"), 6, 0, 0)
        found = formateaa("x", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("factor"), 4, 0, 0)
        found = formateaa("", 1, 0, 0)
        'SALDO INICIAL
        saldoini = 0
        saldoant = 0
        xcosto = 0

        If mytablez.State = 1 Then mytablez.Close
        mytablez.Open "Select * from dsaldoini where local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "' and fecha='" & Format(fechai, "YYYYMMDD") & "'", cn, adOpenStatic, adLockOptimistic
        Do

            If mytablez.EOF Then Exit Do
            saldoini = saldoini + Val("" & mytablez.Fields("cantidad")) * Val("" & mytablez.Fields("factor"))
            saldoant = saldoant + Val("" & mytablez.Fields("saldoant")) * Val("" & mytablez.Fields("factor"))
            xcosto = Val(Format(Val("" & mytablez.Fields("precio")), "0.00000"))
            mytablez.MoveNext
        Loop
        mytablez.Close
        'MsgBox ""
        suma1 = suma1 + saldoini
        buf = "" & saldoini 'calcula_saldo(saldoini, Val("" & mytablex.Fields("factor")))

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)

        suma2 = suma2 + saldoant
        buf = "" & saldoant 'calcula_saldo(saldoant, Val("" & mytablex.Fields("factor")))

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)

        buf = "" & xcosto
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        sobrante = saldoini
        faltante = saldoini
        sdx = Val(Format(xcosto * sobrante, "0.00000"))
        buf = "" & sdx
        suma3 = suma3 + sdx
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)

        sdx = Val(Format(xcosto * faltante, "0.00"))
        buf = Format(sdx, "0.00")
        suma4 = suma4 + sdx
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 2, 0)

        nlineas
        '---------------------------------------
sigueme121:
seguy7:
        mytablex.MoveNext
    Loop
    found = formateaa(" Total ", 64, 0, 0)
    bufx = "" & suma1
    found = formateaa(bufx, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    bufx = "" & suma2
    found = formateaa(bufx, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("", 11, 0, 0)

    suma3 = Val(Format(suma3, "0.00"))
    suma4 = Val(Format(suma4, "0.00"))
    bufx = "" & suma3
    found = formateaa(bufx, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    bufx = "" & suma4
    found = formateaa(bufx, 10, 0, 1)
    found = formateaa("", 1, 2, 0)

End Sub

Sub nlineas()
    contlin = contlin + 1

    If contlin > Val(nrolineas) Then
        If opcion2 = "1" Then
            cabecera_kardex

        End If

        If opcion2 = "2" Then
            cabecera_saldoini

        End If

        If opcion2 = "3" Then
            cabecera_lineas

        End If

        If opcion2 = "4" Then
            cabecera_saldo

        End If

        If opcion2 = "6" Then
            cabecera_saldo1

        End If

        If opcion2 = "7" Then
            cabecera_saldo2

        End If

        If opcion2 = "8" Then
            cabecera_saldo8

        End If

    End If

End Sub

Sub reporte_lineas()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    Dim sdx3     As Double

    sdx3 = 0

    found = sql_producto(mytablex)

    If found = 0 Then
        mytablex.Close
    
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    Command1.Visible = True
    '------------------------------------
    cabecera_lineas
    cuerpo_programa_lineas mytablex
    Command1.Visible = False
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
     
    'genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1
    found = valida_wordpad(FileName)
    Command1.Visible = False

End Sub

Sub imprime_lineas()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "Select * from linea", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do

        found = formateaa("", 62, 0, 0)
        'found = formateaa("Saldo", 10, 0, 0)
        found = formateaa("" & mytablex.Fields("Linea"), 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("Descripcio"), 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("", 10, 0, 0)

        found = formateaa("" & mytablex.Fields("t1"), 5, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("t2"), 5, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("t3"), 5, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("t4"), 5, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("t5"), 5, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("t6"), 5, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("t7"), 5, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("t8"), 5, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("t9"), 5, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("t10"), 5, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("t11"), 5, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("t12"), 5, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("t13"), 5, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("t14"), 5, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("t15"), 5, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("t16"), 5, 0, 0)
        found = formateaa("", 1, 2, 0)
        nlineas
        mytablex.MoveNext
    Loop
    '------------------------------------- ------------
    mytablex.Close
 
End Sub

Function busca_familia(buf As String) As String

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "Select * from familia where familia='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_familia = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close
 
End Function

Function busca_paridad() As Double

    Dim sdx As Double

    sdx = 1

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from parame where codigo='01' ", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        sdx = Val("" & mytablex.Fields("parivta"))

        If sdx <= 0 Then
            sdx = 1

        End If

    End If

    busca_paridad = sdx
    mytablex.Close

End Function

Function busca_subfamilia(buf As String, buf1 As String) As String

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from subfamil where  familia='" & buf & "' and subfamilia='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_subfamilia = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close

End Function

Sub cuerpo_programa_saldoexcell(mytablex As ADODB.Recordset)

    Dim mytabley  As New ADODB.Recordset

    Dim mytablez  As New ADODB.Recordset

    Dim sw        As Integer

    Dim temp      As String

    Dim buf       As String

    Dim sw1       As Integer

    Dim temp1     As String

    Dim buf1      As String

    Dim bufx      As String

    Dim saldoini  As Double

    Dim saldoindx As Double

    Dim sdx       As Double

    Dim sdx1      As Double

    Dim sdx2      As Double

    Dim found     As Integer

    Dim v, h As Double

    Dim xprecio As Double

    Dim vr

    saldoindx = 0
    sdx1 = 0
    sdx2 = 0
    suma1 = 0
    suma2 = 0
    sw1 = 0

    ''''19/09/2017 kenyo Mejora Reporte Saldo Actual
    Dim Heading(12) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    On Error GoTo cmd5612_err

    Heading(1) = "Familia"
    Heading(2) = "Producto"
    Heading(3) = "Descripcio"
    Heading(4) = "Unidad"
    Heading(5) = "Factor"
    Heading(6) = "Saldo"
    Heading(7) = "Costo"
    Heading(8) = "Total"
    Heading(9) = "Subfamilia"
    Heading(10) = "Categoria"
    
    '''21/09/2017 kenyo Reporte de Stock minimo Ticket
    Heading(11) = "Minimo"
    '''21/09/2017 kenyo Reporte de Stock minimo Ticket
    
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_Excel(11, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook

    ''''09/10/2017 kenyo Testing Reportes
    objExcel.ActiveSheet.Cells(1, 3) = "                                                  REPORTE DE SALDO ACTUAL                                           "
    objExcel.ActiveSheet.Cells(1, 3).Font.bold = True
    objExcel.ActiveSheet.Cells(1, 3).Font.Size = 14
    objExcel.ActiveSheet.Cells(1, 3).Font.color = RGB(0, 112, 184)
    ''''09/10/2017 kenyo Testing Reportes

    ''''19/09/2017 kenyo Mejora Reporte Saldo Actual
    v = 4
    ''''19/09/2017 kenyo Mejora Reporte Saldo Actual

    h = 1

    Command1.Visible = True
    Do

        If mytablex.EOF Then Exit Do
        If proveedor <> "%" Then
            found = ver_proveedor("" & mytablex.Fields("producto"))

            If found = 0 Then GoTo seguy8

        End If

        vr = DoEvents()

        If Command1.Visible = False Then Exit Do

        '------------- verificamos la condicion
        'If Combo1 <> "TODOS" Then

        xprecio = Val(Format(Val("" & mytablex.Fields("costop")), "0.00000"))

        'verificamos que tipo de costeo
        If mytablez.State = 1 Then mytablez.Close
        If Trim(quecosto) = "COSTOULTIMO" Then
            xprecio = Val(Format(Val("" & mytablex.Fields("costou")), "0.00000"))

        End If

        If Trim(quecosto) = "COSTOPROMEDIO" Then
            xprecio = Val(Format(Val("" & mytablex.Fields("costop")), "0.00000"))

        End If

        If Trim(quecosto) = "PRECIOVENTA" Then
            mytablez.Open "Select * from precios where local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic

            If mytablez.RecordCount > 0 Then
                xprecio = Val(Format("" & mytablez.Fields("pventa1"), "0.00"))

            End If

            mytablez.Close

        End If

        If conigv = "" Or conigv = "S" Then
            xprecio = xprecio

        End If

        If conigv = "N" Then
            If Val("" & mytablex.Fields("igv")) > 0 Then
                xprecio = xprecio / (1 + Val("" & mytablex.Fields("igv")) / 100)
                xprecio = Val(Format(xprecio, "0.00000"))

            End If

        End If

        saldoini = 0

        If mytablez.State = 1 Then mytablez.Close
        mytablez.Open "Select * from almacen where local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "'", cn, adOpenStatic, adLockOptimistic

        If mytablez.RecordCount > 0 Then
            saldoini = Val("" & mytablez.Fields("saldo"))

            If Combo1 = "SALDO>0" Then
                If saldoini > 0 Then GoTo sigueme

            End If

            If Combo1 = "SALDO>=0" Then
                If saldoini >= 0 Then GoTo sigueme

            End If

            If Combo1 = "SALDO<0" Then
                If saldoini < 0 Then GoTo sigueme

            End If

            If Combo1 = "SALDO<=0" Then
                If saldoini <= 0 Then GoTo sigueme

            End If

            If Combo1 = "SALDO=0" Then
                If saldoini = 0 Then GoTo sigueme
     
                '  MsgBox ("Aqui")
   
                'If saldoini < (Val("" & mytablez.Fields("minimo"))) Then GoTo sigueme
      
                'If saldoini = 0 Then GoTo sigueme
                'If saldoini = 0 Then GoTo sigueme
                ' If mytablesm.State = 1 Then mytablesm.Close
                'mytablesm.Open "select a.*   from producto AS P, ALMACEN AS A WHERE P.PRODUCTO=A.PRODUCTO AND A.SALDO < P.MINIMO AND P.MINIMO>0  where local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "'", cn, adOpenStatic, adLockOptimistic
 
            End If

            If Combo1 = "TODOS" Then
                GoTo sigueme

            End If

            GoTo sigueme1
        Else
            GoTo sigueme1

        End If

        'End If
sigueme:
        'mytablez.Close

        '''09/10/2017 kenyo Testing Reportes
        'objExcel.ActiveSheet.Cells(4, 1) = " "
        '''09/10/2017 kenyo Testing Reportes

        buf = "" & mytablex.Fields("familia")

        If sw = 0 Then
            sw = 1
            '  objExcel.ActiveSheet.Cells(v, 1) = " "
            '''09/10/2017 kenyo Testing Reportes
            v = v + 1
            objExcel.ActiveSheet.Cells(v - 1, h) = " "
            objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("familia")
            objExcel.ActiveSheet.Cells(v, h).Font.bold = True
            objExcel.ActiveSheet.Cells(v, h).Font.color = RGB(62, 95, 138)
            objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_familia("" & mytablex.Fields("familia"))
            objExcel.ActiveSheet.Cells(v, h + 1).Font.bold = True
            objExcel.ActiveSheet.Cells(v, h + 1).Font.color = RGB(62, 95, 138)
            '''09/10/2017 kenyo Testing Reportes
            objExcel.ActiveSheet.Cells(v, h + 2) = ""
            objExcel.ActiveSheet.Cells(v, h + 3) = ""
            objExcel.ActiveSheet.Cells(v, h + 4) = ""
            objExcel.ActiveSheet.Cells(v, h + 5) = ""
            objExcel.ActiveSheet.Cells(v, h + 6) = ""
            objExcel.ActiveSheet.Cells(v, h + 7) = ""
            v = v + 1
            temp = "" & mytablex.Fields("familia")

        End If

        If "" & mytablex.Fields("familia") <> temp Then
            temp = "" & mytablex.Fields("familia")
            v = v + 1
            objExcel.ActiveSheet.Cells(v - 1, h) = " "
            objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("familia")
            objExcel.ActiveSheet.Cells(v, h).Font.bold = True
            objExcel.ActiveSheet.Cells(v, h).Font.color = RGB(62, 95, 138)
            objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_familia("" & mytablex.Fields("familia"))
            objExcel.ActiveSheet.Cells(v, h + 1).Font.bold = True
            objExcel.ActiveSheet.Cells(v, h + 1).Font.color = RGB(62, 95, 138)
                       
            objExcel.ActiveSheet.Cells(v, h + 2) = ""
            objExcel.ActiveSheet.Cells(v, h + 3) = ""
            objExcel.ActiveSheet.Cells(v, h + 4) = ""
            objExcel.ActiveSheet.Cells(v, h + 5) = ""
            objExcel.ActiveSheet.Cells(v, h + 6) = ""
            objExcel.ActiveSheet.Cells(v, h + 7) = ""
            v = v + 1

        End If
   
        ''''19/09/2017 kenyo Mejora Reporte Saldo Actual
        'If sw1 = 0 Then
        '         objExcel.ActiveSheet.Cells(v, h) = "" & mytablex.Fields("familia")
        '            objExcel.ActiveSheet.Cells(v, h + 1) = "" & mytablex.Fields("subfamilia")
        '               objExcel.ActiveSheet.Cells(v, h + 2) = ""
        '            objExcel.ActiveSheet.Cells(v, h + 3) = ""
        '            objExcel.ActiveSheet.Cells(v, h + 4) = ""
        '            objExcel.ActiveSheet.Cells(v, h + 5) = ""
        '            objExcel.ActiveSheet.Cells(v, h + 6) = ""
        '            objExcel.ActiveSheet.Cells(v, h + 7) = ""
        '            v = v + 1
        '   temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")
        '   sw1 = 1
        'End If

        'If temp1 <> "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia") Then
        '         objExcel.ActiveSheet.Cells(v, h) = "" & mytablex.Fields("familia")
        '            objExcel.ActiveSheet.Cells(v, h + 1) = "" & mytablex.Fields("subfamilia")
        '            objExcel.ActiveSheet.Cells(v, h + 2) = ""
        '            objExcel.ActiveSheet.Cells(v, h + 3) = ""
        '            objExcel.ActiveSheet.Cells(v, h + 4) = ""
        '            objExcel.ActiveSheet.Cells(v, h + 5) = ""
        '            objExcel.ActiveSheet.Cells(v, h + 6) = ""
        '            objExcel.ActiveSheet.Cells(v, h + 7) = ""
        '            v = v + 1
        '   temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")
        'End If
        ''''19/09/2017 kenyo Mejora Reporte Saldo Actual

        'SALDO ALMACEN
        saldoini = 0

        If mytablez.State = 1 Then mytablez.Close
        mytablez.Open "Select * from almacen where local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "'", cn, adOpenStatic, adLockOptimistic

        If mytablez.RecordCount > 0 Then
            saldoini = Val("" & mytablez.Fields("saldo"))

        End If

        mytablez.Close
        saldoindx = saldoini
        ''''19/09/2017 kenyo Mejora Reporte Saldo Actual

        objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("familia")
        objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("producto")
        objExcel.ActiveSheet.Cells(v, h + 2) = "" & mytablex.Fields("descripcio")
        objExcel.ActiveSheet.Cells(v, h + 3) = "" & mytablex.Fields("unidad")
        objExcel.ActiveSheet.Cells(v, h + 4) = "" & mytablex.Fields("factor")

        objExcel.ActiveSheet.Cells(v, h + 5) = saldoindx
        objExcel.ActiveSheet.Cells(v, h + 5).Font.bold = True
        objExcel.ActiveSheet.Cells(v, h + 6) = xprecio
        sdx = xprecio * saldoindx
        suma1 = suma1 + sdx
        suma2 = suma2 + saldoini
        objExcel.ActiveSheet.Cells(v, h + 7) = sdx
        objExcel.ActiveSheet.Cells(v, h + 9) = ""
        objExcel.ActiveSheet.Cells(v, h + 9) = "" & mytablex.Fields("categoria")
        objExcel.ActiveSheet.Cells(v, h + 8) = "" & mytablex.Fields("subfamilia")

        '''21/09/2017 kenyo Reporte de Stock minimo Ticket
        objExcel.ActiveSheet.Cells(v, h + 10) = "" & mytablex.Fields("minimo")
        '''21/09/2017 kenyo Reporte de Stock minimo Ticket

        v = v + 1
        '--------------------------------------
sigueme1:
seguy8:
        mytablex.MoveNext
    Loop

    objExcel.ActiveSheet.Cells(v, h + 1) = ""
    objExcel.ActiveSheet.Cells(v, h + 2) = ""
    objExcel.ActiveSheet.Cells(v, h + 3) = "GRAN TOTAL >>>"
    objExcel.ActiveSheet.Cells(v, h + 4) = ""
    objExcel.ActiveSheet.Cells(v, h + 5) = "" & suma2
    objExcel.ActiveSheet.Cells(v, h + 6) = ""
    objExcel.ActiveSheet.Cells(v, h + 7) = "" & suma1
    objExcel.ActiveSheet.Cells(v, h + 8) = ""
    v = v + 1

    Dim I As Integer

    For I = 4 To 8
        objExcel.ActiveSheet.Cells(v - 1, I).Font.bold = True
        objExcel.ActiveSheet.Cells(v - 1, I).Interior.color = RGB(248, 243, 53)
    Next
    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    Exit Sub
cmd5612_err:
    MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub cuerpo_programa_saldoexcellRotacion(mytablex As ADODB.Recordset)

    Dim mytabley  As New ADODB.Recordset

    Dim mytablez  As New ADODB.Recordset

    Dim sw        As Integer

    Dim temp      As String

    Dim buf       As String

    Dim sw1       As Integer

    Dim temp1     As String

    Dim buf1      As String

    Dim bufx      As String

    Dim saldoini  As Double

    Dim saldoindx As Double

    Dim sdx       As Double

    Dim sdx1      As Double

    Dim sdx2      As Double

    Dim found     As Integer

    Dim v, h As Double

    Dim xprecio As Double

    Dim vr

    saldoindx = 0
    sdx1 = 0
    sdx2 = 0
    suma1 = 0
    suma2 = 0
    sw1 = 0

    Dim Heading(9) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    On Error GoTo cmd5612_err
   
    Heading(1) = "Producto"
    Heading(2) = "Descripcio"
    Heading(3) = "Unidad"
    Heading(4) = "Factor"
    Heading(5) = "Saldo"
    Heading(6) = "Costo"
    Heading(7) = "Total"
    Heading(8) = "Familia"

    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_Excel(9, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    v = 5
    h = 1

    Command1.Visible = True
    Do

        If mytablex.EOF Then Exit Do
        If proveedor <> "%" Then
            found = ver_proveedor("" & mytablex.Fields("producto"))

            If found = 0 Then GoTo seguy8

        End If

        vr = DoEvents()

        If Command1.Visible = False Then Exit Do

        '------------- verificamos la condicion
        'If Combo1 <> "TODOS" Then

        xprecio = Val(Format(Val("" & mytablex.Fields("costop")), "0.00000"))

        'verificamos que tipo de costeo
        If mytablez.State = 1 Then mytablez.Close
        If Trim(quecosto) = "COSTOULTIMO" Then
            xprecio = Val(Format(Val("" & mytablex.Fields("costou")), "0.00000"))

        End If

        If Trim(quecosto) = "COSTOPROMEDIO" Then
            xprecio = Val(Format(Val("" & mytablex.Fields("costop")), "0.00000"))

        End If

        If Trim(quecosto) = "PRECIOVENTA" Then
            mytablez.Open "Select * from precios where local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic

            If mytablez.RecordCount > 0 Then
                xprecio = Val(Format("" & mytablez.Fields("pventa1"), "0.00"))

            End If

            mytablez.Close

        End If

        If conigv = "" Or conigv = "S" Then
            xprecio = xprecio

        End If

        If conigv = "N" Then
            If Val("" & mytablex.Fields("igv")) > 0 Then
                xprecio = xprecio / (1 + Val("" & mytablex.Fields("igv")) / 100)
                xprecio = Val(Format(xprecio, "0.00000"))

            End If

        End If

        saldoini = 0

        If mytablez.State = 1 Then mytablez.Close
        mytablez.Open "Select * from almacen where local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "'", cn, adOpenStatic, adLockOptimistic

        If mytablez.RecordCount > 0 Then
            saldoini = Val("" & mytablez.Fields("saldo"))

            If Combo1 = "SALDO>0" Then
                If saldoini > 0 Then GoTo sigueme

            End If

            If Combo1 = "SALDO>=0" Then
                If saldoini >= 0 Then GoTo sigueme

            End If

            If Combo1 = "SALDO<0" Then
                If saldoini < 0 Then GoTo sigueme

            End If

            If Combo1 = "SALDO<=0" Then
                If saldoini <= 0 Then GoTo sigueme

            End If

            If Combo1 = "SALDO=0" Then
                If saldoini = 0 Then GoTo sigueme

            End If

            If Combo1 = "TODOS" Then
                GoTo sigueme

            End If

            GoTo sigueme1
        Else
            GoTo sigueme1

        End If

        'End If
sigueme:
        'mytablez.Close

        buf = "" & mytablex.Fields("familia")

        If sw = 0 Then
            sw = 1
            objExcel.ActiveSheet.Cells(v, h) = "" & mytablex.Fields("familia")
            objExcel.ActiveSheet.Cells(v, h + 1) = busca_familia("" & mytablex.Fields("familia"))
            objExcel.ActiveSheet.Cells(v, h + 2) = ""
            objExcel.ActiveSheet.Cells(v, h + 3) = ""
            objExcel.ActiveSheet.Cells(v, h + 4) = ""
            objExcel.ActiveSheet.Cells(v, h + 5) = ""
            objExcel.ActiveSheet.Cells(v, h + 6) = ""
            objExcel.ActiveSheet.Cells(v, h + 7) = ""

            v = v + 1
            temp = "" & mytablex.Fields("familia")

        End If

        If "" & mytablex.Fields("familia") <> temp Then
            temp = "" & mytablex.Fields("familia")
            objExcel.ActiveSheet.Cells(v, h) = "" & mytablex.Fields("familia")
            objExcel.ActiveSheet.Cells(v, h + 1) = busca_familia("" & mytablex.Fields("familia"))
            objExcel.ActiveSheet.Cells(v, h + 2) = ""
            objExcel.ActiveSheet.Cells(v, h + 3) = ""
            objExcel.ActiveSheet.Cells(v, h + 4) = ""
            objExcel.ActiveSheet.Cells(v, h + 5) = ""
            objExcel.ActiveSheet.Cells(v, h + 6) = ""
            objExcel.ActiveSheet.Cells(v, h + 7) = ""
            v = v + 1

        End If

        If sw1 = 0 Then
            objExcel.ActiveSheet.Cells(v, h) = "" & mytablex.Fields("familia")
            objExcel.ActiveSheet.Cells(v, h + 1) = "" & mytablex.Fields("subfamilia")
            objExcel.ActiveSheet.Cells(v, h + 2) = ""
            objExcel.ActiveSheet.Cells(v, h + 3) = ""
            objExcel.ActiveSheet.Cells(v, h + 4) = ""
            objExcel.ActiveSheet.Cells(v, h + 5) = ""
            objExcel.ActiveSheet.Cells(v, h + 6) = ""
            objExcel.ActiveSheet.Cells(v, h + 7) = ""
            v = v + 1
            temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")
            sw1 = 1

        End If

        If temp1 <> "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia") Then
            objExcel.ActiveSheet.Cells(v, h) = "" & mytablex.Fields("familia")
            objExcel.ActiveSheet.Cells(v, h + 1) = "" & mytablex.Fields("subfamilia")
            objExcel.ActiveSheet.Cells(v, h + 2) = ""
            objExcel.ActiveSheet.Cells(v, h + 3) = ""
            objExcel.ActiveSheet.Cells(v, h + 4) = ""
            objExcel.ActiveSheet.Cells(v, h + 5) = ""
            objExcel.ActiveSheet.Cells(v, h + 6) = ""
            objExcel.ActiveSheet.Cells(v, h + 7) = ""
            v = v + 1
            temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")

        End If

        'SALDO ALMACEN
        saldoini = 0

        If mytablez.State = 1 Then mytablez.Close
        mytablez.Open "Select * from almacen where local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "'", cn, adOpenStatic, adLockOptimistic

        If mytablez.RecordCount > 0 Then
            saldoini = Val("" & mytablez.Fields("saldo"))

        End If

        mytablez.Close
        saldoindx = saldoini
        objExcel.ActiveSheet.Cells(v, h) = "" & mytablex.Fields("producto")
        objExcel.ActiveSheet.Cells(v, h + 1) = "" & mytablex.Fields("descripcio")
        objExcel.ActiveSheet.Cells(v, h + 2) = "" & mytablex.Fields("unidad")
        objExcel.ActiveSheet.Cells(v, h + 3) = "" & mytablex.Fields("factor")

        objExcel.ActiveSheet.Cells(v, h + 4) = saldoindx
        objExcel.ActiveSheet.Cells(v, h + 5) = xprecio
        sdx = xprecio * saldoindx
        suma1 = suma1 + sdx
        suma2 = suma2 + saldoini
        objExcel.ActiveSheet.Cells(v, h + 6) = sdx
        objExcel.ActiveSheet.Cells(v, h + 7) = ""
        objExcel.ActiveSheet.Cells(v, h + 7) = "" & mytablex.Fields("Familia")
        v = v + 1
        '--------------------------------------
sigueme1:
seguy8:
        mytablex.MoveNext
    Loop
    objExcel.ActiveSheet.Cells(v, h) = ""
    objExcel.ActiveSheet.Cells(v, h + 1) = ""
    objExcel.ActiveSheet.Cells(v, h + 2) = ""
    objExcel.ActiveSheet.Cells(v, h + 3) = ""
    objExcel.ActiveSheet.Cells(v, h + 4) = "" & suma2
    objExcel.ActiveSheet.Cells(v, h + 5) = ""
    objExcel.ActiveSheet.Cells(v, h + 6) = "" & suma1
    objExcel.ActiveSheet.Cells(v, h + 7) = ""
    v = v + 1

    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    Exit Sub
cmd5612_err:
    MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
    Exit Sub

End Sub

''''19/09/2017 kenyo Mejora Reporte Saldo Actual
Sub reporte_saldoexcell()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    ''''21/09/2017 kenyo Reporte de Stock minimo Ticket
    'found = sql_producto(mytablex)
    If Combo2 = "Debajo de Stock minimo" Then
        found = sql_productominimo(mytablex)
    Else
        found = sql_producto(mytablex)

    End If

    ''''21/09/2017 kenyo Reporte de Stock minimo Ticket

    If found = 0 Then
        mytablex.Close
        'mytablez.Close
        Exit Sub

    End If

    cuerpo_programa_saldoexcell mytablex
    Command1.Visible = False
    mytablex.Close
    'mytablez.Close

End Sub

Sub reporte_sinrotacionexcell()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    found = sql_producto2(mytablex)

    If found = 0 Then
        mytablex.Close
        'mytablez.Close
        Exit Sub

    End If

    cuerpo_programa_saldoexcellRotacion mytablex
    Command1.Visible = False
    mytablex.Close
    'mytablez.Close

End Sub

Sub reporte_saldoex8()  'reporte a un periodo

    Dim txcambio As Double

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    On Error GoTo cmd34444_err

    If Not IsDate(fechai) Then
        fechai.SetFocus
        Exit Sub

    End If

    If Not IsDate(fechaf) Then
        fechaf.SetFocus
        Exit Sub

    End If

    found = condiciona_mensual(mytablex)

    If found = 0 Then
        mytablex.Close
        Exit Sub

    End If

    Command1.Visible = True
    txcambio = busca_paridad()
    
    '----------------------------------------------------
    cuerpo_saldo8ex mytablex, txcambio
    mytablex.Close
    Command1.Visible = False
    Exit Sub
cmd34444_err:
    MsgBox "Aviso reporte_saldoex8" + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub cuerpo_saldo8ex(mytablex As ADODB.Recordset, txcambio As Double)

    Dim I           As Integer

    Dim mytabley    As New ADODB.Recordset

    Dim tipo_cambio As Double

    Dim sw          As Integer

    Dim temp        As String

    Dim buf         As String

    Dim sw1         As Integer

    Dim temp1       As String

    Dim buff        As String

    Dim buf1        As String

    Dim bufx        As String

    Dim saldoini    As Double

    Dim saldoindx   As Double

    Dim sdx         As Double

    Dim sdx1        As Double

    Dim sdx2        As Double

    Dim sdx3        As Double

    Dim found       As Integer

    Dim pventa1     As String

    Dim pventa2     As String

    Dim pventa3     As String

    Dim pventa4     As String

    Dim pventa5     As String

    Dim unidad1     As String

    Dim unidad2     As String

    Dim unidad3     As String

    Dim unidad4     As String

    Dim factor1     As String

    Dim factor2     As String

    Dim factor3     As String

    Dim factor4     As String

    Dim sbuf        As String

    Dim vr

    Dim v           As Integer

    Dim h           As Integer

    Dim Heading(10) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    On Error GoTo cmd781117_err

    sbuf = ""
    saldoindx = 0
    sdx1 = 0
    sdx2 = 0
    suma1 = 0
    sw1 = 0
    suma1 = 0
    sdx = 0

    ''' kenyo 23/08/2017 Mejora reporte saldos a un periodo
    Dim xprecio As Double

    Dim sumac   As Double

    Dim sdxt    As Double

    sumac = 0
    sdxt = 0

    Dim JK As Integer

    JK = 0
    ''' kenyo 23/08/2017 Mejora reporte saldos a un periodo

    'Command1.Visible = True
    '--------------- cabecera excell
    Heading(1) = "Producto"
    Heading(2) = "Descripcio"
    Heading(3) = "Unidad"
    Heading(4) = "Factor"
    Heading(5) = "Saldo"

    ''' kenyo 23/08/2017 Mejora reporte saldos a un periodo
    Heading(6) = "Costo"
    Heading(7) = "Total"
    Heading(8) = "Familia"
    Heading(9) = "SubFamilia"
    ''' kenyo 23/08/2017 Mejora reporte saldos a un periodo

    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_SaldosPeriodo(9, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    '--- fin cabecera excell

    ''' kenyo 23/08/2017 Mejora reporte saldos a un periodo
    'v = 5
    v = 4
    ''' kenyo 23/08/2017 Mejora reporte saldos a un periodo

    h = 1
    tipo_cambio = txcambio
    Do

        If mytablex.EOF Then Exit Do
        vr = DoEvents()

        If Command1.Visible = False Then
            Exit Do

        End If

        If mytabley.State = 1 Then mytabley.Close
        mytabley.Open "Select * from producto where producto='" & "" & mytablex.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic

        If mytabley.RecordCount = 0 Then
            GoTo seguy

        End If

        '------------- verificamos la condicion
        buf = "" & mytablex.Fields("producto")

        If sw = 0 Then
            sw = 1
            objExcel.ActiveSheet.Cells(v, h) = "" & mytabley.Fields("producto")
            objExcel.ActiveSheet.Cells(v, h + 1) = "" & mytabley.Fields("descripcio")
            objExcel.ActiveSheet.Cells(v, h + 2) = "" & mytabley.Fields("unidad")
            objExcel.ActiveSheet.Cells(v, h + 3) = "" & mytabley.Fields("factor")

            ''' kenyo 23/08/2017 Mejora reporte saldos a un periodo
            xprecio = Val(Format(Val("" & mytabley.Fields("costou")), "0.00000"))
            objExcel.ActiveSheet.Cells(v, h + 5) = "" & xprecio
            objExcel.ActiveSheet.Cells(v, h + 7) = "" & mytabley.Fields("familia")
            objExcel.ActiveSheet.Cells(v, h + 8) = "" & mytabley.Fields("subfamilia")
            ''' kenyo 23/08/2017 Mejora reporte saldos a un periodo

            temp = "" & mytablex.Fields("producto")

        End If

        If "" & mytablex.Fields("producto") <> temp Then
            buf = "" & sdx 'calcula_saldo(sdx, Val("" & mytabley.Fields("factor")))
            objExcel.ActiveSheet.Cells(v, h + 4) = "" & buf
            v = v + 1
            sdx = 0
            temp = "" & mytablex.Fields("producto")
            objExcel.ActiveSheet.Cells(v, h) = "" & mytabley.Fields("producto")
            objExcel.ActiveSheet.Cells(v, h + 1) = "" & mytabley.Fields("descripcio")
            objExcel.ActiveSheet.Cells(v, h + 2) = "" & mytabley.Fields("unidad")
            objExcel.ActiveSheet.Cells(v, h + 3) = "" & mytabley.Fields("factor")
   
            ''' kenyo 23/08/2017 Mejora reporte saldos a un periodo
            objExcel.ActiveSheet.Cells(v, h + 5) = "" & mytabley.Fields("costou")
            objExcel.ActiveSheet.Cells(v, h + 7) = "" & mytabley.Fields("familia")
            objExcel.ActiveSheet.Cells(v, h + 8) = "" & mytabley.Fields("subfamilia")
            ''' kenyo 23/08/2017 Mejora reporte saldos a un periodo
   
        End If

        If "" & mytablex.Fields("acu") = "A" Or "" & mytablex.Fields("acu") = "B" Or "" & mytablex.Fields("acu") = "C" Or "" & mytablex.Fields("acu") = "D" Or "" & mytablex.Fields("acu") = "G" Then     'ventas
            sdx = sdx - Val("" & mytablex.Fields("xcanti"))
            suma1 = suma1 - Val("" & mytablex.Fields("xcanti"))

        End If

        If "" & mytablex.Fields("acu") = "J" Or "" & mytablex.Fields("acu") = "K" Or "" & mytablex.Fields("acu") = "L" Or "" & mytablex.Fields("acu") = "M" Or "" & mytablex.Fields("acu") = "P" Then     'COMPRAS
            sdx = sdx + Val("" & mytablex.Fields("xcanti"))
            suma1 = suma1 + Val("" & mytablex.Fields("xcanti"))

        End If

        If "" & mytablex.Fields("acu") = "S" Then      'entradas
            sdx = sdx + Val("" & mytablex.Fields("xcanti"))
            suma1 = suma1 + Val("" & mytablex.Fields("xcanti"))

        End If

        If "" & mytablex.Fields("acu") = "T" Then      'salida
            sdx = sdx - Val("" & mytablex.Fields("xcanti"))
            suma1 = suma1 - Val("" & mytablex.Fields("xcanti"))

        End If

        If "" & mytablex.Fields("acu") = "N" Then      'nota credito compras
            sdx = sdx - Val("" & mytablex.Fields("xcanti"))
            suma1 = suma1 - Val("" & mytablex.Fields("xcanti"))

        End If

        If "" & mytablex.Fields("acu") = "O" Then      'nota DEBITO compras
            sdx = sdx + Val("" & mytablex.Fields("xcanti"))
            suma1 = suma1 + Val("" & mytablex.Fields("xcanti"))

        End If

        If "" & mytablex.Fields("acu") = "E" Then      'nota credito ventas
            sdx = sdx - Val("" & mytablex.Fields("xcanti"))
            suma1 = suma1 - Val("" & mytablex.Fields("xcanti"))

        End If

        If "" & mytablex.Fields("acu") = "F" Then      'nota DEBITO ventas
            sdx = sdx + Val("" & mytablex.Fields("xcanti"))
            suma1 = suma1 + Val("" & mytablex.Fields("xcanti"))

        End If

        ''' kenyo 23/08/2017 Mejora reporte saldos a un periodo
        xprecio = objExcel.ActiveSheet.Cells(v, h + 5)
        objExcel.ActiveSheet.Cells(v, h + 6) = sdx * xprecio
        ''' kenyo 23/08/2017 Mejora reporte saldos a un periodo

seguy:

        mytablex.MoveNext
    Loop
    Command1.Visible = False

    If buf <> "" Then
        buf = calcula_saldo(sdx, Val("" & mytabley.Fields("factor")))

    End If

    buf = "" & sdx
    objExcel.ActiveSheet.Cells(v, h + 4) = "" & buf
    v = v + 1
    buf = "" & suma1
    objExcel.ActiveSheet.Cells(v, h + 4) = "" & buf

    ''' kenyo 23/08/2017 Mejora reporte saldos a un periodo
    'v = v + 1

    For JK = 4 To v
        sumac = sumac + objExcel.ActiveSheet.Cells(JK, 7)
    Next
    objExcel.ActiveSheet.Cells(v, h + 6) = "" & sumac

    objExcel.ActiveSheet.Cells(v, h + 4).Font.bold = True
    objExcel.ActiveSheet.Cells(v, h + 4).Interior.color = RGB(248, 243, 53)
    objExcel.ActiveSheet.Cells(v, h + 6).Font.bold = True
    objExcel.ActiveSheet.Cells(v, h + 6).Interior.color = RGB(248, 243, 53)
    v = v + 1
  
    ''' kenyo 23/08/2017 Mejora reporte saldos a un periodo

    MsgBox "Proceso terminado ", 48, "Aviso"
    Exit Sub
cmd781117_err:
    MsgBox "Error en cuerpo_saldo8ex " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub reporte_saldo8()  'reporte a un periodo

    Dim txcambio As Double

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    If Not IsDate(fechai) Then
        fechai.SetFocus
        Exit Sub

    End If

    If Not IsDate(fechaf) Then
        fechaf.SetFocus
        Exit Sub

    End If

    found = condiciona_mensual(mytablex)

    If found = 0 Then
        mytablex.Close
        Exit Sub

    End If

    Command1.Visible = True
    txcambio = busca_paridad()
    
    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    cabecera_saldo8
    cuerpo_programa_saldo8 mytablex, txcambio
    Command1.Visible = False
    Close #1
    cerrar_archivo
    mytablex.Close
    
    'genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1
    Command1.Visible = False
    found = valida_wordpad(FileName)

End Sub

Sub cabecera_saldo8()

    Dim mytablex As Table

    Dim buf      As String

    Dim I        As Integer

    Dim found    As Integer

    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("ALMACEN : " & bodega, 25, 2, 0)
    found = formateaa("FECHA INICIO :" & fechai, 40, 2, 0)
    found = formateaa("FECHA FINAL  :" & fechaf, 40, 2, 0)
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)
    found = formateaa("Producto ", 11, 0, 0)
    found = formateaa("Descripcio ", 41, 0, 0)
    found = formateaa("Und", 4, 0, 0)
    found = formateaa("Fac", 5, 0, 0)
    found = formateaa("saldo ", 11, 2, 1)
   
    '--------------------------------------
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)

End Sub

Sub cuerpo_programa_saldo8(mytablex As ADODB.Recordset, txcambio As Double)

    Dim I           As Integer

    Dim tipo_cambio As Double

    Dim sw          As Integer

    Dim temp        As String

    Dim buf         As String

    Dim sw1         As Integer

    Dim temp1       As String

    Dim buff        As String

    Dim buf1        As String

    Dim bufx        As String

    Dim saldoini    As Double

    Dim saldoindx   As Double

    Dim sdx         As Double

    Dim sdx1        As Double

    Dim sdx2        As Double

    Dim sdx3        As Double

    Dim found       As Integer

    Dim pventa1     As String

    Dim pventa2     As String

    Dim pventa3     As String

    Dim pventa4     As String

    Dim pventa5     As String

    Dim unidad1     As String

    Dim unidad2     As String

    Dim unidad3     As String

    Dim unidad4     As String

    Dim factor1     As String

    Dim factor2     As String

    Dim factor3     As String

    Dim factor4     As String

    Dim mytabley    As New ADODB.Recordset

    Dim sbuf        As String

    Dim vr

    sbuf = ""
    saldoindx = 0
    sdx1 = 0
    sdx2 = 0
    suma1 = 0
    sw1 = 0
    suma1 = 0
    sdx = 0
    Command1.Visible = True
    tipo_cambio = txcambio
    Do

        If mytablex.EOF Then Exit Do
        If proveedor <> "%" Then
            found = ver_proveedor("" & mytablex.Fields("producto"))

            If found = 0 Then GoTo seguy9

        End If

        vr = DoEvents()

        If Command1.Visible = False Then
            Exit Do

        End If

        If mytabley.State = 1 Then mytabley.Close
        mytabley.Open "Select * from producto where producto='" & "" & mytablex.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic

        If mytabley.RecordCount = 0 Then
            GoTo seguy9

        End If

        '------------- verificamos la condicion
        buf = "" & mytablex.Fields("producto")

        If sw = 0 Then
            sw = 1
            found = formateaa("" & mytablex.Fields("producto"), 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            found = formateaa("" & mytabley.Fields("descripcio"), 40, 0, 0)
            found = formateaa("", 1, 0, 0)
            found = formateaa("" & mytabley.Fields("unidad"), 3, 0, 0)
            found = formateaa("", 1, 0, 0)
            found = formateaa("" & mytabley.Fields("factor"), 4, 0, 0)
            found = formateaa("", 1, 0, 0)
            temp = "" & mytablex.Fields("producto")
            nlineas

        End If

        If "" & mytablex.Fields("producto") <> temp Then
            buf = calcula_saldo(sdx, Val("" & mytabley.Fields("factor")))
            found = formateaa(buf, 10, 2, 1)
            sdx = 0
            nlineas
            temp = "" & mytablex.Fields("producto")
            found = formateaa("" & mytablex.Fields("producto"), 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            found = formateaa("" & mytabley.Fields("descripcio"), 40, 0, 0)
            found = formateaa("", 1, 0, 0)
            found = formateaa("" & mytabley.Fields("unidad"), 3, 0, 0)
            found = formateaa("", 1, 0, 0)
            found = formateaa("" & mytabley.Fields("factor"), 4, 0, 0)
            found = formateaa("", 1, 0, 0)

        End If

        If "" & mytablex.Fields("acu") = "A" Or "" & mytablex.Fields("acu") = "B" Or "" & mytablex.Fields("acu") = "C" Or "" & mytablex.Fields("acu") = "D" Or "" & mytablex.Fields("acu") = "G" Then     'ventas
            sdx = sdx - Val("" & mytablex.Fields("xcanti"))
            suma1 = suma1 - Val("" & mytablex.Fields("xcanti"))

        End If

        If "" & mytablex.Fields("acu") = "J" Or "" & mytablex.Fields("acu") = "K" Or "" & mytablex.Fields("acu") = "L" Or "" & mytablex.Fields("acu") = "M" Or "" & mytablex.Fields("acu") = "P" Then     'COMPRAS
            sdx = sdx + Val("" & mytablex.Fields("xcanti"))
            suma1 = suma1 + Val("" & mytablex.Fields("xcanti"))

        End If

        If "" & mytablex.Fields("acu") = "S" Then      'entradas
            sdx = sdx + Val("" & mytablex.Fields("xcanti"))
            suma1 = suma1 + Val("" & mytablex.Fields("xcanti"))

        End If

        If "" & mytablex.Fields("acu") = "T" Then      'salida
            sdx = sdx - Val("" & mytablex.Fields("xcanti"))
            suma1 = suma1 - Val("" & mytablex.Fields("xcanti"))

        End If

        If "" & mytablex.Fields("acu") = "N" Then      'nota credito compras
            sdx = sdx - Val("" & mytablex.Fields("xcanti"))
            suma1 = suma1 - Val("" & mytablex.Fields("xcanti"))

        End If

        If "" & mytablex.Fields("acu") = "O" Then      'nota DEBITO compras
            sdx = sdx + Val("" & mytablex.Fields("xcanti"))
            suma1 = suma1 + Val("" & mytablex.Fields("xcanti"))

        End If

        If "" & mytablex.Fields("acu") = "E" Then      'nota credito ventas
            sdx = sdx - Val("" & mytablex.Fields("xcanti"))
            suma1 = suma1 - Val("" & mytablex.Fields("xcanti"))

        End If

        If "" & mytablex.Fields("acu") = "F" Then      'nota DEBITO ventas
            sdx = sdx + Val("" & mytablex.Fields("xcanti"))
            suma1 = suma1 + Val("" & mytablex.Fields("xcanti"))

        End If

seguy9:
        mytablex.MoveNext
    Loop
    Command1.Visible = False

    ''' kenyo 23/08/2017 Mejora reporte saldos a un periodo
    'buf = calcula_saldo(sdx, Val("" & mytabley.Fields("factor")))
    If buf <> "" Then
        buf = calcula_saldo(sdx, Val("" & mytabley.Fields("factor")))

    End If

    ''' kenyo 23/08/2017 Mejora reporte saldos a un periodo

    found = formateaa(buf, 10, 2, 1)
    nlineas
    buf = "" & suma1
    found = formateaa(" Total Unidades ", 61, 0, 1)
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)

End Sub

Function busca_paramed(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "Select * from bodega where codigo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_paramed = "" & mytablex.Fields("fecha")

    End If

    mytablex.Close

End Function

Function condiciona_mensual(mytablex As ADODB.Recordset)

    Dim buf As String

    On Error GoTo cmd321_err

    buf = "select producto,acu,sum(cantidad*factor) AS XCANTI from " & xbasedatos & "  where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    
    buf = buf & " and local='" & extra_loquesea(local1) & "'"
    buf = buf & " and bodega='" & extra_loquesea(bodega) & "'"

    If producto <> "%" Then
        buf = buf & " and producto like '" & producto & "'"

    End If

    If descripcio <> "%" Then
        buf = buf & " and descripcio like '" & descripcio & "'"

    End If

    If moneda <> "%" Then
        buf = buf & " and moneda='" & moneda & "'"

    End If

    If igv <> "%" Then
        If igv = "GRAVADO" Then
            buf = buf & " and igv>0 "

        End If

        If igv = "EXENTO" Then
            buf = buf & " and (igv=0 or igv=null) "

        End If

    End If

    If familia <> "%" Then
        buf = buf & " and familia like '" & extra_loquesea(familia) & "'"

    End If

    If subfamilia <> "%" Then
        buf = buf & " and subfamilia like '" & subfamilia & "'"

    End If

    buf = buf & " and (acu='S' or acu='T' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' OR acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='E' or acu='N')"
    buf = buf & " and estado='2'"
    buf = buf & " group by producto,acu "
    buf = buf & " order by producto,acu"
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    condiciona_mensual = 1
    
    'found = kardexactualizasi("" & local1, "%", "" & bodega, "" & fechai, "" & fechaf)
    
    Exit Function
cmd321_err:
    Exit Function

End Function

Function ve_descarga(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM tipo where  tipo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe

        Select Case "" & mytablex.Fields("tipodoc")

            Case "A", "B", "C", "D", "G", "E", "F", "J", "K", "L", "M", "P", "N", "O", "S", "T"
                ve_descarga = 1

        End Select

    End If

    mytablex.Close

End Function

Sub cuerpo_programa_excellini(mytablex As ADODB.Recordset)

    Dim mytabley  As New ADODB.Recordset

    Dim mytablez  As New ADODB.Recordset

    Dim mcanti    As Double

    Dim mmx       As Double

    Dim sw        As Integer

    Dim temp      As String

    Dim buf       As String

    Dim sw1       As Integer

    Dim temp1     As String

    Dim buf1      As String

    Dim bufx      As String

    Dim saldoini  As Double

    Dim saldoindx As Double

    Dim sdx       As Double

    Dim sdx1      As Double

    Dim sdx2      As Double

    Dim sdx3      As Double

    Dim found     As Integer

    Dim vr

    saldoindx = 0
    sdx1 = 0
    sdx2 = 0
    sdx3 = 0
    suma1 = 0
    suma2 = 0
    sw1 = 0
    mcanti = 0
    Command1.Visible = True
    Do

        If mytablex.EOF Then Exit Do
        If proveedor <> "%" Then
            found = ver_proveedor("" & mytablex.Fields("producto"))

            If found = 0 Then GoTo seguy10

        End If

        vr = DoEvents()

        If Command1.Visible = False Then Exit Do

        mmx = 0

        If mytablez.State = 1 Then mytablez.Close
        mytablez.Open "Select * from saldoini where local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "' and fecha='" & Format(fechai, "YYYYMMDD") & "'", cn, adOpenStatic, adLockOptimistic

        If mytablez.RecordCount > 0 Then
            mmx = Val("" & mytablez.Fields("cantidad")) * Val("" & mytablez.Fields("factor")) + Val("" & mytablez.Fields("cantidad1"))

            If Combo1 = "SALDO>0" Then
                If mmx > 0 Then GoTo sigueme116

            End If

            If Combo1 = "SALDO>=0" Then
                If mmx >= 0 Then GoTo sigueme116

            End If

            If Combo1 = "SALDO<0" Then
                If mmx < 0 Then GoTo sigueme116

            End If

            If Combo1 = "SALDO<=0" Then
                If mmx <= 0 Then GoTo sigueme116

            End If

            If Combo1 = "SALDO=0" Then
                If mmx = 0 Then GoTo sigueme116

            End If

            GoTo sigueme212

        End If

sigueme116:
        'mytablez.Close

        buf = "" & mytablex.Fields("familia")

        If sw = 0 Then
            sw = 1
            temp = "" & mytablex.Fields("familia")
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("*", 1, 2, 0)
            nlineas

        End If

        If "" & mytablex.Fields("familia") <> temp Then
            temp = "" & mytablex.Fields("familia")
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas

        End If

        If sw1 = 0 Then
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("-", 1, 0, 0)
            found = formateaa("" & mytablex.Fields("subfamilia"), 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")
            sw1 = 1

        End If

        If temp1 <> "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia") Then
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("-", 1, 0, 0)
            found = formateaa("" & mytablex.Fields("subfamilia"), 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")

        End If

        found = formateaa("" & mytablex.Fields("producto"), 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("descripcio"), 40, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("unidad"), 6, 0, 0)
        found = formateaa("x", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("factor"), 4, 0, 0)
        found = formateaa("", 1, 0, 0)
        'SALDO INICIAL
        saldoini = 0

        If mytablez.State = 1 Then mytablez.Close
        mytablez.Open "Select * from saldoini where local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "' and fecha='" & Format(fechai, "YYYYMMDD") & "'", cn, adOpenStatic, adLockOptimistic

        If mytablez.RecordCount > 0 Then
            saldoini = Val("" & mytablez.Fields("cantidad")) * Val("" & mytablez.Fields("factor")) + Val("" & mytablez.Fields("cantidad1"))

        End If

        mytablez.Close
        suma2 = suma2 + saldoini
        saldoindx = saldoini
        bufx = "" & saldoini
        buf = calcula_saldo(saldoindx, Val("" & mytablex.Fields("factor")))

        If Val(bufx) = 0 Then
            bufx = ""

        End If

        found = formateaa(bufx, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("costop")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        sdx = Val("" & mytablex.Fields("costop")) * Val(bufx)
        buf = "" & sdx
        suma1 = suma1 + sdx
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 2, 0)
        nlineas
        '---------------------------------------
sigueme212:
seguy10:
        mytablex.MoveNext
    Loop
    found = formateaa(" Total Inventario ", 64, 0, 0)
    bufx = "" & suma2
    found = formateaa(bufx, 10, 0, 1)
    found = formateaa("", 12, 0, 0)
    bufx = Format(Val("" & suma1), "0.00")
    found = formateaa(bufx, 10, 2, 1)

End Sub

Sub reporte_saldo_margen()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    found = sql_producto(mytablex)

    If found = 0 Then
        mytablex.Close
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_saldo_margen
    cuerpo_programa_saldo_margen mytablex
    Command1.Visible = False
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
     
    'genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1
    found = valida_wordpad(FileName)

End Sub

Sub cabecera_saldo_margen()

    Dim buf   As String

    Dim I     As Integer

    Dim found As Integer

    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)
   
    found = formateaa(" ", 52, 0, 0)
    found = formateaa("------Costo----------- ", 23, 0, 0)
    found = formateaa("------Pventa---------- ", 23, 0, 0)
    found = formateaa("", 11, 2, 0)
   
    found = formateaa("Producto ", 11, 0, 0)
    found = formateaa("Descripcio ", 41, 0, 0)
    found = formateaa("Unid", 7, 0, 0)
    found = formateaa("Fx", 5, 0, 0)
    found = formateaa("Costou ", 11, 0, 1)
    found = formateaa("Unid", 7, 0, 0)
    found = formateaa("Fx", 5, 0, 0)
    found = formateaa("Pventa ", 11, 0, 1)
    found = formateaa("Margen ", 11, 0, 1)
    found = formateaa("%Margen ", 11, 2, 1)
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)

End Sub

Sub cuerpo_programa_saldo_margen(mytablex As ADODB.Recordset)

    Dim mytabley  As New ADODB.Recordset

    Dim mytablez  As New ADODB.Recordset

    Dim sw        As Integer

    Dim temp      As String

    Dim buf       As String

    Dim sw1       As Integer

    Dim temp1     As String

    Dim buf1      As String

    Dim bufx      As String

    Dim saldoini  As Double

    Dim found     As Integer

    Dim bufund    As String

    Dim buffac    As String

    Dim bufprecio As String

    Dim xventa    As Double

    Dim xcosto    As Double

    Dim xmargen   As Double

    Dim xmargenpo As Double

    Dim vr

    sw1 = 0
    saldoini = 0
    Command1.Visible = True
    Do

        If mytablex.EOF Then Exit Do
        If proveedor <> "%" Then
            found = ver_proveedor("" & mytablex.Fields("producto"))

            If found = 0 Then GoTo seguy11

        End If

        vr = DoEvents()

        If Command1.Visible = False Then Exit Do
        '------------- verificamos la condicion
        'If Combo1 <> "TODOS" Then
        saldoini = 0

        If mytablez.State = 1 Then mytablez.Close
        mytablez.Open "Select * from almacen where local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "'", cn, adOpenStatic, adLockOptimistic

        If mytablez.RecordCount > 0 Then
            saldoini = Val("" & mytablez.Fields("saldo"))

            If Combo1 = "SALDO>0" Then
                If saldoini > 0 Then GoTo siguemee

            End If

            If Combo1 = "SALDO>=0" Then
                If saldoini >= 0 Then GoTo siguemee

            End If

            If Combo1 = "SALDO<0" Then
                If saldoini < 0 Then GoTo siguemee

            End If

            If Combo1 = "SALDO<=0" Then
                If saldoini <= 0 Then GoTo siguemee

            End If

            If Combo1 = "SALDO=0" Then
                If saldoini = 0 Then GoTo siguemee

            End If

            If Combo1 = "TODOS" Then
                GoTo siguemee

            End If

            GoTo sigueme113
            Else: GoTo sigueme113

        End If

siguemee:
        buf = "" & mytablex.Fields("familia")

        If sw = 0 Then
            sw = 1
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            buf = busca_familia("" & mytablex.Fields("familia"))
            found = formateaa(buf, 30, 0, 0)
            found = formateaa("*", 1, 2, 0)
            temp = "" & mytablex.Fields("familia")
            nlineas

        End If

        If "" & mytablex.Fields("familia") <> temp Then
            temp = "" & mytablex.Fields("familia")
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            buf = busca_familia("" & mytablex.Fields("familia"))
            found = formateaa(buf, 30, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas

        End If

        If sw1 = 0 Then
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("-", 1, 0, 0)
            found = formateaa("" & mytablex.Fields("subfamilia"), 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")
            sw1 = 1

        End If

        If temp1 <> "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia") Then
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("-", 1, 0, 0)
            found = formateaa("" & mytablex.Fields("subfamilia"), 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")

        End If

        found = formateaa("" & mytablex.Fields("producto"), 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("descripcio"), 40, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("unidad"), 6, 0, 0)
        found = formateaa("x", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("factor"), 4, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("costou"), 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        xcosto = Val("" & mytablex.Fields("costou")) / Val("" & mytablex.Fields("factor"))

        bufund = ""
        buffac = "1"
        bufprecio = ""

        If mytabley.State = 1 Then mytabley.Close
        mytabley.Open "Select * from precios where local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic

        If mytabley.RecordCount > 0 Then
            bufund = "" & mytabley.Fields("unidad1")
            buffac = "" & mytabley.Fields("factor1")
            bufprecio = "" & mytabley.Fields("pventa1")

        End If

        xventa = 0

        If Val(buffac) > 0 Then
            xventa = Val(bufprecio) / Val(buffac)

        End If

        found = formateaa(bufund, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa(buffac, 4, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa(bufprecio, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        xmargen = xventa - xcosto
        xmargen = Val(Format("" & xmargen, "0.00"))

        found = formateaa("" & xmargen, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        xmargenpo = 0

        If xcosto > 0 Then
            xmargenpo = (xventa - xcosto) * 100 / xcosto

        End If

        xmargenpo = Val(Format("" & xmargenpo, "0.00"))
        found = formateaa("" & xmargenpo, 10, 0, 1)
        found = formateaa("", 1, 2, 0)

        nlineas
sigueme113:
seguy11:
        mytablex.MoveNext
    Loop

End Sub

Function ver_proveedor(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "Select * from codprov where producto='" & buf & "' and codigo='" & "" & proveedor & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        ver_proveedor = 1

    End If

    mytablex.Close

End Function

Sub kardex_sunat_excell()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    ''''21/09/2017 kenyo Mejora reporte entradas salidas, kardex
    'found = sql_producto(mytablex)
    found = sql_productoConMovimiento(mytablex)
    ''''21/09/2017 kenyo Mejora reporte entradas salidas, kardex

    If found = 0 Then
        mytablex.Close
        Exit Sub

    End If

    cuerpo_kardex_sunat1 mytablex
    Command1.Visible = False

End Sub

Sub cuerpo_kardex_sunat(mytablex As ADODB.Recordset)

    Dim mytabley As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset

    Dim sw       As Integer

    Dim temp     As String

    Dim buf      As String

    Dim found    As Integer

    Dim v, h As Double

    Dim vr

    Dim I           As Integer

    Dim sdx         As Double

    Dim XCantidad   As Double

    Dim xsaldo      As Double

    Dim xcosto      As Double

    Dim xent        As Double

    Dim xsal        As Double

    Dim txent       As Double

    Dim txsal       As Double

    Dim ttxtot      As Double

    Dim Heading(18) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    On Error GoTo cmdd5612_err
    
    Heading(1) = "Fecha"
    Heading(2) = "Tipo"
    Heading(3) = "Serie"
    Heading(4) = "Numero"
    
    Heading(5) = "Operacion"
    Heading(6) = "Cantidad"
    Heading(7) = "CostoUnitario"
    Heading(8) = "CostoTotal"
    Heading(9) = "Cantidad"
    Heading(10) = "CostoUnitario"
    Heading(11) = "CostoTotal"
    Heading(12) = "Cantidad"
    Heading(13) = "CostoUnitario"
    Heading(14) = "CostoTotal"
    
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    '-------------------------------------------------------
    'With objExcel.ActiveSheet
    '.Cells(2, 2) = "Periodo"
    '.Cells(3, 2) = "Ruc"
    '.Cells(4, 2) = "Apellidos y Nombres"
    '.Cells(5, 2) = "Establecimiento"
    '.Cells(6, 2) = "Codigo Existencia:"
    '.Cells(7, 2) = "Tipo:"
    '.Cells(8, 2) = "Descripcio:"
    '.Cells(9, 2) = "Codigo de la unidad de Medida"
    '.Cells(10, 2) = "Metodo de la evaluacion"
        
    'For i = 1 To 17 Step 1
    '    .Cells(12, i) = Heading(i)
    'Next i
    
    'End With
    '-------------------------------------------------------
    v = 1
    h = 1
    Command1.Visible = True
    Do

        If mytablex.EOF Then Exit Do

        '--------------------------------
        '--------------------------------

        vr = DoEvents()

        If Command1.Visible = False Then Exit Do

        '------------- verificamos la condicion
        If mytabley.State = 1 Then
            mytabley.Close
            Set mytabley = Nothing

        End If

        buf = "select * from " & xbasedatos & " where "
        buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
        buf = buf & " and local='" & extra_loquesea(local1) & "'"
        buf = buf & " and producto='" & "" & mytablex.Fields("producto") & "'"
        buf = buf & " and bodega='" & extra_loquesea(bodega) & "'"

        If VENTANEGRA = "S" Then
            buf = buf & " and acu<>'G'"

        End If

        buf = buf & " and (acu='S' or acu='T' or acu='C' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='E' or acu='N')"
        buf = buf & " and estado='2'"
        buf = buf & " order by fecha,hora"
        mytabley.Open buf, cn, adOpenStatic, adLockOptimistic
        xcosto = 0
        XCantidad = 0
        xsaldo = 0
        xent = 0
        xsal = 0
        txent = 0
        txsal = 0
        ttxtot = 0

        If mytabley.RecordCount = 0 Then GoTo ksigue

        With objExcel.ActiveSheet
            .Cells(v, 2) = "Periodo"
            .Cells(v, 5) = "'" & Format(Month(fechai), "00")
            v = v + 1
            .Cells(v, 2) = "Ruc"
            .Cells(v, 5) = "'" & busca_local(0)
            v = v + 1
            .Cells(v, 2) = "Apellidos y Nombres"
            .Cells(v, 5) = "'" & busca_local(1)
            v = v + 1
            .Cells(v, 2) = "Establecimiento"
            v = v + 1
            .Cells(v, 2) = "Codigo Existencia:"
            .Cells(v, 5) = "'" & mytablex.Fields("producto")
            v = v + 1
            .Cells(v, 2) = "Tipo:"
            v = v + 1
            .Cells(v, 2) = "Descripcio:"
            .Cells(v, 5) = "'" & mytablex.Fields("descripcio")
            v = v + 1
            .Cells(v, 2) = "Codigo de la unidad de Medida"
            .Cells(v, 5) = "'" & mytablex.Fields("unidad")
            v = v + 1
            .Cells(v, 2) = "Metodo de la evaluacion"
            v = v + 1
            .Cells(v, 6) = "Entradas"
            .Cells(v, 9) = "Salidas"
            .Cells(v, 12) = "Saldo"
            v = v + 1
    
            For I = 1 To 17 Step 1
                .Cells(v, I) = Heading(I)
            Next I

            v = v + 1

        End With

        Do

            If mytabley.EOF Then Exit Do

            objExcel.ActiveSheet.Cells(v, 1) = "'" & mytabley.Fields("fecha")
            objExcel.ActiveSheet.Cells(v, 2) = "'" & mytabley.Fields("tipo")
            objExcel.ActiveSheet.Cells(v, 3) = "'" & mytabley.Fields("serie")
            objExcel.ActiveSheet.Cells(v, 4) = "'" & mytabley.Fields("NUmero")

            objExcel.ActiveSheet.Cells(v, 5) = "'" & mytabley.Fields("Tipo")

            If "" & mytabley.Fields("acu") = "T" Or "" & mytabley.Fields("acu") = "A" Or "" & mytabley.Fields("acu") = "B" Or "" & mytabley.Fields("acu") = "C" Or "" & mytabley.Fields("acu") = "D" Or "" & mytabley.Fields("acu") = "G" Then
                objExcel.ActiveSheet.Cells(v, 9) = Val("" & mytabley.Fields("Cantidad"))
                objExcel.ActiveSheet.Cells(v, 10) = Val("" & mytablex.Fields("costou"))
                objExcel.ActiveSheet.Cells(v, 11) = Val("" & mytabley.Fields("Cantidad")) * Val("" & mytablex.Fields("costou"))
                XCantidad = Val("" & mytabley.Fields("cantidad"))
                xcosto = Val("" & mytabley.Fields("Cantidad")) * Val("" & mytablex.Fields("costou"))
                xsaldo = xsaldo - XCantidad
                xsal = xsal + XCantidad
                txsal = txsal + xcosto

            End If

            If "" & mytabley.Fields("acu") = "S" Or "" & mytabley.Fields("acu") = "J" Or "" & mytabley.Fields("acu") = "K" Or "" & mytabley.Fields("acu") = "L" Or "" & mytabley.Fields("acu") = "M" Or "" & mytabley.Fields("acu") = "P" Then
                objExcel.ActiveSheet.Cells(v, 6) = Val("" & mytabley.Fields("Cantidad"))
                objExcel.ActiveSheet.Cells(v, 7) = Val("" & mytablex.Fields("costou"))
                objExcel.ActiveSheet.Cells(v, 8) = Val("" & mytabley.Fields("Cantidad")) * Val("" & mytablex.Fields("costou"))
                XCantidad = Val("" & mytabley.Fields("cantidad"))
                xcosto = Val("" & mytabley.Fields("Cantidad")) * Val("" & mytablex.Fields("costou"))
                xsaldo = xsaldo + XCantidad
                xent = xent + XCantidad
                txent = txent + xcosto

            End If

            objExcel.ActiveSheet.Cells(v, 12) = xsaldo
            objExcel.ActiveSheet.Cells(v, 13) = Val("" & mytablex.Fields("costou"))
            objExcel.ActiveSheet.Cells(v, 14) = xsaldo * Val("" & mytablex.Fields("costou"))
            ttxtot = ttxtot + xsaldo * Val("" & mytablex.Fields("costou"))
            v = v + 1
            mytabley.MoveNext
        Loop
        objExcel.ActiveSheet.Cells(v, 9) = xsal
        objExcel.ActiveSheet.Cells(v, 6) = xent

        objExcel.ActiveSheet.Cells(v, 11) = txsal
        objExcel.ActiveSheet.Cells(v, 8) = txent
        objExcel.ActiveSheet.Cells(v, 14) = ttxtot

        v = v + 1
        mytabley.Close
ksigue:
        mytablex.MoveNext
    Loop
    'mytablex.Close

    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    Exit Sub
cmdd5612_err:
    MsgBox "Error en exportacion excell" + error$, 48, "Aviso"

End Sub

''''28/09/2017 kenyo Mejora formato Kardex Sunat
    
'Sub cuerpo_kardex_sunat1(mytablex As ADODB.Recordset)
'Dim mytabley As New ADODB.Recordset
'Dim mytablez As New ADODB.Recordset
'Dim sw As Integer
'Dim temp As String
'Dim buf As String
'Dim found As Integer
'Dim v, h As Double
'Dim vr
'Dim i As Integer
'Dim sdx As Double
'Dim XCantidad As Double
'Dim xsaldo As Double
'Dim xcosto As Double
'Dim xent As Double
'Dim xsal As Double
'Dim txent As Double
'Dim txsal As Double
'Dim ttxtot As Double
'Dim xprecio As Double
'
'''26/06/2017 kenyo costo kardex
'Dim xtotal As Double
'''26/06/2017 kenyo costo kardex
'
'
'    Dim Heading(18) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion
'    On Error GoTo cmd1d5612_err
'
'    Heading(1) = "Fecha"
'    Heading(2) = "Tipo"
'    Heading(3) = "Serie"
'    Heading(4) = "Numero"
'
'    Heading(5) = "API"
'    Heading(6) = "Temperatura"
'
'    Heading(7) = "Operacion"
'    Heading(8) = "Cantidad"
'    Heading(9) = "CostoUnitario"
'    Heading(10) = "CostoTotal"
'    Heading(11) = "Cantidad"
'    Heading(12) = "CostoUnitario"
'    Heading(13) = "CostoTotal"
'    Heading(14) = "Cantidad"
'    Heading(15) = "CostoUnitario"
'    Heading(16) = "CostoTotal"
'
'    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
'    '-------------------------------------------------------
'    'With objExcel.ActiveSheet
'    '.Cells(2, 2) = "Periodo"
'    '.Cells(3, 2) = "Ruc"
'    '.Cells(4, 2) = "Apellidos y Nombres"
'    '.Cells(5, 2) = "Establecimiento"
'    '.Cells(6, 2) = "Codigo Existencia:"
'    '.Cells(7, 2) = "Tipo:"
'    '.Cells(8, 2) = "Descripcio:"
'    '.Cells(9, 2) = "Codigo de la unidad de Medida"
'    '.Cells(10, 2) = "Metodo de la evaluacion"
'
'    'For i = 1 To 17 Step 1
'    '    .Cells(12, i) = Heading(i)
'    'Next i
'
''End With
''-------------------------------------------------------
'v = 1
'h = 1
'Command1.Visible = True
'Do
'If mytablex.EOF Then Exit Do
'
''--------------------------------
''--------------------------------
'
'vr = DoEvents()
'If Command1.Visible = False Then Exit Do
''------------- verificamos la condicion
'If mytabley.State = 1 Then
'   mytabley.Close
'   Set mytabley = Nothing
'End If
'
'buf = "select * from " & xbasedatos & " where "
'buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
'buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
'buf = buf & " and local='" & extra_loquesea(local1) & "'"
'buf = buf & " and producto='" & "" & mytablex.Fields("producto") & "'"
'buf = buf & " and bodega='" & extra_loquesea(bodega) & "'"
'If VENTANEGRA = "S" Then
'   buf = buf & " and acu<>'G' AND acu<>'P' "
'End If
'
'buf = buf & " and (acu='S' or acu='T' or acu='C' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='E' or acu='N')"
'buf = buf & " and estado='2'"
'buf = buf & " order by fecha,hora"
'mytabley.Open buf, cn, adOpenStatic, adLockOptimistic
'xcosto = 0
'XCantidad = 0
'xsaldo = 0
'xent = 0
'xsal = 0
'txent = 0
'txsal = 0
'ttxtot = 0
'
'''26/06/2017 kenyo costo kardex
'xtotal = 0
'''26/06/2017 kenyo costo kardex
'
'
''If mytabley.RecordCount = 0 Then GoTo ksigue
'With objExcel.ActiveSheet
'    .Cells(v, 2) = "Periodo"
'    .Cells(v, 5) = "'" & Format(Month(fechai), "00")
'    v = v + 1
'    .Cells(v, 2) = "Ruc"
'    .Cells(v, 5) = "'" & busca_empresas(0)
'    v = v + 1
'    .Cells(v, 2) = "Apellidos y Nombres"
'    .Cells(v, 5) = "'" & busca_empresas(1)
'    v = v + 1
'    .Cells(v, 2) = "Establecimiento"
'    v = v + 1
'    .Cells(v, 2) = "Codigo Existencia:"
'    .Cells(v, 5) = "'" & mytablex.Fields("producto")
'    v = v + 1
'    .Cells(v, 2) = "Tipo:"
'    v = v + 1
'    .Cells(v, 2) = "Descripcio:"
'    .Cells(v, 5) = "'" & mytablex.Fields("descripcio")
'    v = v + 1
'    .Cells(v, 2) = "Codigo de la unidad de Medida"
'    .Cells(v, 5) = "'" & mytablex.Fields("unidad")
'    v = v + 1
'    .Cells(v, 2) = "Metodo de la evaluacion"
'    v = v + 1
'.Cells(v, 8) = "Entradas"
'.Cells(v, 11) = "Salidas"
'.Cells(v, 14) = "Saldo"
'v = v + 1
'
'    For i = 1 To 17 Step 1
'        .Cells(v, i) = Heading(i)
'    Next i
'    v = v + 1
'End With
'
''---------------------------------------------------------
''inventario inicial
'XCantidad = 0
'xprecio = 0
'If mytablez.State = 1 Then mytablez.Close
'mytablez.Open "Select * from dsaldoini where local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "' and fecha='" & fechai & "'", cn, adOpenStatic, adLockOptimistic
'Do
'If mytablez.EOF Then Exit Do
'   XCantidad = XCantidad + Val("" & mytablez.Fields("cantidad")) * Val("" & mytablez.Fields("factor"))
'   xprecio = Val(Format(Val("" & mytablez.Fields("precio")), "0.000000"))
'mytablez.MoveNext
'Loop
'mytablez.Close
'xprecio = Val(Format(Val("" & mytablex.Fields("costou")), "0.000000"))
'If conigv = "" Or conigv = "S" Then
'   xprecio = xprecio
'End If
'
'If conigv = "N" Then
'   If Val("" & mytablex.Fields("igv")) > 0 Then
'      xprecio = xprecio / (1 + (Val("" & mytablex.Fields("igv")) / 100))
'      xprecio = Val(Format(xprecio, "0.00000"))
'   End If
'End If
'
'''26/06/2017 kenyo costo kardex
'
''objExcel.ActiveSheet.Cells(v, 8) = XCantidad
''objExcel.ActiveSheet.Cells(v, 9) = xprecio
''objExcel.ActiveSheet.Cells(v, 10) = XCantidad * xprecio
'
'''26/06/2017 kenyo costo kardex
'
'
'
'
'xcosto = XCantidad * xprecio
'xsaldo = xsaldo + XCantidad
'xent = xent + XCantidad
'txent = txent + xcosto
'
'''26/06/2017 kenyo costo kardex
'
''objExcel.ActiveSheet.Cells(v, 14) = xsaldo
''objExcel.ActiveSheet.Cells(v, 15) = xprecio
''objExcel.ActiveSheet.Cells(v, 16) = xsaldo * xprecio
'
'''26/06/2017 kenyo costo kardex
'
'ttxtot = ttxtot + xsaldo * xprecio
'v = v + 1
'
'Do
'If mytabley.EOF Then Exit Do
'objExcel.ActiveSheet.Cells(v, 1) = "'" & mytabley.Fields("fecha")
'objExcel.ActiveSheet.Cells(v, 2) = "'" & mytabley.Fields("tipo")
'objExcel.ActiveSheet.Cells(v, 3) = "'" & mytabley.Fields("serie")
'objExcel.ActiveSheet.Cells(v, 4) = "'" & mytabley.Fields("Numero")
'objExcel.ActiveSheet.Cells(v, 7) = "'" & mytabley.Fields("Tipo")
'
'XCantidad = Val("" & mytabley.Fields("cantidad")) * Val("" & mytabley.Fields("factor"))
'
''19/06/2017 kenyo NOTA DE CREDITO
''If "" & mytabley.Fields("acu") = "T" Or "" & mytabley.Fields("acu") = "A" Or "" & mytabley.Fields("acu") = "B" Or "" & mytabley.Fields("acu") = "C" Or "" & mytabley.Fields("acu") = "D" Or "" & mytabley.Fields("acu") = "G" Then
'If "" & mytabley.Fields("acu") = "T" Or "" & mytabley.Fields("acu") = "A" Or "" & mytabley.Fields("acu") = "B" Or "" & mytabley.Fields("acu") = "C" Or "" & mytabley.Fields("acu") = "D" Or "" & mytabley.Fields("acu") = "G" Or "" & mytabley.Fields("acu") = "E" Then
''19/06/2017 kenyo NOTA DE CREDITO
'
'objExcel.ActiveSheet.Cells(v, 11) = XCantidad
'
''19/06/2017 kenyo NOTA DE CREDITO
''objExcel.ActiveSheet.Cells(v, 12) =xprecio
''objExcel.ActiveSheet.Cells(v, 13) = XCantidad * xprecio
'
'''''25/09/2017 kenyo Testing Kardex Sunat
'    'objExcel.ActiveSheet.Cells(v, 12) = xtotal / xsaldo
'If xtotal = 0 And xsaldo = 0 Then
'objExcel.ActiveSheet.Cells(v, 12) = 0
'ElseIf xtotal > 0 And xsaldo = 0 Then
'objExcel.ActiveSheet.Cells(v, 12) = 0
'Else
'objExcel.ActiveSheet.Cells(v, 12) = xtotal / xsaldo
'End If
'''''25/09/2017 kenyo Testing Kardex Sunat
'
'
'objExcel.ActiveSheet.Cells(v, 13) = XCantidad * xprecio
''19/06/2017 kenyo NOTA DE CREDITO
'
'
''XCantidad = Val("" & mytabley.Fields("cantidad")) * Val("" & mytabley.Fields("factor"))
'
'xcosto = XCantidad * xprecio
'xsaldo = xsaldo - XCantidad
'xsal = xsal + XCantidad
'txsal = txsal + xcosto
'End If
'If "" & mytabley.Fields("acu") = "S" Or "" & mytabley.Fields("acu") = "J" Or "" & mytabley.Fields("acu") = "K" Or "" & mytabley.Fields("acu") = "L" Or "" & mytabley.Fields("acu") = "M" Or "" & mytabley.Fields("acu") = "P" Then
'objExcel.ActiveSheet.Cells(v, 8) = XCantidad
'
'''26/06/2017 kenyo costo kardex
'
'    'objExcel.ActiveSheet.Cells(v, 9) = xprecio
'    'objExcel.ActiveSheet.Cells(v, 10) = XCantidad * xprecio
'    'xcosto = XCantidad * xprecio
'
'objExcel.ActiveSheet.Cells(v, 9) = "'" & mytabley.Fields("precio")
'objExcel.ActiveSheet.Cells(v, 10) = XCantidad * mytabley.Fields("precio")
'
'xcosto = XCantidad * mytabley.Fields("precio")
'''26/06/2017 kenyo costo kardex
'
'
'
'xsaldo = xsaldo + XCantidad
'xent = xent + XCantidad
'
'txent = txent + xcosto
'
'
'xtotal = xtotal + xcosto
'
'End If
'
'objExcel.ActiveSheet.Cells(v, 14) = xsaldo
'
'''26/06/2017 kenyo costo kardex
''objExcel.ActiveSheet.Cells(v, 15) = xprecio
''objExcel.ActiveSheet.Cells(v, 16) = xsaldo * xprecio
'
'
'''''25/09/2017 kenyo Testing Kardex Sunat
'    'objExcel.ActiveSheet.Cells(v, 15) = xtotal / xsaldo
'If xtotal = 0 And xsaldo = 0 Then
'objExcel.ActiveSheet.Cells(v, 15) = 0
'ElseIf xtotal > 0 And xsaldo = 0 Then
'objExcel.ActiveSheet.Cells(v, 15) = 0
'Else
'objExcel.ActiveSheet.Cells(v, 15) = xtotal / xsaldo
'End If
'
'''''25/09/2017 kenyo Testing Kardex Sunat
'
'
'objExcel.ActiveSheet.Cells(v, 16) = xtotal
'''26/06/2017 kenyo costo kardex
'
'
'ttxtot = ttxtot + xsaldo * xprecio
'v = v + 1
'mytabley.MoveNext
'Loop
'objExcel.ActiveSheet.Cells(v, 11) = xsal
'objExcel.ActiveSheet.Cells(v, 8) = xent
'
'objExcel.ActiveSheet.Cells(v, 13) = txsal
'objExcel.ActiveSheet.Cells(v, 10) = txent
'
'
'''26/06/2017 kenyo costo kardex
'    'objExcel.ActiveSheet.Cells(v, 16) = ttxtot
'''26/06/2017 kenyo costo kardex
'
'v = v + 1
'mytabley.Close
'ksigue:
'mytablex.MoveNext
'Loop
''mytablex.Close
'
'Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
'Exit Sub
'cmd1d5612_err:
'MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
'End Sub
Sub cuerpo_kardex_sunat1(mytablex As ADODB.Recordset)

    Dim mytabley As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset

    Dim sw       As Integer

    Dim temp     As String

    Dim buf      As String

    Dim found    As Integer

    Dim v, h As Double

    Dim vr

    Dim I           As Integer

    Dim sdx         As Double

    Dim XCantidad   As Double

    Dim xsaldo      As Double

    Dim xcosto      As Double

    Dim xent        As Double

    Dim xsal        As Double

    Dim txent       As Double

    Dim txsal       As Double

    Dim ttxtot      As Double

    Dim xprecio     As Double

    '' 04/01/2018 Correcion costo con o sin igv en reporte kardex Sunat
    Dim xpreciosin  As Double

    '' 04/01/2018 Correcion costo con o sin igv en reporte kardex Sunat

    Dim xpreciosal  As Double

    Dim xtotal      As Double

    Dim Heading(18) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    On Error GoTo cmd1d5612_err
    
    Heading(1) = "Fecha"
    Heading(2) = "Tipo"
    Heading(3) = "Serie"
    Heading(4) = "Numero"
    
    Heading(5) = "OPERACIÓN"
    Heading(6) = "Cantidad"
    Heading(7) = "CostoU."
    Heading(8) = "CostoTotal"
    Heading(9) = "Cantidad"
    Heading(10) = "CostoU."
    Heading(11) = "CostoTotal"
    Heading(12) = "Cantidad"
    Heading(13) = "CostoU."
    Heading(14) = "CostoTotal"
    
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_ordenTamaño
    v = 1
    h = 1
    Command1.Visible = True
    Do

        If mytablex.EOF Then Exit Do

        '--------------------------------
        '--------------------------------

        vr = DoEvents()

        If Command1.Visible = False Then Exit Do

        '------------- verificamos la condicion
        If mytabley.State = 1 Then
            mytabley.Close
            Set mytabley = Nothing

        End If

        buf = "select * from " & xbasedatos & " where "
        buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
        buf = buf & " and local='" & extra_loquesea(local1) & "'"
        buf = buf & " and producto='" & "" & mytablex.Fields("producto") & "'"
        buf = buf & " and bodega='" & extra_loquesea(bodega) & "'"

        If VENTANEGRA = "S" Then
            buf = buf & " and acu<>'G' AND acu<>'P' "

        End If

        '07/08/2018 No descuenta stock en guia de remision
        buf = buf & " AND (L4<>'N' or L4 IS null) "
        '07/08/2018 No descuenta stock en guia de remision

        '' 29/11/2017 Correción  General del Stock
        buf = buf & " and acu1='' "
        '' 29/11/2017 Correción  General del Stock

        ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
        'buf = buf & " and (acu='S' or acu='T' or acu='C' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='E' or acu='N')"
        buf = buf & " and (acu='S' or acu='T' or acu='C' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='E' or acu='F')"
        ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

        buf = buf & " and estado='2'"
        buf = buf & " order by fecha,hora"
        mytabley.Open buf, cn, adOpenStatic, adLockOptimistic
        xcosto = 0
        XCantidad = 0
        xsaldo = 0
        xent = 0
        xsal = 0
        txent = 0
        txsal = 0
        ttxtot = 0
        xtotal = 0
        xpreciosal = 0

        v = v + 2

        With objExcel.ActiveSheet
            .Cells(v, 1) = "PERIODO:                                                                    "
            .Cells(v, 5) = "'" & Format(Month(fechai), "00")
            .Cells(v, 1).Font.bold = True

            For I = 1 To 14
                .Cells(v, I).Interior.color = RGB(255, 255, 255)
            Next
            v = v + 1
    
            .Cells(v, 1) = "RUC:                                                                         "
            .Cells(v, 5) = "'" & busca_local(0)
            .Cells(v, 1).Font.bold = True

            For I = 1 To 14
                .Cells(v, I).Interior.color = RGB(255, 255, 255)
            Next
            v = v + 1
    
            .Cells(v, 1) = "APELLIDOS Y NOMBRES, RAZON SOCIAL:                                                                   "
            .Cells(v, 1).Font.bold = True

            For I = 1 To 14
                .Cells(v, I).Interior.color = RGB(255, 255, 255)
            Next
            .Cells(v, 5) = "'" & busca_local(1)
            v = v + 1
    
            .Cells(v, 1) = "ESTABLECIMIENTO:                                                                   "
            .Cells(v, 1).Font.bold = True

            For I = 1 To 14
                .Cells(v, I).Interior.color = RGB(255, 255, 255)
            Next
            v = v + 1
    
            .Cells(v, 1) = "CÓDIGO DE EXISTENCIA:                                                                   "
            .Cells(v, 5) = "'" & mytablex.Fields("producto")
            .Cells(v, 1).Font.bold = True

            For I = 1 To 14
                .Cells(v, I).Interior.color = RGB(255, 255, 255)
            Next
            v = v + 1
    
            .Cells(v, 1) = "TIPO:                                                                             "
            .Cells(v, 1).Font.bold = True

            For I = 1 To 14
                .Cells(v, I).Interior.color = RGB(255, 255, 255)
            Next
            v = v + 1
    
            .Cells(v, 1) = "DESCRIPCIÓN:                                                                       "
            .Cells(v, 5) = "'" & mytablex.Fields("descripcio")
            .Cells(v, 1).Font.bold = True

            For I = 1 To 14
                .Cells(v, I).Interior.color = RGB(255, 255, 255)
            Next
            v = v + 1
    
            .Cells(v, 1) = "CÓDIGO DE LA UNIDAD DE MEDIDA:                                                                     "
            .Cells(v, 1).Font.bold = True

            For I = 1 To 14
                .Cells(v, I).Interior.color = RGB(255, 255, 255)
            Next
            .Cells(v, 5) = "'" & mytablex.Fields("unidad")
            v = v + 1
    
            .Cells(v, 1) = "MÉTODO DE EVALUACIÓN:                                     "
            .Cells(v, 1).Font.bold = True

            For I = 1 To 14
                .Cells(v, I).Interior.color = RGB(255, 255, 255)
            Next
            v = v + 1
            
            .Cells(v, 1) = "                             DOCUMENTO                          "
            .Cells(v, 5) = "    TIPO DE            "
            .Cells(v, 6) = "                       ENTRADAS              "
            .Cells(v, 9) = "                         SALIDAS               "
            .Cells(v, 12) = "                          SALDO                 "

            With objExcel.ActiveSheet
                .Range(.Cells(v, 1), .Cells(v + 1, 14)).Interior.color = RGB(192, 192, 250)
                .Range(.Cells(v, 1), .Cells(v, 14)).Font.bold = True
                .Range(.Cells(v, 5), .Cells(v + 1, 5)).Font.bold = True

            End With

            v = v + 1
      
            For I = 1 To 17 Step 1
                .Cells(v, I) = Heading(I)
            Next I

            v = v + 1
     
        End With
  
        '---------------------------------------------------------
        'inventario inicial
        XCantidad = 0
        xprecio = 0

        If mytablez.State = 1 Then mytablez.Close
        mytablez.Open "Select * from dsaldoini where local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "' and fecha='" & fechai & "'", cn, adOpenStatic, adLockOptimistic
        Do

            If mytablez.EOF Then Exit Do
            XCantidad = XCantidad + Val("" & mytablez.Fields("cantidad")) * Val("" & mytablez.Fields("factor"))
            xprecio = Val(Format(Val("" & mytablez.Fields("precio")), "0.000000"))

            mytablez.MoveNext
        Loop
        mytablez.Close
        xprecio = Val(Format(Val("" & mytablex.Fields("costou")), "0.000000"))

        If conigv = "" Or conigv = "S" Then
            xprecio = xprecio

        End If

        If conigv = "N" Then
            If Val("" & mytablex.Fields("igv")) > 0 Then
                xprecio = xprecio / (1 + (Val("" & mytablex.Fields("igv")) / 100))
                xprecio = Val(Format(xprecio, "0.00000"))

            End If

        End If

        xcosto = XCantidad * xprecio
        xsaldo = xsaldo + XCantidad
        xent = xent + XCantidad
        txent = txent + xcosto

        ttxtot = ttxtot + xsaldo * xprecio

        '' 16/01/2018 Stock Inicial en Kardex Producto y Sunat
        'v = v + 1
        If ChkSaldoInicial.Value = 1 Then
            objExcel.ActiveSheet.Cells(v, 3) = "       SALDO INICIAL     "

            Dim stockinicial      As String

            Dim Costostockinicial As String
        
            stockinicial = 0
            Costostockinicial = 0

            Dim fechafinal As String

            fechafinal = DateAdd("d", -1, fechai)
            fechafinal = Format(DateAdd("d", -1, fechai), "DD/MM/YYYY")
        
            If CVDate(fechai) <= CVDate(fechafinal) Then
                stockinicial = 0
                objExcel.ActiveSheet.Cells(v, h + 5) = "0"
            Else
                stockinicial = ObtieneStockInicial(extra_loquesea(local1), mytablex.Fields("producto"), extra_loquesea(bodega), fechainicial, fechafinal)
                objExcel.ActiveSheet.Cells(v, 6) = stockinicial
                objExcel.ActiveSheet.Cells(v, 7) = xprecio
                Costostockinicial = Val(Format(stockinicial * xprecio, "0.00000"))
                objExcel.ActiveSheet.Cells(v, 8) = Costostockinicial
                objExcel.ActiveSheet.Cells(v, 12) = stockinicial
          
                objExcel.ActiveSheet.Cells(v, 3).Font.bold = True
                objExcel.ActiveSheet.Cells(v, 5).Font.bold = True
                objExcel.ActiveSheet.Cells(v, 6).Font.bold = True
                objExcel.ActiveSheet.Cells(v, 7).Font.bold = True
                objExcel.ActiveSheet.Cells(v, 8).Font.bold = True
                objExcel.ActiveSheet.Cells(v, 12).Font.bold = True
            
                objExcel.ActiveSheet.Cells(v, 3).Font.color = RGB(254, 0, 0)
                objExcel.ActiveSheet.Cells(v, 5).Font.color = RGB(254, 0, 0)
                objExcel.ActiveSheet.Cells(v, 6).Font.color = RGB(254, 0, 0)
                objExcel.ActiveSheet.Cells(v, 7).Font.color = RGB(254, 0, 0)
                objExcel.ActiveSheet.Cells(v, 8).Font.color = RGB(254, 0, 0)
                objExcel.ActiveSheet.Cells(v, 12).Font.color = RGB(254, 0, 0)
        
                xsaldo = xsaldo + stockinicial
            
            End If

            v = v + 1

        End If

        '' 16/01/2018 Stock Inicial en Kardex Producto y Sunat

        Do

            If mytabley.EOF Then Exit Do

            With objExcel.ActiveSheet
                .Range(.Cells(v - 2, 1), .Cells(v, 14)).Borders.LineStyle = xlContinuous

            End With

            '' 16/01/2018 Stock Inicial en Kardex Producto y Sunat
            ' With objExcel.ActiveSheet
            '.Range(.Cells(v, 1), .Cells(v, 14)).Interior.color = RGB(160, 160, 160)
            'End With

            With objExcel.ActiveSheet
                .Range(.Cells(v, 1), .Cells(v, 14)).Interior.color = RGB(160, 160, 160)

            End With

            '' 16/01/2018 Stock Inicial en Kardex Producto y Sunat

            objExcel.ActiveSheet.Cells(v, 1) = "'" & mytabley.Fields("fecha")
            objExcel.ActiveSheet.Cells(v, 2) = "'" & mytabley.Fields("tipo")
            objExcel.ActiveSheet.Cells(v, 3) = "'" & mytabley.Fields("serie")
            objExcel.ActiveSheet.Cells(v, 4) = "'" & mytabley.Fields("Numero")

            objExcel.ActiveSheet.Cells(v, 5) = "'" & busca_tipooperacion(mytabley.Fields("tipo"))

            XCantidad = Val("" & mytabley.Fields("cantidad")) * Val("" & mytabley.Fields("factor"))

            ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
            'If "" & mytabley.Fields("acu") = "T" Or "" & mytabley.Fields("acu") = "A" Or "" & mytabley.Fields("acu") = "B" Or "" & mytabley.Fields("acu") = "C" Or "" & mytabley.Fields("acu") = "D" Or "" & mytabley.Fields("acu") = "G" Or "" & mytabley.Fields("acu") = "E" Then

            '07/08/2018 Nota de credito final
            'If "" & mytabley.Fields("acu") = "T" Or "" & mytabley.Fields("acu") = "A" Or "" & mytabley.Fields("acu") = "B" Or "" & mytabley.Fields("acu") = "C" Or "" & mytabley.Fields("acu") = "D" Or "" & mytabley.Fields("acu") = "G" Or "" & mytabley.Fields("acu") = "F" Then
            If "" & mytabley.Fields("acu") = "T" Or "" & mytabley.Fields("acu") = "A" Or "" & mytabley.Fields("acu") = "B" Or "" & mytabley.Fields("acu") = "C" Or "" & mytabley.Fields("acu") = "D" Or "" & mytabley.Fields("acu") = "G" Or "" & mytabley.Fields("acu") = "E" Then
                '07/08/2018 Nota de credito final

                ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

                objExcel.ActiveSheet.Cells(v, 9) = XCantidad

                If xtotal = 0 And xsaldo = 0 Then
                    objExcel.ActiveSheet.Cells(v, 10) = 0
                ElseIf xtotal > 0 And xsaldo = 0 Then
                    objExcel.ActiveSheet.Cells(v, 10) = 0
                Else
                    objExcel.ActiveSheet.Cells(v, 10) = Val(objExcel.ActiveSheet.Cells(v - 1, 13))  ' Costo de salida

                End If
    
                objExcel.ActiveSheet.Cells(v, 11) = XCantidad * Val(objExcel.ActiveSheet.Cells(v - 1, 13))

                xcosto = XCantidad * Val(objExcel.ActiveSheet.Cells(v - 1, 13))
                xsaldo = xsaldo - XCantidad
    
                xsal = xsal + XCantidad
                txsal = txsal + xcosto

            End If

            ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
            'If "" & mytabley.Fields("acu") = "S" Or "" & mytabley.Fields("acu") = "J" Or "" & mytabley.Fields("acu") = "K" Or "" & mytabley.Fields("acu") = "L" Or "" & mytabley.Fields("acu") = "M" Or "" & mytabley.Fields("acu") = "P" Then

            '07/08/2018 Nota de credito final
            'If "" & mytabley.Fields("acu") = "S" Or "" & mytabley.Fields("acu") = "J" Or "" & mytabley.Fields("acu") = "K" Or "" & mytabley.Fields("acu") = "L" Or "" & mytabley.Fields("acu") = "M" Or "" & mytabley.Fields("acu") = "P" Or "" & mytabley.Fields("acu") = "E" Then
            If "" & mytabley.Fields("acu") = "S" Or "" & mytabley.Fields("acu") = "J" Or "" & mytabley.Fields("acu") = "K" Or "" & mytabley.Fields("acu") = "L" Or "" & mytabley.Fields("acu") = "M" Or "" & mytabley.Fields("acu") = "P" Or "" & mytabley.Fields("acu") = "F" Then
                '07/08/2018 Nota de credito final

                ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

                xpreciosin = 0
                xpreciosin = Val(Format(Val("" & mytabley.Fields("precio")), "0.000000"))
     
                If conigv = "" Or conigv = "S" Then
                    xpreciosin = xpreciosin

                End If

                If conigv = "N" Then
                    If Val("" & mytablex.Fields("igv")) > 0 Then
                        xpreciosin = xpreciosin / (1 + (Val("" & mytablex.Fields("igv")) / 100))
                        xpreciosin = Val(Format(xpreciosin, "0.00000"))

                    End If

                End If

                objExcel.ActiveSheet.Cells(v, 6) = XCantidad
                objExcel.ActiveSheet.Cells(v, 7) = xpreciosin  ' Costo de enrtrada
        
                xpreciosal = xpreciosal + xpreciosin
    
                '11/06/2018 Actualiza Precio Promedio Ponderado Masivo
                'objExcel.ActiveSheet.Cells(v, 8) = XCantidad * xpreciosin
                objExcel.ActiveSheet.Cells(v, 8) = mytabley.Fields("total")
                '11/06/2018 Actualiza Precio Promedio Ponderado Masivo
    
                xcosto = XCantidad * xpreciosin
                xsaldo = xsaldo + XCantidad
    
                xent = xent + XCantidad
                txent = txent + xcosto
    
                '' 16/01/2018 Stock Inicial en Kardex Producto y Sunat
                If ChkSaldoInicial.Value = 1 Then
                    ' xent = xent + XCantidad
                    xent = xent + stockinicial
                    txent = txent + Costostockinicial

                End If

                '' 16/01/2018 Stock Inicial en Kardex Producto y Sunat
  
                xtotal = objExcel.ActiveSheet.Cells(v, 8) + Val(objExcel.ActiveSheet.Cells(v - 1, 14))
                '' 04/01/2018 Correcion costo con o sin igv en reporte kardex Sunat

            End If

            objExcel.ActiveSheet.Cells(v, 12) = xsaldo

            If xtotal = 0 And xsaldo = 0 Then
                objExcel.ActiveSheet.Cells(v, 13) = 0
            ElseIf xtotal > 0 And xsaldo = 0 Then
                objExcel.ActiveSheet.Cells(v, 13) = 0
            Else

                If objExcel.ActiveSheet.Cells(v, 7) <> "" Then
                    If xsaldo <> "0" Then
    
                        '13/08/2018 Integración FE - Pizzeria
                        '11/06/2018 Actualiza Precio Promedio Ponderado Masivo
                        objExcel.ActiveSheet.Cells(v, 13) = (Format(Val(xtotal / xsaldo), "0.00000"))  ' Costo de saldo
                        '11/06/2018 Actualiza Precio Promedio Ponderado Masivo
                        '13/08/2018 Integración FE - Pizzeria
    
                    Else
                        objExcel.ActiveSheet.Cells(v, 13) = objExcel.ActiveSheet.Cells(v, 7)

                    End If

                Else

                    objExcel.ActiveSheet.Cells(v, 13) = Val(objExcel.ActiveSheet.Cells(v - 1, 13))

                End If

            End If

            If objExcel.ActiveSheet.Cells(v, 7) <> "" Then
    
                If xsaldo <> 0 Then
    
                    '13/08/2018 Integración FE - Pizzeria
                    '11/06/2018 Actualiza Precio Promedio Ponderado Masivo
                    'objExcel.ActiveSheet.Cells(v, 14) = xtotal
                    objExcel.ActiveSheet.Cells(v, 14) = objExcel.ActiveSheet.Cells(v, 13) * objExcel.ActiveSheet.Cells(v, 12)
                    '11/06/2018 Actualiza Precio Promedio Ponderado Masivo
                    '13/08/2018 Integración FE - Pizzeria
     
                Else
                    objExcel.ActiveSheet.Cells(v, 14) = 0

                End If

            Else
                objExcel.ActiveSheet.Cells(v, 14) = objExcel.ActiveSheet.Cells(v, 13) * objExcel.ActiveSheet.Cells(v, 12)

            End If

            ttxtot = ttxtot + xsaldo * xprecio
            v = v + 1
            mytabley.MoveNext
        Loop
        objExcel.ActiveSheet.Cells(v, 5) = "TOTALES >"
        objExcel.ActiveSheet.Cells(v, 9) = xsal
        objExcel.ActiveSheet.Cells(v, 6) = xent

        objExcel.ActiveSheet.Cells(v, 6).Interior.color = RGB(248, 243, 53)
        objExcel.ActiveSheet.Cells(v, 9).Interior.color = RGB(248, 243, 53)
        objExcel.ActiveSheet.Cells(v - 1, 12).Interior.color = RGB(248, 243, 53)
        objExcel.ActiveSheet.Cells(v - 1, 12).Font.bold = True

        objExcel.ActiveSheet.Cells(v, 11) = txsal
        objExcel.ActiveSheet.Cells(v, 8) = txent
        objExcel.ActiveSheet.Cells(v, 5).Font.bold = True

        objExcel.ActiveSheet.Cells(v, 6).Font.bold = True
        objExcel.ActiveSheet.Cells(v, 8).Font.bold = True
        objExcel.ActiveSheet.Cells(v, 9).Font.bold = True
        objExcel.ActiveSheet.Cells(v, 11).Font.bold = True
       
        v = v + 1

        mytabley.Close
ksigue:
        mytablex.MoveNext
    Loop

    Set objExcel = Nothing
    Exit Sub
cmd1d5612_err:
    MsgBox "Error en exportacion excell" + error$, 48, "Aviso"

End Sub

Function Formato_ordenTamaño() As Boolean

    Dim I As Integer

    With objExcel.ActiveSheet
        .columns("B").ColumnWidth = 5
        .columns("C").ColumnWidth = 8
  
        .columns("F").ColumnWidth = 9
        .columns("G").ColumnWidth = 9
        .columns("H").ColumnWidth = 10
        .columns("I").ColumnWidth = 9
        .columns("J").ColumnWidth = 9
        .columns("K").ColumnWidth = 10
        .columns("L").ColumnWidth = 9
        .columns("M").ColumnWidth = 9
        .columns("N").ColumnWidth = 10
            
    End With

End Function

Function busca_tipooperacion(sw As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "Select sunatope from tipo where tipo='" & "" & sw & "'  ", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_tipooperacion = "" & mytablex.Fields("sunatope")

    End If

    mytablex.Close

End Function

''''28/09/2017 kenyo Mejora formato Kardex Sunat

Function busca_local(sw As Integer) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "Select codigo1, nombre from tlocal where codigo='" & "" & extra_loquesea(local1) & "'  ", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If sw = 0 Then
            busca_local = "" & mytablex.Fields("codigo1")

        End If

        If sw = 1 Then
            busca_local = "" & mytablex.Fields("nombre")

        End If

    End If

    mytablex.Close

End Function

Function busca_empresas(sw As Integer) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "Select * from empresa where codigo='01' ", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If sw = 0 Then
            busca_empresas = "" & mytablex.Fields("codigo1")

        End If

        If sw = 1 Then
            busca_empresas = "" & mytablex.Fields("nombre")

        End If

    End If

    mytablex.Close

End Function

Sub reporte_saldocf()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    found = sql_producto(mytablex)

    If found = 0 Then
        mytablex.Close
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_saldocf
    cuerpo_programa_saldocf mytablex
    Command1.Visible = False
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
    found = valida_wordpad(FileName)
     
    'genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1

End Sub

Sub cabecera_saldocf()

    Dim mytablex As Table

    Dim buf      As String

    Dim I        As Integer

    Dim found    As Integer

    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("ALMACEN : " & bodega, 25, 2, 0)
    found = formateaa("FECHA INICIO :" & fechai, 40, 2, 0)
    found = formateaa("FECHA FINAL  :" & fechaf, 40, 2, 0)
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)
    found = formateaa("Producto ", 11, 0, 0)
    found = formateaa("Descripcio ", 41, 0, 0)
    found = formateaa("Unid", 5, 0, 0)
    found = formateaa("Fact", 5, 0, 0)
    found = formateaa("Saldo", 8, 0, 0)
    found = formateaa("Cont1", 6, 0, 0)
    found = formateaa("Cont2", 6, 0, 0)
    found = formateaa("Cont3", 6, 2, 0)
    '--------------------------------------
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)

End Sub

Sub cuerpo_programa_saldocf(mytablex As ADODB.Recordset)

    Dim I         As Integer

    Dim mytabley  As New ADODB.Recordset

    Dim mytableyy As New ADODB.Recordset

    Dim sw        As Integer

    Dim temp      As String

    Dim buf       As String

    Dim sw1       As Integer

    Dim temp1     As String

    Dim buff      As String

    Dim buf1      As String

    Dim bufx      As String

    Dim saldoini  As Double

    Dim saldoindx As Double

    Dim sdx       As Double

    Dim sdx1      As Double

    Dim sdx2      As Double

    Dim sdx3      As Double

    Dim found     As Integer

    Dim sbuf      As String

    Dim dbuf      As String

    Dim xsaldo    As Double

    Dim vr

    sbuf = ""
    sw1 = 0
    xsaldo = 0
    Command1.Visible = True
    Do

        If mytablex.EOF Then Exit Do
        vr = DoEvents()

        xsaldo = 0
        dbuf = "select * from almacen  where producto='" & "" & mytablex.Fields("producto") & "' and local='" & extra_loquesea(local1) & "' and bodega='" & extra_loquesea(bodega) & "'"
        mytableyy.Open dbuf, cn, adOpenStatic, adLockOptimistic

        If mytableyy.RecordCount > 0 Then
            xsaldo = Val("" & mytableyy.Fields("saldo"))

        End If

        mytableyy.Close

        If Command1.Visible = False Then Exit Do

        '----------------------------------------------
        If proveedor <> "%" Then
            found = ver_proveedor("" & mytablex.Fields("producto"))

            If found = 0 Then GoTo seguy126

        End If

        '----------------------------------------------
        '------------- verificamos la condicion
        buf = "" & mytablex.Fields("familia")

        If sw = 0 Then
            sw = 1
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            buf = busca_familia("" & mytablex.Fields("familia"))
            found = formateaa(buf, 30, 0, 0)
            found = formateaa("*", 1, 2, 0)
            temp = "" & mytablex.Fields("familia")
            nlineas

        End If

        If "" & mytablex.Fields("familia") <> temp Then
            temp = "" & mytablex.Fields("familia")
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            buf = busca_familia("" & mytablex.Fields("familia"))
            found = formateaa(buf, 30, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas

        End If

        If sw1 = 0 Then
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("-", 1, 0, 0)
            found = formateaa("" & mytablex.Fields("subfamilia"), 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")
            sw1 = 1

        End If

        If temp1 <> "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia") Then
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("-", 1, 0, 0)
            found = formateaa("" & mytablex.Fields("subfamilia"), 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")

        End If

        found = formateaa("" & mytablex.Fields("producto"), 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("descripcio"), 40, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("unidad"), 4, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("factor"), 4, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor")))
        found = formateaa(buf, 8, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("_____", 6, 0, 0)
        found = formateaa("_____", 6, 0, 0)
        found = formateaa("_____", 6, 2, 0)
        nlineas
seguy126:
        mytablex.MoveNext
    Loop

End Sub

Private Sub VENTANEGRA_Click()

    If VENTANEGRA = "" Then
        VENTANEGRA = "S"
        Exit Sub

    End If

    If VENTANEGRA = "S" Then
        VENTANEGRA = ""
        Exit Sub

    End If

End Sub

Sub reporte_kardex_excell()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    ''''21/09/2017 kenyo Mejora reporte entradas salidas, kardex
    '   found = sql_producto(mytablex)
    found = sql_productoConMovimiento(mytablex)
    ''''21/09/2017 kenyo Mejora reporte entradas salidas, kardex

    If found = 0 Then
        mytablex.Close
        Exit Sub

    End If
    
    cuerpo_programa_kardex_excell mytablex
    Command1.Visible = False
    '------------------------------------
    mytablex.Close

End Sub

'''10/08/2017 kenyo Mejor Kardex Producto
Function Formato_orden(Num_Campos As Integer, Nombre_Campos() As String) As Boolean

    Dim I As Integer

    With objExcel.ActiveSheet
        'MsgBox Num_Campos
        'Formato de las celdas de los titulos
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Font.bold = True
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Interior.color = RGB(192, 192, 250)

        For I = 1 To Num_Campos Step 1
            .Cells(3, I) = Nombre_Campos(I)
        Next I
            
    End With

End Function

'''10/08/2017 kenyo Mejor Kardex Producto

Sub cuerpo_programa_kardex_excell(mytablex As ADODB.Recordset)

    Dim mytabley  As New ADODB.Recordset

    Dim sw        As Integer

    Dim temp      As String

    Dim buf       As String

    Dim mytablez  As New ADODB.Recordset

    Dim sw1       As Integer

    Dim temp1     As String

    Dim buf1      As String

    Dim bufx      As String

    Dim saldoini  As Double

    Dim saldoindx As Double

    Dim sdx       As Double

    Dim sdx1      As Double

    Dim sdx2      As Double

    Dim xbuf      As String

    Dim xcosto    As Double

    Dim mytablera As New ADODB.Recordset

    Dim found     As Integer

    Dim XCantidad As Double

    Dim vr

    Dim nsw         As Integer

    Dim Heading(20) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    Dim v           As Long

    Dim h           As Long

    saldoindx = 0
    sdx1 = 0
    sdx2 = 0
    sw1 = 0
    xcosto = 0

    '''10/08/2017 kenyo Mejor Kardex Producto
    Dim saldof As Double

    saldof = 0
    '''10/08/2017 kenyo Mejor Kardex Producto
      
    Command1.Visible = True

    On Error GoTo cmd561245_err
    
    '10/08/2017 kenyo Mejor Kardex Producto
   
    Heading(1) = "Familia"
    Heading(2) = "Subfamilia"
    Heading(3) = "Producto"
    Heading(4) = "Descripcion"
    
    Heading(5) = "Tipo"
    Heading(6) = "Serie"
    Heading(7) = "Numero"
    Heading(8) = "Fecha"
    Heading(9) = "Alm."
    Heading(10) = "Unidad"
    Heading(11) = "Factor"
    Heading(12) = "Entrada"
    Heading(13) = "Salida"
    Heading(14) = "Saldo"
    Heading(15) = "valor"
    Heading(16) = "Total"
    Heading(17) = "Hora"
    Heading(18) = "Nombre"
    Heading(19) = "Observa"
    
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_kardexx(19, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
  
    '''10/08/2017 kenyo Mejor Kardex Producto
    Call Formato_orden(19, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    '''10/08/2017 kenyo Mejor Kardex Producto
  
    v = 5
    h = 1

    Do

        If mytablex.EOF Then Exit Do
        If proveedor <> "%" Then
            found = ver_proveedor("" & mytablex.Fields("producto"))

            If found = 0 Then GoTo seguy2n

        End If

        vr = DoEvents()

        If Command1.Visible = False Then Exit Do

        buf = "" & mytablex.Fields("familia")

        If sw = 0 Then
            sw = 1
            temp = "" & mytablex.Fields("familia")
            objExcel.ActiveSheet.Cells(v - 1, h) = " "
            objExcel.ActiveSheet.Cells(v, h) = "" & mytablex.Fields("familia")
            objExcel.ActiveSheet.Cells(v, h).Font.bold = True
            objExcel.ActiveSheet.Cells(v, h).Font.color = RGB(62, 95, 138)

            If vesubfamilia = "N" Then
                objExcel.ActiveSheet.Cells(v, h + 2) = busca_familia("" & mytablex.Fields("familia"))
                objExcel.ActiveSheet.Cells(v, h + 2).Font.bold = True
                objExcel.ActiveSheet.Cells(v, h + 2).Font.color = RGB(62, 95, 138)
            Else
                objExcel.ActiveSheet.Cells(v, h + 1) = busca_familia("" & mytablex.Fields("familia"))
                objExcel.ActiveSheet.Cells(v, h + 1).Font.bold = True
                objExcel.ActiveSheet.Cells(v, h + 1).Font.color = RGB(62, 95, 138)

            End If
 
            '
            '  objExcel.ActiveSheet.Cells(v, i).Interior.color = RGB(248, 243, 53)
      
            v = v + 1
   
        End If

        If "" & mytablex.Fields("familia") <> temp Then
            temp = "" & mytablex.Fields("familia")
            objExcel.ActiveSheet.Cells(v - 1, h) = " "
            objExcel.ActiveSheet.Cells(v, h) = "" & mytablex.Fields("familia")
            objExcel.ActiveSheet.Cells(v, h).Font.bold = True
            objExcel.ActiveSheet.Cells(v, h).Font.color = RGB(62, 95, 138)

            If vesubfamilia = "N" Then
                objExcel.ActiveSheet.Cells(v, h + 2) = busca_familia("" & mytablex.Fields("familia"))
                objExcel.ActiveSheet.Cells(v, h + 2).Font.bold = True
                objExcel.ActiveSheet.Cells(v, h + 2).Font.color = RGB(62, 95, 138)
            Else
                objExcel.ActiveSheet.Cells(v, h + 1) = busca_familia("" & mytablex.Fields("familia"))
                objExcel.ActiveSheet.Cells(v, h + 1).Font.bold = True
                objExcel.ActiveSheet.Cells(v, h + 1).Font.color = RGB(62, 95, 138)

            End If
   
            v = v + 1

        End If

        If vesubfamilia = "S" Then
            If sw1 = 0 Then

                '       objExcel.ActiveSheet.Cells(v, h) = "" & mytablex.Fields("familia")
                '       objExcel.ActiveSheet.Cells(v, h + 1) = "" & mytablex.Fields("subfamilia")
                '        objExcel.ActiveSheet.Cells(v, h).Font.bold = True
                '        objExcel.ActiveSheet.Cells(v, h + 1).Font.bold = True
                '      v = v + 1
                '       temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")
                '       sw1 = 1
            End If

        End If

        If vesubfamilia = "S" Then
            If temp1 <> "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia") Then

                '   objExcel.ActiveSheet.Cells(v, h) = "" & mytablex.Fields("familia")
                '   objExcel.ActiveSheet.Cells(v, h + 1) = "" & mytablex.Fields("subfamilia")
                '    objExcel.ActiveSheet.Cells(v, h).Font.bold = True
                '     objExcel.ActiveSheet.Cells(v, h + 1).Font.bold = True
                '   v = v + 1
                '   temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")
            End If

        End If
  
        '''09/10/2017 kenyo Testing Reportes
        objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("familia")
  
        If vesubfamilia = "S" Then
            objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("subfamilia")

        End If
    
        objExcel.ActiveSheet.Cells(v, h + 2) = "'" & mytablex.Fields("producto")
        objExcel.ActiveSheet.Cells(v, h + 3) = "" & mytablex.Fields("descripcio")
        objExcel.ActiveSheet.Cells(v, h).Font.bold = True
        objExcel.ActiveSheet.Cells(v, h + 1).Font.bold = True
        objExcel.ActiveSheet.Cells(v, h + 2).Font.bold = True
        objExcel.ActiveSheet.Cells(v, h + 3).Font.bold = True
    
        ''' 16/01/2018 Stock Inicial en Kardex Producto y Sunat
        If ChkSaldoInicial.Value = 1 Then

            Dim stockinicial As String

            stockinicial = 0

            Dim fechafinal As String

            fechafinal = DateAdd("d", -1, fechai)
            fechafinal = Format(DateAdd("d", -1, fechai), "DD/MM/YYYY")
        
            If CVDate(fechai) <= CVDate(fechafinal) Then
                stockinicial = 0
                objExcel.ActiveSheet.Cells(v, h + 13) = "0"
            Else
                stockinicial = ObtieneStockInicial(extra_loquesea(local1), mytablex.Fields("producto"), extra_loquesea(bodega), fechainicial, fechafinal)
                objExcel.ActiveSheet.Cells(v, h + 13) = stockinicial
          
            End If
        
            objExcel.ActiveSheet.Cells(v, h + 11) = stockinicial
            objExcel.ActiveSheet.Cells(v, h + 8) = "Stock Inicial >"
        
            objExcel.ActiveSheet.Cells(v, h + 8).Font.bold = True
            objExcel.ActiveSheet.Cells(v, h + 11).Font.bold = True
            objExcel.ActiveSheet.Cells(v, h + 13).Font.bold = True
        
            objExcel.ActiveSheet.Cells(v, h + 8).Font.color = RGB(254, 0, 0)
            objExcel.ActiveSheet.Cells(v, h + 11).Font.color = RGB(254, 0, 0)
            objExcel.ActiveSheet.Cells(v, h + 13).Font.color = RGB(254, 0, 0)
       
        End If
    
        ''' 16/01/2018 Stock Inicial en Kardex Producto y Sunat
  
        '''09/10/2017 kenyo Testing Reportes

        '''09/10/2017 kenyo Testing Reportes
        'objExcel.ActiveSheet.Cells(v, h + 5) = "'" & mytablex.Fields("unidad")
        'objExcel.ActiveSheet.Cells(v, h + 6) = "'" & mytablex.Fields("factor")
        '''09/10/2017 kenyo Testing Reportes
 
        'SALDO INICIAL
        saldoini = 0
        xcosto = 0

        If mytablez.State = 1 Then mytablez.Close
        mytablez.Open "Select * from dsaldoini where local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "' and fecha='" & fechai & "'", cn, adOpenStatic, adLockOptimistic
        Do

            If mytablez.EOF Then Exit Do
            saldoini = saldoini + Val("" & mytablez.Fields("cantidad")) * Val("" & mytablez.Fields("factor"))
            xcosto = Val(Format(Val("" & mytablez.Fields("precio")), "0.000000"))
            mytablez.MoveNext
        Loop
        mytablez.Close

        ''10/08/2017 kenyo Mejor Kardex Producto
        If Trim(quecosto) = "COSTOULTIMO" Then
            xcosto = Val(Format(Val("" & mytablex.Fields("costou")), "0.000000"))

        End If

        If Trim(quecosto) = "COSTOPROMEDIO" Then
            xcosto = Val(Format(Val("" & mytablex.Fields("COSTOP")), "0.000000"))

        End If

        '''10/08/2017 kenyo Mejor Kardex Producto

        If conigv = "" Or conigv = "S" Then
            xcosto = xcosto

        End If

        If conigv = "N" Then
            If Val("" & mytablex.Fields("igv")) > 0 Then
                xcosto = xcosto / (1 + (Val("" & mytablex.Fields("igv")) / 100))
                xcosto = Val(Format(xcosto, "0.00000"))

            End If

        End If

        bufx = "" & saldoini

        saldoindx = saldoini

        ''' 16/01/2018 Stock Inicial en Kardex Producto y Sunat
        If ChkSaldoInicial.Value = 1 Then
            saldoindx = stockinicial

        End If

        ''' 16/01/2018 Stock Inicial en Kardex Producto y Sunat

        xbuf = "" & saldoindx 'calcula_saldo(saldoini, Val("" & mytablex.Fields("factor")))

        '''10/08/2017 kenyo Mejor Kardex Producto
        'objExcel.ActiveSheet.Cells(v, h + 7) = "" & xbuf
        'objExcel.ActiveSheet.Cells(v, h + 10) = xcosto
        'sdx = Val("" & mytablex.Fields("costou")) * Val(bufx)
        'objExcel.ActiveSheet.Cells(v, h + 11) = "" & sdx
        '''10/08/2017 kenyo Mejor Kardex Producto

        v = v + 1
        sdx2 = 0
        sdx1 = 0
        nsw = 0
        '-------ahora las transacciones------------
        buf = "select * from " & xbasedatos & " where "
        buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
        buf = buf & " and local='" & extra_loquesea(local1) & "'"
        buf = buf & " and producto='" & "" & mytablex.Fields("producto") & "'"
        buf = buf & " and bodega='" & extra_loquesea(bodega) & "'"

        If VENTANEGRA = "S" Then
            buf = buf & " and acu<>'G' AND acu<>'P' "

        End If

        '07/08/2018 No descuenta stock en guia de remision
        buf = buf & " AND (L4<>'N' or L4 IS null) "
        '07/08/2018 No descuenta stock en guia de remision

        ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
        '  buf = buf & " and (acu='1' or acu='S' or acu='T' or acu='C' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='E' or acu='N')"
        buf = buf & " and (acu='1' or acu='S' or acu='T' or acu='C' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='E' or acu='N' or acu='F')"
        ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

        buf = buf & " and acu1=''"
        buf = buf & " and estado='2'"
        buf = buf & " order by fecha,hora"

        'MsgBox buf
        If mytabley.State = 1 Then mytabley.Close
        mytabley.Open buf, cn, adOpenStatic, adLockOptimistic
        Do

            If mytabley.EOF Then Exit Do
            vr = DoEvents()

            If Command1.Visible = False Then Exit Do
            Command1.Caption = " " & mytabley.Fields("local") & " " & mytabley.Fields("tipo") & " " & mytabley.Fields("serie") & " " & mytabley.Fields("numero") & " " & mytabley.Fields("fecha")

            XCantidad = Val("" & mytabley.Fields("cantidad")) * Val("" & mytabley.Fields("factor"))

            ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
            '19/06/2017 kenyo NOTA DE CREDITO
            'If "" & mytabley.Fields("acu") = "1" Or "" & mytabley.Fields("acu") = "T" Or "" & mytabley.Fields("acu") = "A" Or "" & mytabley.Fields("acu") = "B" Or "" & mytabley.Fields("acu") = "C" Or "" & mytabley.Fields("acu") = "D" Or "" & mytabley.Fields("acu") = "G" Or "" & mytabley.Fields("acu") = "N"  Then
            'If "" & mytabley.Fields("acu") = "1" Or "" & mytabley.Fields("acu") = "T" Or "" & mytabley.Fields("acu") = "A" Or "" & mytabley.Fields("acu") = "B" Or "" & mytabley.Fields("acu") = "C" Or "" & mytabley.Fields("acu") = "D" Or "" & mytabley.Fields("acu") = "G" Or "" & mytabley.Fields("acu") = "N" Or "" & mytabley.Fields("acu") = "E" Then
            'If "" & mytabley.Fields("acu") = "1" Or "" & mytabley.Fields("acu") = "T" Or "" & mytabley.Fields("acu") = "A" Or "" & mytabley.Fields("acu") = "B" Or "" & mytabley.Fields("acu") = "C" Or "" & mytabley.Fields("acu") = "D" Or "" & mytabley.Fields("acu") = "G" Or "" & mytabley.Fields("acu") = "N" Or "" & mytabley.Fields("acu") = "F" Then

            '07/08/2018 Nota de credito final
            If "" & mytabley.Fields("acu") = "1" Or "" & mytabley.Fields("acu") = "T" Or "" & mytabley.Fields("acu") = "A" Or "" & mytabley.Fields("acu") = "B" Or "" & mytabley.Fields("acu") = "C" Or "" & mytabley.Fields("acu") = "D" Or "" & mytabley.Fields("acu") = "G" Or "" & mytabley.Fields("acu") = "N" Or "" & mytabley.Fields("acu") = "E" Then
                '07/08/2018 Nota de credito final

                '19/06/2017 kenyo NOTA DE CREDITO
                ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

                objExcel.ActiveSheet.Cells(v, h) = "'" & mytabley.Fields("familia")

                If vesubfamilia = "S" Then
                    objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytabley.Fields("subfamilia")

                End If

                objExcel.ActiveSheet.Cells(v, h + 2) = "'" & mytabley.Fields("producto")
                objExcel.ActiveSheet.Cells(v, h + 3) = "'" & mytabley.Fields("descripcio")

                objExcel.ActiveSheet.Cells(v, h + 4) = "'" & mytabley.Fields("tipo")
                objExcel.ActiveSheet.Cells(v, h + 5) = "'" & mytabley.Fields("serie")
                objExcel.ActiveSheet.Cells(v, h + 6) = "'" & mytabley.Fields("numero")
                objExcel.ActiveSheet.Cells(v, h + 7) = "'" & Format("" & mytabley.Fields("fecha"), "dd/mm/yyyy")
                objExcel.ActiveSheet.Cells(v, h + 8) = "'" & mytabley.Fields("bodega")
                objExcel.ActiveSheet.Cells(v, h + 9) = "'" & mytabley.Fields("unidad")
                objExcel.ActiveSheet.Cells(v, h + 10) = "'" & mytabley.Fields("factor")
                objExcel.ActiveSheet.Cells(v, h + 11) = ""
                objExcel.ActiveSheet.Cells(v, h + 12) = XCantidad
  
                saldoindx = saldoindx - XCantidad
   
                buf = "" & saldoindx 'calcula_saldo(saldoindx, Val("" & mytabley.Fields("factor")))
                sdx2 = sdx2 + XCantidad
                objExcel.ActiveSheet.Cells(v, h + 13) = "" & buf
   
                '''10/08/2017 kenyo Mejor Kardex Producto
                'objExcel.ActiveSheet.Cells(v, h + 10) = "" & xcosto
                'sdx = xcosto * Val("" & mytabley.Fields("cantidad"))
                'buf = Format(sdx, "0.00")
                'objExcel.ActiveSheet.Cells(v, h + 11) = "" & buf
                '''10/08/2017 kenyo Mejor Kardex Producto
   
                objExcel.ActiveSheet.Cells(v, h + 16) = "" & mytabley.Fields("hora")
   
                ''' 30/11/2017 Correción  General del Sistema Parte I ' Muestra observacion en kardex producto
   
                objExcel.ActiveSheet.Cells(v, h + 17) = "" & busca_nombre(mytabley)

                If mytablez.State = 1 Then mytablez.Close
                mytablez.Open "Select observa from factura where local='" & extra_loquesea(local1) & "' and tipo='" & mytabley.Fields("tipo") & "' and serie='" & mytabley.Fields("serie") & "' and numero='" & mytabley.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic

                If mytablez.RecordCount > 0 Then
                    objExcel.ActiveSheet.Cells(v, h + 18) = "" & mytablez.Fields("observa")

                End If

                mytablez.Close
                ''' 30/11/2017 Correción  General del Sistema Parte I
   
                v = v + 1

            End If
 
            ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
            '19/06/2017 kenyo NOTA DE CREDITO
            ' If "" & mytabley.Fields("acu") = "S" Or "" & mytabley.Fields("acu") = "J" Or "" & mytabley.Fields("acu") = "K" Or "" & mytabley.Fields("acu") = "L" Or "" & mytabley.Fields("acu") = "M" Or "" & mytabley.Fields("acu") = "P" Or "" & mytabley.Fields("acu") = "E" Then
   
            'If "" & mytabley.Fields("acu") = "S" Or "" & mytabley.Fields("acu") = "J" Or "" & mytabley.Fields("acu") = "K" Or "" & mytabley.Fields("acu") = "L" Or "" & mytabley.Fields("acu") = "M" Or "" & mytabley.Fields("acu") = "P" Then
        
            '07/08/2018 Nota de credito final
            ' If "" & mytabley.Fields("acu") = "S" Or "" & mytabley.Fields("acu") = "J" Or "" & mytabley.Fields("acu") = "K" Or "" & mytabley.Fields("acu") = "L" Or "" & mytabley.Fields("acu") = "M" Or "" & mytabley.Fields("acu") = "P" Or "" & mytabley.Fields("acu") = "E" Then
            If "" & mytabley.Fields("acu") = "S" Or "" & mytabley.Fields("acu") = "J" Or "" & mytabley.Fields("acu") = "K" Or "" & mytabley.Fields("acu") = "L" Or "" & mytabley.Fields("acu") = "M" Or "" & mytabley.Fields("acu") = "P" Or "" & mytabley.Fields("acu") = "F" Then
    
                '07/08/2018 Nota de credito final
    
                '19/06/2017 kenyo NOTA DE CREDITO
                ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

                objExcel.ActiveSheet.Cells(v, h + 2) = "'" & mytabley.Fields("producto")
                objExcel.ActiveSheet.Cells(v, h + 3) = "'" & mytabley.Fields("descripcio")
                objExcel.ActiveSheet.Cells(v, h) = "'" & mytabley.Fields("familia")
  
                If vesubfamilia = "S" Then
                    objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytabley.Fields("subfamilia")

                End If
   
                objExcel.ActiveSheet.Cells(v, h + 4) = "'" & mytabley.Fields("tipo")
                objExcel.ActiveSheet.Cells(v, h + 5) = "'" & mytabley.Fields("serie")
                objExcel.ActiveSheet.Cells(v, h + 6) = "'" & mytabley.Fields("numero")
                objExcel.ActiveSheet.Cells(v, h + 7) = "'" & Format("" & mytabley.Fields("fecha"), "dd/mm/yyyy")
                objExcel.ActiveSheet.Cells(v, h + 8) = "'" & mytabley.Fields("bodega")
                objExcel.ActiveSheet.Cells(v, h + 9) = "'" & mytabley.Fields("unidad")
                objExcel.ActiveSheet.Cells(v, h + 10) = "'" & mytabley.Fields("factor")
                objExcel.ActiveSheet.Cells(v, h + 11) = "" & mytabley.Fields("cantidad")
                objExcel.ActiveSheet.Cells(v, h + 12) = ""
   
                'xcosto = 0
                'If Val("" & mytabley.Fields("precio")) = 0 Then
                '   xcosto = mytablex.Fields("costou")
                'End If
   
                saldoindx = saldoindx + XCantidad
                'buf = saldoindx
                buf = "" & saldoindx 'calcula_saldo(saldoindx, Val("" & mytabley.Fields("factor")))
                sdx1 = sdx1 + XCantidad
   
                '' 16/01/2018 Stock Inicial en Kardex Producto y Sunat
                If ChkSaldoInicial.Value = 1 Then
                    sdx1 = sdx1 + stockinicial

                End If

                '' 16/01/2018 Stock Inicial en Kardex Producto y Sunat
   
                objExcel.ActiveSheet.Cells(v, h + 13) = "" & buf
   
                '''10/08/2017 kenyo Mejor Kardex Producto
                'objExcel.ActiveSheet.Cells(v, h + 10) = "" & xcosto
                'sdx = xcosto * saldoindx
                'buf = Format(sdx, "0.00")
                'objExcel.ActiveSheet.Cells(v, h + 11) = "" & buf
                '''10/08/2017 kenyo Mejor Kardex Producto
  
                objExcel.ActiveSheet.Cells(v, h + 16) = "" & mytabley.Fields("hora")
                objExcel.ActiveSheet.Cells(v, h + 17) = "" & busca_nombre(mytabley)
   
                If mytablez.State = 1 Then mytablez.Close
                mytablez.Open "Select observa from factura where local='" & extra_loquesea(local1) & "' and tipo='" & mytabley.Fields("tipo") & "' and serie='" & mytabley.Fields("serie") & "' and numero='" & mytabley.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic

                If mytablez.RecordCount > 0 Then
                    objExcel.ActiveSheet.Cells(v, h + 18) = "" & mytablez.Fields("observa")

                End If

                mytablez.Close
   
                v = v + 1

            End If

siguiente_buscan:
            mytabley.MoveNext
        Loop

        buf = "" & sdx1
        objExcel.ActiveSheet.Cells(v, h + 11) = "" & buf
   
        buf = "" & sdx2
        objExcel.ActiveSheet.Cells(v, h + 12) = "" & buf
   
        '''10/08/2017 kenyo Mejor Kardex Producto
        saldof = sdx1 - sdx2
    
        objExcel.ActiveSheet.Cells(v, h + 13) = "" & Format(saldof, "0.00000")  ' Saldo Final
        objExcel.ActiveSheet.Cells(v, h + 14) = "" & xcosto    ' Costo Ultimo de producto
        objExcel.ActiveSheet.Cells(v, h + 15) = "" & Format(xcosto * saldof, "0.00000")  ' Costo Valorizado
   
        Dim I As Integer

        For I = 12 To 16
            objExcel.ActiveSheet.Cells(v, I).Font.bold = True
            objExcel.ActiveSheet.Cells(v, I).Interior.color = RGB(248, 243, 53)
        Next
        '''10/08/2017 kenyo Mejor Kardex Producto
    
        v = v + 1
        sdx1 = 0
        sdx2 = 0
     
        '''10/08/2017 kenyo Mejor Kardex Producto
        saldof = 0
        '''10/08/2017 kenyo Mejor Kardex Producto

        '---------------------------------------
seguy2n:
        mytablex.MoveNext
    Loop
    Exit Sub
cmd561245_err:
    MsgBox "Aviso en cuerpo excell " + error$, 48, "Aviso"
    Exit Sub

End Sub

Function Formato_kardexx(Num_Campos As Integer, Nombre_Campos() As String) As Boolean

    Dim I As Integer

    With objExcel.ActiveSheet
        'MsgBox Num_Campos
        'Formato de las celdas de los titulos
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 8)).Font.bold = True
        
        For I = 1 To Num_Campos Step 1
            .Cells(3, I) = Nombre_Campos(I)
        Next I

        'hasta aki pa colocar los titulos
        
        'a partir de aki ta claro que es pa darle el ancho a las celdas ;-)
       
        .columns("A").ColumnWidth = 8
        
        If repinv.vesubfamilia = "S" Then
            .columns("B").ColumnWidth = 9
        Else
            .columns("B").ColumnWidth = 0

        End If
        
        .columns("C").ColumnWidth = 10
        .columns("D").ColumnWidth = 35
        .columns("E").ColumnWidth = 4
        .columns("F").ColumnWidth = 5
        .columns("G").ColumnWidth = 8
        .columns("H").ColumnWidth = 10
        .columns("I").ColumnWidth = 4
        .columns("J").ColumnWidth = 6.5
        .columns("K").ColumnWidth = 5.5
        
        .columns("L").ColumnWidth = 8
        .columns("M").ColumnWidth = 8
        .columns("N").ColumnWidth = 8
        .columns("O").ColumnWidth = 8
        .columns("P").ColumnWidth = 8
    
    End With

End Function

Function busca_nombre(mytabley As ADODB.Recordset) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select nombre from factura where local='" & "" & mytabley.Fields("local") & "' and tipo='" & "" & mytabley.Fields("tipo") & "' and serie='" & "" & mytabley.Fields("serie") & "' and numero='" & "" & mytabley.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_nombre = Trim("" & mytablex.Fields("nombre"))

    End If

    mytablex.Close
 
End Function

Sub excel_entrada_salida()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    ''''21/09/2017 kenyo Mejora reporte entradas salidas, kardex
    'found = sql_producto(mytablex)
    found = sql_productoConMovimiento(mytablex)
    ''''21/09/2017 kenyo Mejora reporte entradas salidas, kardex

    If found = 0 Then
        mytablex.Close
        Exit Sub

    End If

    cuerpo_programa_entrada_salida mytablex
    Command1.Visible = False
    mytablex.Close

End Sub

'' 16/01/2018 Stock Inicial en Kardex Producto y Sunat

''''21/09/2017 kenyo Mejora reporte entradas salidas, kardex
'Sub cuerpo_programa_entrada_salida(mytablex As ADODB.Recordset)
'Dim mytabley As New ADODB.Recordset
'Dim mytablez As New ADODB.Recordset
'Dim sw As Integer
'Dim temp As String
'Dim buf As String
'Dim sw1 As Integer
'Dim temp1 As String
'Dim buf1 As String
'Dim bufx As String
'Dim found As Integer
'Dim v, h As Double
'Dim vr
'Dim xventa As Double
'Dim xcompra As Double
'Dim xentrada As Double
'Dim xsalida As Double
'Dim xnotae As Double
'Dim xnotas As Double
'Dim sdx As Double
'
'Dim xprecio As Double
'
'
'sw1 = 0
'
'    Dim Heading(15) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion
'    On Error GoTo cmd561212_err
'
'
'    ''''21/09/2017 kenyo Mejora reporte entradas salidas, kardex
'     Heading(1) = "Familia"
'     Heading(2) = "SubFamilia"
'    Heading(3) = "Cod. Producto"
'    Heading(4) = "Descripcion"
'    Heading(5) = "Unidad"
'    Heading(6) = "Factor"
'    Heading(7) = "Compras"
'    Heading(8) = "Entradas"
'    Heading(9) = "Ventas"
'    Heading(10) = "Salidas"
'    Heading(11) = "NCredVta"
'    Heading(12) = "NCredCom"
'    Heading(13) = "Saldo"
'
'    Heading(14) = "Costo"
'    Heading(15) = "Total"
'
'If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
'Call Formato_ExcelEntradasSalidas(15, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
'v = 5
'h = 1
'Command1.Visible = True
'Do
'
'If mytablex.EOF Then Exit Do
'If proveedor <> "%" Then
'   found = ver_proveedor("" & mytablex.Fields("producto"))
'   If found = 0 Then GoTo seguy88
'End If
'vr = DoEvents()
'If Command1.Visible = False Then Exit Do
'buf = "" & mytablex.Fields("familia")
'If sw = 0 Then
'   sw = 1
'    objExcel.ActiveSheet.Cells(4, 1) = " "
'
'            objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("familia")
'             objExcel.ActiveSheet.Cells(v, h).Font.bold = True
'            objExcel.ActiveSheet.Cells(v, h).Font.color = RGB(62, 95, 138)
'
'            If vesubfamilia = "S" Then
'                objExcel.ActiveSheet.Cells(v, h + 1) = busca_familia("" & mytablex.Fields("familia"))
'                objExcel.ActiveSheet.Cells(v, h + 1).Font.bold = True
'                objExcel.ActiveSheet.Cells(v, h + 1).Font.color = RGB(62, 95, 138)
'                objExcel.ActiveSheet.Cells(v, h + 2) = ""
'            Else
'                objExcel.ActiveSheet.Cells(v, h + 2) = busca_familia("" & mytablex.Fields("familia"))
'                objExcel.ActiveSheet.Cells(v, h + 2).Font.bold = True
'                objExcel.ActiveSheet.Cells(v, h + 2).Font.color = RGB(62, 95, 138)
'            End If
'
'
'            objExcel.ActiveSheet.Cells(v, h + 3) = ""
'            objExcel.ActiveSheet.Cells(v, h + 4) = ""
'            objExcel.ActiveSheet.Cells(v, h + 5) = ""
'            objExcel.ActiveSheet.Cells(v, h + 6) = ""
'            objExcel.ActiveSheet.Cells(v, h + 7) = ""
'            v = v + 1
'   temp = "" & mytablex.Fields("familia")
'End If
'If "" & mytablex.Fields("familia") <> temp Then
'   temp = "" & mytablex.Fields("familia")
'            v = v + 1
'            objExcel.ActiveSheet.Cells(v - 1, h) = " "
'              objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("familia")
'             objExcel.ActiveSheet.Cells(v, h).Font.bold = True
'            objExcel.ActiveSheet.Cells(v, h).Font.color = RGB(62, 95, 138)
'
'            If vesubfamilia = "S" Then
'                objExcel.ActiveSheet.Cells(v, h + 1) = busca_familia("" & mytablex.Fields("familia"))
'                objExcel.ActiveSheet.Cells(v, h + 1).Font.bold = True
'                objExcel.ActiveSheet.Cells(v, h + 1).Font.color = RGB(62, 95, 138)
'                objExcel.ActiveSheet.Cells(v, h + 2) = ""
'            Else
'                objExcel.ActiveSheet.Cells(v, h + 2) = busca_familia("" & mytablex.Fields("familia"))
'                objExcel.ActiveSheet.Cells(v, h + 2).Font.bold = True
'                objExcel.ActiveSheet.Cells(v, h + 2).Font.color = RGB(62, 95, 138)
'            End If
'
'
'            objExcel.ActiveSheet.Cells(v, h + 3) = ""
'            objExcel.ActiveSheet.Cells(v, h + 4) = ""
'            objExcel.ActiveSheet.Cells(v, h + 5) = ""
'            objExcel.ActiveSheet.Cells(v, h + 6) = ""
'            objExcel.ActiveSheet.Cells(v, h + 7) = ""
'            v = v + 1
'   End If
'
'''''21/09/2017 kenyo Mejora reporte entradas salidas, kardex
'If vesubfamilia = "S" Then
'    If sw1 = 0 Then
'                objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("familia")
'                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("subfamilia")
'
'                objExcel.ActiveSheet.Cells(v, h).Font.bold = True
'                objExcel.ActiveSheet.Cells(v, h + 1).Font.bold = True
'
'
'                objExcel.ActiveSheet.Cells(v, h + 2) = ""
'                objExcel.ActiveSheet.Cells(v, h + 3) = ""
'                objExcel.ActiveSheet.Cells(v, h + 4) = ""
'                objExcel.ActiveSheet.Cells(v, h + 5) = ""
'                objExcel.ActiveSheet.Cells(v, h + 6) = ""
'                objExcel.ActiveSheet.Cells(v, h + 7) = ""
'                v = v + 1
'       temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")
'       sw1 = 1
'    End If
'End If
'
'
'If vesubfamilia = "S" Then
'    If temp1 <> "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia") Then
'             objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("familia")
'                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("subfamilia")
'                objExcel.ActiveSheet.Cells(v, h).Font.bold = True
'                objExcel.ActiveSheet.Cells(v, h + 1).Font.bold = True
'                objExcel.ActiveSheet.Cells(v, h + 2) = ""
'                objExcel.ActiveSheet.Cells(v, h + 3) = ""
'                objExcel.ActiveSheet.Cells(v, h + 4) = ""
'                objExcel.ActiveSheet.Cells(v, h + 5) = ""
'                objExcel.ActiveSheet.Cells(v, h + 6) = ""
'                objExcel.ActiveSheet.Cells(v, h + 7) = ""
'                v = v + 1
'       temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")
'    End If
'End If
'
'''''21/09/2017 kenyo Mejora reporte entradas salidas, kardex
''SALDO ALMACEN
'xventa = 0
'xcompra = 0
'xentrada = 0
'xsalida = 0
'xnotae = 0
'xnotas = 0
'If mytablez.State = 1 Then mytablez.Close
'buf = "select * from DETALLE where "
'buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
'buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
'
''''06/08/2017 kenyo Testing Completo al Sistema
'        'buf = buf & " and local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "'"
'buf = buf & " and local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "'"
''''06/08/2017 kenyo Testing Completo al Sistema
'
'
''''27/07/2017 kenyo Testing Completo al Sistema
'buf = buf & " and estado='2' "
''''27/07/2017 kenyo Testing Completo al Sistema
'
''''21/08/2017 kenyo Guia de Salida con Factura
'buf = buf & " and acu1='' "
''''21/08/2017 kenyo Guia de Salida con Factura
'
'
'
'mytablez.Open buf, cn, adOpenStatic, adLockOptimistic
'Do
'If mytablez.EOF Then Exit Do
'If "" & mytablez.Fields("acu") = "A" Or "" & mytablez.Fields("acu") = "B" Or "" & mytablez.Fields("acu") = "C" Or "" & mytablez.Fields("acu") = "D" Or "" & mytablez.Fields("acu") = "G" Then 'VENTAS
'xventa = xventa + Val("" & mytablez.Fields("cantidad")) * Val("" & mytablez.Fields("factor"))
'End If
'If "" & mytablez.Fields("acu") = "J" Or "" & mytablez.Fields("acu") = "K" Or "" & mytablez.Fields("acu") = "L" Or "" & mytablez.Fields("acu") = "M" Or "" & mytablez.Fields("acu") = "P" Then 'COMPRAS
'xcompra = xcompra + Val("" & mytablez.Fields("cantidad")) * Val("" & mytablez.Fields("factor"))
'End If
'If "" & mytablez.Fields("acu") = "S" Then  'entrada
'xentrada = xentrada + Val("" & mytablez.Fields("cantidad")) * Val("" & mytablez.Fields("factor"))
'End If
'If "" & mytablez.Fields("acu") = "T" Then  'salida
'xsalida = xsalida + Val("" & mytablez.Fields("cantidad")) * Val("" & mytablez.Fields("factor"))
'End If
'If "" & mytablez.Fields("acu") = "E" Then  'Nota credito entrada
'xnotae = xnotae + Val("" & mytablez.Fields("cantidad")) * Val("" & mytablez.Fields("factor"))
'End If
'If "" & mytablez.Fields("acu") = "N" Then  'Nota CREDITO SALIDA
'xnotas = xnotas + Val("" & mytablez.Fields("cantidad")) * Val("" & mytablez.Fields("factor"))
'End If
'
'
'mytablez.MoveNext
'Loop
'mytablez.Close
'
''''21/09/2017 kenyo Mejora reporte entradas salidas, kardex
'objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("familia")
'
'objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("subfamilia")
'
'objExcel.ActiveSheet.Cells(v, h + 2) = "'" & mytablex.Fields("producto")
'objExcel.ActiveSheet.Cells(v, h + 3) = "" & mytablex.Fields("descripcio")
'objExcel.ActiveSheet.Cells(v, h + 4) = "" & mytablex.Fields("unidad")
'objExcel.ActiveSheet.Cells(v, h + 5) = "" & mytablex.Fields("factor")
'objExcel.ActiveSheet.Cells(v, h + 6) = xcompra
'objExcel.ActiveSheet.Cells(v, h + 7) = xentrada
'objExcel.ActiveSheet.Cells(v, h + 8) = xventa
'objExcel.ActiveSheet.Cells(v, h + 9) = xsalida
'objExcel.ActiveSheet.Cells(v, h + 10) = xnotae
'objExcel.ActiveSheet.Cells(v, h + 11) = xnotas
'
'
'
'
'
''19/06/2017 kenyo NOTA DE CREDITO
''sdx = xcompra + xentrada - xventa - xsalida + xnotae - xnotas
'sdx = xcompra + xentrada - xventa - xsalida - xnotae - xnotas
''19/06/2017 kenyo NOTA DE CREDITO
'
'
'
'objExcel.ActiveSheet.Cells(v, h + 12) = sdx
'
'
'
'If Trim(quecosto) = "COSTOULTIMO" Then
'    xprecio = Val(Format(Val("" & mytablex.Fields("costou")), "0.00"))
'End If
'If Trim(quecosto) = "COSTOPROMEDIO" Then
'    xprecio = Val(Format(Val("" & mytablex.Fields("costop")), "0.00"))
'End If
'
'objExcel.ActiveSheet.Cells(v, h + 13) = xprecio
'objExcel.ActiveSheet.Cells(v, h + 14) = "" & Format(Val(xprecio) * sdx, "0.00000")  ' Costo Valorizado
'
'
'v = v + 1
''--------------------------------------
'sigueme1:
'seguy88:
'mytablex.MoveNext
'Loop
'
'        ''20/07/2017 kenyo reporte de entradas Salidas  en el modulo de reportes. Correcion de formato
'            'objExcel.ActiveSheet.Cells(v, h) = "" & suma2
'            'objExcel.ActiveSheet.Cells(v, h + 1) = "" & suma1
'            objExcel.ActiveSheet.Cells(v, h) = ""
'            objExcel.ActiveSheet.Cells(v, h + 1) = ""
'        ''20/07/2017 kenyo reporte de entradas Salidas  en el modulo de reportes. Correcion de formato
'
'
'            objExcel.ActiveSheet.Cells(v, h + 2) = ""
'            objExcel.ActiveSheet.Cells(v, h + 3) = ""
'            objExcel.ActiveSheet.Cells(v, h + 4) = ""
'            objExcel.ActiveSheet.Cells(v, h + 5) = ""
'            objExcel.ActiveSheet.Cells(v, h + 6) = ""
'            objExcel.ActiveSheet.Cells(v, h + 7) = ""
'            v = v + 1
'
'Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
'Exit Sub
'cmd561212_err:
'MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
'Exit Sub
'
'End Sub

Sub cuerpo_programa_entrada_salida(mytablex As ADODB.Recordset)

    Dim mytabley As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset

    Dim sw       As Integer

    Dim temp     As String

    Dim buf      As String

    Dim sw1      As Integer

    Dim temp1    As String

    Dim buf1     As String

    Dim bufx     As String

    Dim found    As Integer

    Dim v, h As Double

    Dim vr

    Dim xventa   As Double

    Dim xcompra  As Double

    Dim xentrada As Double

    Dim xsalida  As Double

    Dim xnotae   As Double

    Dim xnotas   As Double

    Dim sdx      As Double

    Dim xprecio  As Double

    sw1 = 0

    Dim Heading(16) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    On Error GoTo cmd561212_err
   
    Heading(1) = "Familia"
    Heading(2) = "SubFamilia"
    Heading(3) = "Cod. Producto"
    Heading(4) = "Descripcion"
    Heading(5) = "Unidad"
    Heading(6) = "Factor"
    Heading(7) = "SaldoInicial"
    Heading(8) = "Compras"
    Heading(9) = "Entradas"
    Heading(10) = "Ventas"
    Heading(11) = "Salidas"
    
    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
    'Heading(12) = "NCredVta"
    'Heading(13) = "NCredCom"
    Heading(12) = "NCredVta"
    Heading(13) = "NDebVta"
    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
    
    Heading(14) = "Saldo"
    
    Heading(15) = "Costo"
    Heading(16) = "Total"
    
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_ExcelEntradasSalidas(16, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    v = 5
    h = 1
    Command1.Visible = True
    Do

        If mytablex.EOF Then Exit Do
        If proveedor <> "%" Then
            found = ver_proveedor("" & mytablex.Fields("producto"))

            If found = 0 Then GoTo seguy88

        End If

        vr = DoEvents()

        If Command1.Visible = False Then Exit Do
        buf = "" & mytablex.Fields("familia")

        If sw = 0 Then
            sw = 1
            objExcel.ActiveSheet.Cells(4, 1) = " "
   
            objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("familia")
            objExcel.ActiveSheet.Cells(v, h).Font.bold = True
            objExcel.ActiveSheet.Cells(v, h).Font.color = RGB(62, 95, 138)
            
            If vesubfamilia = "S" Then
                objExcel.ActiveSheet.Cells(v, h + 1) = busca_familia("" & mytablex.Fields("familia"))
                objExcel.ActiveSheet.Cells(v, h + 1).Font.bold = True
                objExcel.ActiveSheet.Cells(v, h + 1).Font.color = RGB(62, 95, 138)
                objExcel.ActiveSheet.Cells(v, h + 2) = ""
            Else
                objExcel.ActiveSheet.Cells(v, h + 2) = busca_familia("" & mytablex.Fields("familia"))
                objExcel.ActiveSheet.Cells(v, h + 2).Font.bold = True
                objExcel.ActiveSheet.Cells(v, h + 2).Font.color = RGB(62, 95, 138)

            End If
     
            objExcel.ActiveSheet.Cells(v, h + 3) = ""
            objExcel.ActiveSheet.Cells(v, h + 4) = ""
            objExcel.ActiveSheet.Cells(v, h + 5) = ""
            objExcel.ActiveSheet.Cells(v, h + 6) = ""
            objExcel.ActiveSheet.Cells(v, h + 7) = ""
            v = v + 1
            temp = "" & mytablex.Fields("familia")

        End If

        If "" & mytablex.Fields("familia") <> temp Then
            temp = "" & mytablex.Fields("familia")
            v = v + 1
            objExcel.ActiveSheet.Cells(v - 1, h) = " "
            objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("familia")
            objExcel.ActiveSheet.Cells(v, h).Font.bold = True
            objExcel.ActiveSheet.Cells(v, h).Font.color = RGB(62, 95, 138)
            
            If vesubfamilia = "S" Then
                objExcel.ActiveSheet.Cells(v, h + 1) = busca_familia("" & mytablex.Fields("familia"))
                objExcel.ActiveSheet.Cells(v, h + 1).Font.bold = True
                objExcel.ActiveSheet.Cells(v, h + 1).Font.color = RGB(62, 95, 138)
                objExcel.ActiveSheet.Cells(v, h + 2) = ""
            Else
                objExcel.ActiveSheet.Cells(v, h + 2) = busca_familia("" & mytablex.Fields("familia"))
                objExcel.ActiveSheet.Cells(v, h + 2).Font.bold = True
                objExcel.ActiveSheet.Cells(v, h + 2).Font.color = RGB(62, 95, 138)

            End If
     
            objExcel.ActiveSheet.Cells(v, h + 3) = ""
            objExcel.ActiveSheet.Cells(v, h + 4) = ""
            objExcel.ActiveSheet.Cells(v, h + 5) = ""
            objExcel.ActiveSheet.Cells(v, h + 6) = ""
            objExcel.ActiveSheet.Cells(v, h + 7) = ""
            v = v + 1

        End If

        ''''21/09/2017 kenyo Mejora reporte entradas salidas, kardex
        If vesubfamilia = "S" Then
            If sw1 = 0 Then
                objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("familia")
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("subfamilia")
                
                objExcel.ActiveSheet.Cells(v, h).Font.bold = True
                objExcel.ActiveSheet.Cells(v, h + 1).Font.bold = True
           
                objExcel.ActiveSheet.Cells(v, h + 2) = ""
                objExcel.ActiveSheet.Cells(v, h + 3) = ""
                objExcel.ActiveSheet.Cells(v, h + 4) = ""
                objExcel.ActiveSheet.Cells(v, h + 5) = ""
                objExcel.ActiveSheet.Cells(v, h + 6) = ""
                objExcel.ActiveSheet.Cells(v, h + 7) = ""
                v = v + 1
                temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")
                sw1 = 1

            End If

        End If

        If vesubfamilia = "S" Then
            If temp1 <> "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia") Then
                objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("familia")
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("subfamilia")
                objExcel.ActiveSheet.Cells(v, h).Font.bold = True
                objExcel.ActiveSheet.Cells(v, h + 1).Font.bold = True
                objExcel.ActiveSheet.Cells(v, h + 2) = ""
                objExcel.ActiveSheet.Cells(v, h + 3) = ""
                objExcel.ActiveSheet.Cells(v, h + 4) = ""
                objExcel.ActiveSheet.Cells(v, h + 5) = ""
                objExcel.ActiveSheet.Cells(v, h + 6) = ""
                objExcel.ActiveSheet.Cells(v, h + 7) = ""
                v = v + 1
                temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")

            End If

        End If

        xventa = 0
        xcompra = 0
        xentrada = 0
        xsalida = 0
        xnotae = 0
        xnotas = 0

        If mytablez.State = 1 Then mytablez.Close
        buf = "select * from DETALLE where "
        buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

        buf = buf & " and local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "'"
        buf = buf & " and estado='2' "
        buf = buf & " and acu1='' "

        '07/08/2018 No descuenta stock en guia de remision
        buf = buf & " AND (L4<>'N' or L4 IS null) "
        '07/08/2018 No descuenta stock en guia de remision

        mytablez.Open buf, cn, adOpenStatic, adLockOptimistic
        Do

            If mytablez.EOF Then Exit Do
            If "" & mytablez.Fields("acu") = "A" Or "" & mytablez.Fields("acu") = "B" Or "" & mytablez.Fields("acu") = "C" Or "" & mytablez.Fields("acu") = "D" Or "" & mytablez.Fields("acu") = "G" Then 'VENTAS
                xventa = xventa + Val("" & mytablez.Fields("cantidad")) * Val("" & mytablez.Fields("factor"))

            End If

            If "" & mytablez.Fields("acu") = "J" Or "" & mytablez.Fields("acu") = "K" Or "" & mytablez.Fields("acu") = "L" Or "" & mytablez.Fields("acu") = "M" Or "" & mytablez.Fields("acu") = "P" Then 'COMPRAS
                xcompra = xcompra + Val("" & mytablez.Fields("cantidad")) * Val("" & mytablez.Fields("factor"))

            End If

            If "" & mytablez.Fields("acu") = "S" Then  'entrada
                xentrada = xentrada + Val("" & mytablez.Fields("cantidad")) * Val("" & mytablez.Fields("factor"))

            End If

            If "" & mytablez.Fields("acu") = "T" Then  'salida
                xsalida = xsalida + Val("" & mytablez.Fields("cantidad")) * Val("" & mytablez.Fields("factor"))

            End If

            If "" & mytablez.Fields("acu") = "E" Then  'Nota credito entrada
                xnotae = xnotae + Val("" & mytablez.Fields("cantidad")) * Val("" & mytablez.Fields("factor"))

            End If

            ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
            'If "" & mytablez.Fields("acu") = "N" Then  'Nota CREDITO SALIDA
            If "" & mytablez.Fields("acu") = "F" Then  'Nota DEBITO
                xnotas = xnotas + Val("" & mytablez.Fields("cantidad")) * Val("" & mytablez.Fields("factor"))

            End If

            ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

            mytablez.MoveNext
        Loop
        mytablez.Close

        objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("familia")

        objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("subfamilia")

        objExcel.ActiveSheet.Cells(v, h + 2) = "'" & mytablex.Fields("producto")
        objExcel.ActiveSheet.Cells(v, h + 3) = "" & mytablex.Fields("descripcio")
        objExcel.ActiveSheet.Cells(v, h + 4) = "" & mytablex.Fields("unidad")
        objExcel.ActiveSheet.Cells(v, h + 5) = "" & mytablex.Fields("factor")

        '' 16/01/2018 Stock Inicial en Kardex Producto y Sunat
        If ChkSaldoInicial.Value = 1 Then

            Dim stockinicial As String

            stockinicial = 0

            Dim fechafinal As String

            fechafinal = DateAdd("d", -1, fechai)
            fechafinal = Format(DateAdd("d", -1, fechai), "DD/MM/YYYY")
        
            If CVDate(fechai) <= CVDate(fechafinal) Then
                stockinicial = 0
                objExcel.ActiveSheet.Cells(v, h + 6) = "0"
            Else
                stockinicial = ObtieneStockInicial(extra_loquesea(local1), mytablex.Fields("producto"), extra_loquesea(bodega), fechainicial, fechafinal)
                objExcel.ActiveSheet.Cells(v, h + 6) = stockinicial

            End If

        End If
    
        '' 16/01/2018 Stock Inicial en Kardex Producto y Sunat

        objExcel.ActiveSheet.Cells(v, h + 7) = xcompra
        objExcel.ActiveSheet.Cells(v, h + 8) = xentrada
        objExcel.ActiveSheet.Cells(v, h + 9) = xventa
        objExcel.ActiveSheet.Cells(v, h + 10) = xsalida
        objExcel.ActiveSheet.Cells(v, h + 11) = xnotae
        objExcel.ActiveSheet.Cells(v, h + 12) = xnotas

        ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
        'sdx = xcompra + xentrada - xventa - xsalida - xnotae - xnotas
        '25/06/2018 Testing Almacen General
        
        '07/08/2018 Nota de credito final
        'sdx = xcompra + xentrada - xventa - xsalida + xnotae - xnotas
        sdx = xcompra + xentrada - xventa - xsalida - xnotae + xnotas
        '07/08/2018 Nota de credito final
        
        '25/06/2018 Testing Almacen General

        ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

        '' 16/01/2018 Stock Inicial en Kardex Producto y Sunat
        If ChkSaldoInicial.Value = 1 Then
            sdx = sdx + stockinicial

        End If
 
        '' 16/01/2018 Stock Inicial en Kardex Producto y Sunat

        objExcel.ActiveSheet.Cells(v, h + 13) = sdx

        If Trim(quecosto) = "COSTOULTIMO" Then
            xprecio = Val(Format(Val("" & mytablex.Fields("costou")), "0.00"))

        End If

        If Trim(quecosto) = "COSTOPROMEDIO" Then
            xprecio = Val(Format(Val("" & mytablex.Fields("costop")), "0.00"))

        End If

        objExcel.ActiveSheet.Cells(v, h + 14) = xprecio
        objExcel.ActiveSheet.Cells(v, h + 15) = "" & Format(Val(xprecio) * sdx, "0.00000")  ' Costo Valorizado

        v = v + 1
        '--------------------------------------
sigueme1:
seguy88:
        mytablex.MoveNext
    Loop

    objExcel.ActiveSheet.Cells(v, h) = ""
    objExcel.ActiveSheet.Cells(v, h + 1) = ""
  
    objExcel.ActiveSheet.Cells(v, h + 2) = ""
    objExcel.ActiveSheet.Cells(v, h + 3) = ""
    objExcel.ActiveSheet.Cells(v, h + 4) = ""
    objExcel.ActiveSheet.Cells(v, h + 5) = ""
    objExcel.ActiveSheet.Cells(v, h + 6) = ""
    objExcel.ActiveSheet.Cells(v, h + 7) = ""
    v = v + 1

    Set objExcel = Nothing
    Exit Sub
cmd561212_err:
    MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
    Exit Sub

End Sub

'' 16/01/2018 Stock Inicial en Kardex Producto y Sunat

'11/06/2018 Actualiza Precio Promedio Ponderado Masivo
Private Sub ppp_Click()

    If opcion2 = "100" Then 'reporte de kardex

        Dim found    As Integer

        Dim mytablex As New ADODB.Recordset

        Dim buf      As String

        found = sql_productoConMovimientoSinReceta(mytablex)

        If found = 0 Then
            mytablex.Close
            Exit Sub

        End If

        cuerpo_kardex_sunat69 mytablex
        Command1.Visible = False
        Exit Sub

    End If

End Sub

'11/06/2018 Actualiza Precio Promedio Ponderado Masivo

'11/06/2018 Actualiza Precio Promedio Ponderado Masivo
Function sql_productoConMovimientoSinReceta(mytablex As ADODB.Recordset)

    Dim buf      As String

    Dim buf2     As String

    Dim buftotal As String

    buf = "select * from producto where producto.producto not  in  (select producto from receta) and producto.producto in (select producto from detalle d where ESTADO='2' AND  d.fecha>='" & Format(fechai, "YYYYMMDD") & "' and d.fecha<='" & Format(fechaf, "YYYYMMDD") & "' and local='" & extra_loquesea(local1) & "'"

    If bodega = "%" Then
        buf = buf & "  ) "
    Else
        buf = buf & " and bodega='" & extra_loquesea(bodega) & "' )   "

    End If

    buf = buf & "   and producto like '" & producto & "'"

    If Barras <> "%" Then
        buf = buf & " and barras like '" & Barras & "'"

    End If

    If descripcio <> "%" Then
        buf = buf & " and descripcio like '" & descripcio & "'"

    End If

    If familia <> "%" Then
        buf = buf & " and familia like '" & extra_loquesea(familia) & "'"

    End If

    If subfamilia <> "%" Then
        buf = buf & " and subfamilia like '" & subfamilia & "'"

    End If

    If seccion <> "%" Then
        buf = buf & " and seccion like '" & seccion & "'"

    End If

    If categoria <> "%" Then
        buf = buf & " and categoria like '" & categoria & "'"

    End If

    If linea <> "%" Then
        buf = buf & " and linea like '" & linea & "'"

    End If

    If color <> "%" Then
        buf = buf & " and color like '" & color & "'"

    End If

    If marca <> "%" Then
        buf = buf & " and marca like '" & marca & "'"

    End If

    If fechavi <> "%" And fechavf <> "%" Then
        If IsDate(fechavi) And IsDate(fechavf) Then
            buf = buf & "  fechavence>='" & Format(fechavi, "YYYYMMDD") & "'"
            buf = buf & " and fechavence<='" & Format(fechavf, "YYYYMMDD") & "' "

        End If

    End If

    If igv = "EXENTO" Then
        buf = buf & " and igv=0"

    End If

    If igv = "GRAVADO" Then
        buf = buf & " and igv>0"

    End If

    buf2 = "select * from producto where  producto.costoreceta='N'  and producto.producto in (select producto from detalle d where ESTADO='2' AND  d.fecha>='" & Format(fechai, "YYYYMMDD") & "' and d.fecha<='" & Format(fechaf, "YYYYMMDD") & "' and local='" & extra_loquesea(local1) & "'"

    If bodega = "%" Then
        buf2 = buf2 & "  ) "
    Else
        buf2 = buf2 & " and bodega='" & extra_loquesea(bodega) & "' )   "

    End If

    buf2 = buf2 & "   and producto like '" & producto & "'"

    If Barras <> "%" Then
        buf2 = buf2 & " and barras like '" & Barras & "'"

    End If

    If descripcio <> "%" Then
        buf2 = buf2 & " and descripcio like '" & descripcio & "'"

    End If

    If familia <> "%" Then
        buf2 = buf2 & " and familia like '" & extra_loquesea(familia) & "'"

    End If

    If subfamilia <> "%" Then
        buf2 = buf2 & " and subfamilia like '" & subfamilia & "'"

    End If

    If seccion <> "%" Then
        buf2 = buf2 & " and seccion like '" & seccion & "'"

    End If

    If categoria <> "%" Then
        buf2 = buf2 & " and categoria like '" & categoria & "'"

    End If

    If linea <> "%" Then
        buf2 = buf2 & " and linea like '" & linea & "'"

    End If

    If color <> "%" Then
        buf2 = buf2 & " and color like '" & color & "'"

    End If

    If marca <> "%" Then
        buf2 = buf2 & " and marca like '" & marca & "'"

    End If

    If fechavi <> "%" And fechavf <> "%" Then
        If IsDate(fechavi) And IsDate(fechavf) Then
            buf2 = buf2 & "  fechavence>='" & Format(fechavi, "YYYYMMDD") & "'"
            buf2 = buf2 & " and fechavence<='" & Format(fechavf, "YYYYMMDD") & "' "

        End If

    End If

    If igv = "EXENTO" Then
        buf2 = buf2 & " and igv=0"

    End If

    If igv = "GRAVADO" Then
        buf2 = buf2 & " and igv>0"

    End If

    buftotal = ""
    buftotal = buf & " union all  " & buf2
    buftotal = buftotal & " order by familia,Subfamilia,descripcio"

    mytablex.Open buftotal, cn, adOpenStatic, adLockOptimistic
    sql_productoConMovimientoSinReceta = 1

End Function

'11/06/2018 Actualiza Precio Promedio Ponderado Masivo

'11/06/2018 Actualiza Precio Promedio Ponderado Masivo
Sub cuerpo_kardex_sunat69(mytablex As ADODB.Recordset)

    Dim cantidadF   As Double

    Dim CostoPromF  As Double

    Dim CostoTotalF As Double

    Dim XCantidad   As Double

    XCantidad = 0
    cantidadF = 0
    CostoPromF = 0
    CostoTotalF = 0

    Dim mytabley   As New ADODB.Recordset

    Dim mytablez   As New ADODB.Recordset

    Dim sw         As Integer

    Dim temp       As String

    Dim buf        As String

    Dim found      As Integer

    Dim xpreciosin As Double

    Dim xpreciosal As Double

    Dim xtotal     As Double

    On Error GoTo cmd1d5612_err

    Command1.Visible = True
    Do

        If mytablex.EOF Then Exit Do

        If Command1.Visible = False Then Exit Do

        '------------- verificamos la condicion
        If mytabley.State = 1 Then
            mytabley.Close
            Set mytabley = Nothing

        End If

        buf = "select * from " & xbasedatos & " where "
        buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
        buf = buf & " and local='" & extra_loquesea(local1) & "'"
        buf = buf & " and producto='" & "" & mytablex.Fields("producto") & "'"
        buf = buf & " and bodega='" & extra_loquesea(bodega) & "'"

        If VENTANEGRA = "S" Then
            buf = buf & " and acu<>'G' AND acu<>'P' "

        End If

        buf = buf & " and acu1='' "
        buf = buf & " and (acu='S' or acu='T' or acu='C' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='E' or acu='N')"
        buf = buf & " and estado='2'"
        buf = buf & " order by fecha,hora"
        mytabley.Open buf, cn, adOpenStatic, adLockOptimistic

        CostoPromF = busca_costoiProducto(mytablex.Fields("producto"))
        CostoPromF = Format((CostoPromF), "0.00000")

        Do

            If mytabley.EOF Then Exit Do
    
            XCantidad = Val("" & mytabley.Fields("cantidad")) * Val("" & mytabley.Fields("factor"))
    
            ' SALIDAS
            If "" & mytabley.Fields("acu") = "T" Or "" & mytabley.Fields("acu") = "A" Or "" & mytabley.Fields("acu") = "B" Or "" & mytabley.Fields("acu") = "C" Or "" & mytabley.Fields("acu") = "D" Or "" & mytabley.Fields("acu") = "G" Or "" & mytabley.Fields("acu") = "E" Then
                cantidadF = cantidadF - XCantidad
                CostoTotalF = cantidadF * CostoPromF

            End If
    
            ' ENTRADAS
            If "" & mytabley.Fields("acu") = "S" Or "" & mytabley.Fields("acu") = "J" Or "" & mytabley.Fields("acu") = "K" Or "" & mytabley.Fields("acu") = "L" Or "" & mytabley.Fields("acu") = "M" Or "" & mytabley.Fields("acu") = "P" Then
       
                cantidadF = cantidadF + XCantidad
                CostoPromF = Val(Format(Val("" & mytabley.Fields("precio")), "0.00000"))

                If CostoTotalF > 0 Then
                    CostoTotalF = CostoTotalF + mytabley.Fields("total")

                End If
    
                If CostoTotalF > 0 Then
                    CostoPromF = Val(Format((CostoTotalF / cantidadF), "0.00000"))

                End If
        
            End If
    
            mytabley.MoveNext
        Loop
    
        'MsgBox (mytablex.Fields("producto") & " - " & CostoPromF)
        cn.Execute ("UPDATE PRODUCTO SET COSTOP= '" & CostoPromF & "'  WHERE  PRODUCTO='" & mytablex.Fields("producto") & "'")
    
        mytabley.Close
    
ksigue:
        mytablex.MoveNext
  
    Loop

    'MsgBox (cantidadF)

    MsgBox ("Costos Actualizados")

    Exit Sub
cmd1d5612_err:
    MsgBox "Error en exportacion excell" + error$, 48, "Aviso"

End Sub

'11/06/2018 Actualiza Precio Promedio Ponderado Masivo

'11/06/2018 Actualiza Precio Promedio Ponderado Masivo
Function busca_costoiProducto(producto As String) As String

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    buf = "select  TOP 1 PRECIO FROM detalle  where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    buf = buf & " and local='" & extra_loquesea(local1) & "'"
    buf = buf & " and producto='" & "" & producto & "'"
    buf = buf & " and bodega='" & extra_loquesea(bodega) & "'"

    buf = buf & " and acu1='' "
    buf = buf & " and (acu='J' or acu='K')"
    buf = buf & " and estado='2'"
    buf = buf & " order by fecha,hora"

    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_costoiProducto = "" & mytablex.Fields("PRECIO")
    Else
        busca_costoiProducto = 0

    End If

    mytablex.Close

End Function

'11/06/2018 Actualiza Precio Promedio Ponderado Masivo

