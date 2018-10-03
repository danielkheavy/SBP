VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form xprodet 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Productos y Mercaderias"
   ClientHeight    =   10470
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   19755
   DrawMode        =   11  'Not Xor Pen
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10470
   ScaleWidth      =   19755
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command25 
      Caption         =   "Command25"
      Height          =   495
      Left            =   14760
      TabIndex        =   131
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid dbgrid9 
      Height          =   4215
      Left            =   14760
      TabIndex        =   129
      Top             =   1200
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   7435
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   14737632
      Enabled         =   0   'False
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
         Name            =   "Arial Narrow"
         Size            =   9.75
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
            LCID            =   13322
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
            LCID            =   13322
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
   Begin VB.Frame Frame6 
      BackColor       =   &H8000000B&
      Caption         =   "Cambiar Precios de Todos los producto Seleccionados"
      Height          =   5280
      Left            =   0
      TabIndex        =   116
      Top             =   600
      Visible         =   0   'False
      Width           =   10755
      Begin VB.TextBox clave 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   2880
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   127
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox valoropera 
         Height          =   495
         Left            =   2880
         MaxLength       =   6
         TabIndex        =   126
         Top             =   1800
         Width           =   1695
      End
      Begin VB.ComboBox Combo5 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   124
         Top             =   1320
         Width           =   3975
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   120
         Top             =   840
         Width           =   3975
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   118
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label Label29 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clave de paso"
         Height          =   495
         Left            =   120
         TabIndex        =   128
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label Label30 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Operacion"
         Height          =   495
         Left            =   120
         TabIndex        =   125
         Top             =   1800
         Width           =   2775
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cambiar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6960
         TabIndex        =   123
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6960
         TabIndex        =   122
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label26 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero de Precio"
         Height          =   495
         Left            =   120
         TabIndex        =   121
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lista Precios"
         Height          =   495
         Left            =   120
         TabIndex        =   119
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label24 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Operacion"
         Height          =   495
         Left            =   120
         TabIndex        =   117
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cambiar Codigo de Producto x Otro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   0
      TabIndex        =   104
      Top             =   600
      Visible         =   0   'False
      Width           =   10695
      Begin VB.TextBox codigon 
         Height          =   615
         Left            =   2880
         MaxLength       =   15
         TabIndex        =   110
         Top             =   2040
         Width           =   3255
      End
      Begin VB.TextBox codigoa 
         Height          =   615
         Left            =   2880
         MaxLength       =   15
         TabIndex        =   109
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCTO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   151
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label33"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         TabIndex        =   150
         Top             =   840
         Width           =   990
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6240
         TabIndex        =   108
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6240
         TabIndex        =   107
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label21 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   106
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Label Label20 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo Antiguo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   105
         Top             =   1440
         Width           =   2655
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   3375
      Left            =   3600
      TabIndex        =   94
      Top             =   5880
      Visible         =   0   'False
      Width           =   5295
      Begin VB.Label procesos 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   95
         Top             =   480
         Width           =   3135
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Busqueda"
      Height          =   9735
      Left            =   -1200
      TabIndex        =   88
      Top             =   2040
      Visible         =   0   'False
      Width           =   14535
      Begin VB.CommandButton Command13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   6840
         TabIndex        =   91
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox cadena 
         Height          =   375
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   90
         Top             =   240
         Width           =   3375
      End
      Begin VB.ComboBox xbuffer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   89
         Top             =   240
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid dbgrid3 
         Height          =   8895
         Left            =   45
         TabIndex        =   92
         Top             =   825
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   15690
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
      Begin VB.Label counter 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   9960
         TabIndex        =   100
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Condicion"
         Height          =   375
         Left            =   120
         TabIndex        =   93
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Stock Almacenes"
      Height          =   4815
      Left            =   0
      TabIndex        =   83
      Top             =   600
      Visible         =   0   'False
      Width           =   10695
      Begin VB.TextBox fechaf 
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
         Left            =   5280
         MaxLength       =   10
         TabIndex        =   138
         Top             =   3120
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox fechai 
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
         Left            =   5280
         MaxLength       =   10
         TabIndex        =   135
         Top             =   2640
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton cmdCancelar 
         BackColor       =   &H00808080&
         Caption         =   "Cerrar Ventana"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5280
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.CommandButton Command20 
         BackColor       =   &H00C0C0C0&
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
         Height          =   855
         Left            =   5280
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   1680
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid dbgrid2 
         Height          =   3135
         Left            =   120
         TabIndex        =   86
         Top             =   360
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   5530
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
      Begin VB.Label productop 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   137
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label contador 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   136
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label stknom 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   360
         Width           =   105
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Consulta de Codigos de Barras"
      Height          =   5175
      Left            =   4680
      TabIndex        =   78
      Top             =   10080
      Visible         =   0   'False
      Width           =   10695
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3300
         Left            =   120
         TabIndex        =   81
         Top             =   840
         Width           =   6015
      End
      Begin VB.TextBox cbarras 
         Height          =   375
         Left            =   1440
         MaxLength       =   18
         TabIndex        =   80
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton Command11 
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
         Left            =   7920
         MaskColor       =   &H00FFFFFF&
         Picture         =   "xprodet.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ingrese Codigo"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   82
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Refresca"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   13635
      MaskColor       =   &H00FFFFFF&
      Picture         =   "xprodet.frx":07AE
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   90
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid dbGrid1 
      Height          =   8055
      Left            =   0
      TabIndex        =   41
      Top             =   600
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   14208
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777215
      HeadLines       =   1
      RowHeight       =   18
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
         Name            =   "Courier New"
         Size            =   9
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
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   18720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6720
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   18720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7200
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   18720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8640
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   18720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8280
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   14715
      TabIndex        =   35
      Top             =   11280
      Visible         =   0   'False
      Width           =   14775
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Casillero vacio Hab.Edicion"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12600
         TabIndex        =   37
         Top             =   0
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label kproducto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF00&
         DataField       =   "producto"
         DataSource      =   "Data1"
         Height          =   195
         Left            =   12240
         TabIndex        =   36
         Top             =   0
         Width           =   45
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   18720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7920
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   19695
      TabIndex        =   19
      Top             =   0
      Width           =   19755
      Begin VB.CommandButton Command28 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Comisión"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11240
         TabIndex        =   153
         Top             =   0
         Width           =   810
      End
      Begin VB.CommandButton Command27 
         BackColor       =   &H00808080&
         Caption         =   "Kardex Sunat"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2540
         MaskColor       =   &H00000000&
         TabIndex        =   152
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdEntradasSalidas 
         BackColor       =   &H00808080&
         Caption         =   "Entradas Salidas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   870
         MaskColor       =   &H00000000&
         TabIndex        =   149
         Top             =   0
         Width           =   810
      End
      Begin VB.CommandButton Command24 
         BackColor       =   &H00808080&
         Caption         =   "Cod Proveedor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5320
         TabIndex        =   115
         Top             =   0
         Width           =   825
      End
      Begin VB.CommandButton Command23 
         BackColor       =   &H00808080&
         Caption         =   "Orden Prn"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9920
         TabIndex        =   114
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Command22 
         BackColor       =   &H00808080&
         Caption         =   "Sema"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   17235
         TabIndex        =   113
         Top             =   60
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSDBGrid.DBGrid DBGrid5 
         Height          =   30
         Left            =   0
         OleObjectBlob   =   "xprodet.frx":0F5C
         TabIndex        =   103
         Top             =   600
         Width           =   11535
      End
      Begin VB.CommandButton Command21 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Grupos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   16620
         TabIndex        =   101
         Top             =   45
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command18 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cambios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9200
         TabIndex        =   99
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton Command17 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tecla"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   15480
         TabIndex        =   77
         Top             =   60
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton Command16 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Combo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12050
         TabIndex        =   75
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H00808080&
         Caption         =   "Conex"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8600
         TabIndex        =   74
         Top             =   0
         Width           =   600
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Copia Lista"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7980
         TabIndex        =   54
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00808080&
         Caption         =   "Eti queta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7500
         TabIndex        =   49
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton Command19 
         BackColor       =   &H00FFFF00&
         Caption         =   "CostoImporta."
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   8760
         Width           =   1095
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00808080&
         Caption         =   "Ver CBarra"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   18120
         TabIndex        =   32
         Top             =   60
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Print CBarra"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6870
         TabIndex        =   30
         Top             =   0
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00808080&
         Caption         =   "Recalculo de Saldos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   6140
         TabIndex        =   26
         Top             =   0
         Width           =   740
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Compra - Venta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3880
         TabIndex        =   25
         Top             =   0
         Width           =   690
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Stock"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         MaskColor       =   &H00004080&
         TabIndex        =   24
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Regula riza"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10535
         TabIndex        =   23
         Top             =   0
         Width           =   690
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00808080&
         Caption         =   "Gráfico"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   22
         Top             =   0
         Width           =   765
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00808080&
         Caption         =   "Kardex Producto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1700
         MaskColor       =   &H00000000&
         TabIndex        =   21
         Top             =   0
         Width           =   840
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00808080&
         Caption         =   "Receta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3280
         TabIndex        =   20
         Top             =   0
         Width           =   600
      End
      Begin VB.CommandButton Command26 
         BackColor       =   &H00C0C0C0&
         Caption         =   "pRece ta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   14850
         TabIndex        =   132
         Top             =   45
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label nro_registros 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   12720
         TabIndex        =   76
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   18720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7560
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   9735
      Left            =   11520
      ScaleHeight     =   9675
      ScaleWidth      =   3120
      TabIndex        =   1
      Top             =   600
      Width           =   3180
      Begin VB.PictureBox frmOtros 
         BackColor       =   &H00FFFFFF&
         Height          =   2535
         Left            =   120
         ScaleHeight     =   2475
         ScaleWidth      =   2955
         TabIndex        =   139
         TabStop         =   0   'False
         Top             =   6120
         Visible         =   0   'False
         Width           =   3015
         Begin VB.ComboBox sexo 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   143
            Top             =   360
            Width           =   1935
         End
         Begin VB.ComboBox proyecto 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   142
            Top             =   1080
            Width           =   1935
         End
         Begin VB.ComboBox talla 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   141
            Top             =   0
            Width           =   1935
         End
         Begin VB.ComboBox procedencia 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   140
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label lblSexo 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Sexo"
            Height          =   375
            Left            =   0
            TabIndex        =   147
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblTalla 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Talla"
            Height          =   375
            Left            =   0
            TabIndex        =   146
            Top             =   0
            Width           =   855
         End
         Begin VB.Label lblProcedencia 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Proced."
            Height          =   375
            Left            =   0
            TabIndex        =   145
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblProyecto 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Proyecto"
            Height          =   375
            Left            =   0
            TabIndex        =   144
            Top             =   1080
            Width           =   855
         End
      End
      Begin VB.ComboBox SEINVENTARIA 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   134
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox percepcion 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   111
         Text            =   "%"
         Top             =   8640
         Width           =   375
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   98
         Top             =   9000
         Width           =   855
      End
      Begin VB.TextBox criterio 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         MaxLength       =   5
         TabIndex        =   97
         Text            =   "%"
         Top             =   9000
         Width           =   735
      End
      Begin VB.ComboBox combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   96
         Top             =   9000
         Width           =   1215
      End
      Begin VB.TextBox f1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         MaxLength       =   1
         TabIndex        =   72
         Text            =   "%"
         Top             =   8640
         Width           =   375
      End
      Begin VB.TextBox fechavpi 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         MaxLength       =   10
         TabIndex        =   69
         Text            =   "%"
         Top             =   7920
         Width           =   1575
      End
      Begin VB.TextBox fechavpf 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         MaxLength       =   10
         TabIndex        =   68
         Text            =   "%"
         Top             =   8280
         Width           =   1575
      End
      Begin VB.ComboBox diauso 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox fechavf 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         MaxLength       =   10
         TabIndex        =   63
         Text            =   "%"
         Top             =   7560
         Width           =   1575
      End
      Begin VB.TextBox fechavi 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         MaxLength       =   10
         TabIndex        =   61
         Text            =   "%"
         Top             =   7200
         Width           =   1575
      End
      Begin VB.ComboBox vecaja 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   6840
         Width           =   1935
      End
      Begin VB.ComboBox activo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   6480
         Width           =   1935
      End
      Begin VB.ComboBox Peso 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   6120
         Width           =   1935
      End
      Begin VB.ComboBox bodega 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   5760
         Width           =   1935
      End
      Begin VB.ComboBox proveedor 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   5400
         Width           =   1935
      End
      Begin VB.TextBox monedav 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   47
         Text            =   "%"
         Top             =   4320
         Width           =   375
      End
      Begin VB.TextBox oferta 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         MaxLength       =   1
         TabIndex        =   45
         Text            =   "%"
         Top             =   4320
         Width           =   375
      End
      Begin VB.TextBox Barras 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         MaxLength       =   18
         TabIndex        =   43
         Text            =   "%"
         Top             =   3960
         Width           =   1575
      End
      Begin VB.ComboBox local1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   4680
         Width           =   1935
      End
      Begin VB.ComboBox igv 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   2880
         Width           =   1095
      End
      Begin VB.ComboBox ordenado 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   5040
         Width           =   1935
      End
      Begin VB.ComboBox color 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2160
         Width           =   1935
      End
      Begin VB.ComboBox subfamilia 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox familia 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   0
         Width           =   1935
      End
      Begin VB.ComboBox seccion 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   1935
      End
      Begin VB.ComboBox marca 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         TabIndex        =   5
         Text            =   "marca"
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox linea 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1800
         Width           =   1935
      End
      Begin VB.ComboBox categoria 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox producto 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         MaxLength       =   15
         TabIndex        =   2
         Text            =   "%"
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox descripcio 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         MaxLength       =   60
         TabIndex        =   0
         Text            =   "%"
         Top             =   3600
         Width           =   1575
      End
      Begin ChamaleonButton.ChameleonBtn mas 
         Height          =   465
         Left            =   2640
         TabIndex        =   148
         ToolTipText     =   "Más Opciones"
         Top             =   8520
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   820
         BTYPE           =   4
         TX              =   "+"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   4210752
         BCOLO           =   4210752
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "xprodet.frx":192F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label32 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inventario"
         Height          =   375
         Left            =   2040
         TabIndex        =   133
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label f 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Percepcio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1320
         TabIndex        =   112
         Top             =   8640
         Width           =   855
      End
      Begin VB.Label f 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VariaCost"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   73
         Top             =   8640
         Width           =   855
      End
      Begin VB.Label fechavx 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FVariapIn"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   71
         Top             =   7920
         Width           =   855
      End
      Begin VB.Label fechavx 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FVariapFn"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   70
         Top             =   8280
         Width           =   855
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DiaUso"
         Height          =   375
         Left            =   120
         TabIndex        =   66
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label fechavx 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaVence"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   64
         Top             =   7560
         Width           =   855
      End
      Begin VB.Label fechavx 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaVence"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   62
         Top             =   7200
         Width           =   855
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ve Caja"
         Height          =   375
         Left            =   120
         TabIndex        =   60
         Top             =   6840
         Width           =   855
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Activo"
         Height          =   375
         Left            =   120
         TabIndex        =   58
         Top             =   6480
         Width           =   855
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Peso"
         Height          =   375
         Left            =   120
         TabIndex        =   56
         Top             =   6120
         Width           =   855
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Almacen"
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Top             =   5760
         Width           =   855
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Proveedor"
         Height          =   375
         Left            =   120
         TabIndex        =   51
         Top             =   5400
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MonedaV"
         Height          =   375
         Index           =   4
         Left            =   1320
         TabIndex        =   48
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Oferta"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   46
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Barras"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   44
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PrecioLocal"
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Igv"
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label buffer 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2640
         TabIndex        =   31
         Top             =   7800
         Width           =   105
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ordenado"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   5040
         Width           =   855
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Color"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Subfamilia"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Familia"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Seccion"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Marca"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Linea"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Categoria"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Producto"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcio"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   3600
         Width           =   855
      End
   End
   Begin MSDataGridLib.DataGrid dbgrid4 
      Height          =   1815
      Left            =   0
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   8640
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   3201
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12648447
      HeadLines       =   1
      RowHeight       =   17
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   18
      BeginProperty Column00 
         DataField       =   "Local"
         Caption         =   "Lista"
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
         DataField       =   "fechavp"
         Caption         =   "Fechavp"
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
         DataField       =   "Unidad1"
         Caption         =   "Und"
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
      BeginProperty Column03 
         DataField       =   "Factor1"
         Caption         =   "Fac"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Pventa1"
         Caption         =   "Pventa1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Margen1"
         Caption         =   "Margen1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Unidad2"
         Caption         =   "Und"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Factor2"
         Caption         =   "Factor2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Pventa2"
         Caption         =   "Pventa2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Margen2"
         Caption         =   "Margen2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "Unidad3"
         Caption         =   "Und"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "Factor3"
         Caption         =   "Fac"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "Pventa3"
         Caption         =   "Pventa3"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "Margen3"
         Caption         =   "Margen3"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "Unidad4"
         Caption         =   "Und"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "Factor4"
         Caption         =   "Fac"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "Pventa4"
         Caption         =   "Pventa4"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column17 
         DataField       =   "Margen4"
         Caption         =   "Margen4"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   555.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   555.024
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column17 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label31 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STOCK EXISTENCIAS"
      Height          =   375
      Left            =   14760
      TabIndex        =   130
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label flag 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   19440
      TabIndex        =   102
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label esactivo 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   600
      TabIndex        =   27
      Top             =   8640
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Menu dk343 
      Caption         =   "&Nuevo"
   End
   Begin VB.Menu Modif34 
      Caption         =   "&Modifica"
   End
   Begin VB.Menu fk3434 
      Caption         =   "&Borra"
   End
   Begin VB.Menu Zom82 
      Caption         =   "&Zoom"
   End
   Begin VB.Menu dlo2323 
      Caption         =   "&Excell"
      Begin VB.Menu dk89231 
         Caption         =   "&1.Precios"
      End
      Begin VB.Menu preciogeneral 
         Caption         =   "&2.PreciosGeneral"
      End
      Begin VB.Menu dj7833re 
         Caption         =   "&3.Receta"
      End
      Begin VB.Menu fdk883 
         Caption         =   "&5.MinimosMaximos"
      End
   End
   Begin VB.Menu dk89331 
      Caption         =   "&Reporte"
   End
   Begin VB.Menu dkj8383 
      Caption         =   "&Conectividad"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu LI8912 
      Caption         =   "&CambiaCodigo"
   End
   Begin VB.Menu fk88555re 
      Caption         =   "&RecargosDescuentoProductos"
   End
   Begin VB.Menu flo434 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "xprodet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sqldatos     As String

Dim rrproducto   As New ADODB.Recordset

Dim miMarca      As String

Public cnn       As New ADODB.Connection

Dim ultimo_costo As Double

Private Sub Barras_KeyPress(KeyAscii As Integer)

    Dim buf As String

    If KeyAscii = 13 Then
        buf = convierte_barras(Barras)

        If Len(buf) > 0 Then
            Barras = buf

        End If

    End If

End Sub

Private Sub cadena_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        If opcion1 = "1" Then
            Frame3.Visible = False
            dbGrid1.SetFocus
            Exit Sub

        End If

        If opcion1 = "2" Then
            Frame3.Visible = False
            dbGrid1.SetFocus
            Exit Sub

        End If

        If opcion1 = "3" Then
            Frame3.Visible = False
            dbGrid1.SetFocus
            Exit Sub

        End If

    End If

    Command13_Click

End Sub

Private Sub categoria_Click()

    'Command5_Click
End Sub

Private Sub cbarras_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    Dim buf   As String

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        Command11_Click
        Exit Sub

    End If

    If Len(cbarras) = 0 Then
        cbarras.SetFocus
        Exit Sub

    End If

    buf = convierte_barras(cbarras)

    If Len(buf) > 0 Then
        cbarras = buf

    End If

    'MsgBox cbarras

    found = consulta_barras()

    If found = 0 Then
        cbarras = ""
        cbarras.SetFocus
        Exit Sub

    End If

    producto = List1.List(List1.ListIndex)
    found = sql_cabeza(0)
        
    List1.SetFocus

End Sub

Private Sub cmdCommand27_Click()

    ''22/06/2017 kenyo recalculo de saldo automatico desde producto
    Recalculo

    ''22/06/2017 kenyo recalculo de saldo automatico desde producto
End Sub

Private Sub cmdEntradasSalidas_Click()

    ''20/07/2017 kenyo reporte de entradas Salidas  en el modulo de reportes.

    Dim quebusco As String

    On Error GoTo cmd100_err

    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    opcion2 = "44"

    quebusco = Trim("" & dbGrid1.columns(1))
    repinv.producto = "" & dbGrid1.columns(1)
    repinv.excell.Visible = True

    repinv.Label15.Visible = True
    repinv.Label16.Visible = True
    repinv.fechai.Visible = True
    repinv.fechaf.Visible = True
    repinv.fechai.Enabled = True

    ''' 16/01/2018 Stock Inicial en Kardex Producto y Sunat
    repinv.ChkSaldoInicial.Visible = True
    ''' 16/01/2018 Stock Inicial en Kardex Producto y Sunat
    repinv.Show 1
    rrproducto.Find "producto='" & quebusco & "'"

    Exit Sub
cmd100_err:
    Exit Sub

    ''20/07/2017 kenyo reporte de entradas Salidas  en el modulo de reportes.

End Sub

Private Sub codigoa_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    codigon.SetFocus

End Sub

Private Sub Command14_Click()

    Dim quebusco As String

    On Error GoTo cmd1321_err

    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    tcolista.Show 1
    Exit Sub

    quebusco = Trim(dbGrid1.columns(1))
    tdueno.producto = Trim("" & dbGrid1.columns(1))
    tdueno.nproducto = Trim("" & dbGrid1.columns(0))
    tdueno.Show 1
    rrproducto.Find "producto='" & quebusco & "'"

    Exit Sub
cmd1321_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command15_Click()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd8912_err

    flag_clave1 = 0
    tconcla.X = "CONEXION"
    tconcla.Show 1

    If flag_clave1 <> 1 Then  'si es abre cajon
        MsgBox "No tiene permiso de conexion", 48, "Aviso"
        Exit Sub
   
    End If

    mytablex.Open "select * from ip", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If Len(Trim("" & mytablex.Fields("ip"))) = 0 Then
            MsgBox "No existe Ip ", 48, "Aviso"
            mytablex.Close
            Exit Sub

        End If

        Frame4.Visible = True
        found = conexion_cnn("" & mytablex.Fields("ip"))

        If found = 0 Then
            MsgBox "NO se puede conectar..", 48, "Aviso"
            mytablex.Close
            Exit Sub

        End If

        envio_producto
        envio_familia
        envio_subfamilia
        envio_marca
        envio_productb
        Frame4.Visible = False
        cnn.Close

    End If

    mytablex.Close
    Frame4.Visible = False

    Exit Sub
cmd8912_err:
    Frame4.Visible = False
    MsgBox "No se conecta  " + error$, 48, "Aviso"
    Exit Sub

End Sub

Function conexion_cnn(buf As String)

    On Error GoTo cmd9012_err

    cnn.CursorLocation = adUseClient
    cnn.Open "Driver={SQL Server};Server=" & Trim(buf) & ";Database=calipso;uid=sa"
    conexion_cnn = 1
    Exit Function
cmd9012_err:
    Exit Function

End Function

Sub envio_producto()

    Dim found As Integer

    Dim vr

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim sdx      As Double

    On Error GoTo cmd9000_err

    procesos = "Producto"
    vr = DoEvents
    sdx = 0
    mytablex.Open "select * from producto ", cnn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    cn.Execute ("delete from producto")
    mytabley.Open "select * from producto ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        sdx = sdx + 1
        procesos = " " & sdx
        vr = DoEvents
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 1
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Update
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close
    Exit Sub
cmd9000_err:
    MsgBox "NO existe proceso " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub envio_familia()

    Dim found As Integer

    Dim vr

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim sdx      As Double

    On Error GoTo cmd90001_err

    procesos = "familia"
    vr = DoEvents
    sdx = 0
    mytablex.Open "select * from familia ", cnn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    cn.Execute ("delete from familia")
    mytabley.Open "select * from familia ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        sdx = sdx + 1
        procesos = " " & sdx
        vr = DoEvents
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 1
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Update
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close
    Exit Sub
cmd90001_err:
    MsgBox "NO existe proceso " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub envio_subfamilia()

    Dim found As Integer

    Dim vr

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim sdx      As Double

    On Error GoTo cmd90002_err

    procesos = "subfamil"
    vr = DoEvents
    sdx = 0
    mytablex.Open "select * from subfamil ", cnn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    cn.Execute ("delete from subfamil")
    mytabley.Open "select * from subfamil ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        sdx = sdx + 1
        procesos = " " & sdx
        vr = DoEvents
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 1
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Update
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close
    Exit Sub
cmd90002_err:
    MsgBox "NO existe proceso " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub envio_marca()

    Dim found As Integer

    Dim vr

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim sdx      As Double

    On Error GoTo cmd90004_err

    procesos = "marca"
    vr = DoEvents
    sdx = 0
    mytablex.Open "select * from marca ", cnn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    cn.Execute ("delete from marca")
    mytabley.Open "select * from marca ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        sdx = sdx + 1
        procesos = " " & sdx
        vr = DoEvents
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 1
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Update
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close
    Exit Sub
cmd90004_err:
    MsgBox "NO existe proceso " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub envio_productb()

    Dim found As Integer

    Dim vr

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim sdx      As Double

    On Error GoTo cmd90006_err

    procesos = "productb"
    vr = DoEvents
    sdx = 0
    mytablex.Open "select * from productb ", cnn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    cn.Execute ("delete from productb")
    mytabley.Open "select * from productb ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        sdx = sdx + 1
        procesos = " " & sdx
        vr = DoEvents
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 1
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Update
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close
    Exit Sub
cmd90006_err:
    MsgBox "NO existe proceso " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command17_Click()

    On Error GoTo cmd901233_err

    TXFTECLA.codigo = Trim("" & dbGrid1.columns(1))
    TXFTECLA.Show 1
    'kardex_sunate
    Exit Sub
cmd901233_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Check1_Click()

    'MsgBox Check1.Value
    If Check1.Value = 1 Then
        dbGrid1.AllowUpdate = False
        Exit Sub

    End If

    If Check1.Value = 0 Then
        dbGrid1.AllowUpdate = True
        'Check1.Caption = "Habilitado Edicion"
        Exit Sub

    End If

End Sub

Private Sub Check1_Validate(Cancel As Boolean)

    'MsgBox "x"
End Sub

Private Sub cmdAddEntry_Click()

End Sub

Private Sub cmdCancelar_Click()
    Frame1.Visible = False
    dbGrid1.SetFocus

End Sub

Private Sub cmdExit_Click()
    flo434_Click

End Sub

Private Sub cmdSave_Click()

End Sub

Private Sub Command1_Click()

    Dim found    As Integer

    Dim quebusco As String

    On Error GoTo cmd321_err

    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    treceta.tiporeceta = "receta"
    quebusco = Trim("" & dbGrid1.columns(1))
    treceta.producto = "" & dbGrid1.columns(1)
    treceta.linea = "" & dbGrid1.columns(10)
    treceta.descripcio = "" & dbGrid1.columns(0)

    treceta.platos = "" & rrproducto.Fields("platos")

    treceta.detalle = Trim("" & rrproducto.Fields("detalle"))
    treceta.nro = "1"
    treceta.Show 1
    found = sql_cabeza(0)
    rrproducto.Find "producto='" & quebusco & "'"

    Exit Sub
cmd321_err:
    MsgBox "Seleccione un Producto", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command10_Click()

    Dim quebusco As String

    On Error GoTo cmd321_err

    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Len(Trim("" & rrproducto.Fields("producto"))) = 0 Then
        MsgBox "Barra no Existe", 48, "Aviso"
        Exit Sub

    End If

    'If Not IsNumeric(Trim("" & rrproducto.Fields("barras"))) Then
    '   MsgBox "Barra debe ser Numerico", 48, "Aviso"
    '   Exit Sub
    'End If
    quebusco = Trim(dbGrid1.columns(1))

    If Len(Trim("" & rrproducto.Fields("descorto"))) = 0 Then
        MsgBox "No existe descripcion corto", 48, "Aviso"
        Exit Sub

    End If

    frmlabel.txtData = Trim("" & dbGrid1.columns(1))

    's|o|p
    's|s|cm|2.5|5.0
    's|f|Arial|8|false|false
    'b|c|10|N133911
    't|133|35|S/ 100.00
    't|10|50|DETTOL J.LI SKIN DOY
    't|10|60|DETTOL J.LI SKIN DOY

    frmlabel.txtLabelDef = frmlabel.txtLabelDef & "s|o|p" & vbcrlf
    frmlabel.txtLabelDef = frmlabel.txtLabelDef & "s|s|cm|2.5|5.0" & vbcrlf
    frmlabel.txtLabelDef = frmlabel.txtLabelDef & "s|f|Arial|8|False|false" & vbcrlf
    frmlabel.txtLabelDef = frmlabel.txtLabelDef & "b|c|10|" & Trim("" & rrproducto.Fields("producto")) & vbcrlf
    frmlabel.txtLabelDef = frmlabel.txtLabelDef & "t|133|35|" & dicmoneda & busca_preciobarra() & vbcrlf
    frmlabel.txtLabelDef = frmlabel.txtLabelDef & "t|10|50|" & Mid$(Trim("" & rrproducto.Fields("descripcio")), 1, 25) & vbcrlf
    frmlabel.txtLabelDef = frmlabel.txtLabelDef & "t|10|60|" & Mid$(Trim("" & rrproducto.Fields("descripcio")), 26, 50)
    frmlabel.Show 1
    rrproducto.Find "producto='" & quebusco & "'"

    'tbarcode.Show 1
    Exit Sub
    'barracod.Barras = "" & dbGrid1.Columns(1)
    'barracod.Show 1
    Exit Sub
cmd321_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command11_Click()
    Frame2.Visible = False
    dbGrid1.SetFocus

End Sub

Private Sub Command12_Click()

    If Frame1.Visible = True Then Exit Sub
    List1.Clear
    Frame2.Visible = True
    producto = "%"
    cbarras = ""
    cbarras.SetFocus

End Sub

Private Sub Command13_Click()
    ejecuta 1

End Sub

Private Sub Command16_Click()

    Dim found    As Integer

    Dim quebusco As String

    On Error GoTo cmd7000_err

    quebusco = Trim("" & dbGrid1.columns(1))
    tabcombo.xproducto = "" & dbGrid1.columns(1)
    tabcombo.xdescripcio = "" & dbGrid1.columns(0)
    tabcombo.Show 1

    rrproducto.Find "producto='" & quebusco & "'"

    Exit Sub
cmd7000_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command18_Click()
    flag_clave1 = 0
    tconcla.X = "CAMBIOS"
    tconcla.Show 1

    If flag_clave1 <> 1 Then  'si es abre cajon
        MsgBox "No tiene permiso de conexion", 48, "Aviso"
        Exit Sub
   
    End If

    tcaprod.tabla = "producto"
    tcaprod.Show 1

End Sub

Private Sub Command2_Click()

    Dim quebusco As String

    On Error GoTo cmd100_err

    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    opcion2 = "1"

    quebusco = Trim("" & dbGrid1.columns(1))
    repinv.producto = "" & dbGrid1.columns(1)
    repinv.excell.Visible = True
    repinv.Label15.Visible = True
    repinv.Label16.Visible = True
    repinv.fechai.Visible = True
    repinv.fechaf.Visible = True
    repinv.fechai.Enabled = True

    '''10/08/2017 kenyo Mejor Kardex Producto
    repinv.quecosto.Visible = True
    repinv.Label33.Visible = True

    '''10/08/2017 kenyo Mejor Kardex Producto

    ''' 16/01/2018 Stock Inicial en Kardex Producto y Sunat
    repinv.ChkSaldoInicial.Visible = True
    ''' 16/01/2018 Stock Inicial en Kardex Producto y Sunat

    repinv.Show 1
    rrproducto.Find "producto='" & quebusco & "'"

    Exit Sub
cmd100_err:
    Exit Sub

End Sub

Private Sub Command20_Click()

    On Error GoTo cmd901260_err

    stknom = "" & dbGrid1.columns(0)
    sql_saldo_locales Trim("" & dbGrid1.columns(1))
    'hacer_sunat
    'sql_saldo_locales Trim("" & dbGrid1.columns(1))
    Exit Sub
cmd901260_err:
    MsgBox "Seleccione un dato", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command21_Click()

    Dim quebusco As String

    On Error GoTo cmd12100_err

    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    quebusco = Trim("" & dbGrid1.columns(1))
    tpereq.buffer = Trim("" & dbGrid1.columns(1))
    tpereq.Show 1
    rrproducto.Find "producto='" & quebusco & "'"
    Exit Sub
cmd12100_err:
    Exit Sub

End Sub

Private Sub Command22_Click()

    Dim found    As Integer

    Dim quebusco As String

    On Error GoTo cmd12321_err

    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    '---------------------------------------
    'abrir a que formulacion de produccion va tener el sistema
    tdiasema.Show 1
    found = sql_cabeza(0)
    '---------------------------------------
    Exit Sub
    quebusco = Trim("" & dbGrid1.columns(1))
    treceta.tiporeceta = "recetapro"
    treceta.producto = "" & dbGrid1.columns(1)
    treceta.linea = "" & dbGrid1.columns(10)
    treceta.descripcio = "" & dbGrid1.columns(0)
    treceta.platos = "" & rrproducto.Fields("platos")
    treceta.detalle = Trim("" & rrproducto.Fields("detalle"))
    treceta.nro = "1"
    treceta.Show 1
    found = sql_cabeza(0)
    rrproducto.Find "producto='" & quebusco & "'"
    Exit Sub
cmd12321_err:
    MsgBox "Seleccione un Producto", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command23_Click()

    Dim quebusco As String

    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    quebusco = Trim("" & dbGrid1.columns(1))
    timpod.producto = "" & dbGrid1.columns(1)
    timpod.Show 1
    rrproducto.Find "producto='" & quebusco & "'"

End Sub

Private Sub Command24_Click()

    Dim quebusco As String

    quebusco = Trim("" & dbGrid1.columns(1))

    On Error GoTo cmd99451_err

    If Frame2.Visible = True Then Exit Sub
    tcodprov.buffer = Trim("" & dbGrid1.columns(1))
    tcodprov.Show 1
    Exit Sub
cmd99451_err:
    MsgBox "Seleccione un Producto", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command25_Click()
    Exit Sub
    CONTADOR_producto

End Sub

Private Sub Command26_Click()

    Dim found    As Integer

    Dim quebusco As String

    On Error GoTo cmd6321_err

    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    '---------------
    quebusco = Trim("" & dbGrid1.columns(1))
    ttrecepr.pproducto = Trim("" & dbGrid1.columns("producto"))
    ttrecepr.Show 1
    '---------------
    found = sql_cabeza(0)
    rrproducto.Find "producto='" & quebusco & "'"

    Exit Sub
cmd6321_err:
    MsgBox "Seleccione un Producto", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command27_Click()

    '' 17/01/2018 Kardex Sunat desde Modulo de Productos
    Dim quebusco As String

    On Error GoTo cmd100_err

    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub

    opcion2 = "100"

    quebusco = Trim("" & dbGrid1.columns(1))
    repinv.producto = "" & dbGrid1.columns(1)

    repinv.Label15.Visible = True
    repinv.Label16.Visible = True
    repinv.fechai.Visible = True
    repinv.fechaf.Visible = True

    repinv.ChkSaldoInicial.Visible = True
    repinv.fechai.BackColor = 8454143
    repinv.fechaf.BackColor = 8454143
    ''' 16/01/2018 Stock Inicial en Kardex Producto y Sunat
    repinv.ChkSaldoInicial.Visible = True
    ''' 16/01/2018 Stock Inicial en Kardex Producto y Sunat
    repinv.Show 1
    rrproducto.Find "producto='" & quebusco & "'"

    Exit Sub
cmd100_err:
    Exit Sub
    '' 17/01/2018 Kardex Sunat desde Modulo de Productos

End Sub

Private Sub Command28_Click()
    '' 29/01/2018 Comisiones por producto por trabajador. Proyectos requerimientos Spa Cañete.
    FrmComisiones.producto = Trim("" & dbGrid1.columns(1))
    FrmComisiones.descripcion = Trim("" & dbGrid1.columns(0))
    FrmComisiones.Show 1
    '' 29/01/2018 Comisiones por producto por trabajador. Proyectos requerimientos Spa Cañete.

End Sub

Private Sub Command3_Click()

    '' 06/01/2018 Actualiza Costos Compra en producto

    Dim quebusco As String

    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    quebusco = Trim("" & dbGrid1.columns(1))
    FrmChart.producto = "" & dbGrid1.columns(1)
    FrmChart.acu = "V"
    FrmChart.docu = "1"
    FrmChart.Show 1
    rrproducto.Find "producto='" & quebusco & "'"
    'ActualizaCostosCompra
    '' 06/01/2018 Actualiza Costos Compra en producto

End Sub

'' 06/01/2018 Actualiza Costos Compra en producto
Sub ActualizaCostosCompra()

    Dim mytablexyz As New ADODB.Recordset

    Do

        If rrproducto.EOF Then Exit Do
   
        mytablexyz.Open "select producto from receta where producto='" & rrproducto.Fields("producto") & "' ", cn, adOpenKeyset, adLockOptimistic

        If mytablexyz.RecordCount > 0 Then
        Else
            cn.Execute ("UPDATE PRODUCTO SET COSTOU=(SELECT TOP 1(PRECIO) FROM DETALLE WHERE (ACU='J' OR ACU='K' OR ACU='P') AND ESTADO='2'  AND PRODUCTO='" & rrproducto.Fields("producto") & "' ORDER BY FECHA DESC) WHERE  PRODUCTO='" & rrproducto.Fields("producto") & "'")

        End If
 
        rrproducto.MoveNext
        mytablexyz.Close
 
    Loop

End Sub

'' 06/01/2018 Actualiza Costos Compra en producto

Private Sub Command4_Click()

    Dim quebusco As String

    On Error GoTo cmd2100_err

    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    quebusco = Trim("" & dbGrid1.columns(1))
    tsiconte.Show 1
    rrproducto.Find "producto='" & quebusco & "'"

    'opcion2 = "1"
    'repsunat.producto = "" & dbGrid1.columns(1)
    'repsunat.Show 1
    Exit Sub
cmd2100_err:
    Exit Sub

End Sub

Private Sub Command5_Click()

    Dim found As Integer

    found = sql_cabeza(1)

End Sub

Sub Recalculo()

    Dim found     As Integer

    Dim mytablex  As New ADODB.Recordset

    Dim mytablez  As New ADODB.Recordset

    Dim mytabley  As New ADODB.Recordset

    Dim mytableb  As New ADODB.Recordset

    Dim buf       As String

    Dim signo     As Double

    Dim saldoini  As Double

    Dim xt1       As Double

    Dim xt2       As Double

    Dim xt3       As Double

    Dim xt4       As Double

    Dim xt5       As Double

    Dim xt6       As Double

    Dim xt7       As Double

    Dim xt8       As Double

    Dim xt9       As Double

    Dim xt10      As Double

    Dim xt11      As Double

    Dim xt12      As Double

    Dim xt13      As Double

    Dim xt14      As Double

    Dim xt15      As Double

    Dim xt16      As Double

    Dim sdx       As Double

    Dim mytablera As New ADODB.Recordset

    'Dim found As Integer
    Dim vr

    Dim sdxt As Double

    If Len(fechai) = 0 Then
        bodega.SetFocus
        Exit Sub

    End If

    If Len(fechai) <> 10 Then
        bodega.SetFocus
        Exit Sub

    End If

    If Not IsDate(fechai) Then
        bodega.SetFocus
        Exit Sub

    End If

    If Len(fechai) <> 10 Then Exit Sub
    If Len(fechaf) <> 10 Then Exit Sub
    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub
    If CVDate(fechaf) < CVDate(fechai) Then Exit Sub
    'sql producto
    actualiza_kardex
    Exit Sub
    sdxt = 0
    buf = "select * from producto where descripcio like '%'"

    If producto <> "%" Then
        buf = buf & " and producto like '" & producto & "'"

    End If

    If familia <> "%" Then
        buf = buf & " and familia like '" & familia & "'"

    End If

    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    '----primero borrando los datos----
    suma1 = 0
    Do

        If mytablex.EOF Then Exit Do
        suma1 = suma1 + 1
        contador = Format(suma1, "0")
        vr = DoEvents()

        If Command2.Visible = False Then Exit Do
        saldoini = 0

        xt1 = 0
        xt2 = 0
        xt3 = 0
        xt4 = 0
        xt5 = 0
        xt6 = 0
        xt7 = 0
        xt8 = 0
        xt9 = 0
        xt10 = 0
        xt11 = 0
        xt12 = 0
        xt13 = 0
        xt14 = 0
        xt15 = 0
        xt16 = 0

        If mytablez.State = 1 Then mytablez.Close
        mytablez.Open "Select * from dsaldoini where local='01' and producto='" & "" & mytablex.Fields("producto") & "' and bodega='01' and fecha='" & Format(fechai, "YYYYMMDD") & "'", cn, adOpenStatic, adLockOptimistic

        If mytablez.RecordCount > 0 Then
            saldoini = Val("" & mytablez.Fields("cantidad")) * Val("" & mytablez.Fields("factor"))
            xt1 = Val("" & mytablez.Fields("t1"))
            xt2 = Val("" & mytablez.Fields("t2"))
            xt3 = Val("" & mytablez.Fields("t3"))
            xt4 = Val("" & mytablez.Fields("t4"))
            xt5 = Val("" & mytablez.Fields("t5"))
            xt6 = Val("" & mytablez.Fields("t6"))
            xt7 = Val("" & mytablez.Fields("t7"))
            xt8 = Val("" & mytablez.Fields("t8"))
            xt9 = Val("" & mytablez.Fields("t9"))
            xt10 = Val("" & mytablez.Fields("t10"))
            xt11 = Val("" & mytablez.Fields("t11"))
            xt12 = Val("" & mytablez.Fields("t12"))
            xt13 = Val("" & mytablez.Fields("t13"))
            xt14 = Val("" & mytablez.Fields("t14"))
            xt15 = Val("" & mytablez.Fields("t15"))
            xt16 = Val("" & mytablez.Fields("t16"))

        End If

        mytablez.Close

        If mytabley.State = 1 Then mytabley.Close
        mytabley.Open "Select * from almacen where local='" & extra_loquesea(local1) & "' and producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "'", cn, adOpenStatic, adLockOptimistic

        If mytabley.RecordCount > 0 Then
            'mytabley.Edit
            sdxt = sdxt + saldoini
            mytabley.Fields("saldo") = saldoini
            mytabley.Fields("t1") = xt1
            mytabley.Fields("t2") = xt2
            mytabley.Fields("t3") = xt3
            mytabley.Fields("t4") = xt4
            mytabley.Fields("t5") = xt5
            mytabley.Fields("t6") = xt6
            mytabley.Fields("t7") = xt7
            mytabley.Fields("t8") = xt8
            mytabley.Fields("t9") = xt9
            mytabley.Fields("t10") = xt10
            mytabley.Fields("t11") = xt11
            mytabley.Fields("t12") = xt12
            mytabley.Fields("t13") = xt13
            mytabley.Fields("t14") = xt14
            mytabley.Fields("t15") = xt15
            mytabley.Fields("t16") = xt16
            mytabley.Update
        Else
            mytabley.AddNew
            sdxt = sdxt + saldoini
            mytabley.Fields("local") = extra_loquesea(local1)
            mytabley.Fields("producto") = "" & mytablex.Fields("producto")
            mytabley.Fields("bodega") = extra_loquesea(bodega)
            mytabley.Fields("saldo") = saldoini
            mytabley.Fields("t1") = xt1
            mytabley.Fields("t2") = xt2
            mytabley.Fields("t3") = xt3
            mytabley.Fields("t4") = xt4
            mytabley.Fields("t5") = xt5
            mytabley.Fields("t6") = xt6
            mytabley.Fields("t7") = xt7
            mytabley.Fields("t8") = xt8
            mytabley.Fields("t9") = xt9
            mytabley.Fields("t10") = xt10
            mytabley.Fields("t11") = xt11
            mytabley.Fields("t12") = xt12
            mytabley.Fields("t13") = xt13
            mytabley.Fields("t14") = xt14
            mytabley.Fields("t15") = xt15
            mytabley.Fields("t16") = xt16
            mytabley.Update

        End If

        mytablex.MoveNext
    Loop
    mytablex.Close

    'ahora ver las transacciones y sumarlos al saldo
    buf = "select * from detalle where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If familia <> "%" Then
        buf = buf & " and familia like '" & familia & "'"

    End If

    If producto <> "%" Then
        buf = buf & " and producto like '" & producto & "'"

    End If

    buf = buf & " and local='01'"
    buf = buf & " and bodega='01'"
    buf = buf & " and (acu='S' or acu='T' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' OR acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='E')"
    'buf = buf & " and (len(acu1)=0 or acu1=null)"
    buf = buf & " and estado='2'"
    'buf = buf & " group by producto,bodega,flage,acu1"

    If mytableb.State = 1 Then mytableb.Close
    mytableb.Open buf, cn, adOpenStatic, adLockOptimistic
    suma1 = 0

    If Command2.Visible = False Then Exit Sub
    Do
        vr = DoEvents()

        If Command2.Visible = False Then Exit Do
        If mytableb.EOF Then Exit Do

        'aqui validamos si se puede actualizar
        If mytablera.State = 1 Then mytablera.Close
        mytablera.Open "select tipo1 from factura where local='01' and tipo='" & "" & mytableb.Fields("tipo") & "' and serie='" & mytableb.Fields("serie") & "' and numero='" & "" & mytableb.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic

        If mytablera.RecordCount > 0 Then
            found = ve_descarga("" & mytablera.Fields("tipo1"))

            If found = 1 Then 'qe no se descarge
                mytablera.Close
                GoTo sisin

            End If

        End If

        mytablera.Close

        suma1 = suma1 + 1
        contador = Format(suma1, "0")
        productop = "" & mytableb.Fields("producto")
        signo = 1

        If "" & mytableb.Fields("acu") = "T" Or "" & mytableb.Fields("acu") = "A" Or "" & mytableb.Fields("acu") = "B" Or "" & mytableb.Fields("acu") = "C" Or "" & mytableb.Fields("acu") = "D" Or "" & mytableb.Fields("acu") = "G" Or "" & mytableb.Fields("acu") = "N" Then
            signo = -1

        End If

        If "" & mytableb.Fields("acu") = "S" Or "" & mytableb.Fields("acu") = "J" Or "" & mytableb.Fields("acu") = "K" Or "" & mytableb.Fields("acu") = "L" Or "" & mytableb.Fields("acu") = "M" Or "" & mytableb.Fields("acu") = "P" Or "" & mytableb.Fields("acu") = "E" Then
            signo = 1

        End If

        'If Val("" & mytableb.Fields("cantidad")) < 0 Then
        '   signo = 1
        'End If
        'ahora en almacenes
        If mytabley.State = 1 Then mytabley.Close
        mytabley.Open "Select * from almacen where local='01' and producto='" & "" & mytableb.Fields("producto") & "' and bodega='01'", cn, adOpenStatic, adLockOptimistic

        If mytabley.RecordCount > 0 Then
            'mytabley.Edit
            sdx = Val("" & mytabley.Fields("saldo")) + signo * Val("" & mytableb.Fields("cantidad")) * Val("" & mytableb.Fields("factor"))
            mytabley.Fields("saldo") = sdx

            'sdxt = sdxt + sdx
            If Len("" & mytableb.Fields("linea")) > 0 Then
                sdx = Val("" & mytabley.Fields("T1")) + signo * Val("" & mytableb.Fields("T1"))
                mytabley.Fields("t1") = sdx
                sdx = Val("" & mytabley.Fields("T2")) + signo * Val("" & mytableb.Fields("T2"))
                mytabley.Fields("t2") = sdx
                sdx = Val("" & mytabley.Fields("T3")) + signo * Val("" & mytableb.Fields("T3"))
                mytabley.Fields("t3") = sdx
                sdx = Val("" & mytabley.Fields("T4")) + signo * Val("" & mytableb.Fields("T4"))
                mytabley.Fields("t4") = sdx
                sdx = Val("" & mytabley.Fields("T5")) + signo * Val("" & mytableb.Fields("T5"))
                mytabley.Fields("t5") = sdx
                sdx = Val("" & mytabley.Fields("T6")) + signo * Val("" & mytableb.Fields("T6"))
                mytabley.Fields("t6") = sdx
                sdx = Val("" & mytabley.Fields("T7")) + signo * Val("" & mytableb.Fields("T7"))
                mytabley.Fields("t7") = sdx
                sdx = Val("" & mytabley.Fields("T8")) + signo * Val("" & mytableb.Fields("T8"))
                mytabley.Fields("t8") = sdx
                sdx = Val("" & mytabley.Fields("T9")) + signo * Val("" & mytableb.Fields("T9"))
                mytabley.Fields("t9") = sdx
                sdx = Val("" & mytabley.Fields("T10")) + signo * Val("" & mytableb.Fields("T10"))
                mytabley.Fields("t10") = sdx
                sdx = Val("" & mytabley.Fields("T11")) + signo * Val("" & mytableb.Fields("T11"))
                mytabley.Fields("t11") = sdx
                sdx = Val("" & mytabley.Fields("T12")) + signo * Val("" & mytableb.Fields("T12"))
                mytabley.Fields("t12") = sdx
                sdx = Val("" & mytabley.Fields("T13")) + signo * Val("" & mytableb.Fields("T13"))
                mytabley.Fields("t13") = sdx
                sdx = Val("" & mytabley.Fields("T14")) + signo * Val("" & mytableb.Fields("T14"))
                mytabley.Fields("t14") = sdx
                sdx = Val("" & mytabley.Fields("T15")) + signo * Val("" & mytableb.Fields("T15"))
                mytabley.Fields("t15") = sdx
                sdx = Val("" & mytabley.Fields("T16")) + signo * Val("" & mytableb.Fields("T16"))
                mytabley.Fields("t16") = sdx

            End If

            mytabley.Update
        Else
            mytabley.AddNew
            mytabley.Fields("producto") = "01"
            mytabley.Fields("bodega") = "01"
            mytabley.Fields("local") = "01"
            sdx = Val("" & mytabley.Fields("saldo")) + signo * Val("" & mytableb.Fields("cantidad")) * Val("" & mytableb.Fields("factor"))
            mytabley.Fields("saldo") = sdx

            'sdxt = sdxt + sdx
            If Len("" & mytableb.Fields("linea")) > 0 Then
                sdx = Val("" & mytabley.Fields("T1")) + signo * Val("" & mytableb.Fields("T1"))
                mytabley.Fields("t1") = sdx
                sdx = Val("" & mytabley.Fields("T2")) + signo * Val("" & mytableb.Fields("T2"))
                mytabley.Fields("t2") = sdx
                sdx = Val("" & mytabley.Fields("T3")) + signo * Val("" & mytableb.Fields("T3"))
                mytabley.Fields("t3") = sdx
                sdx = Val("" & mytabley.Fields("T4")) + signo * Val("" & mytableb.Fields("T4"))
                mytabley.Fields("t4") = sdx
                sdx = Val("" & mytabley.Fields("T5")) + signo * Val("" & mytableb.Fields("T5"))
                mytabley.Fields("t5") = sdx
                sdx = Val("" & mytabley.Fields("T6")) + signo * Val("" & mytableb.Fields("T6"))
                mytabley.Fields("t6") = sdx
                sdx = Val("" & mytabley.Fields("T7")) + signo * Val("" & mytableb.Fields("T7"))
                mytabley.Fields("t7") = sdx
                sdx = Val("" & mytabley.Fields("T8")) + signo * Val("" & mytableb.Fields("T8"))
                mytabley.Fields("t8") = sdx
                sdx = Val("" & mytabley.Fields("T9")) + signo * Val("" & mytableb.Fields("T9"))
                mytabley.Fields("t9") = sdx
                sdx = Val("" & mytabley.Fields("T10")) + signo * Val("" & mytableb.Fields("T10"))
                mytabley.Fields("t10") = sdx
                sdx = Val("" & mytabley.Fields("T11")) + signo * Val("" & mytableb.Fields("T11"))
                mytabley.Fields("t11") = sdx
                sdx = Val("" & mytabley.Fields("T12")) + signo * Val("" & mytableb.Fields("T12"))
                mytabley.Fields("t12") = sdx
                sdx = Val("" & mytabley.Fields("T13")) + signo * Val("" & mytableb.Fields("T13"))
                mytabley.Fields("t13") = sdx
                sdx = Val("" & mytabley.Fields("T14")) + signo * Val("" & mytableb.Fields("T14"))
                mytabley.Fields("t14") = sdx
                sdx = Val("" & mytabley.Fields("T15")) + signo * Val("" & mytableb.Fields("T15"))
                mytabley.Fields("t15") = sdx
                sdx = Val("" & mytabley.Fields("T16")) + signo * Val("" & mytableb.Fields("T16"))
                mytabley.Fields("t16") = sdx

            End If

            mytabley.Update

        End If

sisin:
        mytableb.MoveNext
    Loop
    'MsgBox sdxt
    mytableb.Close
    mytabley.Close
    Command2.Visible = False

End Sub

''22/06/2017 kenyo recalculo de saldo automatico desde producto
Function actualiza_kardex()

    Dim found As Integer

    found = kardexactualiza("01", "" & producto, "01", fechai, fechaf)

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

''22/06/2017 kenyo recalculo de saldo automatico desde producto

Private Sub Command6_Click()
    '''22/06/2017 kenyo recalculo de saldo automatico desde producto
    '' crear fechai,fechaf,contador,productop
    '
    'fechai = "01" & "/" & "01" & "/" & Format(Year(Now), "0000")
    'fechaf = Format(Now, "dd/mm/yyyy")
    'Recalculo
    '
    '''22/06/2017 kenyo recalculo de saldo automatico desde producto

    On Error GoTo cmd451_err

    If Frame2.Visible = True Then Exit Sub

    Frame1.Visible = True
    'ccosto = ""
    DBGrid2.Enabled = True
    stknom = "" & dbGrid1.columns(0)
    sql_saldo_locales Trim("" & dbGrid1.columns(1))
    Exit Sub
cmd451_err:
    MsgBox "Seleccione un Producto", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command7_Click()

    Dim quebusco As String

    On Error GoTo cmd789_err

    quebusco = Trim("" & dbGrid1.columns(1))

    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    procomr.producto = "" & dbGrid1.columns(1)
    procomr.Show 1
    rrproducto.Find "producto='" & quebusco & "'"

    Exit Sub
cmd789_err:
    MsgBox "Seleccione un Producto", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command8_Click()

    Dim quebusco As String

    On Error GoTo cmd301_err

    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    quebusco = Trim("" & dbGrid1.columns(1))
    trecalcu.producto = Trim("" & dbGrid1.columns(1))
    'Cargastk.descripcio = descripcio

    trecalcu.Show 1
    rrproducto.Find "producto='" & quebusco & "'"

    Exit Sub
cmd301_err:
    Exit Sub

End Sub

Private Sub Command9_Click()

    Dim quebusco As String

    On Error GoTo cmd3321_err

    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Len(Trim("" & rrproducto.Fields("producto"))) = 0 Then
        MsgBox "No existe producto", 48, "Aviso"
        Exit Sub

    End If

    'If Not IsNumeric(Trim("" & rrproducto.Fields("barras"))) Then
    '   MsgBox "Barra debe ser Numerico", 48, "Aviso"
    '   Exit Sub
    'End If
    If Len(Trim("" & rrproducto.Fields("descripcio"))) = 0 Then
        MsgBox "No existe descripcion ", 48, "Aviso"
        Exit Sub

    End If

    quebusco = Trim("" & dbGrid1.columns(1))
    frmlabel.txtData = Trim("" & dbGrid1.columns(1))
    frmlabel.txtLabelDef = ""
    frmlabel.txtLabelDef = frmlabel.txtLabelDef & "s|o|p" & vbcrlf
    frmlabel.txtLabelDef = frmlabel.txtLabelDef & "s|s|cm|3.5|5.5" & vbcrlf
    frmlabel.txtLabelDef = frmlabel.txtLabelDef & "s|f|Arial|8|false|false" & vbcrlf
    'frmlabel.txtLabelDef = frmlabel.txtLabelDef & "f" & vbcrlf
    frmlabel.txtLabelDef = frmlabel.txtLabelDef & "t|5|5|" & Mid$(Trim("" & rrproducto.Fields("descripcio")), 1, 30) & vbcrlf
    frmlabel.txtLabelDef = frmlabel.txtLabelDef & "t|5|25|" & Mid$(Trim("" & rrproducto.Fields("descripcio")), 31, 60) & vbcrlf
    frmlabel.txtLabelDef = frmlabel.txtLabelDef & "t|5|40|" & Mid$(Trim("" & rrproducto.Fields("monedav")), 1, 1) & "/." & vbcrlf
    frmlabel.txtLabelDef = frmlabel.txtLabelDef & "s|f|Arial|35|true|false" & vbcrlf
    frmlabel.txtLabelDef = frmlabel.txtLabelDef & "t|25|60|" & busca_preciobarra()
    frmlabel.columna = "1"
    'MsgBox frmlabel.txtLabelDef

    frmlabel.Show 1
    rrproducto.Find "producto='" & quebusco & "'"

    Exit Sub
    Exit Sub
cmd3321_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub dbgrid1_Click()

    On Error GoTo cmd89122_err

    consulta_precios Trim("" & rrproducto.Fields("producto"))
    Exit Sub
cmd89122_err:
    Exit Sub

End Sub

Private Sub dbgrid1_DblClick()
    Zom82_Click

    If Check1.Value = 1 Then

        'visualiza_precios
    End If

End Sub

Private Sub dbgrid1_Error(ByVal DataError As Integer, Response As Integer)
    Response = False

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo cmd45_err

    'consulta_precios Trim("" & dbGrid1.columns(1))
    Exit Sub
cmd45_err:
    Exit Sub

End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)

    Dim buf   As String

    Dim buf2  As String

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then
        If KeyAscii = 8 Then

            'If Len(descripcio) = 0 Then
            '   descripcio = "%"
            '   KeyAscii = 0
            'End If
            If Len(descripcio) > 0 Then
                buf = Mid$(descripcio, 1, Len(descripcio) - 1)
                descripcio = buf
                KeyAscii = 0
            Else
                descripcio = "%"
                KeyAscii = 0
                Exit Sub

            End If

        End If

        buf = Chr(KeyAscii)

        If Chr(KeyAscii) = "*" Then
            buf = "%"
            descripcio = buf

        End If

        If descripcio = "%" Then
            descripcio = ""

        End If

        If KeyAscii <> 13 Then
            descripcio = descripcio + buf

        End If

        buf = descripcio
        'MsgBox buf
        found = sql_cabeza(0)
        Exit Sub

    End If

    If KeyAscii = 13 Then
        Zom82_Click

    End If

End Sub

Private Sub dbgrid1_KeyUp(KeyCode As Integer, Shift As Integer)

    On Error GoTo cmd102_err

    consulta_precios Trim("" & dbGrid1.columns(1))

    If KeyCode = 13 Then
        If Frame1.Visible = True Then Exit Sub
        If Frame2.Visible = True Then Exit Sub
        Exit Sub

    End If

    If KeyCode = &H70 And dbGrid1.Col = 4 Then 'f1
        xbuffer.Clear
        xbuffer.AddItem "Descripcio"
        xbuffer.AddItem "Familia"
        xbuffer.ListIndex = 0
        opcion1 = "1"
        Frame3.Visible = True
        cadena = ""
        cadena.SetFocus
        Command13_Click
        Exit Sub

    End If

    If KeyCode = &H70 And dbGrid1.Col = 3 Then 'f1
        xbuffer.Clear
        xbuffer.AddItem "Descripcio"
        xbuffer.AddItem "Marca"
        xbuffer.ListIndex = 0
        opcion1 = "2"
        Frame3.Visible = True
        cadena = ""
        cadena.SetFocus
        Command13_Click
        Exit Sub

    End If

    If KeyCode = &H70 And dbGrid1.Col = 5 Then 'f1
        xbuffer.Clear
        xbuffer.AddItem "Descripcio"
        xbuffer.AddItem "Subfamilia"
        xbuffer.ListIndex = 0
        opcion1 = "3"
        Frame3.Visible = True
        cadena = ""
        cadena.SetFocus
        Command13_Click
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        Command12_Click
        Exit Sub

    End If

    If KeyCode = &H71 Then  'f2   'cambia precios
        Exit Sub

    End If

    Exit Sub
cmd102_err:
    'MsgBox "Seleccione un producto ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub DBGrid3_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        cadena.SetFocus
        Exit Sub

    End If

    Exit Sub

    If KeyCode = 13 Then
        If opcion1 = "1" Then
            Frame3.Visible = False
            dbGrid1.SetFocus

        End If

        If opcion1 = "2" Then
            Frame3.Visible = False
            dbGrid1.SetFocus

        End If

        If opcion1 = "3" Then
            Frame3.Visible = False
            dbGrid1.SetFocus

        End If

    End If

End Sub

Private Sub dbgrid3_KeyPress(KeyAscii As Integer)

    Dim buf  As String

    Dim buf2 As String

    If KeyAscii <> 13 And KeyAscii <> 27 Then
        If KeyAscii = 8 Then
            If Len(cadena) > 0 Then
                buf = Mid$(cadena, 1, Len(cadena) - 1)
                cadena = buf
                KeyAscii = 0
            Else
                KeyAscii = 0
                Exit Sub

            End If

        End If

        buf = Chr(KeyAscii)

        If Chr(KeyAscii) = "*" Then
            buf = ""
            cadena = buf

        End If

        If KeyAscii <> 13 Then
            cadena = cadena + buf

        End If

        buf = cadena
        ejecuta 0
         
    End If

End Sub

Private Sub DBGrid4_AfterColUpdate(ByVal ColIndex As Integer)

    Select Case ColIndex

        Case 3

        Case 6

        Case 9

        Case 12

        Case 15

        Case 18

        Case 21

        Case 24

        Case 27

        Case 30

        Case Else
            'Cancel = True
            Exit Sub

    End Select

End Sub

Private Sub dbgrid4_BeforeColUpdate(ByVal ColIndex As Integer, _
                                    OldValue As Variant, _
                                    Cancel As Integer)

    Select Case ColIndex

        Case 3

            If Len(DBGrid4.columns(1)) = 0 Or Len(DBGrid4.columns(2)) = 0 Then
                Cancel = True
                Exit Sub

            End If

        Case 6

            If Len(DBGrid4.columns(4)) = 0 Or Len(DBGrid4.columns(5)) = 0 Then
                Cancel = True
                Exit Sub

            End If

        Case 9

            If Len(DBGrid4.columns(7)) = 0 Or Len(DBGrid4.columns(8)) = 0 Then
                Cancel = True
                Exit Sub

            End If

        Case 12

            If Len(DBGrid4.columns(10)) = 0 Or Len(DBGrid4.columns(11)) = 0 Then
                Cancel = True
                Exit Sub

            End If

        Case 15

            If Len(DBGrid4.columns(13)) = 0 Or Len(DBGrid4.columns(14)) = 0 Then
                Cancel = True
                Exit Sub

            End If

        Case 18

            If Len(DBGrid4.columns(17)) = 0 Or Len(DBGrid4.columns(16)) = 0 Then
                Cancel = True
                Exit Sub

            End If

        Case 21

            If Len(DBGrid4.columns(19)) = 0 Or Len(DBGrid4.columns(20)) = 0 Then
                Cancel = True
                Exit Sub

            End If

        Case 24

            If Len(DBGrid4.columns(22)) = 0 Or Len(DBGrid4.columns(23)) = 0 Then
                Cancel = True
                Exit Sub

            End If

        Case 27

            If Len(DBGrid4.columns(26)) = 0 Or Len(DBGrid4.columns(25)) = 0 Then
                Cancel = True
                Exit Sub

            End If

        Case 30

            If Len(DBGrid4.columns(29)) = 0 Or Len(DBGrid4.columns(28)) = 0 Then
                Cancel = True
                Exit Sub

            End If

        Case Else
            Cancel = True
            Exit Sub

    End Select

End Sub

Private Sub descripcio_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    Command5_Click

End Sub

'13/08/2018 Integración FE - Pizzeria
Private Sub dj7833re_Click()

    Dim found      As Integer

    Dim I          As Integer

    Dim v          As Long

    Dim R          As Long

    Dim ih         As Integer

    Dim h          As Integer

    Dim cad        As String

    Dim Tmp        As String

    Dim sw         As Integer

    Dim sdx        As Double

    Dim mytabley   As New ADODB.Recordset

    Dim Heading(8) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    Dim indx       As Long

    On Error GoTo cmd45612_err

    If Frame6.Visible = True Then Exit Sub

    If Frame5.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    If MsgBox("Desea Exportar excel", 1, "Aviso") <> 1 Then Exit Sub

    Heading(1) = "Producto"
    Heading(2) = "Descripcio"
    Heading(3) = "Und"
    Heading(4) = "Factor"
    Heading(5) = "Cantidad"
    Heading(6) = "Costo"
    Heading(7) = "Total"

    If rrproducto.RecordCount = 0 Then Exit Sub
    rrproducto.MoveFirst
    
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    
    '' 11/12/2017 SubReceta
    'Call Formato_Excel(7, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    Call Formato_ExcelReceta(7, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
    objExcel.ActiveSheet.Cells(1, 2) = "                                                     LISTADO DE RECETAS"
    objExcel.ActiveSheet.Cells(1, 2).Font.bold = True
    objExcel.ActiveSheet.Cells(1, 2).Font.Size = 14
    objExcel.ActiveSheet.Cells(1, 2).Font.color = RGB(0, 112, 184)
    '' 11/12/2017 SubReceta
  
    v = 4
    h = 1
    indx = 1
    
    Do

        If rrproducto.EOF Then Exit Do
        If mytabley.State = 1 Then
            mytabley.Close
            Set mytabley = Nothing

        End If
         
        '' 11/12/2017 SubReceta
        'mytabley.Open "Select * from receta where  producto='" & "" & rrproducto.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic
        mytabley.Open "Select * from receta where linea='' and producto='" & "" & rrproducto.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic
        '' 11/12/2017 SubReceta
         
        If mytabley.RecordCount = 0 Then
            mytabley.Close
            GoTo sigame1

        End If
     
        objExcel.ActiveSheet.Cells(v, 1) = "Receta Nro " & indx
        indx = indx + 1
        v = v + 1
   
        objExcel.ActiveSheet.Cells(v, 1) = "" & rrproducto.Fields("Producto")
        objExcel.ActiveSheet.Cells(v, 2) = "" & rrproducto.Fields("Descripcio")
   
        '' 11/12/2017 SubReceta
        objExcel.ActiveSheet.Cells(v, 1).Font.bold = True
        objExcel.ActiveSheet.Cells(v, 2).Font.bold = True
        objExcel.ActiveSheet.Cells(v, 3).Font.bold = True
        objExcel.ActiveSheet.Cells(v, 4).Font.bold = True
   
        objExcel.ActiveSheet.Cells(v, 1).Font.color = RGB(62, 95, 138)
        objExcel.ActiveSheet.Cells(v, 2).Font.color = RGB(62, 95, 138)
        objExcel.ActiveSheet.Cells(v, 3).Font.color = RGB(62, 95, 138)
        objExcel.ActiveSheet.Cells(v, 4).Font.color = RGB(62, 95, 138)
        '' 11/12/2017 SubReceta
   
        ''' 11/12/2017 SubReceta
        If Val(rrproducto.Fields("platos")) > 1 Then
            objExcel.ActiveSheet.Cells(v, 3) = "Porciones"
            objExcel.ActiveSheet.Cells(v, 4) = "'(" & "" & rrproducto.Fields("platos") & ")"

        End If
    
        'objExcel.ActiveSheet.Cells(v, 3) = busca_factorProduccion("" & mytabley.Fields("producto"), 0)
        'objExcel.ActiveSheet.Cells(v, 4) = busca_factorProduccion("" & mytabley.Fields("producto"), 1)
        ''' 11/12/2017 SubReceta
   
        v = v + 1
        mytabley.MoveFirst
        sw = 0
        sdx = 0
        Do

            If mytabley.EOF Then Exit Do
            objExcel.ActiveSheet.Cells(v, h + 0) = "" & mytabley.Fields("Productoi")
            objExcel.ActiveSheet.Cells(v, h + 1) = "" & mytabley.Fields("Descripcio")
            objExcel.ActiveSheet.Cells(v, h + 2) = "" & mytabley.Fields("Unidad")
            objExcel.ActiveSheet.Cells(v, h + 3) = "" & mytabley.Fields("Factor")
            objExcel.ActiveSheet.Cells(v, h + 4) = "" & mytabley.Fields("Cantidad")
            objExcel.ActiveSheet.Cells(v, h + 5) = "" & mytabley.Fields("Precio")
            objExcel.ActiveSheet.Cells(v, h + 6) = "" & mytabley.Fields("Total")
            sdx = sdx + Val("" & mytabley.Fields("Total"))
            v = v + 1
            mytabley.MoveNext
        Loop
        'objExcel.ActiveSheet.Cells(v, h + 6) = Format(Val("" & rrproducto.Fields("costou")), "0.00")

        '' 11/12/2017 SubReceta
        'objExcel.ActiveSheet.Cells(v, h + 6) = Format(sdx, "0.00")
    
        '11/06/2018 Actualiza Precio Promedio Ponderado Masivo
        'objExcel.ActiveSheet.Cells(v, h + 6) = Format(sdx, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 6) = Format(sdx, "0.00000")
        '11/06/2018 Actualiza Precio Promedio Ponderado Masivo
 
        objExcel.ActiveSheet.Cells(v, h + 6).Font.bold = True
        objExcel.ActiveSheet.Cells(v, h + 6).Interior.color = RGB(248, 243, 53)
        '' 11/12/2017 SubReceta

        mytabley.Close
sigame1:
        rrproducto.MoveNext
    Loop
    
    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    Exit Sub
cmd45612_err:
    MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
    Exit Sub

End Sub

'13/08/2018 Integración FE - Pizzeria

Private Sub dk343_Click()

    Dim found As Integer

    If puede_modificar() = 0 Then Exit Sub
    If Frame6.Visible = True Then Exit Sub

    If Frame5.Visible = True Then Exit Sub

    If Frame3.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    tproduct.codigo.Enabled = True
    FLAG = ""
    tproduct.ordename = "NUEVO"
    tproduct.Show 1

    If FLAG = "1" Then
        found = sql_cabeza(0)

    End If

    FLAG = ""

End Sub

Private Sub dk8923_Click()

End Sub

Private Sub dk89231_Click()

    If Frame6.Visible = True Then Exit Sub

    If Frame5.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    If MsgBox("Desea Exportar excel", 1, "Aviso") <> 1 Then Exit Sub
    precio_excell 0

End Sub

Private Sub dk89331_Click()

    On Error GoTo cmd36712_err

    Dim quebusco As String

    If Frame6.Visible = True Then Exit Sub

    If Frame5.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    quebusco = Trim("" & dbGrid1.columns(1))
    reporgen.NAMETABLA = "producto"
    reporgen.Show 1
    rrproducto.Find "producto='" & quebusco & "'"
    Exit Sub
cmd36712_err:
    MsgBox "Seleccion un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub dk898930_Click()

End Sub

Private Sub dkj8383_Click()
    tenvioda.Show 1

End Sub

Private Sub familia_Click()

    If extra_loquesea1(familia) <> "%" Then
        carga_subfamilia

    End If

End Sub

Private Sub fdk883_Click()

    Dim mytablex     As New ADODB.Recordset

    Dim found        As Integer

    Dim I            As Integer

    Dim v            As Long

    Dim R            As Long

    Dim ih           As Integer

    Dim h            As Integer

    Dim vprecios(10) As String

    Dim cad          As String

    Dim Tmp          As String

    Dim sw           As Integer

    Dim sdx          As Double

    Dim xsw          As Integer

    Dim mytabley     As New ADODB.Recordset
    
    Dim Heading(11)  As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    On Error GoTo cmd56120_err

    If Frame6.Visible = True Then Exit Sub

    If Frame5.Visible = True Then Exit Sub
    ordenado.ListIndex = 4
    found = sql_cabeza(0)

    If rrproducto.RecordCount = 0 Then Exit Sub
    
    Heading(1) = "Producto"
    Heading(2) = "Descripcio"
    Heading(3) = "Unidad"
    Heading(4) = "Factor"
    Heading(5) = "Minimo"
    Heading(6) = "Maximo"
    Heading(7) = "Local"
    Heading(8) = "Almacen"
    Heading(9) = "Saldo"
    Heading(10) = "Reponer"

    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_Excel(10, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    v = 5
    h = 1
    rrproducto.MoveFirst
    sw = 0
    Do

        If rrproducto.EOF Then Exit Do
        If mytablex.State = 1 Then
            mytablex.Close
            Set mytablex = Nothing

        End If

        If Val("" & rrproducto.Fields("minimo")) = 0 Or Val("" & rrproducto.Fields("maximo")) = 0 Then
            GoTo ajicali

        End If
     
        cad = "select * from almacen where producto='" & Trim(rrproducto.Fields("producto")) & "'"

        If local1 <> "%" Then
            cad = cad & "' and local='" & local1 & "'"

        End If

        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
     
            If sw = 0 Then
                Tmp = "" & rrproducto.Fields("familia")
                sw = 1
                objExcel.ActiveSheet.Cells(v, h) = "" & rrproducto.Fields("familia")
                v = v + 1

            End If

            If Tmp <> "" & rrproducto.Fields("familia") Then
                Tmp = "" & rrproducto.Fields("familia")
                objExcel.ActiveSheet.Cells(v, h) = "" & rrproducto.Fields("familia")
                v = v + 1

            End If

            '---------------------------------------------
            objExcel.ActiveSheet.Cells(v, h) = "" & rrproducto.Fields("producto")
            objExcel.ActiveSheet.Cells(v, h + 1) = "" & rrproducto.Fields("descripcio")
            objExcel.ActiveSheet.Cells(v, h + 2) = "" & rrproducto.Fields("Und")
            objExcel.ActiveSheet.Cells(v, h + 3) = "" & rrproducto.Fields("Fac")
            objExcel.ActiveSheet.Cells(v, h + 4) = "" & rrproducto.Fields("Minimo")
            objExcel.ActiveSheet.Cells(v, h + 5) = "" & rrproducto.Fields("Maximo")
            'v = v + 1
            xsw = 0
            Do

                If mytablex.EOF Then Exit Do
                objExcel.ActiveSheet.Cells(v, h + 6) = "'" & mytablex.Fields("Local")
                objExcel.ActiveSheet.Cells(v, h + 7) = "'" & mytablex.Fields("bodega")
                objExcel.ActiveSheet.Cells(v, h + 8) = "'" & mytablex.Fields("saldo")
                sdx = 0

                If Val("" & mytablex.Fields("saldo")) <= Val("" & rrproducto.Fields("minimo")) Then
                    sdx = Val("" & rrproducto.Fields("maximo")) - Val("" & mytablex.Fields("saldo"))
                    objExcel.ActiveSheet.Cells(v, h + 9) = sdx
                Else
                    objExcel.ActiveSheet.Cells(v, h + 9) = sdx

                End If

                v = v + 1
                xsw = 1
                mytablex.MoveNext
            Loop

            If xsw = 0 Then
                objExcel.ActiveSheet.Cells(v, h + 9) = "'Sin Almacen"
                v = v + 1

            End If
      
        End If

ajicali:
    
        rrproducto.MoveNext
    Loop
    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    Exit Sub
cmd56120_err:
    MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub fk3434_Click()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd45_err

    If puede_modificar() = 0 Then Exit Sub
    If Frame6.Visible = True Then Exit Sub

    If Frame5.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub

    mytablex.Open "select producto from detalle where producto='" & Trim("" & dbGrid1.columns(1)) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
   
        If MsgBox("Existe Movimiento de Producto,Desea Borra ", 1, "Aviso") <> 1 Then
            mytablex.Close
            Exit Sub

        End If

    End If

    mytablex.Close

    If MsgBox("Desea Borrar " & "" & dbGrid1.columns(1), 1, "Aviso") <> 1 Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    borra_almacen_producto Trim("" & dbGrid1.columns(1))

    found = sql_cabeza(0)
    'Data1.Recordset.Delete
    Exit Sub
cmd45_err:
    MsgBox "Seleccione un producto ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fk78333_Click()

End Sub

Private Sub fk88555re_Click()

    If Frame6.Visible = True Then Exit Sub

    If Frame5.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    Frame6.Visible = True

End Sub

Private Sub flo434_Click()

    If Frame5.Visible = True Then
        Frame5.Visible = False
        Exit Sub

    End If

    If Frame3.Visible = True Then
        cadena_KeyPress 27
        Exit Sub

    End If

    If Frame1.Visible = True Then
        cmdCancelar_Click
        Exit Sub

    End If

    If Frame2.Visible = True Then
        Command11_Click
        Exit Sub

    End If

    xprodet.Hide
    Unload xprodet

End Sub

Sub cargas_iniciales()

    Dim cad      As String

    Dim mytablex As New ADODB.Recordset

    local1.Clear
    'local1.AddItem "*"
    'Set mytablex = mydbxglo.OpenTable("tlocal")
    'Do
    'If mytablex.EOF Then Exit Do
    'local1.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("Nombre")
    'mytablex.MoveNext
    'Loop
    'local1.ListIndex = 0
    'mytablex.Close

    diauso.Clear
    diauso.AddItem "%"
    diauso.AddItem "LUNES"
    diauso.AddItem "MARTES"
    diauso.AddItem "MIERCOLES"
    diauso.AddItem "JUEVES"
    diauso.AddItem "VIERNES"
    diauso.AddItem "SABADO"
    diauso.AddItem "DOMINGO"
    diauso.ListIndex = 0

    local1.Clear
    local1.AddItem "%"
    local1.AddItem "01"
    local1.AddItem "02"
    local1.AddItem "03"
    local1.AddItem "04"
    local1.ListIndex = 0

    familia.Clear
    familia.AddItem "%"

    cad = "SELECT * FROM FAMILIA  order by descripcio "

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open cad, cn, adOpenStatic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        familia.AddItem Trim("" & mytablex.Fields("descripcio")) & "|" & mytablex.Fields("familia")
        mytablex.MoveNext
    Loop
    familia.ListIndex = 0
    mytablex.Close
    Set mytablex = Nothing

    subfamilia.Clear
    subfamilia.AddItem "%"

    cad = "SELECT * FROM subfamil  order by subfamilia "

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open cad, cn, adOpenStatic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        subfamilia.AddItem "" & mytablex.Fields("subfamilia") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    subfamilia.ListIndex = 0
    mytablex.Close
    seccion.Clear
    seccion.AddItem "%"

    cad = "SELECT * FROM seccion  order by seccion "

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open cad, cn, adOpenStatic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        seccion.AddItem "" & mytablex.Fields("seccion") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    seccion.ListIndex = 0
    mytablex.Close

    marca.Clear
    marca.AddItem "%"
    cad = "SELECT * FROM marca  order by marca "

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open cad, cn, adOpenStatic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        marca.AddItem "" & mytablex.Fields("marca") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    marca.ListIndex = 0
    mytablex.Close

    ''18/07/2017 kenyo tienda ropa opciones producto
    sexo.Clear
    sexo.AddItem "%"
    cad = "SELECT * FROM sexo  order by sexo "

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open cad, cn, adOpenStatic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        sexo.AddItem "" & mytablex.Fields("sexo") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    sexo.ListIndex = 0
    mytablex.Close

    talla.Clear
    talla.AddItem "%"
    cad = "SELECT * FROM talla  order by talla "

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open cad, cn, adOpenStatic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        talla.AddItem "" & mytablex.Fields("talla") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    talla.ListIndex = 0
    mytablex.Close

    proyecto.Clear
    proyecto.AddItem "%"
    cad = "SELECT * FROM proyecto  order by proyecto "

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open cad, cn, adOpenStatic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        proyecto.AddItem "" & mytablex.Fields("proyecto") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    proyecto.ListIndex = 0
    mytablex.Close

    procedencia.Clear
    procedencia.AddItem "%"
    cad = "SELECT * FROM procedencia  order by procedencia "

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open cad, cn, adOpenStatic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        procedencia.AddItem "" & mytablex.Fields("procedencia") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    procedencia.ListIndex = 0
    mytablex.Close

    ''18/07/2017 kenyo tienda ropa opciones producto

    categoria.Clear
    categoria.AddItem "%"

    cad = "SELECT * FROM categori  order by categoria "

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open cad, cn, adOpenStatic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        categoria.AddItem "" & mytablex.Fields("categoria") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    categoria.ListIndex = 0
    mytablex.Close

    linea.Clear
    linea.AddItem "%"

    cad = "SELECT * FROM Linea  order by linea "

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open cad, cn, adOpenStatic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        linea.AddItem "" & mytablex.Fields("linea") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    linea.ListIndex = 0
    mytablex.Close
    color.Clear
    color.AddItem "%"

    cad = "SELECT * FROM color  order by color "

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open cad, cn, adOpenStatic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        color.AddItem "" & mytablex.Fields("color") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    color.ListIndex = 0
    mytablex.Close

    proveedor.Clear
    proveedor.AddItem "%"

    cad = "SELECT * FROM proveedo  order by codigo"

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open cad, cn, adOpenStatic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        proveedor.AddItem "" & mytablex.Fields("Codigo") & "|" & mytablex.Fields("Nombre")
        mytablex.MoveNext
    Loop
    proveedor.ListIndex = 0
    mytablex.Close

    bodega.Clear
    bodega.AddItem "%"

    cad = "SELECT * FROM bodega  order by codigo "

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open cad, cn, adOpenStatic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        bodega.AddItem "" & mytablex.Fields("Codigo") & "|" & mytablex.Fields("Nombre")
        mytablex.MoveNext
    Loop
    bodega.ListIndex = 0
    mytablex.Close

    'margen.Clear
    'margen.AddItem "%"

    'cad = "SELECT * FROM margen  order by margen "
    '   If mytablex.State = 1 Then mytablex.Close
    '   mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
    '
    'Do
    'If mytablex.EOF Then Exit Do
    'margen.AddItem "" & mytablex.Fields("margen") & "|" & mytablex.Fields("descripcio")
    'mytablex.MoveNext
    'Loop
    'margen.ListIndex = 0
    'mytablex.Close

    'MsgBox "xx"
End Sub

Private Sub Form_Activate()

    Dim found As Integer

    If esactivo = "" Then
        cargas_iniciales
        found = sql_cabeza(0)
        esactivo = "1"
        ve_permisos
        esactivo = "S"
        carga_campos
        producto = "%"

        'MsgBox "xxx"
    End If

End Sub

Sub carga_campos()

    Dim I As Integer

    Combo3.Clear
    'Combo3.AddItem ""
    Combo3.AddItem "+"
    Combo3.AddItem "-"
    Combo3.ListIndex = 0
   
    Combo2.Clear
    Combo2.AddItem "%"
    Combo2.AddItem "Igual"
    Combo2.AddItem "Distinto"
    Combo2.AddItem "Mayor"
    Combo2.AddItem "Menor"
    Combo2.AddItem "MayorIgual"
    Combo2.AddItem "MenorIgual"
    Combo2.ListIndex = 0

    Combo1.Clear
    Combo1.AddItem "%"

    For I = 0 To rrproducto.Fields.count - 1
        Combo1.AddItem Trim(rrproducto.Fields(I).Name)
    Next I

    Combo1.ListIndex = 0

End Sub

Sub ve_permisos()

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then
        mytablex.Close

    End If

    mytablex.Open "select * from vendedor where codigo='" & Trim(gusuario) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF = True And mytablex.BOF = True Then
        mytablex.Close
        Exit Sub

    End If

    If Mid$("" & mytablex.Fields("RW1"), 1, 1) = "R" Then
        Modif34.Enabled = False
        fk3434.Enabled = False
        dk343.Enabled = False

    End If

    mytablex.Close

End Sub

Function sql_cabeza(sw As Integer)

    On Error GoTo cmd37_err

    Dim indx      As Integer

    Dim buf       As String

    Dim buf1      As String

    Dim queprecio As String

    Dim queunidad As String

    Dim quefactor As String

    Dim xprecio   As String

    Dim fvprecio  As String

    Dim X         As Integer

    'MsgBox "aa"}
    sw = 0
    nro_registros = ""
    queprecio = "precios.pventa1 as Precio "
    queunidad = "precios.unidad1 as Und "
    quefactor = "precios.factor1 as fac "
    fvprecio = "precios.fechavp"
    xprecio = queunidad & "," & quefactor & "," & queprecio
   
    buf = "select Producto.Descripcio,Producto.producto,Producto.Ok as F,Producto.Marca,Unidad as Und,factor as Fac,Costou ,Producto.Monedac as M,Producto.Familia,Producto.Subfamilia,Producto.barras,producto.linea,producto.Seccion,producto.fotonombre,producto.flete,producto.igv,producto.Oferta,Producto.minimo,producto.maximo,producto.fechavence,producto.descorto,producto.monedav,producto.factor,producto.platos,producto.detalle,producto.Color,producto.productoequ,producto.percepcion from producto where descripcio like '%" & descripcio & "%' "

    If proveedor <> "%" Then
        local1.ListIndex = 0
        bodega.ListIndex = 0
        buf = "select  Producto.Descripcio,Producto.producto,Producto.Ok as F,Producto.Marca,Unidad as Und,factor as Fac,Costou ,Producto.Monedac as M,Producto.Familia,Producto.Subfamilia,Producto.barras,producto.linea,producto.Seccion,producto.fotonombre,producto.flete,producto.igv,producto.Oferta,Producto.minimo,producto.maximo,producto.fechavence,producto.descorto,producto.monedav,producto.factor,producto.platos,producto.detalle,producto.Color,producto.productoequ,producto.percepcion from producto INNER JOIN codprov  on  producto.producto=codprov.producto  and codprov.codigo='" & extra_loquesea(proveedor) & "'"

    End If
      
    If bodega <> "%" Then
        local1.ListIndex = 0
        proveedor.ListIndex = 0
        buf = "select  Producto.Descripcio,Producto.producto,Producto.Ok as F,Producto.Marca,Unidad as Und,factor as Fac,Almacen.saldo ,Producto.Monedac as M,Producto.Familia,Producto.Subfamilia,Producto.barras,producto.linea,producto.Seccion,producto.fotonombre,producto.flete,producto.igv,producto.Oferta,Producto.minimo,producto.maximo,producto.fechavence,producto.descorto,producto.monedav,producto.factor,producto.platos,producto.detalle,producto.Color,producto.productoequ,producto.percepcion from producto INNER JOIN almacen  on  producto.producto=almacen.producto  and almacen.bodega='" & extra_loquesea(bodega) & "'"

    End If

    ' buf = "select * from producto "

    'busqueda por varios codigos de barras 07/07/2018
    If Barras <> "%" Then

        'busqueda por varios codigos de barras 07/07/2018
        Dim bus As String

        bus = ""
        bus = consulta_variosCodbarras(Barras)

        If Len(bus) > 0 Then
            'producto = bus
            buf = buf & " and producto.producto like '" & bus & "'"
        Else
            buf = buf & " and producto.barras like '" & Barras & "'"

        End If

    End If

    'busqueda por varios codigos de barras 07/07/2018

    If producto <> "%" Then
        buf = buf & " and producto.producto like '" & producto & "'"

    End If

    If descripcio <> "%" Then
        buf = buf & " and producto.descripcio like  '%" & descripcio & "%'"

        ' '%" & buffer & "%'"
    End If

    '%" & buffer & "%'"
    If igv = "EXENTO" Then
        buf = buf & " and (producto.IGV=0 or producto.IGV=NULL)"

    End If

    If igv = "GRAVADO" Then
        buf = buf & " and producto.IGV>0"

    End If

    If marca <> "%" Then
        buf = buf & " and producto.marca like '" & extra_loquesea(marca) & "'"

    End If

    If activo <> "%" Then
        buf = buf & " and producto.estado like '" & activo & "'"

    End If

    If fechavi <> "%" And fechavf <> "%" Then
        If IsDate(fechavi) And IsDate(fechavf) Then
            buf = buf & "  and producto.fechavence>='" & Format(fechavi, "YYYYMMDD") & "'"
            buf = buf & " and producto.fechavence<='" & Format(fechavf, "YYYYMMDD") & "' "

        End If

    End If

    If local1 <> "%" Then
        If fechavpi <> "%" And fechavpf <> "%" Then
            If IsDate(fechavpi) And IsDate(fechavpf) Then
                buf = buf & "  and precios.fechavp>='" & Format(fechavpi, "YYYYMMDD") & "'"
                buf = buf & " and precios.fechavp<='" & Format(fechavpf, "YYYYMMDD") & "' "

            End If

        End If

    End If

    If f1 = "S" Then
        buf = buf & " and producto.OK='F'"

    End If

    If SEINVENTARIA <> "%" Then
        buf = buf & " and producto.seinventaria='" & SEINVENTARIA & "'"

    End If

    If vecaja <> "%" Then
        buf = buf & " and producto.vecaja like '" & vecaja & "'"

    End If

    If percepcion <> "%" Then
        buf = buf & " and producto.percepcion like '" & percepcion & "'"

    End If

    If familia <> "%" Then
        buf = buf & " and producto.familia like '" & extra_loquesea1(familia) & "'"

    End If

    If subfamilia <> "%" Then
        buf = buf & " and producto.subfamilia like '" & extra_loquesea(subfamilia) & "'"

    End If

    If seccion <> "%" Then
        buf = buf & " and producto.seccion like '" & extra_loquesea(seccion) & "'"

    End If

    If diauso <> "%" Then
        buf = buf & " and producto.dia like '" & extra_loquesea(diauso) & "'"

    End If

    If categoria <> "%" Then
        buf = buf & " and producto.categoria like '" & extra_loquesea(categoria) & "'"

    End If

    ''18/07/2017 kenyo tienda ropa opciones producto
    If sexo <> "%" Then
        buf = buf & " and producto.sexo like '" & extra_loquesea(sexo) & "'"

    End If

    If talla <> "%" Then
        buf = buf & " and producto.talla like '" & extra_loquesea(talla) & "'"

    End If

    If proyecto <> "%" Then
        buf = buf & " and producto.proyecto like '" & extra_loquesea(proyecto) & "'"

    End If

    If procedencia <> "%" Then
        buf = buf & " and producto.procedencia like '" & extra_loquesea(procedencia) & "'"

    End If

    ''18/07/2017 kenyo tienda ropa opciones producto

    If color <> "%" Then
        buf = buf & " and producto.color like '" & extra_loquesea(color) & "'"

    End If

    If oferta <> "%" Then
        buf = buf & " and producto.oferta like '" & oferta & "'"

    End If

    If monedav <> "%" Then
        buf = buf & " and producto.monedav like '" & monedav & "'"

    End If

    If Peso <> "%" Then
        buf = buf & " and producto.peso like '" & Peso & "'"

    End If

nathing:

    If sw = 1 Then

        If Combo1.List(Combo1.ListIndex) <> "%" And Combo2.List(Combo2.ListIndex) <> "%" Then
            'sw = 1
            buf = buf & "and producto." & Combo1.List(Combo1.ListIndex)
            buf = buf & poner_signo(Combo2.List(Combo2.ListIndex))
            buf = buf & "" & criterio.Text

        End If

    End If

    If ordenado <> "%" Then
        buf = buf & " order by " & ordenado

    End If

    sqldatos = buf

    'MsgBox sqldatos
    If rrproducto.State = 1 Then rrproducto.Close
    Set rrproducto = Nothing
    rrproducto.Open buf, cn, adOpenStatic, adLockOptimistic
    'If rrproducto.RecordCount = 0 Then
    '
    '   Exit Function
    'End If
    nro_registros = "" & rrproducto.RecordCount
   
    Set dbGrid1.DataSource = rrproducto
    dbGrid1.Col = 0
    dbGrid1.columns(0).Width = 5000
    dbGrid1.columns(1).Width = 1300
    dbGrid1.columns(2).Width = 250
    dbGrid1.columns(3).Width = 1000
    dbGrid1.columns(4).Width = 900
    dbGrid1.columns(5).Width = 500
    dbGrid1.columns(6).Width = 800
    dbGrid1.columns(7).Width = 400

    If local1 <> "%" Then
        dbGrid1.columns(7).Width = 1200

    End If

    dbGrid1.columns(8).Width = 800
    dbGrid1.columns(9).Width = 800
    dbGrid1.columns(10).Width = 1500
    dbGrid1.columns(11).Width = 800
    dbGrid1.SetFocus
    'dbGrid1.Columns(12).Width = 800
    'dbGrid1.Columns(13).Width = 800
    'End If
    'MsgBox ""
    sql_cabeza = 1

    'If rrproducto.RecordCount > 0 Then
    '   rrproducto.MoveLast
    '   dbGrid1.SetFocus
    'End If
    If rrproducto.RecordCount > 0 Then
        consulta_precios Trim("" & rrproducto.Fields("producto"))

    End If
               
    Exit Function
cmd37_err:
    MsgBox "Error en select " & error$, 48, "Aviso"
    Exit Function

End Function

'busqueda por varios codigos de barras 07/07/2018
Function consulta_variosCodbarras(buf1 As String)

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    consulta_variosCodbarras = ""
    mytablex.Open "select producto from producto where barras='" & "" & buf1 & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytabley.Open "select producto from productb where barras='" & "" & buf1 & "'", cn, adOpenStatic, adLockOptimistic

        If mytabley.RecordCount > 0 Then
            consulta_variosCodbarras = "" & mytabley.Fields("producto")

        End If

        mytabley.Close

    End If

    mytablex.Close
    Exit Function
  
End Function

'busqueda por varios codigos de barras 07/07/2018

Private Sub Form_Load()

    Frame1.Top = 0: Frame1.Left = 0
    Frame2.Top = 0: Frame2.Left = 0
    Frame3.Top = 0: Frame3.Left = 0
    Frame4.Top = 0: Frame4.Left = 0
    Frame5.Top = 0: Frame5.Left = 0
    Frame6.Top = 0: Frame6.Left = 0

    Dim mytablex As New ADODB.Recordset

    SEINVENTARIA.Clear
    SEINVENTARIA.AddItem "%"
    SEINVENTARIA.AddItem "S"
    SEINVENTARIA.ListIndex = 0

    Combo5.Clear
    Combo5.AddItem ""
    Combo5.AddItem "Pventa1,Unidad1,Factor1"
    Combo5.AddItem "Pventa2,Unidad2,Factor2"
    Combo5.AddItem "Pventa3,Unidad3,Factor3"
    Combo5.AddItem "Pventa4,Unidad4,Factor4"
    Combo5.AddItem "Pventa5,Unidad5,Factor5"
    Combo5.AddItem "Pventa6,Unidad6,Factor6"
    Combo5.AddItem "Pventa7,Unidad7,Factor7"
    Combo5.AddItem "Pventa8,Unidad8,Factor8"
    Combo5.AddItem "Pventa9,Unidad9,Factor9"
    Combo5.AddItem "Pventa10,Unidad10,Factor10"
    Combo5.ListIndex = 0

    Combo4.Clear
    Combo4.AddItem ""
    'Combo4.AddItem "00"
    Combo4.AddItem "01"
    Combo4.AddItem "02"
    Combo4.AddItem "03"
    Combo4.AddItem "04"
    Combo4.AddItem "05"
    Combo4.AddItem "06"
    Combo4.AddItem "07"
    Combo4.AddItem "08"
    Combo4.AddItem "09"
    Combo4.AddItem "10"
    Combo4.AddItem "11"
    Combo4.AddItem "00"
    Combo4.ListIndex = 0

    'anno = Format(Year(Now), "0000")
    igv.Clear
    igv.AddItem "%"
    igv.AddItem "EXENTO"
    igv.AddItem "GRAVADO"
    igv.ListIndex = 0

    Peso.AddItem "%"
    Peso.AddItem "S"
    Peso.AddItem "N"
    Peso.ListIndex = 0

    vecaja.AddItem "%"
    vecaja.AddItem "S"
    vecaja.AddItem "N"
    vecaja.ListIndex = 0

    activo.AddItem "S"
    activo.AddItem "N"
    activo.AddItem "%"

    activo.ListIndex = 0

    ordenado.Clear

    ordenado.AddItem "%"
    ordenado.AddItem "producto.Descripcio"
    ordenado.AddItem "producto.Producto"
    ordenado.AddItem "producto.Marca"
    ordenado.AddItem "producto.Familia"
    ordenado.AddItem "producto.SubFamilia"
    ordenado.AddItem "producto.Linea"
    ordenado.AddItem "producto.Categoria"
    ordenado.AddItem "producto.Color"

    ordenado.ListIndex = 0

    'subodega.Clear
    'subodega.AddItem "%"
    'mytablex.Open "SELECT * from bodega ", cn, adOpenKeyset, adLockOptimistic
    'Do
    'If mytablex.EOF Then Exit Do
    'subodega.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
    'mytablex.MoveNext
    'Loop
    'mytablex.Close
    'subodega.ListIndex = 0

    'sulocal.Clear
    'sulocal.AddItem "%"
    'mytablex.Open "SELECT * from tlocal ", cn, adOpenKeyset, adLockOptimistic
    'Do
    'If mytablex.EOF Then Exit Do
    'sulocal.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
    'mytablex.MoveNext
    'Loop
    'mytablex.Close
    'sulocal.ListIndex = 0

End Sub

Private Sub igv_Click()

    'Command5_Click
End Sub

Private Sub kproducto_Change()

    On Error GoTo cmd390_err

    'busca_precioss "" & kproducto
    'consulta_precios "" & kproducto
    Exit Sub
cmd390_err:
    Exit Sub

End Sub

Private Sub Label22_Click()

    Dim mytablex As New ADODB.Recordset

    If Len(Trim(codigoa)) = 0 Then
        codigoa = ""
        codigoa.SetFocus
        Exit Sub

    End If

    If Len(Trim(codigon)) = 0 Then
        codigon = ""
        codigon.SetFocus
        Exit Sub

    End If

    '' 21/11/2017 Validar Cambio de Codigo de Producto
    Dim mytablexyz As New ADODB.Recordset

    mytablexyz.Open "select producto from producto where producto='" & Trim(codigon) & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablexyz.RecordCount > 0 Then
        MsgBox "Ya Existe Código de Producto Usado ", 48, "Aviso"
        mytablexyz.Close
        codigon.SetFocus
        Exit Sub

    End If

    '' 21/11/2017 Validar Cambio de Codigo de Producto

    mytablex.Open "select producto from producto where producto='" & Trim(codigoa) & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        MsgBox "No Existe Producto ", 48, "Aviso"
        mytablex.Close
        codigoa.SetFocus
        Exit Sub

    End If

    mytablex.Fields("producto") = "" & Trim(codigon)
    mytablex.Update
    mytablex.Close
    cn.Execute ("update precios set producto='" & Trim(codigon) & "' where producto='" & Trim(codigoa) & "'")
    cn.Execute ("update receta set producto='" & Trim(codigon) & "' where producto='" & Trim(codigoa) & "'")
    MsgBox "Proceso Realizado ", 48, "Aviso"
    codigoa = ""
    codigon = ""
    codigoa.SetFocus
   
    '' 21/11/2017 Validar Cambio de Codigo de Producto
    Frame5.Visible = False
    '' 21/11/2017 Validar Cambio de Codigo de Producto
   
    Exit Sub

End Sub

Private Sub Label23_Click()
    Frame5.Visible = False

End Sub

Private Sub Label27_Click()
    Frame6.Visible = False

End Sub

Private Sub Label28_Click()

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd90800_err

    If Len(Trim(Combo3.Text)) = 0 Then
        MsgBox "Seleccion Tipo"
        Exit Sub

    End If

    If Len(Trim(Combo4.Text)) = 0 Then
        MsgBox "Seleccion Lista"
        Exit Sub

    End If

    If Len(Trim(Combo5.Text)) = 0 Then
        MsgBox "Seleccion Pventa?"
        Exit Sub

    End If

    If Val(valoropera) <= 0 Then
        MsgBox "Digite un Valor"
        Exit Sub

    End If

    If Len(Trim(clave)) = 0 Then
        MsgBox "Seleccion Clave"
        Exit Sub

    End If

    mytablex.Open "SELECT * FROM vendedor where clave='" & Trim(clave) & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then  'si no existe
        mytablex.Close
        MsgBox "No existe usuario", 48, "Aviso"
        Exit Sub

    End If

    mytablex.Close
    rrproducto.MoveFirst
    Do

        If rrproducto.EOF Then Exit Do

        Select Case Combo5.Text

            Case "Pventa1,Unidad1,Factor1"
                cn.Execute ("update precios set pventa1=pventa1 " & Combo3.Text & " pventa1*" & Val(valoropera) & "/100 where producto='" & Trim("" & rrproducto.Fields("producto")) & "' and local='" & Combo4 & "'")

            Case "Pventa2,Unidad2,factor2"
                cn.Execute ("update precios set pventa2=pventa2 " & Combo3.Text & " pventa2*" & Val(valoropera) & "/100 where producto='" & Trim("" & rrproducto.Fields("producto")) & "' and local='" & Combo4 & "'")

            Case "Pventa3,Unidad3,Factor3"
                cn.Execute ("update precios set pventa3=pventa3 " & Combo3.Text & " pventa3*" & Val(valoropera) & "/100 where producto='" & Trim("" & rrproducto.Fields("producto")) & "' and local='" & Combo4 & "'")

            Case "Pventa4,Unidad4,factor4"
                cn.Execute ("update precios set pventa4=pventa4 " & Combo3.Text & " pventa4*" & Val(valoropera) & "/100 where producto='" & Trim("" & rrproducto.Fields("producto")) & "' and local='" & Combo4 & "'")

            Case "Pventa5,Unidad5,Factor5"
                cn.Execute ("update precios set pventa5=pventa5 " & Combo3.Text & " pventa5*" & Val(valoropera) & "/100 where producto='" & Trim("" & rrproducto.Fields("producto")) & "' and local='" & Combo4 & "'")

            Case "Pventa6,Unidad6,Factor6"
                cn.Execute ("update precios set pventa6=pventa6 " & Combo3.Text & " pventa6*" & Val(valoropera) & "/100 where producto='" & Trim("" & rrproducto.Fields("producto")) & "' and local='" & Combo4 & "'")

            Case "Pventa7,Unidad7,Factor7"
                cn.Execute ("update precios set pventa7=pventa7 " & Combo3.Text & " pventa7*" & Val(valoropera) & "/100 where producto='" & Trim("" & rrproducto.Fields("producto")) & "' and local='" & Combo4 & "'")

            Case "Pventa8,Unidad8,Factor8"
                cn.Execute ("update precios set pventa8=pventa8 " & Combo3.Text & " pventa8*" & Val(valoropera) & "/100 where producto='" & Trim("" & rrproducto.Fields("producto")) & "' and local='" & Combo4 & "'")

            Case "Pventa9,Unidad9,Factor9"
                cn.Execute ("update precios set pventa9=pventa9 " & Combo3.Text & " pventa9*" & Val(valoropera) & "/100 where producto='" & Trim("" & rrproducto.Fields("producto")) & "' and local='" & Combo4 & "'")

            Case "Pventa10,Unidad10,Factor10"
                cn.Execute ("update precios set pventa10=pventa10 " & Combo3.Text & " pventa10*" & Val(valoropera) & "/100 where producto='" & Trim("" & rrproducto.Fields("producto")) & "' and local='" & Combo4 & "'")

        End Select

        rrproducto.MoveNext
    Loop
    MsgBox "Proceso Realizado ", 48, "Aviso"
    dbgrid1_Click
    Label27_Click
    Exit Sub
cmd90800_err:
    MsgBox "Seleccione un datos " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub LI8912_Click()

    If Frame6.Visible = True Then Exit Sub

    If Frame5.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub

    flag_clave1 = 0
    tconcla.X = "CAMBIACODIGO"
    tconcla.Show 1

    If flag_clave1 <> 1 Then  'si es abre cajon
        MsgBox "No tiene permiso ", 48, "Aviso"
        Exit Sub

    End If

    Frame5.Visible = True

    '' 21/11/2017 Validar Cambio de Codigo de Producto
    'codigoa = ""
    'codigon = ""
    'codigoa.SetFocus
    
    codigoa = dbGrid1.columns(1)
    Label33 = dbGrid1.columns(0)
    codigon = ""
    codigon.SetFocus

    '' 21/11/2017 Validar Cambio de Codigo de Producto
End Sub

Private Sub linea_Click()

    'Command5_Click
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        Command11_Click
        Exit Sub

    End If

    If List1.ListIndex >= 0 Then
        producto = "" & List1.List(List1.ListIndex)
        sql_cabeza 0
        Command11_Click

    End If

End Sub

Private Sub local1_Click()

    'Command5_Click
End Sub

Private Sub marca_Click()

    'Command5_Click
End Sub

Private Sub maximo_Change()

End Sub

Private Sub Minimo_Change()

End Sub

Private Sub mas_Click()

    If frmOtros.Visible = True Then
        frmOtros.Visible = False
        Exit Sub

    End If
  
    If frmOtros.Visible = False Then
        frmOtros.Visible = True
        Exit Sub

    End If

End Sub

Private Sub Modif34_Click()

    Dim quebusco As String

    Dim found    As Integer

    On Error GoTo cmd1_err

    If Frame6.Visible = True Then Exit Sub

    If Frame5.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    'miMarca = rrproducto.Bookmark
    quebusco = Trim("" & dbGrid1.columns(1))
    FLAG = ""
    tproduct.codigo = Trim("" & dbGrid1.columns(1))
    tproduct.codigo.Enabled = False
    tproduct.ordename = "MODIFICA"
    'quebusco = rrproducto.AbsolutePosition
    tproduct.Show 1
    'MsgBox "xxx"
    'found = sql_cabeza(0)
    'rrproducto.Find "producto='" & quebusco & "'"
    'rrproducto.Bookmark = miMarca
    'rrproducto.AbsolutePosition = quebusco
    'Recordset.Find Combo1.Text & "='" & Text1.Text & "'", , adSearchForward
    'found = sql_cabeza(0)
    Exit Sub
cmd1_err:
    MsgBox "Elija un Codigo " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub borra_almacen_producto(Tmp As String)

    On Error GoTo cmd34_err

    cn.Execute "DELETE FROM ALMACEN WHERE producto='" & Tmp & "'"
    cn.Execute "DELETE FROM productb WHERE producto='" & Tmp & "'"
    cn.Execute "DELETE FROM codprov WHERE producto='" & Tmp & "'"
    'cn.Execute "DELETE FROM codclie WHERE producto='" & tmp & "'"
    cn.Execute "DELETE FROM precios WHERE producto='" & Tmp & "'"
    cn.Execute "DELETE FROM producto WHERE producto='" & Tmp & "'"
    cn.Execute "DELETE FROM receta WHERE producto='" & Tmp & "'"
    cn.Execute "DELETE FROM dueno WHERE producto='" & Tmp & "'"
    cn.Execute "DELETE FROM COMBINA WHERE producto='" & Tmp & "'"
    
    ''' 29/01/2018 Comisiones por producto por trabajador.
    cn.Execute "DELETE FROM vendedorcomision WHERE producto='" & Tmp & "'"
    ''' 29/01/2018 Comisiones por producto por trabajador.
    
    Exit Sub
cmd34_err:
    MsgBox "Aviso en borra almacen producto " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub sql_saldo_locales(buf As String)

    Dim buf1     As String

    Dim mytabley As New ADODB.Recordset

    Dim mytablex As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset

    On Error GoTo cmd34_err

    mytablex.Open "SELECT * from bodega WHERE 1=2", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        If mytabley.State = 1 Then mytabley.Close
        mytabley.Open "SELECT * from almacen where local='" & mytablex.Fields("local") & "' and producto='" & "" & Trim(buf) & "' AND bodega='" & "" & mytablex.Fields("codigo") & "'", cn, adOpenStatic, adLockOptimistic

        If mytabley.RecordCount = 0 Then
            mytabley.AddNew
            mytabley.Fields("producto") = buf
            mytabley.Fields("local") = mytablex.Fields("local")
            mytabley.Fields("bodega") = "" & mytablex.Fields("codigo")
            mytabley.Fields("minimo") = Val("" & rrproducto.Fields("minimo"))
            mytabley.Fields("maximo") = Val("" & rrproducto.Fields("maximo"))
            mytabley.Fields("saldo") = 0
            mytabley.Update

        End If

        mytabley.Close
        mytablex.MoveNext
    Loop
    mytablex.Close

    buf1 = "select Almacen.saldo,Bodega.nombre,almacen.bodega as Almacen,Almacen.local from almacen,bodega where  almacen.bodega=bodega.codigo and almacen.producto='" & buf & "' order by almacen.bodega"

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open buf1, cn, adOpenKeyset, adLockOptimistic
    Set DBGrid2.DataSource = mytablex
    DBGrid2.columns(0).Width = 1000
    DBGrid2.columns(1).Width = 1000
    DBGrid2.columns(2).Width = 1000
    DBGrid2.columns(3).Width = 2300
    'dbgrid2.columns(4).Width = 1300
   
    'If Val(anno) < 2000 And Val(anno) > 2030 Then
    '   anno = Format(Year(Now), "0000")
    'End If
    'mytablez.Open "SELECT Local,Bodega,anno,Mes,Cantidade,Cantidads,Costo from sisunat where producto='" & Trim(buf) & "' and anno='" & anno & "' ", cn, adOpenKeyset, adLockOptimistic
    '   Set dbgrid15.DataSource = mytablez
    '   dbgrid15.columns(0).Width = 1000
    '   dbgrid15.columns(1).Width = 1000
    '   dbgrid15.columns(2).Width = 1000
    '   dbgrid15.columns(3).Width = 1000
   
    Exit Sub
cmd34_err:
    MsgBox "Aviso en sql saldo locales " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub preciogeneral_Click()

    If Frame6.Visible = True Then Exit Sub

    If Frame5.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    If MsgBox("Desea Exportar excel", 1, "Aviso") <> 1 Then Exit Sub
    precio_excellDetalle 0

End Sub

Private Sub producto_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    Command5_Click

End Sub

Private Sub seccion_Click()

    'Command5_Click
End Sub

Private Sub subfamilia_Click()

    'Command5_Click
End Sub

Private Sub Zom82_Click()

    Dim quebusco As String

    On Error GoTo cmd13_err

    If Frame6.Visible = True Then Exit Sub

    If Frame5.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    quebusco = Trim("" & dbGrid1.columns(1))
    FLAG = ""
    tproduct.codigo = "" & dbGrid1.columns(1)
    tproduct.ordename = "VER"
    tproduct.Show 1
    'found = sql_cabeza(0)
    rrproducto.Find "producto='" & quebusco & "'"
    Exit Sub
cmd13_err:
    MsgBox "Elija un Codigo", 48, "Aviso"
    Exit Sub

End Sub

Sub DrawBarcode(bc_string As String, _
                sDescripcion As String, _
                VLPrecio As String, _
                OBJ As PictureBox)

    Dim Xpos!, Y1!, Y2!, dw%, th!, tw, new_string$

    Dim bc(90) As String

    Dim sAux As String

    Dim I As Byte

    Dim n As Integer

    Dim c As Integer

    Dim bc_pattern$

    bc(1) = "1 1221" 'pre-amble
    bc(2) = "1 1221" 'post-amble
    bc(48) = "11 221" 'dígitos
    bc(49) = "21 112"
    bc(50) = "12 112"
    bc(51) = "22 111"
    bc(52) = "11 212"
    bc(53) = "21 211"
    bc(54) = "12 211"
    bc(55) = "11 122"
    bc(56) = "21 121"
    bc(57) = "12 121"
    'Letras Mayúsculas
    bc(65) = "211 12" 'A
    bc(66) = "121 12" 'B
    bc(67) = "221 11" 'C
    bc(68) = "112 12" 'D
    bc(69) = "212 11" 'E
    bc(70) = "122 11" 'F
    bc(71) = "111 22" 'G
    bc(72) = "211 21" 'H
    bc(73) = "121 21" 'I
    bc(74) = "112 21" 'J
    bc(75) = "2111 2" 'K
    bc(76) = "1211 2" 'L
    bc(77) = "2211 1" 'M
    bc(78) = "1121 2" 'N
    bc(79) = "2121 1" 'O
    bc(80) = "1221 1" 'P
    bc(81) = "1112 2" 'Q
    bc(82) = "2112 1" 'R
    bc(83) = "1212 1" 'S
    bc(84) = "1122 1" 'T
    bc(85) = "2 1112" 'U
    bc(86) = "1 2112" 'V
    bc(87) = "2 2111" 'W
    bc(88) = "1 1212" 'X
    bc(89) = "2 1211" 'Y
    bc(90) = "1 2211" 'Z
    'Misceláneos Caracteres
    bc(32) = "1 2121" 'Espacio
    bc(35) = "" '# no se puede realizar
    bc(36) = "1 1 1 11" '$
    bc(37) = "11 1 1 1" '%
    bc(43) = "1 11 1 1" '+
    bc(45) = "1 1122" '-
    bc(47) = "1 1 11 1" '/
    bc(46) = "2 1121" '.
    bc(64) = "" '@ no se puede realizar
    bc(65) = "1 1221" '*

    bc_string = UCase(bc_string) 'Convertir a mayúsculas

    'Dimensiones
    OBJ.ScaleMode = 2 'Pixeles
    OBJ.Cls
    OBJ.Picture = Nothing
    dw = CInt(OBJ.ScaleHeight / 40) 'Espacio entre barras

    If dw < 1 Then dw = 1
    th = OBJ.TextHeight(bc_string) 'Alto texto
    tw = OBJ.TextWidth(bc_string) 'Ancho texto
    new_string = Chr$(1) & bc_string & Chr$(2) 'Agregar pre-amble, post-amble
    Y1 = OBJ.ScaleTop + 12
    Y2 = OBJ.ScaleTop + OBJ.ScaleHeight - 1.5 * th
    OBJ.Width = 1.1 * Len(new_string) * (15 * dw) * OBJ.Width / OBJ.ScaleWidth

    'Dibujar cada caracter en el string barcode
    Xpos = OBJ.ScaleLeft

    For n = 1 To Len(new_string)
        c = Asc(Mid(new_string, n, 1))

        If c > 90 Then c = 0
        bc_pattern$ = bc(c)

        'Dibujar cada barra
        For I = 1 To Len(bc_pattern$)

            Select Case Mid(bc_pattern$, I, 1)

                Case " "
                    'Espacio
                    OBJ.Line (Xpos, Y1)-(Xpos + 1 * dw, Y2), &HFFFFFF, BF
                    Xpos = Xpos + dw

                Case "1"
                    'Espacio
                    OBJ.Line (Xpos, Y1)-(Xpos + 1 * dw, Y2), &HFFFFFF, BF
                    Xpos = Xpos + dw
                    'Línea
                    OBJ.Line (Xpos, Y1)-(Xpos + 1 * dw, Y2), &H0&, BF
                    Xpos = Xpos + dw

                Case "2"
                    'Espacio
                    OBJ.Line (Xpos, Y1)-(Xpos + 1 * dw, Y2), &HFFFFFF, BF
                    Xpos = Xpos + dw
                    'Ancho línea
                    OBJ.Line (Xpos, Y1)-(Xpos + 2 * dw, Y2), &H0&, BF
                    Xpos = Xpos + 2 * dw

            End Select

        Next
    Next

    'Mas espacio
    OBJ.Line (Xpos, Y1)-(Xpos + 1 * dw, Y2), &HFFFFFF, BF
    Xpos = Xpos + dw

    'Medida final y tamaño
    OBJ.Width = (Xpos + dw) * OBJ.Width / OBJ.ScaleWidth
    OBJ.CurrentX = 1
    OBJ.CurrentY = 1

    If VLPrecio = "0.00" Then VLPrecio = ""
    If Xpos - OBJ.TextWidth(VLPrecio) - 10 < OBJ.TextWidth(sDescripcion) Then
        sAux = ""

        For I = 1 To Len(sDescripcion)

            If Xpos - OBJ.TextWidth(VLPrecio) - 10 < OBJ.TextWidth(sAux) Then
                Exit For
            Else
                sAux = sAux & Mid(sDescripcion, I, 1)

            End If

        Next I

        OBJ.Print sAux
    Else
        OBJ.Print sDescripcion

    End If

    OBJ.CurrentX = Xpos - OBJ.TextWidth(VLPrecio)
    OBJ.CurrentY = 1
    OBJ.Print VLPrecio
    OBJ.CurrentX = (OBJ.ScaleWidth - tw) / 2
    OBJ.CurrentY = Y2 + 0.25 * th
    OBJ.Print bc_string

    'Copiar a clipboard
    OBJ.Picture = OBJ.Image
    Clipboard.Clear
    Clipboard.SetData OBJ.Image, 2

End Sub

Sub carga_inicios()

    Dim found As Integer

    found = sql_cabeza(0)
        
End Sub

'' 01/12/2017 Mejora reporte lista de precios

'Sub precio_excell(asw As Integer)
'
' Dim mytablex As New ADODB.Recordset
' Dim found As Integer
' Dim i As Integer
' Dim v As Long
' Dim R As Long
' Dim ih As Integer
' Dim h As Integer
' Dim vprecios(10) As String
' Dim cad As String
' Dim Tmp As String
' Dim buf As String
' Dim sw As Integer
' Dim sdx As Double
' Dim xcosto As Double
' Dim mytabley As New ADODB.Recordset
'    Dim Heading(10) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion
'    On Error GoTo cmd5612_err
'    ordenado.ListIndex = 4
'    found = sql_cabeza(0)
'    If rrproducto.RecordCount = 0 Then Exit Sub
'
'
'    Heading(1) = "Producto"
'    Heading(2) = "Descripcio"
'    Heading(3) = "Lista"
'    Heading(4) = "Und"
'    Heading(5) = "Factor"
'    Heading(6) = "PVenta"
'    Heading(7) = "M"
'    Heading(8) = "Costo"
'    Heading(9) = "Ganancia"
'    Heading(10) = "Stock"
'
'    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
'    Call Formato_ExcelListaPrecios(10, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
'
'v = 5
'h = 1
'rrproducto.MoveFirst
'sw = 0
'Do
'   If rrproducto.EOF Then Exit Do
'     If sw = 0 Then
'        Tmp = "" & rrproducto.Fields("familia")
'        sw = 1
'        objExcel.ActiveSheet.Cells(v, h) = "" & rrproducto.Fields("familia")
'        v = v + 1
'     End If
'     If Tmp <> "" & rrproducto.Fields("familia") Then
'        Tmp = "" & rrproducto.Fields("familia")
'        objExcel.ActiveSheet.Cells(v, h) = "" & rrproducto.Fields("familia")
'        v = v + 1
'     End If
'     cad = "select * from precios where producto='" & rrproducto.Fields("producto") & "'"
'     If local1 <> "%" Then
'        cad = cad & "' and local='" & local1 & "'"
'     End If
'
'
'     If mytablex.State = 1 Then mytablex.Close
'     mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
'     If mytablex.RecordCount > 0 Then
'            objExcel.ActiveSheet.Cells(v, h) = "" & rrproducto.Fields("producto")
'            objExcel.ActiveSheet.Cells(v, h + 1) = "" & rrproducto.Fields("descripcio")
'
'               Do
'               If mytablex.EOF Then Exit Do
'                  objExcel.ActiveSheet.Cells(v, h + 2) = "" & mytablex.Fields("Local")
'                  objExcel.ActiveSheet.Cells(v, h + 3) = "" & mytablex.Fields("unidad1")
'                  objExcel.ActiveSheet.Cells(v, h + 4) = "" & mytablex.Fields("factor1")
'                  objExcel.ActiveSheet.Cells(v, h + 5) = "" & mytablex.Fields("pventa1")
'                  objExcel.ActiveSheet.Cells(v, h + 6) = "" & rrproducto.Fields("monedav")
'
'                  xcosto = Val("" & rrproducto.Fields("costou"))
'                  If "" & rrproducto.Fields("monedav") = "S" Then
'                     xcosto = Val("" & rrproducto.Fields("costou"))
'                  End If
'                  If "" & rrproducto.Fields("monedav") = "D" Then
'                     xcosto = Val("" & rrproducto.Fields("costou"))
'                  End If
'                  objExcel.ActiveSheet.Cells(v, h + 7) = xcosto
'
'                  sdx = 0
'                  If xcosto > 0 Then
'                     sdx = (Val("" & mytablex.Fields("pventa1")) - xcosto) * 100 / xcosto
'                     objExcel.ActiveSheet.Cells(v, h + 8) = "'" & Format(sdx, "0.00") & "%"
'                     Else
'                     objExcel.ActiveSheet.Cells(v, h + 8) = "RevisaCosto"
'                  End If
'
'
'
'
'                     ih = 1
'                     If mytabley.State = 1 Then
'                        mytabley.Close
'                        Set mytabley = Nothing
'                     End If
'                     buf = "select * from almacen where producto='" & "" & rrproducto.Fields("producto") & "'"
'                     If bodega <> "%" Then
'                        buf = buf & " and bodega='" & extra_loquesea(bodega) & "'"
'                     End If
'                     sdx = 0
'                     mytabley.Open buf, cn, adOpenStatic, adLockOptimistic
'                     If mytabley.RecordCount > 0 Then
'                        Do
'                        If mytabley.EOF Then Exit Do
'                        sdx = sdx + Val("" & mytabley.Fields("saldo"))
'                        mytabley.MoveNext
'                        Loop
'                     End If
'                     mytabley.Close
'                     objExcel.ActiveSheet.Cells(v, h + 9) = sdx
'
'                  v = v + 1
'                  If Len("" & mytablex.Fields("unidad2")) > 0 And Val("" & mytablex.Fields("factor2")) > 0 Then
'                  objExcel.ActiveSheet.Cells(v, h + 2) = "*" & mytablex.Fields("Local")
'                  objExcel.ActiveSheet.Cells(v, h + 3) = "" & mytablex.Fields("unidad2")
'                  objExcel.ActiveSheet.Cells(v, h + 4) = "" & mytablex.Fields("factor2")
'                  objExcel.ActiveSheet.Cells(v, h + 5) = "" & mytablex.Fields("pventa2")
'                  objExcel.ActiveSheet.Cells(v, h + 6) = "" & rrproducto.Fields("monedav")
'                  v = v + 1
'                  End If
'                  If Len("" & mytablex.Fields("unidad3")) > 0 And Val("" & mytablex.Fields("factor3")) > 0 Then
'                  objExcel.ActiveSheet.Cells(v, h + 2) = "*" & mytablex.Fields("Local")
'                  objExcel.ActiveSheet.Cells(v, h + 3) = "" & mytablex.Fields("unidad3")
'                  objExcel.ActiveSheet.Cells(v, h + 4) = "" & mytablex.Fields("factor3")
'                  objExcel.ActiveSheet.Cells(v, h + 5) = "" & mytablex.Fields("pventa3")
'                  objExcel.ActiveSheet.Cells(v, h + 6) = "" & rrproducto.Fields("monedav")
'                  v = v + 1
'                  End If
'                  If Len("" & mytablex.Fields("unidad4")) > 0 And Val("" & mytablex.Fields("factor4")) > 0 Then
'                  objExcel.ActiveSheet.Cells(v, h + 2) = "*" & mytablex.Fields("Local")
'                  objExcel.ActiveSheet.Cells(v, h + 3) = "" & mytablex.Fields("unidad4")
'                  objExcel.ActiveSheet.Cells(v, h + 4) = "" & mytablex.Fields("factor4")
'                  objExcel.ActiveSheet.Cells(v, h + 5) = "" & mytablex.Fields("pventa4")
'                  objExcel.ActiveSheet.Cells(v, h + 6) = "" & rrproducto.Fields("monedav")
'                  v = v + 1
'                  End If
'                  If Len("" & mytablex.Fields("unidad5")) > 0 And Val("" & mytablex.Fields("factor5")) > 0 Then
'                  objExcel.ActiveSheet.Cells(v, h + 2) = "*" & mytablex.Fields("Local")
'                  objExcel.ActiveSheet.Cells(v, h + 3) = "" & mytablex.Fields("unidad5")
'                  objExcel.ActiveSheet.Cells(v, h + 4) = "" & mytablex.Fields("factor5")
'                  objExcel.ActiveSheet.Cells(v, h + 5) = "" & mytablex.Fields("pventa5")
'                  objExcel.ActiveSheet.Cells(v, h + 6) = "" & rrproducto.Fields("monedav")
'                  v = v + 1
'                  End If
'               mytablex.MoveNext
'               Loop
'         End If
' rrproducto.MoveNext
'Loop
'Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
'Exit Sub
'cmd5612_err:
'MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
'Exit Sub
'End Sub

Sub precio_excell(asw As Integer)

    Dim mytablex     As New ADODB.Recordset

    Dim found        As Integer

    Dim I            As Integer

    Dim v            As Long

    Dim R            As Long

    Dim ih           As Integer

    Dim h            As Integer

    Dim vprecios(10) As String

    Dim cad          As String

    Dim Tmp          As String

    Dim buf          As String

    Dim sw           As Integer

    Dim sdx          As Double

    Dim xcosto       As Double

    Dim mytabley     As New ADODB.Recordset

    Dim Heading(14)  As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    On Error GoTo cmd5612_err

    ordenado.ListIndex = 4
    found = sql_cabeza(0)

    If rrproducto.RecordCount = 0 Then Exit Sub
    
    Heading(1) = "Familia"
    Heading(2) = "Producto"
    Heading(3) = "Descripcio"
    Heading(4) = "Lista"
    Heading(5) = "Und"
    Heading(6) = "Factor"
    Heading(7) = "PVenta"
    Heading(8) = "M"
    Heading(9) = "Costo"
    Heading(10) = "Ganancia"
    Heading(11) = "Stock"
    Heading(12) = "Barras"
    Heading(13) = "Marca"
    Heading(14) = "Subfamilia"
    
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_ExcelListaPrecios(14, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
    objExcel.ActiveSheet.Cells(1, 3) = "                                LISTA DE PRECIOS DE PRODUCTOS"
    objExcel.ActiveSheet.Cells(1, 3).Font.bold = True
    objExcel.ActiveSheet.Cells(1, 3).Font.Size = 14
    objExcel.ActiveSheet.Cells(1, 3).Font.color = RGB(0, 112, 184)
    
    v = 5
    h = 1
    rrproducto.MoveFirst
    sw = 0
    Do

        If rrproducto.EOF Then Exit Do
     
        If sw = 0 Then
            Tmp = "" & rrproducto.Fields("familia")
            sw = 1
            objExcel.ActiveSheet.Cells(v - 1, h) = " "
            objExcel.ActiveSheet.Cells(v, h) = "" & rrproducto.Fields("familia")
            objExcel.ActiveSheet.Cells(v, h).Font.bold = True
            objExcel.ActiveSheet.Cells(v, h).Font.color = RGB(62, 95, 138)
        
            v = v + 1

        End If

        If Tmp <> "" & rrproducto.Fields("familia") Then
        
            Tmp = "" & rrproducto.Fields("familia")
            v = v + 1
            objExcel.ActiveSheet.Cells(v - 1, h) = " "
            objExcel.ActiveSheet.Cells(v, h) = "" & rrproducto.Fields("familia")
            objExcel.ActiveSheet.Cells(v, h).Font.bold = True
            objExcel.ActiveSheet.Cells(v, h).Font.color = RGB(62, 95, 138)
         
            v = v + 1

        End If

        cad = "select * from precios where producto='" & rrproducto.Fields("producto") & "'"

        If local1 <> "%" Then
            cad = cad & "' and local='" & local1 & "'"

        End If
     
        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            objExcel.ActiveSheet.Cells(v, h) = "" & rrproducto.Fields("familia")
            objExcel.ActiveSheet.Cells(v, h + 1) = "" & rrproducto.Fields("producto")
            objExcel.ActiveSheet.Cells(v, h + 2) = "" & rrproducto.Fields("descripcio")
            objExcel.ActiveSheet.Cells(v, h + 11) = "'" & rrproducto.Fields("barras")
            objExcel.ActiveSheet.Cells(v, h + 12) = "" & rrproducto.Fields("marca")
            objExcel.ActiveSheet.Cells(v, h + 13) = "" & rrproducto.Fields("subfamilia")
            
            Do

                If mytablex.EOF Then Exit Do
                objExcel.ActiveSheet.Cells(v, h + 3) = "'" & mytablex.Fields("Local")
                objExcel.ActiveSheet.Cells(v, h + 4) = "" & mytablex.Fields("unidad1")
                objExcel.ActiveSheet.Cells(v, h + 5) = "" & mytablex.Fields("factor1")
                objExcel.ActiveSheet.Cells(v, h + 6) = "" & mytablex.Fields("pventa1")
                objExcel.ActiveSheet.Cells(v, h + 7) = "" & rrproducto.Fields("monedav")
                  
                xcosto = Val("" & rrproducto.Fields("costou"))
                  
                If "" & rrproducto.Fields("monedav") = "S" Then
                    xcosto = Val("" & rrproducto.Fields("costou"))

                End If

                If "" & rrproducto.Fields("monedav") = "D" Then
                    xcosto = Val("" & rrproducto.Fields("costou"))

                End If

                objExcel.ActiveSheet.Cells(v, h + 8) = xcosto
                  
                sdx = 0

                If xcosto > 0 Then
                    sdx = (Val("" & mytablex.Fields("pventa1")) - xcosto) * 100 / xcosto
                    objExcel.ActiveSheet.Cells(v, h + 9) = "'" & Format(sdx, "0.00") & "%"
                Else
                    objExcel.ActiveSheet.Cells(v, h + 9) = "RevisaCosto"

                End If
                  
                ih = 1

                If mytabley.State = 1 Then
                    mytabley.Close
                    Set mytabley = Nothing

                End If

                buf = "select * from almacen where producto='" & "" & rrproducto.Fields("producto") & "'"

                If bodega <> "%" Then
                    buf = buf & " and bodega='" & extra_loquesea(bodega) & "'"

                End If

                sdx = 0
                mytabley.Open buf, cn, adOpenStatic, adLockOptimistic

                If mytabley.RecordCount > 0 Then
                    Do

                        If mytabley.EOF Then Exit Do
                        sdx = sdx + Val("" & mytabley.Fields("saldo"))
                        mytabley.MoveNext
                    Loop

                End If

                mytabley.Close
                objExcel.ActiveSheet.Cells(v, h + 10) = sdx
                  
                v = v + 1

                If Len("" & mytablex.Fields("unidad2")) > 0 And Val("" & mytablex.Fields("factor2")) > 0 Then
                    objExcel.ActiveSheet.Cells(v, h + 3) = "*" & mytablex.Fields("Local")
                    objExcel.ActiveSheet.Cells(v, h + 4) = "" & mytablex.Fields("unidad2")
                    objExcel.ActiveSheet.Cells(v, h + 5) = "" & mytablex.Fields("factor2")
                    objExcel.ActiveSheet.Cells(v, h + 6) = "" & mytablex.Fields("pventa2")
                    objExcel.ActiveSheet.Cells(v, h + 7) = "" & rrproducto.Fields("monedav")
                    v = v + 1

                End If

                If Len("" & mytablex.Fields("unidad3")) > 0 And Val("" & mytablex.Fields("factor3")) > 0 Then
                    objExcel.ActiveSheet.Cells(v, h + 3) = "*" & mytablex.Fields("Local")
                    objExcel.ActiveSheet.Cells(v, h + 4) = "" & mytablex.Fields("unidad3")
                    objExcel.ActiveSheet.Cells(v, h + 5) = "" & mytablex.Fields("factor3")
                    objExcel.ActiveSheet.Cells(v, h + 6) = "" & mytablex.Fields("pventa3")
                    objExcel.ActiveSheet.Cells(v, h + 7) = "" & rrproducto.Fields("monedav")
                    v = v + 1

                End If

                If Len("" & mytablex.Fields("unidad4")) > 0 And Val("" & mytablex.Fields("factor4")) > 0 Then
                    objExcel.ActiveSheet.Cells(v, h + 3) = "*" & mytablex.Fields("Local")
                    objExcel.ActiveSheet.Cells(v, h + 4) = "" & mytablex.Fields("unidad4")
                    objExcel.ActiveSheet.Cells(v, h + 5) = "" & mytablex.Fields("factor4")
                    objExcel.ActiveSheet.Cells(v, h + 6) = "" & mytablex.Fields("pventa4")
                    objExcel.ActiveSheet.Cells(v, h + 7) = "" & rrproducto.Fields("monedav")
                    v = v + 1

                End If

                If Len("" & mytablex.Fields("unidad5")) > 0 And Val("" & mytablex.Fields("factor5")) > 0 Then
                    objExcel.ActiveSheet.Cells(v, h + 3) = "*" & mytablex.Fields("Local")
                    objExcel.ActiveSheet.Cells(v, h + 4) = "" & mytablex.Fields("unidad5")
                    objExcel.ActiveSheet.Cells(v, h + 5) = "" & mytablex.Fields("factor5")
                    objExcel.ActiveSheet.Cells(v, h + 6) = "" & mytablex.Fields("pventa5")
                    objExcel.ActiveSheet.Cells(v, h + 7) = "" & rrproducto.Fields("monedav")
                    v = v + 1

                End If

                mytablex.MoveNext
            Loop

        End If

        rrproducto.MoveNext
    Loop
    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    Exit Sub
cmd5612_err:
    MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
    Exit Sub

End Sub

'' 01/12/2017 Mejora reporte lista de precios

'' 01/12/2017 Mejora reporte lista de precios
'Sub precio_excellDetalle(asw As Integer)
'
' Dim mytablex As New ADODB.Recordset
' Dim found As Integer
' Dim i As Integer
' Dim v As Long
' Dim R As Long
' Dim ih As Integer
' Dim h As Integer
' Dim vprecios(23) As String
' Dim cad As String
' Dim Tmp As String
' Dim buf As String
' Dim sw As Integer
' Dim sdx As Double
' Dim xcosto As Double
' Dim mytabley As New ADODB.Recordset
'    Dim Heading(23) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion
'    On Error GoTo cmd5612_err
'    ordenado.ListIndex = 4
'    found = sql_cabeza(0)
'    If rrproducto.RecordCount = 0 Then Exit Sub
'
'    Heading(1) = "Producto"
'    Heading(2) = "Descripcio"
'    Heading(3) = "Lista"
'
'    Heading(4) = "Und"
'    Heading(5) = "Factor"
'    Heading(6) = "PVenta"
'
'    Heading(7) = "Und2"
'    Heading(8) = "Factor2"
'    Heading(9) = "PVenta2"
'
'    Heading(10) = "Und3"
'    Heading(11) = "Factor3"
'    Heading(12) = "PVenta3"
'
'    Heading(13) = "Und4"
'    Heading(14) = "Factor4"
'    Heading(15) = "PVenta4"
'
'    Heading(16) = "Und5"
'    Heading(17) = "Factor5"
'    Heading(18) = "PVenta5"
'
'
'    Heading(19) = "M"
'    Heading(20) = "Costo"
'    Heading(21) = "Ganancia"
'    Heading(22) = "Stock"
'
'
'
'    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
'    Call Formato_Excel(23, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
'
'v = 5
'h = 1
'rrproducto.MoveFirst
'sw = 0
'Do
'   If rrproducto.EOF Then Exit Do
'     If sw = 0 Then
'        Tmp = "" & rrproducto.Fields("familia")
'        sw = 1
'        objExcel.ActiveSheet.Cells(v, h) = "" & rrproducto.Fields("familia")
'        v = v + 1
'     End If
'     If Tmp <> "" & rrproducto.Fields("familia") Then
'        Tmp = "" & rrproducto.Fields("familia")
'        objExcel.ActiveSheet.Cells(v, h) = "" & rrproducto.Fields("familia")
'        v = v + 1
'     End If
'     cad = "select * from precios where producto='" & rrproducto.Fields("producto") & "'"
'     If local1 <> "%" Then
'        cad = cad & "' and local='" & local1 & "'"
'     End If
'
'
'     If mytablex.State = 1 Then mytablex.Close
'     mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
'     If mytablex.RecordCount > 0 Then
'            objExcel.ActiveSheet.Cells(v, h) = "" & rrproducto.Fields("producto")
'            objExcel.ActiveSheet.Cells(v, h + 1) = "" & rrproducto.Fields("descripcio")
'
'               Do
'               If mytablex.EOF Then Exit Do
'                  objExcel.ActiveSheet.Cells(v, h + 2) = "" & mytablex.Fields("Local")
'                  objExcel.ActiveSheet.Cells(v, h + 3) = "" & mytablex.Fields("unidad1")
'                  objExcel.ActiveSheet.Cells(v, h + 4) = "" & mytablex.Fields("factor1")
'                  objExcel.ActiveSheet.Cells(v, h + 5) = "" & mytablex.Fields("pventa1")
'
'                  objExcel.ActiveSheet.Cells(v, h + 6) = "" & mytablex.Fields("unidad2")
'                  objExcel.ActiveSheet.Cells(v, h + 7) = "" & mytablex.Fields("factor2")
'                  objExcel.ActiveSheet.Cells(v, h + 8) = "" & mytablex.Fields("pventa2")
'
'                  objExcel.ActiveSheet.Cells(v, h + 9) = "" & mytablex.Fields("unidad3")
'                  objExcel.ActiveSheet.Cells(v, h + 10) = "" & mytablex.Fields("factor3")
'                  objExcel.ActiveSheet.Cells(v, h + 11) = "" & mytablex.Fields("pventa3")
'
'                objExcel.ActiveSheet.Cells(v, h + 12) = "" & mytablex.Fields("unidad4")
'                  objExcel.ActiveSheet.Cells(v, h + 13) = "" & mytablex.Fields("factor4")
'                  objExcel.ActiveSheet.Cells(v, h + 14) = "" & mytablex.Fields("pventa4")
'
'                  objExcel.ActiveSheet.Cells(v, h + 15) = "" & mytablex.Fields("unidad5")
'                  objExcel.ActiveSheet.Cells(v, h + 16) = "" & mytablex.Fields("factor5")
'                  objExcel.ActiveSheet.Cells(v, h + 17) = "" & mytablex.Fields("pventa5")
'
'
'                  objExcel.ActiveSheet.Cells(v, h + 18) = "" & rrproducto.Fields("monedav")
'
'
'
'                  xcosto = Val("" & rrproducto.Fields("costou"))
'                  If "" & rrproducto.Fields("monedav") = "S" Then
'                     xcosto = Val("" & rrproducto.Fields("costou"))
'                  End If
'                  If "" & rrproducto.Fields("monedav") = "D" Then
'                     xcosto = Val("" & rrproducto.Fields("costou"))
'                  End If
'                  objExcel.ActiveSheet.Cells(v, h + 19) = xcosto
'
'                  sdx = 0
'                  If xcosto > 0 Then
'                     sdx = (Val("" & mytablex.Fields("pventa1")) - xcosto) * 100 / xcosto
'                     objExcel.ActiveSheet.Cells(v, h + 20) = "'" & Format(sdx, "0.00") & "%"
'                     Else
'                     objExcel.ActiveSheet.Cells(v, h + 20) = "RevisaCosto"
'                  End If
'
'
'
'
'                     ih = 1
'                     If mytabley.State = 1 Then
'                        mytabley.Close
'                        Set mytabley = Nothing
'                     End If
'                     buf = "select * from almacen where producto='" & "" & rrproducto.Fields("producto") & "'"
'                     If bodega <> "%" Then
'                        buf = buf & " and bodega='" & extra_loquesea(bodega) & "'"
'                     End If
'                     sdx = 0
'                     mytabley.Open buf, cn, adOpenStatic, adLockOptimistic
'                     If mytabley.RecordCount > 0 Then
'                        Do
'                        If mytabley.EOF Then Exit Do
'                        sdx = sdx + Val("" & mytabley.Fields("saldo"))
'                        mytabley.MoveNext
'                        Loop
'                     End If
'                     mytabley.Close
'                     objExcel.ActiveSheet.Cells(v, h + 21) = sdx
'                  v = v + 1
'
'               mytablex.MoveNext
'               Loop
'         End If
' rrproducto.MoveNext
'Loop
'Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
'Exit Sub
'cmd5612_err:
'MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
'Exit Sub
'End Sub

Sub precio_excellDetalle(asw As Integer)

    Dim mytablex     As New ADODB.Recordset

    Dim found        As Integer

    Dim I            As Integer

    Dim v            As Long

    Dim R            As Long

    Dim ih           As Integer

    Dim h            As Integer

    Dim vprecios(24) As String

    Dim cad          As String

    Dim Tmp          As String

    Dim buf          As String

    Dim sw           As Integer

    Dim sdx          As Double

    Dim xcosto       As Double

    Dim mytabley     As New ADODB.Recordset

    Dim Heading(24)  As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    On Error GoTo cmd5612_err

    ordenado.ListIndex = 4
    found = sql_cabeza(0)

    If rrproducto.RecordCount = 0 Then Exit Sub
    
    Heading(1) = "Familia"
    Heading(2) = "Producto"
    Heading(3) = "Descripción"
    Heading(4) = "Lista"
    
    Heading(5) = "Und"
    Heading(6) = "Factor"
    Heading(7) = "PVenta"
    
    Heading(8) = "Und2"
    Heading(9) = "Factor2"
    Heading(10) = "PVenta2"
    
    Heading(11) = "Und3"
    Heading(12) = "Factor3"
    Heading(13) = "PVenta3"
    
    Heading(14) = "Und4"
    Heading(15) = "Factor4"
    Heading(16) = "PVenta4"
    
    Heading(17) = "Und5"
    Heading(18) = "Factor5"
    Heading(19) = "PVenta5"
    
    Heading(20) = "M"
    Heading(21) = "Costo"
    Heading(22) = "Ganancia"
    Heading(23) = "Stock"
    
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_ExcelListaPrecios(23, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
    v = 5
    h = 1
    rrproducto.MoveFirst
    sw = 0
    Do

        If rrproducto.EOF Then Exit Do
        '
        '     If sw = 0 Then
        '        Tmp = "" & rrproducto.Fields("familia")
        '        sw = 1
        '        objExcel.ActiveSheet.Cells(v, h) = "" & rrproducto.Fields("familia")
        '        v = v + 1
        '     End If
        '
        '     If Tmp <> "" & rrproducto.Fields("familia") Then
        '        Tmp = "" & rrproducto.Fields("familia")
        '        objExcel.ActiveSheet.Cells(v, h) = "" & rrproducto.Fields("familia")
        '        v = v + 1
        '     End If
     
        If sw = 0 Then
            Tmp = "" & rrproducto.Fields("familia")
            sw = 1
            objExcel.ActiveSheet.Cells(v - 1, h) = " "
            objExcel.ActiveSheet.Cells(v, h) = "" & rrproducto.Fields("familia")
            objExcel.ActiveSheet.Cells(v, h).Font.bold = True
            objExcel.ActiveSheet.Cells(v, h).Font.color = RGB(62, 95, 138)
        
            v = v + 1

        End If

        If Tmp <> "" & rrproducto.Fields("familia") Then
        
            Tmp = "" & rrproducto.Fields("familia")
            v = v + 1
            objExcel.ActiveSheet.Cells(v - 1, h) = " "
            objExcel.ActiveSheet.Cells(v, h) = "" & rrproducto.Fields("familia")
            objExcel.ActiveSheet.Cells(v, h).Font.bold = True
            objExcel.ActiveSheet.Cells(v, h).Font.color = RGB(62, 95, 138)
         
            v = v + 1

        End If

        cad = "select * from precios where producto='" & rrproducto.Fields("producto") & "'"

        If local1 <> "%" Then
            cad = cad & "' and local='" & local1 & "'"

        End If
     
        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            
            objExcel.ActiveSheet.Cells(v, h) = "" & rrproducto.Fields("familia")
            objExcel.ActiveSheet.Cells(v, h + 1) = "" & rrproducto.Fields("producto")
            objExcel.ActiveSheet.Cells(v, h + 2) = "" & rrproducto.Fields("descripcio")
            
            Do

                If mytablex.EOF Then Exit Do
                objExcel.ActiveSheet.Cells(v, h + 3) = "'" & mytablex.Fields("Local")
                objExcel.ActiveSheet.Cells(v, h + 4) = "" & mytablex.Fields("unidad1")
                objExcel.ActiveSheet.Cells(v, h + 5) = "" & mytablex.Fields("factor1")
                objExcel.ActiveSheet.Cells(v, h + 6) = "" & mytablex.Fields("pventa1")
                  
                objExcel.ActiveSheet.Cells(v, h + 7) = "" & mytablex.Fields("unidad2")
                objExcel.ActiveSheet.Cells(v, h + 8) = "" & mytablex.Fields("factor2")
                objExcel.ActiveSheet.Cells(v, h + 9) = "" & mytablex.Fields("pventa2")
                  
                objExcel.ActiveSheet.Cells(v, h + 10) = "" & mytablex.Fields("unidad3")
                objExcel.ActiveSheet.Cells(v, h + 11) = "" & mytablex.Fields("factor3")
                objExcel.ActiveSheet.Cells(v, h + 12) = "" & mytablex.Fields("pventa3")
                  
                objExcel.ActiveSheet.Cells(v, h + 13) = "" & mytablex.Fields("unidad4")
                objExcel.ActiveSheet.Cells(v, h + 14) = "" & mytablex.Fields("factor4")
                objExcel.ActiveSheet.Cells(v, h + 15) = "" & mytablex.Fields("pventa4")
                  
                objExcel.ActiveSheet.Cells(v, h + 16) = "" & mytablex.Fields("unidad5")
                objExcel.ActiveSheet.Cells(v, h + 17) = "" & mytablex.Fields("factor5")
                objExcel.ActiveSheet.Cells(v, h + 18) = "" & mytablex.Fields("pventa5")
                  
                objExcel.ActiveSheet.Cells(v, h + 19) = "" & rrproducto.Fields("monedav")
                  
                xcosto = Val("" & rrproducto.Fields("costou"))

                If "" & rrproducto.Fields("monedav") = "S" Then
                    xcosto = Val("" & rrproducto.Fields("costou"))

                End If

                If "" & rrproducto.Fields("monedav") = "D" Then
                    xcosto = Val("" & rrproducto.Fields("costou"))

                End If

                objExcel.ActiveSheet.Cells(v, h + 20) = xcosto
                  
                sdx = 0

                If xcosto > 0 Then
                    sdx = (Val("" & mytablex.Fields("pventa1")) - xcosto) * 100 / xcosto
                    objExcel.ActiveSheet.Cells(v, h + 21) = "'" & Format(sdx, "0.00") & "%"
                Else
                    objExcel.ActiveSheet.Cells(v, h + 21) = "RevisaCosto"

                End If
                  
                ih = 1

                If mytabley.State = 1 Then
                    mytabley.Close
                    Set mytabley = Nothing

                End If

                buf = "select * from almacen where producto='" & "" & rrproducto.Fields("producto") & "'"

                If bodega <> "%" Then
                    buf = buf & " and bodega='" & extra_loquesea(bodega) & "'"

                End If

                sdx = 0
                mytabley.Open buf, cn, adOpenStatic, adLockOptimistic

                If mytabley.RecordCount > 0 Then
                    Do

                        If mytabley.EOF Then Exit Do
                        sdx = sdx + Val("" & mytabley.Fields("saldo"))
                        mytabley.MoveNext
                    Loop

                End If

                mytabley.Close
                objExcel.ActiveSheet.Cells(v, h + 22) = sdx
                v = v + 1

                mytablex.MoveNext
            Loop

        End If

        rrproducto.MoveNext
    Loop
    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    Exit Sub
cmd5612_err:
    MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
    Exit Sub

End Sub

Function busca_precios(mytablex As Table, buf As String, vprecios() As String)
    mytablex.Seek "=", buf, "01"

    If Not mytablex.NoMatch Then
        vprecios(1) = "" & mytablex.Fields("pventa1")
        vprecios(2) = "" & mytablex.Fields("pventa2")
        vprecios(3) = "" & mytablex.Fields("pventa3")
        vprecios(4) = "" & mytablex.Fields("pventa4")
        vprecios(5) = "" & mytablex.Fields("pventa5")
        vprecios(6) = "" & mytablex.Fields("pventa6")
        vprecios(7) = "" & mytablex.Fields("pventa7")
        vprecios(8) = "" & mytablex.Fields("pventa8")
        vprecios(9) = "" & mytablex.Fields("pventa9")
        vprecios(10) = "" & mytablex.Fields("pventa10")
        busca_precios = 1

    End If

    '------------------------------------- ------------

End Function

Sub precio_saldo()

    Dim mytabley     As New ADODB.Recordset

    Dim mytablex     As New ADODB.Recordset

    Dim v            As Integer

    Dim R            As Long

    Dim h            As Integer

    Dim found        As Integer

    Dim I            As Integer

    Dim j            As Integer
  
    Dim xalmacen(50) As String

    Dim Heading(20)  As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    On Error GoTo cmd56121_err

    'Data1.Refresh
    mytabley.Open sqldatos, cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount = 0 Then
        Exit Sub

    End If

    I = 1
    mytablex.Open "select * from bodega ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        xalmacen(I) = "" & mytablex.Fields("codigo")
        Heading(I + 2) = Mid$("" & mytablex.Fields("Nombre"), 1, 6)
        I = I + 1
        mytablex.MoveNext
    Loop
    mytablex.Close
    Heading(1) = "PRODUCTO"
    Heading(2) = "DESCRIPCIO"
    'MsgBox i
   
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    
    Call Formato_Excel(I + 1, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    'Exit Sub
    v = 5
    h = 1
    mytabley.MoveFirst

    Do

        If mytabley.EOF Then Exit Do
        objExcel.ActiveSheet.Cells(v, h) = "" & mytabley.Fields("producto")
        objExcel.ActiveSheet.Cells(v, h + 1) = "" & mytabley.Fields("descripcio")

        For j = 1 To I
            objExcel.ActiveSheet.Cells(v, j + 2) = busca_saldo("" & mytabley.Fields("producto"), xalmacen(j))
        Next j

        v = v + 1
        mytabley.MoveNext
    Loop
    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    Exit Sub
cmd56121_err:
    MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
    Exit Sub

End Sub

Function busca_saldo(xproducto As String, xbodega As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from almacen where local='01' and producto='" & xproducto & "' and bodega='" & xbodega & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_saldo = "" & mytablex.Fields("saldo")

    End If

End Function

Sub pone_registro_precios(buf As String)

End Sub

Function consulta_barras()

    Dim found    As Integer

    Dim buf      As String

    Dim sw       As Integer

    Dim mytablex As New ADODB.Recordset

    List1.Clear
    found = cargar_productosx()
   
    mytablex.Open "SELECT * FROM productb where  barras='" & Trim(cbarras) & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.EOF = True And mytablex.BOF = True Then
        Exit Function

    End If
   
    Do

        If mytablex.EOF Then Exit Do
        List1.AddItem "" & mytablex.Fields("producto")
        found = 1
        mytablex.MoveNext
    Loop

    If found = 1 Then
        List1.ListIndex = 0

    End If

    mytablex.Close
    consulta_barras = found

End Function

Function cargar_productosx()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM producto where  barras='" & Trim(cbarras) & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        List1.AddItem "" & mytablex.Fields("producto")
        cargar_productosx = 1

    End If

    mytablex.Close

End Function

Sub busca_precioss(buf As String)

End Sub

Sub visualiza_precios()

    Dim buf As String

    On Error GoTo cmd678_err

    buf = "" & dbGrid1.columns(1)
    pone_registro_precios buf
    Exit Sub
cmd678_err:
    MsgBox "Seleccione un dato " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub ejecuta(sw As Integer)

    Dim rconsulta1 As New ADODB.Recordset

    Dim buf        As String

    If opcion1 = "1" Then
        If Len(cadena) = 0 Then
            buf = "select Descripcio,familia from familia "
        Else
            buf = "select Descripcio,familia from familia where " & xbuffer & " like '%" & cadena & "%'"

        End If

    End If

    If opcion1 = "2" Then
        If Len(cadena) = 0 Then
            buf = "select Descripcio,Marca from marca  "
        Else
            buf = "select Descripcio,marca from marca where " & xbuffer & " like '%" & cadena & "%'"

        End If

    End If

    If opcion1 = "3" Then
        If Len(cadena) = 0 Then
            buf = "select Descripcio,Subfamilia from subfamilia where familia='" & "" & DBGrid2.columns("familia") & "'"
        Else
            buf = "select Descripcio,Subfamilia from Subfamilia where familia='" & "" & DBGrid2.columns("familia ") & "' and " & xbuffer & " like '%" & cadena & "%'"

        End If

    End If

    'MsgBox buf
    If rconsulta1.State = 1 Then
        rconsulta1.Close

    End If

    rconsulta1.Open buf, cn, adOpenStatic, adLockOptimistic

    If rconsulta1.EOF = True And rconsulta1.BOF = True Then
     
    End If

    Set dbgrid3.DataSource = rconsulta1
    dbgrid3.columns(0).Width = 4000
    dbgrid3.columns(1).Width = 2000

    If rconsulta1.RecordCount = 0 Then
        cadena.SetFocus
        Exit Sub

    End If

    If sw = 1 Then
        dbgrid3.SetFocus

    End If

End Sub

Sub consulta_precios(buf1 As String)

    Dim buf As String

    On Error GoTo cmdo0_err

    Dim mytablex  As New ADODB.Recordset

    Dim mytablexx As New ADODB.Recordset

    consulta_pmax
    mytablexx.Open "select * from producto where producto='" & "" & rrproducto.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic

    If mytablexx.RecordCount = 0 Then
        mytablexx.Close
        Exit Sub

    End If

    buf = "select * "
    buf = buf & "  from precios where producto='" & buf1 & "' order by local"
 
    If mytablex.State = 1 Then
        mytablex.Close

    End If

    'MsgBox buf
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    Set DBGrid4.DataSource = mytablex
      
    If mytablex.RecordCount > 0 Then
        mytablex.MoveFirst
        Do

            If mytablex.EOF Then Exit Do
            If Val("" & mytablex.Fields("pventa1")) > 0 Then

                'calcula_margenes mytablex, mytablexx
            End If

            mytablex.Update
            mytablex.MoveNext
        Loop

    End If

    'mytablex.Close
    mytablexx.Close
   
    Exit Sub
cmdo0_err:
    MsgBox "Aviso " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub calcula_margenes(mytablex As ADODB.Recordset, mytablexx As ADODB.Recordset)

    Dim sdx     As Double

    Dim sdx1    As Double

    Dim sdx2    As Double

    Dim acostou As String

    Dim factor  As String

    Dim costou  As String

    On Error GoTo cmd909090_err

    'Set dbgrid4.DataSource = mytablex
    acostou = Format(Val("" & mytablexx.Fields("costou")), "0.00")
    costou = acostou
    factor = Format(Val("" & mytablexx.Fields("factor")), "0.00")

    If "" & mytablexx.Fields("monedac") <> "S" And "" & mytablexx.Fields("monedac") <> "D" Then Exit Sub
    If "" & mytablexx.Fields("monedav") <> "S" And "" & mytablexx.Fields("monedav") <> "D" Then Exit Sub
    If "" & mytablexx.Fields("monedac") = "S" Then
        If "" & mytablexx.Fields("monedav") = "D" Then
            sdx = Val(acostou) / busca_cambio()

            If sdx <= 0 Then
                sdx = 1

            End If

            acostou = "" & sdx

        End If

    End If

    If "" & mytablexx.Fields("monedac") = "D" Then
        If "" & mytablexx.Fields("monedav") = "S" Then
            sdx = Val(acostou) * busca_cambio()

            If sdx <= 0 Then
                sdx = 1

            End If

            acostou = "" & sdx

        End If

    End If

    If Val(acostou) > 0 And Val("" & mytablex.Fields("pventa1")) > 0 And Val("" & mytablex.Fields("factor1")) > 0 Then 'calculando margenes
        sdx = (Val(acostou))
        sdx = sdx * Val("" & mytablex.Fields("factor1"))
        sdx1 = Val("" & mytablex.Fields("pventa1")) '/ val(""&mytablex.fields("factor1"))
        sdx2 = (sdx1 - sdx) * 100 / sdx
        mytablex.Fields("margen1") = Format(sdx2, "0.00")
        GoTo siguiente1

    End If
       
    If Val("" & mytablex.Fields("margen1")) > 0 And Val("" & mytablex.Fields("pventa1")) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val("" & mytablex.Fields("margen1")) / 100
        sdx = sdx * Val("" & mytablex.Fields("factor1"))
        mytablex.Fields("pventa1") = Format(sdx, "0.00")
        GoTo siguiente1

    End If

    If Val(acostou) <= 0 And Val("" & mytablex.Fields("pventa1")) > 0 And Val("" & mytablex.Fields("margen1")) > 0 Then
        sdx = Val("" & mytablex.Fields("pventa1")) / (1 + (Val("" & mytablex.Fields("margen1")) / 100))
        'costou = Format(sdx, "0.0000")
        GoTo siguiente1

    End If
       
siguiente1:

    If Val(acostou) > 0 And Val("" & mytablex.Fields("pventa2")) > 0 And Val("" & mytablex.Fields("factor2")) > 0 Then 'calculando margenes
        sdx = (Val(acostou))
        sdx = sdx * Val("" & mytablex.Fields("factor2"))
        sdx1 = Val("" & mytablex.Fields("pventa2")) '/ val(""&mytablex.fields("factor2"))
        sdx2 = (Val(sdx1) - sdx) * 100 / sdx
        mytablex.Fields("margen2") = Format(sdx2, "0.00")
        GoTo siguiente2

    End If

    If Val("" & mytablex.Fields("margen2")) > 0 And Val("" & mytablex.Fields("pventa2")) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val("" & mytablex.Fields("margen2")) / 100
        sdx = sdx * Val("" & mytablex.Fields("factor2"))
        mytablex.Fields("pventa2") = Format(sdx, "0.00")
        GoTo siguiente2

    End If

    If Val(acostou) <= 0 And Val("" & mytablex.Fields("pventa2")) > 0 And Val("" & mytablex.Fields("margen2")) > 0 Then
        sdx = Val("" & mytablex.Fields("pventa2")) / (1 + (Val("" & mytablex.Fields("margen2")) / 100))
        'costou = Format(sdx, "0.0000")
        GoTo siguiente2

    End If

siguiente2:

    If Val(acostou) > 0 And Val("" & mytablex.Fields("pventa3")) > 0 And Val("" & mytablex.Fields("factor3")) > 0 Then 'calculando margenes
        sdx = (Val(acostou))  '/ Val(factor))
        sdx = sdx * Val("" & mytablex.Fields("factor3"))
        sdx1 = Val("" & mytablex.Fields("pventa3")) '/ val(""&mytablex.fields("factor1"))
        sdx2 = (sdx1 - sdx) * 100 / sdx
        mytablex.Fields("margen3") = Format(sdx2, "0.00")
        GoTo siguiente3

    End If

    If Val("" & mytablex.Fields("margen3")) > 0 And Val("" & mytablex.Fields("pventa3")) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val("" & mytablex.Fields("margen3")) / 100
        sdx = sdx * Val("" & mytablex.Fields("factor3"))
        mytablex.Fields("pventa3") = Format(sdx, "0.00")
        GoTo siguiente3

    End If

    If Val(acostou) <= 0 And Val("" & mytablex.Fields("pventa3")) > 0 And Val("" & mytablex.Fields("margen3")) > 0 Then
        sdx = Val("" & mytablex.Fields("pventa3")) / (1 + (Val("" & mytablex.Fields("margen3")) / 100))
        'costou = Format(sdx, "0.0000")
        GoTo siguiente2

    End If

siguiente3:

    If Val(acostou) > 0 And Val("" & mytablex.Fields("pventa4")) > 0 And Val("" & mytablex.Fields("factor4")) > 0 Then 'calculando margenes
        sdx = (Val(acostou)) '/ Val(factor))
        sdx = sdx * Val("" & mytablex.Fields("factor4"))
        sdx1 = Val("" & mytablex.Fields("pventa4")) '/ val(""&mytablex.fields("factor1"))
        sdx2 = (sdx1 - sdx) * 100 / sdx
        mytablex.Fields("margen4") = Format(sdx2, "0.00")
        GoTo siguiente4

    End If

    If Val("" & mytablex.Fields("margen4")) > 0 And Val("" & mytablex.Fields("pventa4")) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val("" & mytablex.Fields("margen4")) / 100
        sdx = sdx * Val("" & mytablex.Fields("factor4"))
        mytablex.Fields("pventa4") = Format(sdx, "0.00")
        GoTo siguiente4

    End If

    If Val(acostou) <= 0 And Val("" & mytablex.Fields("pventa4")) > 0 And Val("" & mytablex.Fields("margen4")) > 0 Then
        sdx = Val("" & mytablex.Fields("pventa4")) / (1 + (Val("" & mytablex.Fields("margen4")) / 100))
        'costou = Format(sdx, "0.0000")
        GoTo siguiente4

    End If

siguiente4:

    If Val(acostou) > 0 And Val("" & mytablex.Fields("pventa5")) > 0 And Val("" & mytablex.Fields("factor5")) > 0 Then 'calculando margenes
        sdx = (Val(acostou) / Val(factor))
        sdx = sdx * Val("" & mytablex.Fields("factor5"))
        sdx1 = Val("" & mytablex.Fields("pventa5")) '/ val(""&mytablex.fields("factor1"))
        sdx2 = (sdx1 - sdx) * 100 / sdx
        mytablex.Fields("margen5") = Format(sdx2, "0.00")
        GoTo siguiente5

    End If

    If Val("" & mytablex.Fields("margen5")) > 0 And Val("" & mytablex.Fields("pventa5")) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val("" & mytablex.Fields("margen5")) / 100
        sdx = sdx * Val("" & mytablex.Fields("factor5"))
        mytablex.Fields("pventa5") = Format(sdx, "0.00")
        GoTo siguiente5

    End If

    If Val(acostou) <= 0 And Val("" & mytablex.Fields("pventa5")) > 0 And Val("" & mytablex.Fields("margen5")) > 0 Then
        sdx = Val("" & mytablex.Fields("pventa5")) / (1 + (Val("" & mytablex.Fields("margen5")) / 100))
        'costou = Format(sdx, "0.0000")
        GoTo siguiente5

    End If

siguiente5:

    If Val(acostou) > 0 And Val("" & mytablex.Fields("pventa6")) > 0 And Val("" & mytablex.Fields("factor6")) > 0 Then 'calculando margenes
        sdx = (Val(acostou) / Val(factor))
        sdx = sdx * Val("" & mytablex.Fields("factor6"))
        sdx1 = Val("" & mytablex.Fields("pventa6")) '/ val(""&mytablex.fields("factor1"))
        sdx2 = (sdx1 - sdx) * 100 / sdx
        mytablex.Fields("margen6") = Format(sdx2, "0.00")
        GoTo siguiente6

    End If

    If Val("" & mytablex.Fields("margen6")) > 0 And Val("" & mytablex.Fields("pventa6")) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val("" & mytablex.Fields("margen6")) / 100
        sdx = sdx * Val("" & mytablex.Fields("factor6"))
        mytablex.Fields("pventa6") = Format(sdx, "0.00")
        GoTo siguiente6

    End If

    If Val(acostou) <= 0 And Val("" & mytablex.Fields("pventa6")) > 0 And Val("" & mytablex.Fields("margen6")) > 0 Then
        sdx = Val("" & mytablex.Fields("pventa6")) / (1 + (Val("" & mytablex.Fields("margen6")) / 100))
        'costou = Format(sdx, "0.0000")
        GoTo siguiente6

    End If

siguiente6:

    If Val(acostou) > 0 And Val("" & mytablex.Fields("pventa7")) > 0 And Val("" & mytablex.Fields("factor7")) > 0 Then 'calculando margenes
        sdx = (Val(acostou) / Val(factor))
        sdx = sdx * Val("" & mytablex.Fields("factor7"))
        sdx1 = Val("" & mytablex.Fields("pventa7")) '/ val(""&mytablex.fields("factor1"))
        sdx2 = (sdx1 - sdx) * 100 / sdx
        mytablex.Fields("margen7") = Format(sdx2, "0.00")
        GoTo siguiente7

    End If

    If Val("" & mytablex.Fields("margen7")) > 0 And Val("" & mytablex.Fields("pventa7")) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val("" & mytablex.Fields("margen7")) / 100
        sdx = sdx * Val("" & mytablex.Fields("factor7"))
        mytablex.Fields("pventa7") = Format(sdx, "0.00")
        GoTo siguiente7

    End If

    If Val(acostou) <= 0 And Val("" & mytablex.Fields("pventa7")) > 0 And Val("" & mytablex.Fields("margen7")) > 0 Then
        sdx = Val("" & mytablex.Fields("pventa7")) / (1 + (Val("" & mytablex.Fields("margen7")) / 100))
        'costou = Format(sdx, "0.0000")
        GoTo siguiente7

    End If

siguiente7:

    If Val(costou) > 0 And Val("" & mytablex.Fields("pventa8")) > 0 And Val("" & mytablex.Fields("factor8")) > 0 Then 'calculando margenes
        sdx = (Val(acostou) / Val(factor))
        sdx = sdx * Val("" & mytablex.Fields("factor8"))
        sdx1 = Val("" & mytablex.Fields("pventa8")) '/ val(""&mytablex.fields("factor1"))
        sdx2 = (sdx1 - sdx) * 100 / sdx
        mytablex.Fields("margen8") = Format(sdx2, "0.00")
        GoTo siguiente8

    End If

    If Val("" & mytablex.Fields("margen8")) > 0 And Val("" & mytablex.Fields("pventa8")) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val("" & mytablex.Fields("margen8")) / 100
        sdx = sdx * Val("" & mytablex.Fields("factor8"))
        mytablex.Fields("pventa8") = Format(sdx, "0.00")
        GoTo siguiente8

    End If

    If Val(acostou) <= 0 And Val("" & mytablex.Fields("pventa8")) > 0 And Val("" & mytablex.Fields("margen8")) > 0 Then
        sdx = Val("" & mytablex.Fields("pventa8")) / (1 + (Val("" & mytablex.Fields("margen8")) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente8

    End If

siguiente8:

    If Val(acostou) > 0 And Val("" & mytablex.Fields("pventa9")) > 0 And Val("" & mytablex.Fields("factor9")) > 0 Then 'calculando margenes
        sdx = (Val(acostou) / Val(factor))
        sdx = sdx * Val("" & mytablex.Fields("factor9"))
        sdx1 = Val("" & mytablex.Fields("pventa9")) '/ val(""&mytablex.fields("factor1"))
        sdx2 = (sdx1 - sdx) * 100 / sdx
        mytablex.Fields("margen9") = Format(sdx2, "0.00")
        GoTo siguiente9

    End If

    If Val("" & mytablex.Fields("margen9")) > 0 And Val("" & mytablex.Fields("pventa9")) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val("" & mytablex.Fields("margen9")) / 100
        sdx = sdx * Val("" & mytablex.Fields("factor9"))
        mytablex.Fields("pventa9") = Format(sdx, "0.00")
        GoTo siguiente9

    End If

    If Val(acostou) <= 0 And Val("" & mytablex.Fields("pventa9")) > 0 And Val("" & mytablex.Fields("margen9")) > 0 Then
        sdx = Val("" & mytablex.Fields("pventa9")) / (1 + (Val("" & mytablex.Fields("margen9")) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente9

    End If

siguiente9:

    If Val(acostou) > 0 And Val("" & mytablex.Fields("pventa10")) > 0 And Val("" & mytablex.Fields("factor10")) > 0 Then 'calculando margenes
        sdx = (Val(acostou) / Val(factor))
        sdx = sdx * Val("" & mytablex.Fields("factor10"))
        sdx1 = Val("" & mytablex.Fields("pventa10")) '/ val(""&mytablex.fields("factor1"))
        sdx2 = (sdx1 - sdx) * 100 / sdx
        mytablex.Fields("margen10") = Format(sdx2, "0.00")
        GoTo siguiente10

    End If

    If Val("" & mytablex.Fields("margen10")) > 0 And Val("" & mytablex.Fields("pventa10")) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val("" & mytablex.Fields("margen10")) / 100
        sdx = sdx * Val("" & mytablex.Fields("factor10"))
        mytablex.Fields("pventa2") = Format(sdx, "0.00")
        GoTo siguiente10

    End If

    If Val(acostou) <= 0 And Val("" & mytablex.Fields("pventa10")) > 0 And Val("" & mytablex.Fields("margen10")) > 0 Then
        sdx = Val("" & mytablex.Fields("pventa10")) / (1 + (Val("" & mytablex.Fields("margen10")) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente10

    End If

siguiente10:

    If Val(acostou) > 0 And Val("" & mytablex.Fields("pventa11")) > 0 And Val("" & mytablex.Fields("factor1")) > 0 Then 'calculando margenes
        sdx = (Val(acostou) / Val(factor))
        sdx = sdx * Val("" & mytablex.Fields("factor1"))
        sdx1 = Val("" & mytablex.Fields("pventa11")) '/ val(""&mytablex.fields("factor1"))
        sdx2 = (sdx1 - sdx) * 100 / sdx
        mytablex.Fields("margen11") = Format(sdx2, "0.00")
        GoTo siguiente11

    End If

    If Val("" & mytablex.Fields("margen11")) > 0 And Val("" & mytablex.Fields("pventa11")) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val("" & mytablex.Fields("margen11")) / 100
        sdx = sdx * Val("" & mytablex.Fields("factor1"))
        mytablex.Fields("pventa11") = Format(sdx, "0.00")
        GoTo siguiente11

    End If

    If Val(acostou) <= 0 And Val("" & mytablex.Fields("pventa11")) > 0 And Val("" & mytablex.Fields("margen11")) > 0 Then
        sdx = Val("" & mytablex.Fields("pventa11")) / (1 + (Val("" & mytablex.Fields("margen11")) / 100))
        sdx = sdx * Val("" & mytablex.Fields("factor1"))
        costou = Format(sdx, "0.0000")
        GoTo siguiente11

    End If

siguiente11:

    If Val(acostou) > 0 And Val("" & mytablex.Fields("pventa12")) > 0 And Val("" & mytablex.Fields("factor1")) > 0 Then 'calculando margenes
        sdx = (Val(acostou) / Val(factor))
        sdx = sdx * Val("" & mytablex.Fields("factor1"))
        sdx1 = Val("" & mytablex.Fields("pventa12")) '/ val(""&mytablex.fields("factor1"))
        sdx2 = (sdx1 - sdx) * 100 / sdx
        mytablex.Fields("margen12") = Format(sdx2, "0.00")
        GoTo siguiente12

    End If

    If Val("" & mytablex.Fields("margen12")) > 0 And Val("" & mytablex.Fields("pventa12")) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val("" & mytablex.Fields("margen12")) / 100
        sdx = sdx * Val("" & mytablex.Fields("factor1"))
        mytablex.Fields("pventa12") = Format(sdx, "0.00")
        GoTo siguiente12

    End If

    If Val(acostou) <= 0 And Val("" & mytablex.Fields("pventa12")) > 0 And Val("" & mytablex.Fields("margen12")) > 0 Then
        sdx = Val("" & mytablex.Fields("pventa12")) / (1 + (Val("" & mytablex.Fields("margen12")) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente12

    End If

siguiente12:

    If Val(acostou) > 0 And Val("" & mytablex.Fields("pventa13")) > 0 And Val("" & mytablex.Fields("factor1")) > 0 Then 'calculando margenes
        sdx = (Val(acostou) / Val(factor))
        sdx = sdx * Val("" & mytablex.Fields("factor1"))
        sdx1 = Val("" & mytablex.Fields("pventa13")) '/ val(""&mytablex.fields("factor1"))
        sdx2 = (sdx1 - sdx) * 100 / sdx
        mytablex.Fields("margen13") = Format(sdx2, "0.00")
        GoTo siguiente13

    End If

    If Val("" & mytablex.Fields("margen13")) > 0 And Val("" & mytablex.Fields("pventa13")) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val("" & mytablex.Fields("margen13")) / 100
        sdx = sdx * Val("" & mytablex.Fields("factor1"))
        mytablex.Fields("pventa13") = Format(sdx, "0.00")
        GoTo siguiente13

    End If

    If Val(acostou) <= 0 And Val("" & mytablex.Fields("pventa13")) > 0 And Val("" & mytablex.Fields("margen13")) > 0 Then
        sdx = Val("" & mytablex.Fields("pventa13")) / (1 + (Val("" & mytablex.Fields("margen13")) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente13

    End If

siguiente13:

    If Val(acostou) > 0 And Val("" & mytablex.Fields("pventa14")) > 0 And Val("" & mytablex.Fields("factor1")) > 0 Then 'calculando margenes
        sdx = (Val(acostou) / Val(factor))
        sdx = sdx * Val("" & mytablex.Fields("factor1"))
        sdx1 = Val("" & mytablex.Fields("pventa14")) '/ val(""&mytablex.fields("factor1"))
        sdx2 = (sdx1 - sdx) * 100 / sdx
        mytablex.Fields("margen14") = Format(sdx2, "0.00")
        GoTo siguiente14

    End If

    If Val("" & mytablex.Fields("margen14")) > 0 And Val("" & mytablex.Fields("pventa14")) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val("" & mytablex.Fields("margen14")) / 100
        sdx = sdx * Val("" & mytablex.Fields("factor1"))
        mytablex.Fields("pventa14") = Format(sdx, "0.00")
        GoTo siguiente14

    End If

    If Val(acostou) <= 0 And Val("" & mytablex.Fields("pventa14")) > 0 And Val("" & mytablex.Fields("margen14")) > 0 Then
        sdx = Val("" & mytablex.Fields("pventa14")) / (1 + (Val("" & mytablex.Fields("margen14")) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente14

    End If

siguiente14:

    If Val(acostou) > 0 And Val("" & mytablex.Fields("pventa15")) > 0 And Val("" & mytablex.Fields("factor1")) > 0 Then 'calculando margenes
        sdx = (Val(acostou) / Val(factor))
        sdx = sdx * Val("" & mytablex.Fields("factor1"))
        sdx1 = Val("" & mytablex.Fields("pventa15")) '/ val(""&mytablex.fields("factor1"))
        sdx2 = (sdx1 - sdx) * 100 / sdx
        mytablex.Fields("margen15") = Format(sdx2, "0.00")
        GoTo siguiente15

    End If

    If Val("" & mytablex.Fields("margen15")) > 0 And Val("" & mytablex.Fields("pventa15")) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val("" & mytablex.Fields("margen15")) / 100
        sdx = sdx * Val("" & mytablex.Fields("factor1"))
        mytablex.Fields("pventa15") = Format(sdx, "0.00")
        GoTo siguiente10

    End If

    If Val(acostou) <= 0 And Val("" & mytablex.Fields("pventa15")) > 0 And Val("" & mytablex.Fields("margen15")) > 0 Then
        sdx = Val("" & mytablex.Fields("pventa15")) / (1 + (Val("" & mytablex.Fields("margen15")) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente15

    End If

siguiente15:
    Exit Sub
cmd909090_err:
    MsgBox "Aviso en calcula margenes"
    Exit Sub

    'cospaqu = Format(Val(costou) * Val(factor))
    'cospaqp = Format(Val(costop) * Val(factor))
    'cospaqi = Format(Val(costoini) * Val(factor))
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

Sub poner_inicio()

End Sub

Function busca_preciobarra() As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from precios where producto='" & "" & rrproducto.Fields("producto") & "' and local='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_preciobarra = Format(Val("" & mytablex.Fields("pventa1")), "0.00")

    End If

    mytablex.Close

End Function

Sub procesar_datos()

    Dim buf As String

    Dim ipx As String

    ipx = "192.168.1.3"
    cn.Execute ("delete from producto ")
    buf = "INSERT INTO calipso.dbo.producto "
    buf = buf & " SELECT     * "
    buf = buf & " From " & ipx & "dbo.producto"
    cn.Execute (buf)
    MsgBox "Exitoso"

End Sub

Sub kardex_sunate()

    Dim quebusco As String

    opcion2 = "100"
    repinv.excell.Visible = True
    repinv.Label15.Visible = True
    repinv.Label16.Visible = True
    repinv.fechai.Visible = True
    repinv.fechaf.Visible = True
    repinv.fechai.Enabled = True

    quebusco = Trim("" & dbGrid1.columns(1))
    repinv.producto = "" & dbGrid1.columns(1)
    repinv.Label15.Visible = True
    repinv.Label16.Visible = True
    repinv.fechai.Visible = True
    repinv.fechaf.Visible = True
    repinv.fechai.Enabled = True
    repinv.Show 1
    rrproducto.Find "producto='" & quebusco & "'"

End Sub

Function poner_signo(buf As String) As String

    Select Case buf

        Case "Igual"
            poner_signo = "="

        Case "Distinto"
            poner_signo = "<>"

        Case "Mayor"
            poner_signo = ">"

        Case "Menor"
            poner_signo = "<"

        Case "MayorIgual"
            poner_signo = ">="

        Case "MenorIgual"
            poner_signo = "<="

        Case "TodasPosibles"
            poner_signo = " Like "

        Case "Y"
            poner_signo = " and "

        Case "O"
            poner_signo = " or "

    End Select

End Function

Function convierte_barras(buf) As String

    Dim buf1 As String

    Dim I    As Integer

    Dim sdx  As Integer

    'Exit Sub
    If flag_denisse = "0" Then Exit Function

    If Len(Trim(buf)) = 0 Then Exit Function
    buf1 = ""
    sdx = 18 - Len(buf)

    For I = 1 To sdx
        buf1 = buf1 & "0"
    Next I

    buf1 = buf1 & buf
    convierte_barras = buf1

End Function

Function consulta_pmax()

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select Local,Bodega as Almacen,Saldo from almacen where producto='" & rrproducto.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic
    Set dbgrid9.DataSource = mytablex

End Function

Sub carga_subfamilia()

    Dim cad      As String

    Dim mytablex As New ADODB.Recordset

    subfamilia.Clear
    subfamilia.AddItem "%"
    cad = "SELECT * FROM subfamil where familia='" & extra_loquesea1(familia) & "' order by subfamilia "

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        subfamilia.AddItem "" & mytablex.Fields("subfamilia") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    subfamilia.ListIndex = 0
    mytablex.Close

End Sub

Sub CONTADOR_producto()

    'Dim mytablex As New ADODB.Recordset
    'Dim mytabley As New ADODB.Recordset
    Dim sdx As Double

    Dim vr

    sdx = 1
    'rrproducto.Requery
    'mytablex.Open "select  producto from producto where producto like '%' ", cn, adOpenStatic, adLockOptimistic
    Do

        If rrproducto.EOF Then Exit Do
        vr = DoEvents()
        Command25.Caption = "" & sdx
        rrproducto.Fields("producto") = Trim("" & sdx)
        rrproducto.Update
        'MsgBox "ABC"
        '------------------------------
        cambia_precio rrproducto, sdx
        '-------------------------------
        sdx = sdx + 1
        rrproducto.MoveNext
    Loop

End Sub

Sub cambia_precio(mytablex As ADODB.Recordset, sdx As Double)

    Dim mytabley As New ADODB.Recordset

    mytabley.Open "select producto from precios where producto='" & "" & mytablex.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic
    Do

        If mytabley.EOF Then Exit Do
        mytabley.Fields("producto") = Trim("" & sdx)
        mytabley.Update
        mytabley.MoveNext
    Loop
    mytabley.Close

End Sub

Function puede_modificar()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM vendedor where codigo='" & gusuario & "' and modificaproducto='N'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        mytablex.Close
        Exit Function

    End If

    mytablex.Close
    puede_modificar = 1

End Function

'' 11/12/2017 SubReceta
Function busca_factorProduccion(xproducto As String, tipo As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select unidadp,factorp from producto where producto='" & xproducto & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If tipo = 0 Then
            busca_factorProduccion = "" & mytablex.Fields("unidadp")
        ElseIf tipo = 1 Then
            busca_factorProduccion = "" & mytablex.Fields("factorp")

        End If

    End If

End Function

'' 11/12/2017 SubReceta
